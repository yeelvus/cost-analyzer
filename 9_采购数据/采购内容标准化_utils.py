import importlib.util
import os
from pathlib import Path

import openpyxl


CURRENT_DIR = Path(__file__).resolve().parent
CODE_ROOT = CURRENT_DIR.parent
UNIT_MAP_UTILS_PATH = CODE_ROOT / "0_项目基本信息" / "unit_map_utils.py"
STANDARDIZATION_PIPELINE_DIR = CURRENT_DIR / "物品名称规格标准化"

_PIPELINE_CONTEXT = None


def _load_module_from_path(module_name: str, file_path: Path):
    spec = importlib.util.spec_from_file_location(module_name, str(file_path))
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_unit_map_utils = _load_module_from_path("unit_map_utils", UNIT_MAP_UTILS_PATH)


def standardize_unit(value, auto_prompt: bool = True):
    return _unit_map_utils.standardize_unit(value, auto_prompt=auto_prompt)


def resolve_pending_unit_mappings() -> int:
    return int(_unit_map_utils.resolve_pending_unit_mappings())


def _get_pipeline_context():
    global _PIPELINE_CONTEXT
    if _PIPELINE_CONTEXT is not None:
        return _PIPELINE_CONTEXT

    pipeline_script = STANDARDIZATION_PIPELINE_DIR / "0_统一标准化流水线.py"
    if not pipeline_script.exists():
        _PIPELINE_CONTEXT = {"ok": False, "reason": f"未找到标准化流水线脚本: {pipeline_script}"}
        return _PIPELINE_CONTEXT

    spec = importlib.util.spec_from_file_location("_standardization_pipeline", str(pipeline_script))
    if spec is None or spec.loader is None:
        _PIPELINE_CONTEXT = {"ok": False, "reason": f"无法加载标准化流水线: {pipeline_script}"}
        return _PIPELINE_CONTEXT

    pipeline_module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(pipeline_module)

    rule_names = getattr(pipeline_module, "DEFAULT_RULES", [])
    if not rule_names:
        _PIPELINE_CONTEXT = {"ok": False, "reason": "标准化流水线 DEFAULT_RULES 为空"}
        return _PIPELINE_CONTEXT

    rule_modules = []
    for rule_name in rule_names:
        rule_file = STANDARDIZATION_PIPELINE_DIR / rule_name
        if not rule_file.exists():
            continue
        try:
            rule_module = pipeline_module.load_module_from_file(rule_file)
            rule_modules.append((rule_file.name, rule_module))
        except Exception:
            continue

    _PIPELINE_CONTEXT = {
        "ok": True,
        "pipeline_module": pipeline_module,
        "rule_modules": rule_modules,
    }
    return _PIPELINE_CONTEXT


def apply_standardization_pipeline(excel_path: str) -> None:
    context = _get_pipeline_context()
    if not context.get("ok"):
        print(f"  ⚠️ {context.get('reason', '标准化流水线不可用，已跳过')}")
        return

    workbook = openpyxl.load_workbook(excel_path)
    pipeline_module = context["pipeline_module"]
    for rule_name, rule_module in context["rule_modules"]:
        try:
            stats = pipeline_module.run_rule_on_workbook(rule_module, workbook, rule_name)
            print(
                f"  ✓ {rule_name}: "
                f"命中 {stats['matched']} 行, 更新 {stats['updated']} 行, 未识别 {stats['unresolved']} 行"
            )
        except Exception as exc:
            print(f"  ✗ 规则执行失败 {rule_name}: {exc}")

    workbook.save(excel_path)
