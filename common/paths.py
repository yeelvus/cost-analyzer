# -*- coding: utf-8 -*-
"""路径工具：数据根、0_Raw / 0_Std / 0_Sum / 0_html，以及忽略序号前缀的目录匹配。"""

from __future__ import annotations

import re
from pathlib import Path

from settings import CODE_ROOT, DATA_ROOT, REPORT_ROOT, code_path, data_path, report_path

# 0_  1_  10_  3.1_  12-
_ORDER_PREFIX_RE = re.compile(r"^(?:\d+(?:\.\d+)?[_-])+")
_NON_KEY_RE = re.compile(r"[^\w\u4e00-\u9fff]+", re.UNICODE)

EXCEL_EXTS = {".xlsx", ".xls", ".xlsm"}
_MONTH_TAIL_RE = re.compile(r"(\d{4}年\d{1,2}月)$")

# 在 0_Raw 下新建目录时用的正式名（带序号，便于浏览）
RAW_CREATE_NAMES: dict[str, str] = {
    "raw": "0_Raw",
    "成本报公司": "1_成本报公司",
    "成本统计收集": "2_成本统计收集",
    "预算部分": "3_预算部分",
    "责任成本": "3.1_责任成本",
    "分包合同": "5.1_分包合同",
    "采购台账": "9_采购台账",
    "商混站": "9.1_商混站",
    "材料欠款": "9.1_材料欠款",
    "机具费": "10_机具费",
    "人工费": "11_人工费",
    "人工费台账": "1_人工费台账",
    "机票": "2_机票",
    "财务付款": "12_财务付款",
    "html": "04_html",
    "机具费分析": "机具费分析",
    "采购数据分析": "采购数据分析",
    "综合成本看板": "综合成本看板",
    "财务付款分析": "财务付款",
    "sum": "0_Sum",
    "std": "0_Std",
    "采购部采购台账extract": "B_采购部采购台账_extract",
    "采购部采购台账std": "C_采购部采购台账_std",
    "采购部采购台账stdsplit": "F_采购部采购台账_std_split",
    "机具费rawsplit": "C_机具费_raw_split",
    "机具费stdsplit": "E_机具费_std_split",
    "人工费台账std": "B_人工费台账_std",
    "飞签机票std": "B_飞签机票_std",
    "财务付款rawstd": "A1_财务付款_raw_std",
    "财务付款classified": "A2_财务付款_classified",
    "财务付款std": "B_财务付款_std",
    "财务付款学习池": "D_财务付款_学习池",
}


def strip_order_prefix(name: str) -> str:
    """去掉目录/文件名开头的 0_、1_、3.1_、12- 这类序号。"""
    return _ORDER_PREFIX_RE.sub("", str(name or "").strip())


def dir_key(name: str) -> str:
    """用于比较的规范化名字：去序号、去 emoji/符号、小写。"""
    text = strip_order_prefix(name)
    text = _NON_KEY_RE.sub("", text)
    return text.lower()


def names_match(actual: str, logical: str) -> bool:
    a, b = dir_key(actual), dir_key(logical)
    if not a or not b:
        return False
    return a == b or b in a or a in b


def month_label_from_dirname(name: str) -> str | None:
    """1_2026年08月 / 2026年08月 → 2026年08月。"""
    text = strip_order_prefix(name)
    m = _MONTH_TAIL_RE.search(text)
    return m.group(1) if m else None


def find_named_dir(base: Path, logical: str) -> Path | None:
    """在 base 下找与 logical 匹配的子目录（忽略序号前缀）。"""
    if not base.is_dir() or not logical:
        return None
    target = dir_key(logical)
    if not target:
        return None
    exact: list[Path] = []
    fuzzy: list[Path] = []
    for child in sorted(base.iterdir(), key=lambda p: p.name):
        if not child.is_dir():
            continue
        if child.name.startswith(".") or child.name.startswith("~$"):
            continue
        key = dir_key(child.name)
        if key == target:
            exact.append(child)
        elif target in key or key in target:
            fuzzy.append(child)
    if exact:
        return exact[0]
    if len(fuzzy) == 1:
        return fuzzy[0]
    return None


def find_child_dir(base: Path, keyword: str) -> Path:
    """兼容旧接口：关键字匹配子目录；找不到则返回 base/keyword。"""
    found = find_named_dir(base, keyword)
    if found is not None:
        return found
    if base.is_dir():
        for child in sorted(base.iterdir()):
            if child.is_dir() and keyword in child.name:
                return child
    return base / keyword


def find_script(script_dir: Path, keyword: str) -> Path:
    if script_dir.is_dir():
        for script in sorted(script_dir.iterdir()):
            if script.is_file() and script.suffix == ".py" and keyword in script.name:
                return script
    return script_dir / keyword


def resolve_under(base: Path, *parts: str, create: bool = False) -> Path:
    """从 base 起按段解析，每段都忽略序号前缀。"""
    cur = base
    for i, part in enumerate(parts):
        if not part:
            continue
        found = find_named_dir(cur, part) if cur.is_dir() else None
        if found is not None:
            cur = found
            continue
        create_name = RAW_CREATE_NAMES.get(dir_key(part), part)
        cur = cur / create_name
        if create and i == len(parts) - 1:
            cur.mkdir(parents=True, exist_ok=True)
        elif create:
            cur.mkdir(parents=True, exist_ok=True)
    return cur


def _dir_has_entries(path: Path) -> bool:
    if not path.is_dir():
        return False
    try:
        next(path.iterdir())
        return True
    except StopIteration:
        return False


def discover_raw_root(data_root: Path | None = None) -> Path:
    """定位 Raw / 0_Raw；没有则使用 0_Raw（不自动创建）。"""
    root = Path(data_root or DATA_ROOT)
    found = find_named_dir(root, "Raw")
    if found is not None:
        return found
    return root / "0_Raw"


def ensure_raw_root() -> Path:
    root = discover_raw_root()
    root.mkdir(parents=True, exist_ok=True)
    return root


RAW_ROOT = discover_raw_root()


def discover_html_root(data_root: Path | None = None) -> Path:
    """定位 04_html / html / 0_html；优先 04_html。"""
    root = Path(data_root or DATA_ROOT)
    preferred = root / "04_html"
    if preferred.is_dir():
        return preferred
    found = find_named_dir(root, "html")
    if found is not None:
        return found
    return preferred


def ensure_html_root() -> Path:
    root = discover_html_root()
    root.mkdir(parents=True, exist_ok=True)
    return root


HTML_ROOT = discover_html_root()


def html_data_path(*parts: str, create: bool = True) -> Path:
    """HTML 输出目录：04_html（也可写成 html / 0_html）。"""
    root = ensure_html_root() if create else discover_html_root()
    if not parts:
        if create:
            root.mkdir(parents=True, exist_ok=True)
        return root
    return resolve_under(root, *parts, create=create)


def discover_sum_root(data_root: Path | None = None) -> Path:
    """定位 Sum / 0_Sum；没有则使用 0_Sum。"""
    root = Path(data_root or DATA_ROOT)
    found = find_named_dir(root, "Sum")
    if found is not None:
        return found
    return root / "0_Sum"


def ensure_sum_root() -> Path:
    root = discover_sum_root()
    root.mkdir(parents=True, exist_ok=True)
    return root


SUM_ROOT = discover_sum_root()


def sum_data_path(
    *parts: str,
    create: bool = True,
    legacy: Path | str | None = None,
) -> Path:
    """合并汇总 Excel 目录：0_Sum（也可写成 Sum / 1_Sum）。

    写出用 create=True（始终落在 0_Sum）。
    读取可传 legacy：0_Sum 尚空时回退旧 sum 目录。
    """
    root = discover_sum_root()
    preferred = resolve_under(root, *parts, create=False) if parts else root
    if _dir_has_entries(preferred):
        return preferred
    if legacy:
        legacy_path = Path(legacy)
        if _dir_has_entries(legacy_path):
            return legacy_path
    if create:
        root.mkdir(parents=True, exist_ok=True)
        preferred = resolve_under(root, *parts, create=True) if parts else root
        preferred.mkdir(parents=True, exist_ok=True)
    return preferred


def discover_std_root(data_root: Path | None = None) -> Path:
    """定位 Std / 0_Std；没有则使用 0_Std。"""
    root = Path(data_root or DATA_ROOT)
    found = find_named_dir(root, "Std")
    if found is not None:
        return found
    return root / "0_Std"


def ensure_std_root() -> Path:
    root = discover_std_root()
    root.mkdir(parents=True, exist_ok=True)
    return root


STD_ROOT = discover_std_root()


def std_data_path(
    *parts: str,
    create: bool = True,
    legacy: Path | str | None = None,
) -> Path:
    """标准化中间结果：0_Std（也可写成 Std / 1_Std）。

    写出用 create=True（始终落在 0_Std）。
    读取可传 legacy：0_Std 尚空时回退旧 std 目录。
    """
    root = discover_std_root()
    preferred = resolve_under(root, *parts, create=False) if parts else root
    if _dir_has_entries(preferred):
        return preferred
    if not create and legacy:
        legacy_path = Path(legacy)
        if _dir_has_entries(legacy_path):
            return legacy_path
        if legacy_path.is_dir():
            return legacy_path
    if create:
        root.mkdir(parents=True, exist_ok=True)
        preferred = resolve_under(root, *parts, create=True) if parts else root
        preferred.mkdir(parents=True, exist_ok=True)
    return preferred


def raw_data_path(*parts: str, legacy: Path | str | None = None, create: bool = True) -> Path:
    """原始数据目录：优先 0_Raw 下（忽略序号）；若尚未迁移则回退 legacy。"""
    raw_root = discover_raw_root()
    preferred = resolve_under(raw_root, *parts, create=False)
    if _dir_has_entries(preferred):
        return preferred
    if legacy:
        legacy_path = Path(legacy)
        if _dir_has_entries(legacy_path):
            return legacy_path
        if legacy_path.is_dir() and not create:
            return legacy_path
    if create:
        raw_root.mkdir(parents=True, exist_ok=True)
        preferred = resolve_under(raw_root, *parts, create=True)
        preferred.mkdir(parents=True, exist_ok=True)
    return preferred


def iter_excel_files(root: Path | str, *, recursive: bool = True) -> list[Path]:
    """列出 Excel；默认递归，因此 1_月份 子目录也能扫到。跳过 ~$ 锁文件。"""
    base = Path(root)
    if not base.is_dir():
        return []
    iterator = base.rglob("*") if recursive else base.iterdir()
    files: list[Path] = []
    for p in iterator:
        if not p.is_file():
            continue
        if p.name.startswith("~$") or p.name.startswith("."):
            continue
        if p.suffix.lower() not in EXCEL_EXTS:
            continue
        files.append(p)
    return sorted(files)


def find_month_dirs(parent: Path) -> list[tuple[str, Path]]:
    """父目录下的月份文件夹，1_2026年08月 与 2026年08月 同等对待。"""
    if not parent.is_dir():
        return []
    found: list[tuple[str, Path]] = []
    for child in sorted(parent.iterdir(), key=lambda p: p.name):
        if not child.is_dir():
            continue
        label = month_label_from_dirname(child.name)
        if label:
            found.append((label, child))
    return found


__all__ = [
    "CODE_ROOT",
    "DATA_ROOT",
    "REPORT_ROOT",
    "RAW_ROOT",
    "HTML_ROOT",
    "SUM_ROOT",
    "STD_ROOT",
    "EXCEL_EXTS",
    "code_path",
    "data_path",
    "report_path",
    "raw_data_path",
    "html_data_path",
    "sum_data_path",
    "std_data_path",
    "discover_raw_root",
    "ensure_raw_root",
    "discover_html_root",
    "ensure_html_root",
    "discover_sum_root",
    "ensure_sum_root",
    "discover_std_root",
    "ensure_std_root",
    "strip_order_prefix",
    "dir_key",
    "names_match",
    "month_label_from_dirname",
    "find_named_dir",
    "find_child_dir",
    "find_script",
    "resolve_under",
    "iter_excel_files",
    "find_month_dirs",
]
