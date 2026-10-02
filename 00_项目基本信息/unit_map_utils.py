"""单位标准化工具模块。

unit_map.json 使用 "标准单位 -> 别名列表" 结构。
示例: {"m2": ["m2", "平", "平方米"]}
"""

import json
import os
import sys

_THIS_DIR = os.path.dirname(os.path.abspath(__file__))
_CODE_ROOT = os.path.dirname(_THIS_DIR)
if _CODE_ROOT not in sys.path:
    sys.path.insert(0, _CODE_ROOT)
from settings import UNIT_MAP_PATH  # noqa: E402

UNIT_MAP_JSON_PATH = str(UNIT_MAP_PATH)
INVALID_KEY = "__invalid__"

# 无效单位：直接返回空字符串（兜底）
_HARDCODED_INVALID = {"上", "想", "干", "寸"}

# 模块级缓存: 标准单位 -> 别名列表
_unit_map_cache: dict | None = None
# 反向索引缓存: 别名 -> 标准单位
_alias_to_canonical_cache: dict | None = None
# 无效词集合缓存
_invalid_units_cache: set | None = None
# 运行期覆盖: 未落库但本次运行生效的映射
_runtime_overrides: dict = {}
# 待处理未知单位池
_pending_unknowns: set = set()
# 本次运行中用户选择「跳过」的单位（不写入 JSON，但不再重复询问）
_skipped_unknowns: set = set()


def _rebuild_reverse_index(mapping: dict) -> None:
    global _alias_to_canonical_cache, _invalid_units_cache

    alias_to_canonical = {}
    invalid_units = set(_HARDCODED_INVALID)

    for canonical, aliases in mapping.items():
        if canonical == INVALID_KEY:
            if isinstance(aliases, list):
                invalid_units.update(str(x).strip() for x in aliases if str(x).strip())
            continue
        if not isinstance(aliases, list):
            continue
        for alias in aliases:
            alias_s = str(alias).strip()
            if alias_s:
                alias_to_canonical[alias_s] = canonical

    _alias_to_canonical_cache = alias_to_canonical
    _invalid_units_cache = invalid_units


def load_unit_map() -> dict:
    global _unit_map_cache
    if _unit_map_cache is None:
        with open(UNIT_MAP_JSON_PATH, encoding="utf-8") as f:
            _unit_map_cache = json.load(f)
        _rebuild_reverse_index(_unit_map_cache)
    return _unit_map_cache


def save_unit_map(mapping: dict) -> None:
    global _unit_map_cache
    _unit_map_cache = mapping
    _rebuild_reverse_index(mapping)
    with open(UNIT_MAP_JSON_PATH, "w", encoding="utf-8") as f:
        json.dump(mapping, f, ensure_ascii=False, indent=2)


def _show_batch_unit_dialog(unknown_units: list, known_canonicals: list) -> dict:
    """单窗口批量处理未知单位，返回决策字典。

    返回格式: {unknown: (canonical, save_flag)}
      - canonical == ""      -> 标记无效
      - canonical == unknown -> 跳过（保持原值）
    """
    try:
        import tkinter as tk
        from tkinter import ttk
    except ImportError:
        print("  [警告] tkinter 不可用，未知单位保持原值")
        return {u: (u, False) for u in unknown_units}

    result = {}

    root = tk.Tk()
    root.title("未知单位批量匹配")
    root.geometry("880x520")
    root.resizable(False, False)
    root.attributes("-topmost", True)
    root.lift()
    root.focus_force()

    tk.Label(root, text=f"待匹配未知单位共 {len(unknown_units)} 个", font=("", 14, "bold")).pack(pady=(14, 6))
    tk.Label(root, text="多选左侧单位后，右侧输入标准单位并点击“应用到选中”", fg="#666").pack(pady=(0, 10))

    body = tk.Frame(root)
    body.pack(fill="both", expand=True, padx=12)

    left = tk.Frame(body)
    left.pack(side="left", fill="both", expand=True)
    right = tk.Frame(body)
    right.pack(side="left", fill="y", padx=(12, 0))

    listbox = tk.Listbox(left, selectmode=tk.EXTENDED, width=40, height=22)
    for u in sorted(unknown_units):
        listbox.insert(tk.END, u)
    listbox.pack(side="left", fill="both", expand=True)
    scrollbar = tk.Scrollbar(left, orient="vertical", command=listbox.yview)
    scrollbar.pack(side="left", fill="y")
    listbox.config(yscrollcommand=scrollbar.set)

    candidates = sorted(set(v for v in known_canonicals if v))
    var = tk.StringVar(value="")
    combo = ttk.Combobox(right, textvariable=var, values=candidates, width=28, font=("", 12))
    combo.pack(pady=(8, 8))
    tk.Label(right, text="所有选择均自动保存到 JSON", fg="#999", font=("", 9)).pack(anchor="w", pady=(0, 8))

    def _selected_units():
        return [listbox.get(i) for i in listbox.curselection()]

    def _remove_selected():
        for i in reversed(listbox.curselection()):
            listbox.delete(i)

    def _apply_selected():
        canonical = var.get().strip()
        if not canonical:
            return
        selected = _selected_units()
        for unit in selected:
            result[unit] = (canonical, True)
        _remove_selected()

    def _mark_invalid_selected():
        selected = _selected_units()
        for unit in selected:
            result[unit] = ("", True)
        _remove_selected()

    def _skip_selected():
        """保留原值：也写入 JSON，以自身作为 canonical，下次不再询问。"""
        selected = _selected_units()
        for unit in selected:
            result[unit] = (unit, True)
        _remove_selected()

    tk.Button(right, text="应用到选中", command=_apply_selected, bg="#27ae60", fg="white", width=20).pack(pady=4)
    tk.Button(right, text="标记选中为无效", command=_mark_invalid_selected, bg="#e67e22", fg="white", width=20).pack(pady=4)
    tk.Button(right, text="跳过选中（本次保留原值）", command=_skip_selected, width=20).pack(pady=4)

    def _finish():
        # 未处理的默认为保留原值并写入 JSON
        remaining = listbox.get(0, tk.END)
        for unit in remaining:
            if unit not in result:
                result[unit] = (unit, True)
        root.destroy()

    tk.Button(root, text="完成", command=_finish, width=18).pack(pady=12)

    root.mainloop()
    return result


def resolve_pending_unit_mappings() -> int:
    """对当前暂存池中的未知单位进行一次性批量匹配。返回处理数量。"""
    if not _pending_unknowns:
        return 0

    mapping = load_unit_map()
    known_canonicals = sorted(k for k in mapping.keys() if k and k != INVALID_KEY)
    unknowns = sorted(_pending_unknowns)
    decisions = _show_batch_unit_dialog(unknowns, known_canonicals)

    for unit in unknowns:
        canonical, do_save = decisions.get(unit, (unit, True))

        if canonical == "":
            _runtime_overrides[unit] = ""
            invalid_list = mapping.get(INVALID_KEY, [])
            if unit not in invalid_list:
                invalid_list.append(unit)
                mapping[INVALID_KEY] = invalid_list
            continue

        # canonical == unit 时以自身为 canonical 写入，下次直接命中
        _runtime_overrides[unit] = canonical
        aliases = mapping.get(canonical, [])
        if not isinstance(aliases, list):
            aliases = []
        if canonical not in aliases:
            aliases.insert(0, canonical)
        if unit not in aliases:
            aliases.append(unit)
        mapping[canonical] = aliases

    save_unit_map(mapping)
    _pending_unknowns.clear()
    return len(unknowns)


def standardize_unit(unit_str: str, *, auto_prompt: bool = True) -> str:
    """将原始单位字符串标准化。

    规则（按优先级）：
    1. 空值 / 纯数字 → ""
    2. 硬编码无效词 → ""
    3. 在 unit_map.json 中 → 返回映射值（可能为 ""）
    4. 本次已被跳过的未知单位 → 保持原值
    5. 未知单位 + auto_prompt=True → 弹对话框，用户确认后写回 JSON
    6. 未知单位 + auto_prompt=False → 返回原值
    """
    if not unit_str or str(unit_str).strip() == "":
        return ""
    unit_str = str(unit_str).strip()

    # 纯数字过滤
    try:
        float(unit_str)
        return ""
    except ValueError:
        pass

    load_unit_map()

    # 无效词
    if _invalid_units_cache is not None and unit_str in _invalid_units_cache:
        return ""

    if _alias_to_canonical_cache is not None and unit_str in _alias_to_canonical_cache:
        return _alias_to_canonical_cache[unit_str]

    if unit_str in _runtime_overrides:
        return _runtime_overrides[unit_str]

    # ── 未知单位 ──────────────────────────────────────
    if unit_str in _skipped_unknowns:
        return unit_str  # 本次已跳过，不再询问

    if not auto_prompt:
        return unit_str

    # 改为批量模式：先暂存，稍后统一弹窗处理。
    _pending_unknowns.add(unit_str)
    return unit_str
