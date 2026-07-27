import json
import os
import re
from difflib import SequenceMatcher

import pandas as pd

try:
    import tkinter as tk
    from tkinter import ttk
except Exception:
    tk = None
    ttk = None

# 运行期缓存：避免每次 match 都重建别名表
_ALIAS_CACHE: dict[int, dict] = {}
_FUZZY_CACHE: dict[tuple[int, str], dict | None] = {}


def _clean_name(text):
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""
    s = re.sub(r"（.*?）|\(.*?\)", "", s)
    s = s.replace("泰国", "").replace("项目", "").replace("工程", "")
    s = re.sub(r"\s+", "", s)
    return s.strip().lower()


def _normalize_key(text):
    if text is None:
        return ""
    return re.sub(r"\s+", "", str(text)).strip().lower()


def _invalidate_db_caches(db=None):
    """映射变更后清缓存。"""
    global _ALIAS_CACHE, _FUZZY_CACHE
    if db is None:
        _ALIAS_CACHE.clear()
        _FUZZY_CACHE.clear()
        return
    _ALIAS_CACHE.pop(id(db), None)
    # 模糊缓存按 db id 前缀清理
    kid = id(db)
    drop = [k for k in _FUZZY_CACHE if k[0] == kid]
    for k in drop:
        _FUZZY_CACHE.pop(k, None)


def _cached_alias_map(db) -> dict:
    key = id(db)
    m = _ALIAS_CACHE.get(key)
    if m is None:
        m = _build_alias_map(db)
        _ALIAS_CACHE[key] = m
    return m


def _iter_project_alias_pairs(db):
    """从 projects 结构生成 (alias, standard_name) 对。"""
    projects = db.get("projects", {}) or {}
    for std_name, meta in projects.items():
        std_name = str(std_name).strip()
        if not std_name:
            continue

        yield std_name, std_name
        for alias in meta.get("aliases", []) or []:
            alias = str(alias).strip()
            if alias:
                yield alias, std_name


def _build_alias_map(db):
    """统一别名映射来源：使用 projects.aliases。"""
    alias_map = {}

    projects = db.get("projects", {}) or {}

    # 1) 显式 aliases 优先，避免被同名“伪项目”覆盖。
    for std_name, meta in projects.items():
        std_name = str(std_name).strip()
        if not std_name:
            continue
        for alias in meta.get("aliases", []) or []:
            alias = str(alias).strip()
            if alias:
                alias_map.setdefault(alias, std_name)

    # 2) 标准名自身兜底映射。
    for std_name in projects.keys():
        std_name = str(std_name).strip()
        if std_name:
            alias_map.setdefault(std_name, std_name)

    return alias_map


def _pick_status_column(columns) -> str | None:
    """项目信息表状态列优先顺序：项目状态 > 备注 > 状态 > target_status。"""
    cols = [str(c).strip() for c in columns]
    for name in ("项目状态", "备注", "状态", "target_status"):
        if name in cols:
            return name
    return None


def _clean_field(value, default: str = "") -> str:
    if value is None:
        return default
    s = str(value).strip()
    if not s or s.lower() in {"nan", "none", "nat"}:
        return default
    return s


def build_standard_projects_from_dataframe(info_df):
    """
    从项目基本信息表构建标准项目列表。
    状态字段读取「项目状态」（兼容旧表「备注」），不再误用空备注覆盖为「需手动核对」。
    """
    status_col = _pick_status_column(info_df.columns)
    rows = []
    for _, row in info_df.iterrows():
        name = _clean_field(row.get("项目名称", ""))
        if not name:
            continue
        code = _clean_field(row.get("项目编码", ""), default="需手动核对")
        if status_col:
            status = _clean_field(row.get(status_col, ""), default="需手动核对")
        else:
            status = "需手动核对"
        rows.append({
            "project_name": name,
            "project_code": code,
            "target_status": status,
        })
    return rows


def load_or_init_db(db_path, standard_projects=None):
    """
    加载匹配库，并可选合并标准项目列表。

    合并规则（避免每次运行把状态刷成「需手动核对」）：
    - 新项目：写入编码/状态（缺省才用「需手动核对」）
    - 已有项目：仅当本次传入了「有意义」的编码/状态时才覆盖
      （调用方若只传 project_name+project_code、不传 target_status，则保留库内原状态）
    """
    db = {
        "version": 1,
        "projects": {},
        "unmatched_aliases": [],
    }

    if os.path.exists(db_path):
        try:
            with open(db_path, "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                db.update({
                    "version": loaded.get("version", 1),
                    "projects": loaded.get("projects", {}) or {},
                    "unmatched_aliases": loaded.get("unmatched_aliases", []) or [],
                })
        except Exception:
            pass

    db["unmatched_aliases"] = [
        str(name).strip() for name in db.get("unmatched_aliases", [])
        if str(name).strip()
    ]

    for item in (standard_projects or []):
        std_name = _clean_field(item.get("project_name", ""))
        if not std_name:
            continue

        # 调用方是否显式提供了字段（用于区分「未传」与「传了空」）
        has_code_key = "project_code" in item
        has_status_key = "target_status" in item
        code = _clean_field(item.get("project_code", "")) if has_code_key else ""
        status = _clean_field(item.get("target_status", "")) if has_status_key else ""

        if std_name not in db["projects"]:
            db["projects"][std_name] = {
                "project_code": code or "需手动核对",
                "target_status": status or "需手动核对",
                "aliases": [],
            }
        else:
            meta = db["projects"][std_name]
            # 有明确编码且不是占位时才覆盖
            if has_code_key and code and code != "需手动核对":
                meta["project_code"] = code
            elif not _clean_field(meta.get("project_code", "")):
                meta["project_code"] = code or "需手动核对"

            # 有明确状态时才覆盖；未传 target_status 则保留原值
            if has_status_key and status:
                meta["target_status"] = status
            elif not _clean_field(meta.get("target_status", "")):
                meta["target_status"] = "需手动核对"

            meta.setdefault("aliases", [])

        if std_name in db["unmatched_aliases"]:
            db["unmatched_aliases"].remove(std_name)

    return db


def save_db(db_path, db):
    # 匹配库不再维护 updated_at，避免每次运行都触发无意义变更。
    db.pop("updated_at", None)
    os.makedirs(os.path.dirname(db_path), exist_ok=True)
    with open(db_path, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=2)


def get_project_options(db):
    return sorted(db.get("projects", {}).keys())


def get_project_lookup(db):
    lookup = {}
    for name, meta in db.get("projects", {}).items():
        lookup[name] = (
            meta.get("project_code", "需手动核对"),
            meta.get("target_status", "需手动核对"),
        )
    return lookup


def add_mapping(db, match_name, standard_name):
    if not match_name or not standard_name:
        return

    match_name = str(match_name).strip()
    standard_name = str(standard_name).strip()
    if not match_name or not standard_name:
        return

    if standard_name not in db["projects"]:
        db["projects"][standard_name] = {
            "project_code": "需手动核对",
            "target_status": "需手动核对",
            "aliases": [],
        }

    aliases = db["projects"][standard_name].setdefault("aliases", [])
    if match_name not in aliases and match_name != standard_name:
        aliases.append(match_name)
    clear_unmatched_aliases(db, match_name)
    _invalidate_db_caches(db)


def mark_unmatched(db, raw_name):
    if raw_name is None:
        return

    raw_name = str(raw_name).strip()
    if not raw_name:
        return

    std_name = _cached_alias_map(db).get(raw_name)
    if std_name and std_name in db.get("projects", {}):
        return

    unmatched_aliases = db.setdefault("unmatched_aliases", [])
    # 归一化去重，避免同名不同空格反复写入
    raw_key = _normalize_key(raw_name)
    for name in unmatched_aliases:
        if _normalize_key(name) == raw_key:
            return
    unmatched_aliases.append(raw_name)
    _invalidate_db_caches(db)


def clear_unmatched_aliases(db, raw_name=None):
    unmatched_aliases = db.setdefault("unmatched_aliases", [])
    if raw_name is None:
        unmatched_aliases.clear()
        return

    raw_key = _normalize_key(raw_name)
    db["unmatched_aliases"] = [
        name for name in unmatched_aliases
        if _normalize_key(name) != raw_key
    ]


def is_unmatched_alias(db, raw_name):
    if raw_name is None:
        return False

    raw_key = _normalize_key(raw_name)
    if not raw_key:
        return False

    for name in db.get("unmatched_aliases", []):
        if _normalize_key(name) == raw_key:
            return True

    return False


def _result(raw_name, std_name, meta, source):
    return {
        "raw_name": raw_name,
        "project_name": std_name,
        "project_code": meta.get("project_code", "需手动核对"),
        "target_status": meta.get("target_status", "需手动核对"),
        "source": source,
    }


def _fuzzy_match_project(raw_name: str, db, *, min_ratio: float = 0.82):
    """
    软匹配：清洗后全等 / 唯一子串包含 / 高相似度唯一最优。
    仅在唯一命中时返回，避免误绑。
    """
    cache_key = (id(db), _normalize_key(raw_name))
    if cache_key in _FUZZY_CACHE:
        return _FUZZY_CACHE[cache_key]

    projects = db.get("projects", {}) or {}
    if not projects:
        _FUZZY_CACHE[cache_key] = None
        return None

    raw_clean = _clean_name(raw_name)
    raw_key = _normalize_key(raw_name)
    if not raw_clean and not raw_key:
        _FUZZY_CACHE[cache_key] = None
        return None

    # 清洗后全等（标准名或别名）
    alias_map = _cached_alias_map(db)
    clean_hits = []
    for alias, std in alias_map.items():
        if std not in projects:
            continue
        if _clean_name(alias) == raw_clean or _normalize_key(alias) == raw_key:
            clean_hits.append(std)
    clean_hits = list(dict.fromkeys(clean_hits))
    if len(clean_hits) == 1:
        std = clean_hits[0]
        res = _result(raw_name, std, projects[std], "fuzzy_clean_exact")
        _FUZZY_CACHE[cache_key] = res
        return res

    # 唯一子串包含（长度足够，避免短词误伤）
    if len(raw_clean) >= 4:
        contain = []
        for std in projects.keys():
            sc = _clean_name(std)
            if not sc:
                continue
            if raw_clean in sc or sc in raw_clean:
                # 较短一方至少 4 字，且覆盖率较高
                shorter, longer = (raw_clean, sc) if len(raw_clean) <= len(sc) else (sc, raw_clean)
                if len(shorter) >= 4 and len(shorter) / max(len(longer), 1) >= 0.5:
                    contain.append(std)
        contain = list(dict.fromkeys(contain))
        if len(contain) == 1:
            std = contain[0]
            res = _result(raw_name, std, projects[std], "fuzzy_contain")
            _FUZZY_CACHE[cache_key] = res
            return res

    # 相似度：取唯一最高且 ≥ 阈值
    best_std, best_r, second_r = None, 0.0, 0.0
    for std in projects.keys():
        sc = _clean_name(std) or _normalize_key(std)
        if not sc:
            continue
        r = SequenceMatcher(None, raw_clean or raw_key, sc).ratio()
        if r > best_r:
            second_r = best_r
            best_r = r
            best_std = std
        elif r > second_r:
            second_r = r
    if (
        best_std
        and best_r >= min_ratio
        and (best_r - second_r) >= 0.05  # 与第二名拉开差距，避免双候选
    ):
        res = _result(raw_name, best_std, projects[best_std], f"fuzzy_ratio_{best_r:.2f}")
        _FUZZY_CACHE[cache_key] = res
        return res

    _FUZZY_CACHE[cache_key] = None
    return None


def match_project(raw_name, db, *, fuzzy: bool = True, skip_marked_unmatched: bool = True):
    """
    匹配项目名。

    - fuzzy: 启用清洗/包含/相似度软匹配
    - skip_marked_unmatched: 若曾在弹窗中「保留原名」写入 unmatched_aliases，
      则不再返回 None 触发重复弹窗，而是以原名静默通过（source=unmatched_keep）
    """
    if raw_name is None:
        return None

    raw_name = str(raw_name).strip()
    if not raw_name:
        return None

    projects = db.get("projects", {})
    alias_map = _cached_alias_map(db)

    # 1) Exact alias map.
    std_name = alias_map.get(raw_name)
    if std_name and std_name in projects:
        return _result(raw_name, std_name, projects[std_name], "alias")

    # 2) Normalized exact.
    raw_key = _normalize_key(raw_name)
    for alias, std in alias_map.items():
        if _normalize_key(alias) == raw_key and std in projects:
            meta = projects[std]
            add_mapping(db, raw_name, std)
            return _result(raw_name, std, meta, "normalized_alias")

    # 3) Fuzzy
    if fuzzy:
        hit = _fuzzy_match_project(raw_name, db)
        if hit:
            # 自动写入别名，下次精确命中
            add_mapping(db, raw_name, hit["project_name"])
            return hit

    # 4) 已标记「不匹配/保留原名」→ 静默用原名，避免反复弹窗
    if skip_marked_unmatched and is_unmatched_alias(db, raw_name):
        return {
            "raw_name": raw_name,
            "project_name": raw_name,
            "project_code": "",
            "target_status": "未匹配保留",
            "source": "unmatched_keep",
        }

    return None


def collect_names_needing_prompt(raw_names, db, *, fuzzy: bool = True) -> list[str]:
    """筛出真正需要弹窗的名称（已匹配 / 已保留原名的不再弹出）。"""
    need = []
    seen = set()
    for raw in raw_names:
        name = str(raw or "").strip()
        if not name:
            continue
        key = _normalize_key(name)
        if key in seen:
            continue
        seen.add(key)
        if match_project(name, db, fuzzy=fuzzy, skip_marked_unmatched=True):
            continue
        need.append(name)
    return sorted(need)


def prompt_unmatched_project_mappings(
    unmatched_names,
    db,
    window_title="手动项目匹配",
    db_path=None,
    *,
    interactive: bool = True,
):
    """
    为未匹配项目弹窗选择标准项目，并将映射写入 db。

    interactive=False：不弹窗，全部 mark_unmatched 并返回空映射。
    弹窗内默认选中模糊建议（若有），减少手工选择。
    """
    unmatched_names = [str(n).strip() for n in unmatched_names if str(n).strip()]
    # 再过滤一轮：模糊已能匹配的不再弹
    unmatched_names = collect_names_needing_prompt(unmatched_names, db, fuzzy=True)
    if not unmatched_names:
        return {}

    if not interactive or tk is None or ttk is None:
        for name in unmatched_names:
            mark_unmatched(db, name)
        if db_path and os.path.exists(os.path.dirname(db_path) or "."):
            save_db(db_path, db)
        return {name: "" for name in unmatched_names}

    options = get_project_options(db)
    if not options:
        for name in unmatched_names:
            mark_unmatched(db, name)
        return {name: "" for name in unmatched_names}

    # 预计算建议
    suggestions: dict[str, str] = {}
    for raw in unmatched_names:
        # 临时关闭 skip，仅看 fuzzy（未写入别名）
        hit = _fuzzy_match_project(raw, db, min_ratio=0.72)
        if hit:
            suggestions[raw] = hit["project_name"]

    root = tk.Tk()
    root.withdraw()

    win = tk.Toplevel(root)
    win.title(f"{window_title}（{len(unmatched_names)}项，已自动过滤已知项）")
    win.geometry("980x660")
    win.lift()
    win.attributes("-topmost", True)
    win.after_idle(win.attributes, "-topmost", False)

    tip = (
        "以下为新出现的未匹配项目。下拉可改；灰色建议已预填。"
        "选「保持原名称」后下次不再弹窗。"
    )
    tk.Label(win, text=tip, font=("Arial", 11, "bold"), pady=8, wraplength=900, justify="left").pack()

    canvas = tk.Canvas(win)
    scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas)
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True, padx=10)
    scrollbar.pack(side="right", fill="y")

    blank_option = "（保持原名称，不匹配）"
    display_options = [blank_option] + list(options)
    option_to_standard = {s: s for s in options}

    var_map = {}
    for raw_name in unmatched_names:
        row = tk.Frame(frame, relief=tk.RIDGE, borderwidth=1)
        row.pack(fill="x", padx=8, pady=4)

        tk.Label(row, text=raw_name, font=("Arial", 11, "bold"), width=28, anchor="w").pack(
            side="left", padx=8, pady=6
        )
        tk.Label(row, text="→", font=("Arial", 12)).pack(side="left")

        suggest = suggestions.get(raw_name, "")
        default = suggest if suggest in option_to_standard else blank_option
        var = tk.StringVar(value=default)
        combo = ttk.Combobox(
            row, textvariable=var, values=display_options, width=68, state="readonly"
        )
        combo.pack(side="left", padx=8, pady=6, fill="x", expand=True)
        if suggest:
            tk.Label(row, text=f"建议:{suggest[:20]}", fg="#666", font=("Arial", 9)).pack(
                side="left", padx=4
            )
        var_map[raw_name] = var

    result = {"ok": False}

    def on_confirm():
        result["ok"] = True
        win.destroy()

    def on_skip_all():
        result["ok"] = False
        win.destroy()

    btns = tk.Frame(win)
    btns.pack(pady=10)
    tk.Button(
        btns, text="确定并应用", command=on_confirm, width=16, bg="#4CAF50", fg="white"
    ).pack(side="left", padx=10)
    tk.Button(btns, text="全部保留原名称", command=on_skip_all, width=16).pack(
        side="left", padx=10
    )

    win.protocol("WM_DELETE_WINDOW", on_skip_all)
    root.wait_window(win)
    root.destroy()

    selected_map = {}
    for raw_name, var in var_map.items():
        selected = var.get().strip()
        if not result["ok"] or selected == blank_option or not selected:
            mark_unmatched(db, raw_name)
            selected_map[raw_name] = ""
            continue

        std_name = option_to_standard.get(selected, "")
        selected_map[raw_name] = std_name
        if std_name:
            add_mapping(db, raw_name, std_name)

    if db_path and selected_map:
        parent = os.path.dirname(db_path)
        if parent and os.path.exists(parent):
            save_db(db_path, db)
            print(f"✅ 手动匹配结果已保存: {db_path}")

    return selected_map


def update_db_from_excel(db_path, excel_path,
                         status_column="备注",
                         name_column="项目名称",
                         code_column="项目编码"):
    """以项目编码为键，从 Excel 更新匹配库。"""
    if not os.path.exists(excel_path):
        print(f"❌ Excel 文件不存在: {excel_path}")
        return None

    print(f"正在读取 Excel: {excel_path}")
    df = pd.read_excel(excel_path)
    df = df.dropna(how="all")

    required_cols = [code_column, name_column, status_column]
    for col in required_cols:
        if col not in df.columns:
            print(f"⚠️ Excel 中缺少列 '{col}'，请检查列名是否正确")
            print(f"当前列名: {list(df.columns)}")
            return None

    db = load_or_init_db(db_path)

    updated_count = 0
    new_count = 0

    for _, row in df.iterrows():
        code = str(row.get(code_column, "")).strip()
        if not code or code.lower() in ["nan", "none", ""]:
            continue

        name = str(row.get(name_column, "")).strip()
        status = str(row.get(status_column, "")).strip()
        if status.lower() in ["nan", "none", ""]:
            status = ""

        found = False
        for std_name, meta in list(db["projects"].items()):
            if str(meta.get("project_code", "")).strip() == code:
                old_status = str(meta.get("target_status", "")).strip()
                if old_status != status:
                    meta["target_status"] = status
                    updated_count += 1

                if name and name != std_name and name not in meta.get("aliases", []):
                    aliases = meta.setdefault("aliases", [])
                    aliases.append(name)

                found = True
                break

        if not found:
            if not name:
                name = f"未知项目_{code}"
            db["projects"][name] = {
                "project_code": code,
                "target_status": status,
                "aliases": [],
            }
            new_count += 1
            print(f"✅ 新增项目: {name} (编码: {code}) 状态: {status or '空'}")

    save_db(db_path, db)

    print("\n🎉 更新完成！")
    print(f"   • 更新状态: {updated_count} 条")
    print(f"   • 新增项目: {new_count} 个")
    print(f"   • 当前标准项目总数: {len(db.get('projects', {}))}")
    return db