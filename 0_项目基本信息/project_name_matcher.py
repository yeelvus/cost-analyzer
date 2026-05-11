import json
import os
import re
from datetime import datetime
import pandas as pd

try:
    import tkinter as tk
    from tkinter import ttk
except Exception:
    tk = None
    ttk = None


def _clean_name(text):
    if text is None:
        return ""
    s = str(text).strip()
    if not s:
        return ""
    s = re.sub(r"（.*?）|\(.*?\)", "", s)
    s = s.replace("泰国", "").replace("项目", "").replace("工程", "")
    return s.strip().lower()


def _normalize_key(text):
    if text is None:
        return ""
    return re.sub(r"\s+", "", str(text)).strip().lower()


def _now_text():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


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


def build_standard_projects_from_dataframe(info_df):
    rows = []
    for _, row in info_df.iterrows():
        name = str(row.get("项目名称", "")).strip()
        if not name or name.lower() == "nan":
            continue
        code = str(row.get("项目编码", "需手动核对")).strip()
        status = str(row.get("备注", "需手动核对")).strip()
        rows.append({
            "project_name": name,
            "project_code": code if code and code.lower() != "nan" else "需手动核对",
            "target_status": status if status and status.lower() != "nan" else "需手动核对",
        })
    return rows


def load_or_init_db(db_path, standard_projects=None):
    db = {
        "version": 1,
        "updated_at": _now_text(),
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
                    "updated_at": loaded.get("updated_at", _now_text()),
                    "projects": loaded.get("projects", {}) or {},
                    "unmatched_aliases": loaded.get("unmatched_aliases", []) or [],
                })
        except Exception:
            pass

    db["unmatched_aliases"] = [
        str(name).strip() for name in db.get("unmatched_aliases", [])
        if str(name).strip()
    ]

    # Merge standard projects into DB as base data.
    for item in (standard_projects or []):
        std_name = item["project_name"]
        db["projects"].setdefault(std_name, {
            "project_code": item.get("project_code", "需手动核对"),
            "target_status": item.get("target_status", "需手动核对"),
            "aliases": [],
        })

        # Keep latest standard code/status from source table.
        db["projects"][std_name]["project_code"] = item.get("project_code", "需手动核对")
        db["projects"][std_name]["target_status"] = item.get("target_status", "需手动核对")
        if std_name in db["unmatched_aliases"]:
            db["unmatched_aliases"].remove(std_name)

    db["updated_at"] = _now_text()
    return db


def save_db(db_path, db):
    db["updated_at"] = _now_text()
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


def mark_unmatched(db, raw_name):
    if raw_name is None:
        return

    raw_name = str(raw_name).strip()
    if not raw_name:
        return

    std_name = _build_alias_map(db).get(raw_name)
    if std_name and std_name in db.get("projects", {}):
        return

    unmatched_aliases = db.setdefault("unmatched_aliases", [])
    if raw_name not in unmatched_aliases:
        unmatched_aliases.append(raw_name)


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


def match_project(raw_name, db):
    if raw_name is None:
        return None

    raw_name = str(raw_name).strip()
    if not raw_name:
        return None

    projects = db.get("projects", {})
    alias_map = _build_alias_map(db)

    # 1) Exact alias map.
    std_name = alias_map.get(raw_name)
    if std_name and std_name in projects:
        meta = projects[std_name]
        return {
            "raw_name": raw_name,
            "project_name": std_name,
            "project_code": meta.get("project_code", "需手动核对"),
            "target_status": meta.get("target_status", "需手动核对"),
            "source": "alias",
        }

    # 2) Normalized exact.
    raw_key = _normalize_key(raw_name)
    for alias, std in alias_map.items():
        if _normalize_key(alias) == raw_key and std in projects:
            meta = projects[std]
            add_mapping(db, raw_name, std)
            return {
                "raw_name": raw_name,
                "project_name": std,
                "project_code": meta.get("project_code", "需手动核对"),
                "target_status": meta.get("target_status", "需手动核对"),
                "source": "normalized_alias",
            }

    return None


def prompt_unmatched_project_mappings(unmatched_names, db, window_title="手动项目匹配", db_path=None):
    """
    为未匹配项目弹窗选择标准项目，并将映射写入 db。
    新增 db_path 参数：如果传入，则手动匹配后自动保存到文件。
    """
    unmatched_names = [str(n).strip() for n in unmatched_names if str(n).strip()]
    unmatched_names = sorted(set(unmatched_names))
    if not unmatched_names:
        return {}

    if tk is None or ttk is None:
        return {name: "" for name in unmatched_names}

    options = get_project_options(db)
    if not options:
        return {name: "" for name in unmatched_names}

    root = tk.Tk()
    root.withdraw()

    win = tk.Toplevel(root)
    win.title(f"{window_title}（{len(unmatched_names)}项）")
    win.geometry("960x640")
    win.lift()
    win.attributes("-topmost", True)
    win.after_idle(win.attributes, "-topmost", False)

    tk.Label(
        win,
        text="以下项目未自动匹配，请选择标准项目（可留空保持原名称）",
        font=("Arial", 12, "bold"),
        pady=10,
    ).pack()

    canvas = tk.Canvas(win)
    scrollbar = tk.Scrollbar(win, orient="vertical", command=canvas.yview)
    frame = tk.Frame(canvas)
    frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.pack(side="left", fill="both", expand=True, padx=10)
    scrollbar.pack(side="right", fill="y")

    blank_option = "（保持原名称，不匹配）"
    display_options = [blank_option]
    option_to_standard = {}

    for std_name in options:
        display_options.append(std_name)
        option_to_standard[std_name] = std_name

    var_map = {}
    for raw_name in unmatched_names:
        row = tk.Frame(frame, relief=tk.RIDGE, borderwidth=1)
        row.pack(fill="x", padx=8, pady=4)

        tk.Label(row, text=raw_name, font=("Arial", 11, "bold"), width=30, anchor="w").pack(side="left", padx=8, pady=6)
        tk.Label(row, text="→", font=("Arial", 12)).pack(side="left")

        var = tk.StringVar(value=blank_option)
        combo = ttk.Combobox(row, textvariable=var, values=display_options, width=70, state="readonly")
        combo.pack(side="left", padx=8, pady=6, fill="x", expand=True)
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
    tk.Button(btns, text="确定并应用", command=on_confirm, width=16, bg="#4CAF50", fg="white").pack(side="left", padx=10)
    tk.Button(btns, text="全部保留原名称", command=on_skip_all, width=16).pack(side="left", padx=10)

    win.protocol("WM_DELETE_WINDOW", on_skip_all)
    root.wait_window(win)
    root.destroy()

    # ====================== 核心修改部分 ======================
    selected_map = {}
    any_mapped = False

    for raw_name, var in var_map.items():
        selected = var.get().strip()
        if not result["ok"] or selected == blank_option or not selected:
            # 用户明确选择"保留原名称"：将原名映射到自身，下次运行直接命中，不再弹窗
            add_mapping(db, raw_name, raw_name)
            selected_map[raw_name] = ""
            continue

        std_name = option_to_standard.get(selected, "")
        selected_map[raw_name] = std_name
        if std_name:
            add_mapping(db, raw_name, std_name)
            any_mapped = True

    # 如果进行了任何映射（含"保留原名"），并且传入了 db_path，则立即保存到文件
    if db_path and selected_map and os.path.exists(os.path.dirname(db_path)):
        save_db(db_path, db)
        print(f"✅ 手动匹配结果已自动保存到数据库: {db_path}")

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