# -*- coding: utf-8 -*-
"""raw 多月 Excel 合并 + 列标准化 + 项目匹配。

  B_机具费_raw_sum/机具费_raw合并.xlsx   ← 过程全量合并
  C_机具费_raw_split/{编码}_{项目}_机具费.xlsx  ← 按项目拆分（仅明细）
"""
from __future__ import annotations

import json
import re
from pathlib import Path

import pandas as pd

from machine.config import (
    LEGACY_CLEAN_MERGE_NAME,
    PROJECT_MATCH_DB,
    RAW_DIR,
    RAW_MERGE_NAME,
    RAW_SPLIT_DIR,
    RAW_SUM_DIR,
    ensure_data_dirs,
)


STD_COLUMNS = [
    "序号", "机具种类", "型号规格", "台班费（泰铢）",
    "本次结算台班（小时、天、月）", "加班（小时）", "加班费/超时费",
    "进出场费/运费", "移动费＋拆装费+接管费", "结算费用（泰铢）",
    "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%",
    "本次费用合计（泰铢）", "使用日期", "出租单位", "使用项目", "备注", "月份",
]


def normalize_headers(headers_raw) -> list[str]:
    new_headers = []
    for c in headers_raw:
        c_stripped = str(c).strip() if pd.notna(c) else ""
        if re.match(r"^序号$", c_stripped):
            new_headers.append("序号")
        elif re.match(r"^机具种类$", c_stripped):
            new_headers.append("机具种类")
        elif re.match(r"^型号规格$", c_stripped):
            new_headers.append("型号规格")
        elif "台班费" in c_stripped:
            new_headers.append("台班费（泰铢）")
        elif "结算台班" in c_stripped or "本次结算台班" in c_stripped:
            new_headers.append("本次结算台班（小时、天、月）")
        elif c_stripped == "加班费（小时）":
            new_headers.append("加班费/超时费")
        elif "加班费" in c_stripped and "超时" in c_stripped:
            new_headers.append("加班费/超时费")
        elif "加班" in c_stripped and ("小时" in c_stripped or "(" in c_stripped or "（" in c_stripped):
            new_headers.append("加班（小时）")
        elif c_stripped == "加班费":
            new_headers.append("__drop__")
        elif "进出场" in c_stripped or c_stripped in ("运费＋拆装费", "进出场费/运费"):
            new_headers.append("进出场费/运费")
        elif "移动费" in c_stripped or "拆装费" in c_stripped or "接管费" in c_stripped:
            new_headers.append("移动费＋拆装费+接管费")
        elif "结算费用" in c_stripped or "本次结算费用" in c_stripped:
            new_headers.append("结算费用（泰铢）")
        elif c_stripped == "增值税7%":
            new_headers.append("增值税7%")
        elif "预扣税" in c_stripped:
            new_headers.append("机具租赁预扣税5%" if "5%" in c_stripped else "机具租赁预扣税3%")
        elif "费用合计" in c_stripped:
            new_headers.append("本次费用合计（泰铢）")
        elif "使用时间" in c_stripped or "使用日期" in c_stripped:
            new_headers.append("使用日期")
        elif "出租单位" in c_stripped:
            new_headers.append("出租单位")
        elif "使用项目" in c_stripped:
            new_headers.append("使用项目")
        elif "备注" in c_stripped:
            new_headers.append("备注")
        elif "月份" in c_stripped:
            new_headers.append("月份")
        else:
            print(f"  ⚠️  无法映射列: [{c_stripped}]，跳过")
            new_headers.append("__drop__")
    return new_headers



def normalize_month_value(x) -> str | None:
    """统一月份为 YYYY-MM 文本。"""
    if x is None or (isinstance(x, float) and pd.isna(x)):
        return None
    # datetime / Timestamp
    try:
        import pandas as _pd
        if hasattr(x, "strftime"):
            return x.strftime("%Y-%m")
        ts = _pd.to_datetime(x, errors="coerce")
        if _pd.notna(ts):
            return ts.strftime("%Y-%m")
    except Exception:
        pass
    s = str(x).strip()
    if not s or s.lower() == "nan":
        return None
    # 2026年4月 / 2026-04 / 202604
    m = re.search(r"(20\d{2})\s*[-/年.]\s*(\d{1,2})", s)
    if m:
        return f"{m.group(1)}-{int(m.group(2)):02d}"
    m = re.search(r"^(20\d{2})(0[1-9]|1[0-2])$", s)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    return s


def process_file(filepath: Path) -> pd.DataFrame:
    raw = pd.read_excel(filepath, header=None)
    print(f"  原始 shape: {raw.shape}")
    raw = raw.iloc[:, :19]
    headers_raw = list(raw.iloc[0])
    data = raw.iloc[1:].copy().reset_index(drop=True)
    print(f"  数据行数: {len(data)}")

    new_headers = normalize_headers(headers_raw)
    keep_indices = [i for i, h in enumerate(new_headers) if h != "__drop__"]
    final_headers = [h for h in new_headers if h != "__drop__"]
    data = data.iloc[:, keep_indices].copy()
    data.columns = final_headers

    if "移动费＋拆装费+接管费" not in data.columns:
        data["移动费＋拆装费+接管费"] = None
    for col in STD_COLUMNS:
        if col not in data.columns:
            data[col] = None
    data = data[STD_COLUMNS]
    if "月份" in data.columns:
        data["月份"] = data["月份"].apply(normalize_month_value)
    print(f"  处理后列数: {len(data.columns)}")
    return data


def build_alias_maps(db: dict):
    alias_map, code_map = {}, {}
    for std, meta in (db.get("projects") or {}).items():
        std_s = str(std).strip()
        if not std_s:
            continue
        alias_map[std_s] = std_s
        for a in meta.get("aliases") or []:
            a_s = str(a).strip()
            if a_s:
                alias_map[a_s] = std_s
        code_map[std_s] = meta.get("project_code")
    return alias_map, code_map


def match_by_db(raw_name, alias_map, code_map):
    if raw_name is None:
        return None, None
    s = str(raw_name).strip()
    if not s:
        return None, None
    std = alias_map.get(s)
    if std:
        return std, code_map.get(std)
    key = re.sub(r"\s+", "", s).lower()
    for k, v in alias_map.items():
        if re.sub(r"\s+", "", str(k)).lower() == key:
            return v, code_map.get(v)
    return None, None


def _cleanup_legacy_clean(out_dir: Path) -> None:
    """删除历史遗留的 clean_ 副本，避免汇总误读。"""
    legacy = out_dir / LEGACY_CLEAN_MERGE_NAME
    if legacy.exists():
        try:
            legacy.unlink()
            print(f"  已删除废弃文件: {legacy.name}")
        except Exception as e:
            print(f"  ⚠️ 无法删除 {legacy.name}: {e}")


def run_merge_raw(
    raw_dir: Path | None = None,
    out_dir: Path | None = None,
    *,
    split_by_project: bool = True,
) -> Path:
    """合并 raw → B_B_机具费_raw_sum/机具费_raw合并.xlsx，并按项目拆到 C_机具费_raw_split。"""
    ensure_data_dirs()
    raw_dir = Path(raw_dir or RAW_DIR)
    out_dir = Path(out_dir or RAW_SUM_DIR)
    out_dir.mkdir(parents=True, exist_ok=True)
    _cleanup_legacy_clean(out_dir)
    # 同步清理旧 std 目录下的废弃 clean 副本
    from machine.config import LEGACY_STD_DIR
    _cleanup_legacy_clean(LEGACY_STD_DIR)

    from common.paths import iter_excel_files as _iter_xlsx

    files = [
        f for f in _iter_xlsx(raw_dir, recursive=True)
        if f.suffix.lower() == ".xlsx"
    ]
    print(f"找到 {len(files)} 个 raw 文件: {[f.name for f in files]}")
    if not files:
        raise FileNotFoundError(f"raw 目录无 xlsx: {raw_dir}")

    frames = []
    for f in files:
        print(f"\n📄 处理: {f.name}")
        df = process_file(f)
        print(f"  ✅ 行数: {len(df)}")
        frames.append(df)

    merged = pd.concat(frames, ignore_index=True)
    print(f"\n{'='*60}\n✅ 合并完成！总行数: {len(merged)}")

    # 项目匹配
    try:
        db = json.loads(Path(PROJECT_MATCH_DB).read_text(encoding="utf-8"))
        alias_map, code_map = build_alias_maps(db)
        matcher = lambda raw: match_by_db(raw, alias_map, code_map)
        print(f"✅ 匹配库: {PROJECT_MATCH_DB}")
    except Exception as e:
        print(f"⚠️ 匹配库加载失败: {e}")
        matcher = lambda raw: (None, None)

    if "使用项目" in merged.columns:
        matches = merged["使用项目"].apply(
            lambda x: matcher(x) if pd.notna(x) and str(x).strip() else (None, None)
        )
        merged["项目名称"] = [m[0] for m in matches]
        merged["项目编码"] = [m[1] for m in matches]
        n_ok = merged["项目编码"].notna().sum()
        print(f"  ✅ 已匹配项目编码/名称: {n_ok}/{len(merged)} 行")
    else:
        merged["项目名称"] = None
        merged["项目编码"] = None

    if "月份" in merged.columns:
        merged = merged.sort_values(
            by=[c for c in ["项目编码", "项目名称", "月份"] if c in merged.columns],
            na_position="last",
        ).reset_index(drop=True)

    # 列顺序：项目字段前置，便于过程核对
    front = [c for c in ["项目编码", "项目名称", "使用项目"] if c in merged.columns]
    rest = [c for c in merged.columns if c not in front]
    merged = merged[front + rest]

    out_path = out_dir / RAW_MERGE_NAME
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="机具费合并", index=False)
        ws = writer.sheets["机具费合并"]
        for col_idx, col_name in enumerate(merged.columns, 1):
            max_len = len(str(col_name)) * 2 + 2
            for val in merged.iloc[:80, col_idx - 1].dropna():
                max_len = max(max_len, min(len(str(val)), 40))
            ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 40)

    print(f"\n📁 过程文件: {out_path}")
    if "月份" in merged.columns:
        print("\n◆ 各月份数据量:")
        print(merged["月份"].value_counts().sort_index().to_string())
    if "项目名称" in merged.columns:
        print(f"\n◆ 项目数: {merged['项目名称'].nunique()}")

    if split_by_project:
        from machine.split import split_raw_merge_by_project
        print(f"\n{'='*60}")
        print("按项目拆分 → C_机具费_raw_split/")
        print("=" * 60)
        split_raw_merge_by_project(out_path, RAW_SPLIT_DIR)

    return out_path


def main(argv=None):
    run_merge_raw(split_by_project=True)


if __name__ == "__main__":
    main()
