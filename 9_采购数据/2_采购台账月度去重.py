from __future__ import annotations

import hashlib
import importlib.util
import re
import shutil
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable

import openpyxl
import pandas as pd
from openpyxl.utils import get_column_letter


# 输入：仍然扫描源目录下所有子文件夹中的 Excel 文件
_DEFAULT_INPUT_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/A_采购部采购台账_raw")
_OUTPUT_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/A_采购台账提取")
_OUTPUT_SUBDIR_NAME = "去重分析输出"

RAW_ROOT = _DEFAULT_INPUT_ROOT
STD_ROOT = _OUTPUT_ROOT


AMOUNT_COL_CANDIDATES = ["含税金额", "总金额", "价税合计", "含税总金额"]
MONTH_PATTERN = re.compile(r"((?:19|20)\d{2})年(0?[1-9]|1[0-2])月")


def normalize_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).replace("\n", " ").replace("\r", " ").strip()


def extract_project_name(file_name: str) -> str:
    stem = Path(file_name).stem.strip()
    if "_" in stem:
        return stem.split("_", 1)[0].strip()
    m = re.match(r"^(.+?项目)", stem)
    if m:
        return m.group(1).strip()
    return stem


def extract_month_label(text: str) -> str | None:
    m = MONTH_PATTERN.search(text)
    if not m:
        return None
    y, mm = int(m.group(1)), int(m.group(2))
    return f"{y:04d}-{mm:02d}"


def month_key(month_label: str) -> tuple[int, int]:
    y, m = month_label.split("-")
    return int(y), int(m)


def list_excel_files(folder: Path) -> list[Path]:
    out: list[Path] = []
    for p in folder.iterdir():
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            continue
        out.append(p)
    return sorted(out)


def sha256_file(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as f:
        while True:
            chunk = f.read(1024 * 1024)
            if not chunk:
                break
            h.update(chunk)
    return h.hexdigest()


def _normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    df2 = df.copy()
    df2.columns = [normalize_text(c) for c in df2.columns]
    for col in df2.columns:
        if pd.api.types.is_numeric_dtype(df2[col]):
            df2[col] = pd.to_numeric(df2[col], errors="coerce").round(6)
        else:
            df2[col] = df2[col].map(normalize_text)
    return df2


def _best_amount_column(columns: Iterable[str]) -> str | None:
    cols = [normalize_text(c) for c in columns]
    for std in AMOUNT_COL_CANDIDATES:
        if std in cols:
            return std
    fuzzy: list[str] = []
    for c in cols:
        if ("含税" in c and "金额" in c) or c in {"金额", "总金额"}:
            fuzzy.append(c)
    return fuzzy[0] if fuzzy else None


def _is_sum_formula_for_amount_column(formula_text: object, amount_col_index: int) -> bool:
    if not isinstance(formula_text, str):
        return False

    f = formula_text.upper().replace(" ", "")
    if "SUM(" not in f:
        return False

    col_letter = get_column_letter(amount_col_index)
    if f"{col_letter}:{col_letter}" in f:
        return True

    range_pattern = rf"\$?{col_letter}\$?\d+:\$?{col_letter}\$?\d+"
    return re.search(range_pattern, f) is not None


def _detect_summary_row_indexes(
    formula_ws: openpyxl.worksheet.worksheet.Worksheet | None,
    normalized_columns: Iterable[str],
    amount_col: str | None,
) -> set[int]:
    if formula_ws is None or not amount_col:
        return set()

    cols = [normalize_text(c) for c in normalized_columns]
    amount_col_index = None
    for idx, col_name in enumerate(cols, start=1):
        if col_name == amount_col:
            amount_col_index = idx
            break

    if amount_col_index is None:
        return set()

    summary_indexes: set[int] = set()
    for row_idx in range(2, formula_ws.max_row + 1):
        cell = formula_ws.cell(row_idx, amount_col_index)
        if cell.data_type != "f":
            continue
        if _is_sum_formula_for_amount_column(cell.value, amount_col_index):
            # DataFrame 默认首行表头，Excel 第2行对应 DataFrame 第0行
            summary_indexes.add(row_idx - 2)

    return summary_indexes


def _sheet_metrics(
    df: pd.DataFrame,
    formula_ws: openpyxl.worksheet.worksheet.Worksheet | None = None,
) -> tuple[float, int, str]:
    if df.empty:
        return 0.0, 0, ""

    ndf = _normalize_dataframe(df)
    amount_col = _best_amount_column(ndf.columns)

    summary_row_indexes = _detect_summary_row_indexes(formula_ws, ndf.columns, amount_col)
    if summary_row_indexes:
        valid_indexes = [idx for idx in summary_row_indexes if 0 <= idx < len(ndf)]
        if valid_indexes:
            ndf = ndf.drop(index=valid_indexes).reset_index(drop=True)

    if ndf.empty:
        return 0.0, 0, ""

    amount_sum = 0.0
    row_count = int(len(ndf))
    if amount_col:
        amount_series = pd.to_numeric(ndf[amount_col], errors="coerce")
        amount_sum = float(amount_series.fillna(0).sum())

    # 用值级别签名辅助识别“内容一致但文件元数据不同”的情况
    sig_df = ndf.copy()
    if amount_col and amount_col in sig_df.columns:
        sig_df[amount_col] = pd.to_numeric(sig_df[amount_col], errors="coerce").round(2)
    sig_df = sig_df.fillna("")

    payload = "|".join(sig_df.columns) + "\n"
    for row in sig_df.itertuples(index=False, name=None):
        payload += "\t".join(normalize_text(v) for v in row) + "\n"

    value_hash = hashlib.sha256(payload.encode("utf-8")).hexdigest()
    return amount_sum, row_count, value_hash


def get_excel_metrics(path: Path) -> tuple[float, int, str]:
    xls = pd.ExcelFile(path)
    formula_wb = openpyxl.load_workbook(path, data_only=False)
    total_sum = 0.0
    total_rows = 0
    hash_acc = hashlib.sha256()

    try:
        for sheet in xls.sheet_names:
            try:
                df = pd.read_excel(path, sheet_name=sheet)
            except Exception:
                continue
            formula_ws = formula_wb[sheet] if sheet in formula_wb.sheetnames else None
            s_sum, s_rows, s_hash = _sheet_metrics(df, formula_ws=formula_ws)
            total_sum += s_sum
            total_rows += s_rows
            hash_acc.update(f"{sheet}:{s_hash}".encode("utf-8"))
    finally:
        formula_wb.close()

    return round(total_sum, 2), total_rows, hash_acc.hexdigest()


@dataclass
class FileRecord:
    月份: str
    项目名称: str
    文件名: str
    文件路径: str
    含税金额合计: float
    行数: int
    文件哈希: str
    内容哈希: str


def iter_monthly_excel_files(raw_root: Path) -> list[tuple[Path, str]]:
    candidates: list[tuple[Path, str]] = []
    seen: set[Path] = set()

    for path in sorted(raw_root.rglob("*")):
        if not path.is_file():
            continue
        if path.name.startswith("~$"):
            continue
        if path.suffix.lower() not in {".xlsx", ".xls", ".xlsm"}:
            continue

        month_label = extract_month_label(str(path))
        if not month_label or path in seen:
            continue

        seen.add(path)
        candidates.append((path, month_label))

    return sorted(candidates, key=lambda item: month_key(item[1]))


def scan_raw_files(raw_root: Path) -> list[FileRecord]:
    records: list[FileRecord] = []

    for excel_path, month_label in iter_monthly_excel_files(raw_root):
        project_name = extract_project_name(excel_path.name)
        file_hash = sha256_file(excel_path)
        amount_sum, row_count, value_hash = get_excel_metrics(excel_path)
        records.append(
            FileRecord(
                月份=month_label,
                项目名称=project_name,
                文件名=excel_path.name,
                文件路径=str(excel_path),
                含税金额合计=amount_sum,
                行数=row_count,
                文件哈希=file_hash,
                内容哈希=value_hash,
            )
        )

    return records


def compare_with_previous(df: pd.DataFrame) -> pd.DataFrame:
    out_rows: list[dict] = []

    for project_name, g in df.groupby("项目名称", sort=False):
        g2 = g.sort_values("月份", key=lambda s: s.map(month_key)).reset_index(drop=True)
        prev_row = None
        for _, row in g2.iterrows():
            reason = "首月保留"
            decision = "保留"

            if prev_row is not None:
                amount_diff = abs(float(row["含税金额合计"]) - float(prev_row["含税金额合计"]))
                row_diff = int(row["行数"]) - int(prev_row["行数"])

                if row["文件哈希"] == prev_row["文件哈希"]:
                    reason = "与上月文件字节完全一致"
                    decision = "复用上月，当前月不体现"
                elif row["内容哈希"] == prev_row["内容哈希"]:
                    reason = "与上月内容值一致（忽略文件元数据）"
                    decision = "复用上月，当前月不体现"
                elif amount_diff <= 0.01 and row_diff == 0:
                    reason = "含税金额合计一致且行数一致（疑似未更新）"
                    decision = "建议复核后复用上月"
                elif amount_diff <= 0.01:
                    reason = "含税金额合计一致但行数变化"
                    decision = "建议人工复核"
                else:
                    reason = f"金额变化 {amount_diff:,.2f}"
                    decision = "保留"

            out_rows.append(
                {
                    **row.to_dict(),
                    "对比结论": reason,
                    "处理建议": decision,
                    "是否在最新口径体现": "是" if decision == "保留" else "否",
                }
            )
            prev_row = row

    return pd.DataFrame(out_rows)


def build_month_summary(decision_df: pd.DataFrame) -> pd.DataFrame:
    grp = (
        decision_df.groupby(["月份", "处理建议"], as_index=False)
        .agg(文件数=("文件名", "count"), 项目数=("项目名称", "nunique"))
        .sort_values(["月份", "处理建议"])
    )
    return grp


def _month_key(month_text: str) -> tuple[int, int]:
    y, m = str(month_text).split("-")
    return int(y), int(m)


def clear_directory(dir_path: Path) -> None:
    """清空目录下所有内容（保留目录本身）。"""
    if not dir_path.exists():
        return
    for item in dir_path.iterdir():
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()


def export_one_file_per_project(decision_df: pd.DataFrame, std_root: Path) -> tuple[Path, int, int]:
    """每个项目仅保留1个文件（最新非明确重复），并保持原文件名复制到输出根目录。"""
    out_dir = std_root
    out_dir.mkdir(parents=True, exist_ok=True)
    for item in out_dir.iterdir():
        if item.name == "去重分析输出" and item.is_dir():
            continue
        if item.is_dir():
            shutil.rmtree(item)
        else:
            item.unlink()

    required_cols = {"月份", "项目名称", "文件路径", "处理建议"}
    if not required_cols.issubset(set(decision_df.columns)):
        raise RuntimeError("判定结果缺少必要列，无法导出每项目单文件")

    candidate = decision_df[decision_df["处理建议"] != "复用上月，当前月不体现"].copy()
    if candidate.empty:
        return out_dir, 0, 0

    priority = {
        "保留": 3,
        "建议复核后复用上月": 2,
        "建议人工复核": 1,
    }
    candidate["_month_key"] = candidate["月份"].map(_month_key)
    candidate["_priority"] = candidate["处理建议"].map(priority).fillna(0)

    selected_rows = []
    for _, g in candidate.groupby("项目名称", sort=False):
        g2 = g.sort_values(["_month_key", "_priority"], ascending=[False, False]).reset_index(drop=True)
        selected_rows.append(g2.iloc[0])
    selected = pd.DataFrame(selected_rows).reset_index(drop=True)

    copied = []
    conflicts = []
    used_names: set[str] = set()

    for _, row in selected.iterrows():
        src = Path(str(row["文件路径"]))
        if not src.is_file():
            continue

        # 保持原文件名不变
        dst = out_dir / src.name
        if src.name in used_names or dst.exists():
            conflicts.append(
                {
                    "项目名称": str(row.get("项目名称", "")),
                    "选中月份": str(row.get("月份", "")),
                    "冲突文件名": src.name,
                    "原文件路径": str(src),
                }
            )
            continue

        shutil.copy2(src, dst)
        used_names.add(src.name)
        copied.append(
            {
                "项目名称": str(row.get("项目名称", "")),
                "选中月份": str(row.get("月份", "")),
                "处理建议": str(row.get("处理建议", "")),
                "对比结论": str(row.get("对比结论", "")),
                "原文件名": src.name,
                "拷贝后文件名": dst.name,
                "原文件路径": str(src),
                "拷贝后路径": str(dst),
            }
        )

    analysis_dir = std_root / "去重分析输出"
    analysis_dir.mkdir(parents=True, exist_ok=True)
    copied_df = pd.DataFrame(copied).sort_values(["项目名称"]).reset_index(drop=True)
    copied_df.to_excel(analysis_dir / "A_拷贝清单.xlsx", index=False)

    if conflicts:
        pd.DataFrame(conflicts).to_excel(analysis_dir / "每项目仅1个_同名冲突清单.xlsx", index=False)

    return out_dir, int(len(copied)), int(len(conflicts))


def run(raw_root: Path, output_file: Path) -> None:
    records = scan_raw_files(raw_root)

    if not records:
        decision_df = pd.DataFrame(
            columns=["月份", "项目名称", "文件名", "文件路径", "含税金额合计", "行数", "文件哈希", "内容哈希", "对比结论", "处理建议", "是否在最新口径体现"]
        )
        summary_df = pd.DataFrame(columns=["月份", "处理建议", "文件数", "项目数"])
        print(f"未扫描到 {raw_root} 下的 Excel 文件，已生成空分析结果")
    else:
        df = pd.DataFrame([r.__dict__ for r in records])
        df = df.sort_values(["项目名称", "月份"]).reset_index(drop=True)

        decision_df = compare_with_previous(df)
        summary_df = build_month_summary(decision_df)

    output_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        decision_df.to_excel(writer, index=False, sheet_name="逐文件判定")
        summary_df.to_excel(writer, index=False, sheet_name="月份汇总")

    export_dir, copied_count, conflict_count = export_one_file_per_project(decision_df, STD_ROOT)

    print(f"分析完成: {output_file}")
    print(f"逐文件判定行数: {len(decision_df)}")
    print(f"月份汇总行数: {len(summary_df)}")
    print(f"每项目1个文件输出目录: {export_dir}")
    print(f"每项目1个文件拷贝成功: {copied_count}")
    print(f"每项目1个文件同名冲突: {conflict_count}")
    print("说明:")
    print("1) 只读取目标目录下的顶层 Excel 文件，不再依赖子文件夹")
    print("2) 分析结果和拷贝清单输出到目标目录下的子文件夹：去重分析输出")
    print("3) 处理建议为 复用上月，当前月不体现 的记录，可在最新月份报表中直接剔除")
    print("4) 目标目录根下将只保留去重后的 Excel 文件，分析文件在子文件夹中")


def main() -> None:
    raw_root = RAW_ROOT
    output_dir = STD_ROOT
    output_dir.mkdir(parents=True, exist_ok=True)
    output_file = output_dir / "去重分析输出" / "采购台账_月度去重分析.xlsx"
    run(raw_root=raw_root, output_file=output_file)


if __name__ == "__main__":
    main()
