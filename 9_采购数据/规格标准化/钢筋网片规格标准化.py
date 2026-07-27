#!/usr/bin/env python3
"""
钢筋网片规格标准化。

目标格式（空格分隔三部分）：
    网片间距 钢筋型号 网片大小
示例：
    200x200 Φ8 3x6m
    150x150 Φ4 3x6m
    200x200 Φ6          # 缺尺寸时仍输出已识别部分
    Φ6 3x6m             # 缺间距时仍输出已识别部分

约定：
- 网片间距：统一为毫米整数，如 200x200（原 20*20 / 20cm*20cm / @20*20 等按厘米理解）
- 钢筋型号：Φ + 直径，如 Φ8、Φ11.8
- 网片大小：统一为米，如 3x6m、3.5x6m（原 3000*6000 按毫米转米）
"""

from __future__ import annotations

import argparse
import math
import re
from pathlib import Path

from openpyxl import load_workbook


MESH_NAME_KEYWORDS = ("钢筋网片", "网片")
# 排除非网片物品
EXCLUDE_NAME_KEYWORDS = ("运费", "加工", "安装")


def format_number(value: float) -> str:
    rounded = round(value)
    if math.isclose(value, rounded, abs_tol=1e-9):
        return str(int(rounded))
    return f"{value:.6f}".rstrip("0").rstrip(".")


def normalize_mesh_text(value: object) -> str:
    """规范化文本，但保留空格作为 token 分隔（避免 20 3m 粘成 203m）。"""
    if value is None:
        return ""
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return ""

    text = text.replace("（", "(").replace("）", ")")
    text = text.replace("，", ",").replace("：", ":")
    text = text.replace("×", "*").replace("＊", "*").replace("Ｘ", "*")
    text = text.replace("x", "*").replace("X", "*")
    text = text.replace("＠", "@")
    text = text.replace("φ", "Φ").replace("Ø", "Φ").replace("ø", "Φ")
    text = text.replace("毫米", "mm").replace("米", "m")
    text = text.replace("Ｍ", "m").replace("ｍ", "m")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def _spacing_mm(a: float, b: float, *, force_cm: bool = False) -> tuple[int, int] | None:
    """将一对间距数字转为 mm。10~40 默认按 cm；50~400 按 mm。"""
    if force_cm or (5 <= a <= 40 and 5 <= b <= 40):
        return int(round(a * 10)), int(round(b * 10))
    if 50 <= a <= 500 and 50 <= b <= 500:
        return int(round(a)), int(round(b))
    return None


def _size_m(a: float, b: float) -> tuple[float, float] | None:
    """将一对尺寸数字转为 m。"""
    if 0.5 <= a <= 12 and 0.5 <= b <= 12:
        return a, b
    if 500 <= a <= 15000 and 500 <= b <= 15000:
        return a / 1000.0, b / 1000.0
    return None


def _is_dia(v: float) -> bool:
    return 2.0 <= v <= 20.0


def parse_mesh_spec(spec_value: object) -> dict:
    """
    解析网片规格，返回:
      {
        spacing: (mm, mm) | None,
        diameter: float | None,
        size: (m, m) | None,
        standard: str | None,
      }
    """
    spaced = normalize_mesh_text(spec_value)
    if not spaced:
        return {"spacing": None, "diameter": None, "size": None, "standard": None}

    work = spaced
    spacing: tuple[int, int] | None = None
    diameter: float | None = None
    size: tuple[float, float] | None = None

    # --- 显式带单位的片段优先提取，并从文本中抠掉 ---

    # cm 间距: 15cm*15cm / 20*20cm
    m = re.search(
        r"(\d+(?:\.\d+)?)\s*cm\s*\*\s*(\d+(?:\.\d+)?)\s*cm",
        work,
        re.IGNORECASE,
    )
    if not m:
        m = re.search(
            r"(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*cm",
            work,
            re.IGNORECASE,
        )
    if m:
        spacing = _spacing_mm(float(m.group(1)), float(m.group(2)), force_cm=True)
        work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # mm 间距: 200*200mm
    if spacing is None:
        m = re.search(
            r"(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)\s*mm\b",
            work,
            re.IGNORECASE,
        )
        if m:
            a, b = float(m.group(1)), float(m.group(2))
            if 50 <= a <= 500 and 50 <= b <= 500:
                spacing = (int(round(a)), int(round(b)))
                work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # @间距: @25*25 / @20*20 （<=40 按 cm）
    if spacing is None:
        m = re.search(r"@\s*(\d+(?:\.\d+)?)\s*\*\s*(\d+(?:\.\d+)?)", work)
        if m:
            spacing = _spacing_mm(float(m.group(1)), float(m.group(2)))
            work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # 板尺寸带 m: 3m*6m / 3.5*6m / 3m*6
    m = re.search(
        r"(\d+(?:\.\d+)?)\s*m?\s*\*\s*(\d+(?:\.\d+)?)\s*m\b",
        work,
        re.IGNORECASE,
    )
    if m:
        cand = _size_m(float(m.group(1)), float(m.group(2)))
        if cand:
            size = cand
            work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # 板尺寸 mm 级: 3000*6000
    if size is None:
        m = re.search(r"(?<!\d)(\d{3,5})\s*\*\s*(\d{3,5})(?!\d)", work)
        if m:
            cand = _size_m(float(m.group(1)), float(m.group(2)))
            if cand and cand[0] >= 0.5:
                size = cand
                work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # 直径 Φd
    m = re.search(r"Φ\s*(\d+(?:\.\d+)?)", work, re.IGNORECASE)
    if m:
        diameter = float(m.group(1))
        work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # 直径 dmm
    if diameter is None:
        m = re.search(r"(\d+(?:\.\d+)?)\s*mm\b", work, re.IGNORECASE)
        if m and _is_dia(float(m.group(1))):
            diameter = float(m.group(1))
            work = (work[: m.start()] + " " + work[m.end() :]).strip()

    # 剩余纯数字序列
    rem = re.sub(r"[^0-9.*]", "*", work)
    rem = re.sub(r"\*+", "*", rem).strip("*")
    nums: list[float] = []
    if rem:
        for part in rem.split("*"):
            if not part:
                continue
            try:
                nums.append(float(part))
            except ValueError:
                continue

    spacing, diameter, size = _assign_remaining_numbers(nums, spacing, diameter, size)

    standard = format_mesh_standard(spacing, diameter, size)
    return {
        "spacing": spacing,
        "diameter": diameter,
        "size": size,
        "standard": standard,
    }


def _assign_remaining_numbers(
    nums: list[float],
    spacing: tuple[int, int] | None,
    diameter: float | None,
    size: tuple[float, float] | None,
) -> tuple[tuple[int, int] | None, float | None, tuple[float, float] | None]:
    """
    对剩余数字按常见网片写法模板分配：
      d * s * s * w * h
      d * s * s
      s * s * d
      d * w * h（无间距，如 4*3000*6000 已在前面处理尺寸）
      s * s
      d
    优先匹配「相等的间距对」。
    """
    if not nums:
        return spacing, diameter, size

    n = len(nums)

    def try_spacing_pair(a: float, b: float) -> tuple[int, int] | None:
        # 相等间距优先；20/20、15/15 等按 cm
        if abs(a - b) < 1e-9:
            return _spacing_mm(a, b)
        return _spacing_mm(a, b)

    # 模板：d,s,s,w,h
    if n == 5 and diameter is None and spacing is None and size is None:
        d, a, b, w, h = nums
        sp = try_spacing_pair(a, b)
        sz = _size_m(w, h)
        if _is_dia(d) and sp and sz:
            return sp, d, sz

    # 模板：d,s,s 或 s,s,d
    if n == 3:
        a, b, c = nums
        if diameter is None and spacing is None:
            # d,s,s（s 相等更可信）
            sp = try_spacing_pair(b, c)
            if _is_dia(a) and sp and abs(b - c) < 1e-9:
                return sp, a, size
            # s,s,d
            sp = try_spacing_pair(a, b)
            if sp and _is_dia(c) and abs(a - b) < 1e-9:
                return sp, c, size
        if diameter is None and size is None and spacing is not None:
            # 已有间距，剩余可能 d,w,h 不完整
            if _is_dia(a) and _size_m(b, c):
                return spacing, a, _size_m(b, c)
        if size is None and diameter is not None and spacing is None:
            sp = try_spacing_pair(a, b)
            sz = _size_m(b, c)  # unlikely
            if sp and abs(a - b) < 1e-9:
                # d already set; a,b spacing, c maybe leftover size half
                leftover_size = size
                if _size_m(c, c) is None and n == 3:
                    # only spacing consumed? keep c if size pair incomplete
                    pass
                return sp, diameter, size

    # 模板：d,s,s,w / 四元组 — 常见 d,s,s 后跟不完整
    if n == 4 and diameter is None and spacing is None:
        d, a, b, c = nums
        sp = try_spacing_pair(a, b)
        if _is_dia(d) and sp and abs(a - b) < 1e-9:
            # 第四个可能是尺寸一边，无法单独成立
            return sp, d, size
        # s,s,w,h
        sp = try_spacing_pair(d, a)
        sz = _size_m(b, c)
        if sp and sz and abs(d - a) < 1e-9:
            return sp, diameter, sz

    # 模板：d,s,s 已在 n==3；n==2
    if n == 2:
        a, b = nums
        if spacing is None:
            sp = try_spacing_pair(a, b)
            if sp and (abs(a - b) < 1e-9 or (a >= 50 and b >= 50)):
                return sp, diameter, size
        if size is None:
            sz = _size_m(a, b)
            if sz:
                return spacing, diameter, sz

    # 通用扫描：优先相等间距对，再直径，再尺寸
    used = [False] * n

    if spacing is None:
        for i in range(n - 1):
            if used[i] or used[i + 1]:
                continue
            a, b = nums[i], nums[i + 1]
            if abs(a - b) < 1e-9:
                sp = try_spacing_pair(a, b)
                if sp:
                    spacing = sp
                    used[i] = used[i + 1] = True
                    break

    if spacing is None:
        for i in range(n - 1):
            if used[i] or used[i + 1]:
                continue
            sp = try_spacing_pair(nums[i], nums[i + 1])
            if sp:
                spacing = sp
                used[i] = used[i + 1] = True
                break

    if size is None:
        for i in range(n - 1):
            if used[i] or used[i + 1]:
                continue
            sz = _size_m(nums[i], nums[i + 1])
            if sz:
                size = sz
                used[i] = used[i + 1] = True
                break

    if diameter is None:
        for i, v in enumerate(nums):
            if used[i]:
                continue
            if _is_dia(v):
                diameter = v
                used[i] = True
                break

    return spacing, diameter, size


def format_mesh_standard(
    spacing: tuple[int, int] | None,
    diameter: float | None,
    size: tuple[float, float] | None,
) -> str | None:
    """按「网片间距 钢筋型号 网片大小」拼接；至少识别出一项才返回。"""
    parts: list[str] = []
    if spacing is not None:
        parts.append(f"{spacing[0]}x{spacing[1]}")
    if diameter is not None:
        parts.append(f"Φ{format_number(diameter)}")
    if size is not None:
        parts.append(f"{format_number(size[0])}x{format_number(size[1])}m")
    return " ".join(parts) if parts else None


def standardize_mesh(spec_value: object) -> str | None:
    return parse_mesh_spec(spec_value)["standard"]


def is_mesh_item(item_name: object) -> bool:
    if item_name is None:
        return False
    name = str(item_name).strip()
    if not name:
        return False
    if any(k in name for k in EXCLUDE_NAME_KEYWORDS):
        return False
    return any(k in name for k in MESH_NAME_KEYWORDS)


def get_header_map(worksheet) -> dict[str, int]:
    return {
        str(cell.value).strip(): index
        for index, cell in enumerate(worksheet[1], start=1)
        if cell.value is not None and str(cell.value).strip()
    }


def process_worksheet(worksheet) -> dict:
    """供统一流水线调用：只改「钢筋网片」类物品的规格列。"""
    header_map = get_header_map(worksheet)
    if "物品名称" not in header_map or "规格" not in header_map:
        return {
            "sheet_name": worksheet.title,
            "matched_rows": 0,
            "updated_rows": 0,
            "prefixed_rows": 0,
            "unresolved_rows": [],
            "skipped": True,
        }

    item_col = header_map["物品名称"]
    spec_col = header_map["规格"]

    matched_rows = 0
    updated_rows = 0
    unresolved_rows: list[tuple] = []

    for row_idx in range(2, worksheet.max_row + 1):
        item_name = worksheet.cell(row_idx, item_col).value
        if not is_mesh_item(item_name):
            continue

        matched_rows += 1
        spec_cell = worksheet.cell(row_idx, spec_col)
        old_text = "" if spec_cell.value is None else str(spec_cell.value).strip()
        standardized = standardize_mesh(spec_cell.value)

        if not standardized:
            unresolved_rows.append(
                (row_idx, str(item_name).strip(), old_text, "mesh")
            )
            continue

        if old_text != standardized:
            spec_cell.value = standardized
            updated_rows += 1

    return {
        "sheet_name": worksheet.title,
        "matched_rows": matched_rows,
        "updated_rows": updated_rows,
        "prefixed_rows": 0,
        "unresolved_rows": unresolved_rows,
        "skipped": False,
    }


def preview_dataframe_specs(specs: list[str]) -> None:
    print(f"{'原规格':40s}  {'标准规格':25s}  间距  型号  大小")
    print("-" * 100)
    for s in specs:
        info = parse_mesh_spec(s)
        print(
            f"{s!r:40s}  {str(info['standard'])!r:25s}  "
            f"{info['spacing']!s:12s}  {info['diameter']!s:6s}  {info['size']!s}"
        )


def main():
    parser = argparse.ArgumentParser(description="钢筋网片规格标准化（间距 型号 大小）")
    parser.add_argument("input_file", nargs="?", default="", help="输入 Excel 路径")
    parser.add_argument(
        "--output-file",
        default="",
        help="输出路径；默认在同目录加后缀 _网片规格标准化",
    )
    parser.add_argument(
        "--preview-only",
        action="store_true",
        help="只预览合并表中钢筋网片规格映射，不写文件",
    )
    parser.add_argument(
        "--sheet",
        default="合并数据",
        help="预览/处理时优先使用的工作表名",
    )
    args = parser.parse_args()

    default_merged = Path(
        "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9_采购台账📦/"
        "E_采购部采购台账_sum/采购部采购台账_merged.xlsx"
    )
    input_file = Path(args.input_file).expanduser() if args.input_file else default_merged
    if not input_file.exists():
        raise FileNotFoundError(f"未找到输入文件: {input_file}")

    if args.preview_only:
        import pandas as pd

        df = pd.read_excel(input_file, sheet_name=0)
        mask = df["物品名称"].astype(str).str.strip() == "钢筋网片"
        specs = (
            df.loc[mask, "规格"]
            .fillna("(空)")
            .astype(str)
            .str.strip()
            .value_counts()
            .index.tolist()
        )
        print(f"文件: {input_file}")
        print(f"钢筋网片唯一规格数: {len(specs)}\n")
        preview_dataframe_specs(specs)
        return

    output_file = (
        Path(args.output_file).expanduser().resolve()
        if args.output_file
        else input_file.with_name(f"{input_file.stem}_网片规格标准化{input_file.suffix}")
    )

    print(f"输入: {input_file}")
    print(f"输出: {output_file}")

    wb = load_workbook(input_file)
    total_matched = total_updated = total_unresolved = 0

    for ws in wb.worksheets:
        # 汇总簿里非明细 sheet 可跳过无列的
        result = process_worksheet(ws)
        if result.get("skipped"):
            print(f"  [{ws.title}] 跳过（无物品名称/规格列）")
            continue
        total_matched += result["matched_rows"]
        total_updated += result["updated_rows"]
        total_unresolved += len(result["unresolved_rows"])
        print(
            f"  [{ws.title}] 命中 {result['matched_rows']}，"
            f"更新 {result['updated_rows']}，未识别 {len(result['unresolved_rows'])}"
        )
        for row_idx, name, spec, _ in result["unresolved_rows"][:10]:
            print(f"    未识别 行{row_idx}: {name} | {spec}")

    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
    print(f"\n完成: 命中 {total_matched}，更新 {total_updated}，未识别 {total_unresolved}")
    print(f"已保存: {output_file}")


if __name__ == "__main__":
    main()
