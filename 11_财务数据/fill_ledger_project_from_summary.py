#!/usr/bin/env python3
"""
从明细分类账「摘要」中提取「-」之前的内容，在可识别为项目名称时填入「项目名称」列。

规则：
1. 仅处理「项目名称」为空的行（不覆盖已有值）
2. 取摘要中第一个连接符（- / － / — / –）之前的文本
3. 识别方式（满足其一即可）：
   - 命中项目名称匹配库
   - 与本文件中已有「项目名称」相同或去「项目」后缀后相同
   - 通过启发式判断为合理项目简称（短、含中英文、非日期/业务短语）
4. 写入值优先用摘要前缀的精简形式（去掉末尾「项目」以贴近台账既有写法），
   若匹配库有更合适的短名且与前缀高度相关则仍写摘要前缀（保持原始识别结果）
"""
from __future__ import annotations

import argparse
import re
import sys
from collections import Counter
from pathlib import Path

import pandas as pd

CODE_ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(CODE_ROOT))

from common.project import load_or_init_db, match_project  # noqa: E402

DEFAULT_INPUT = Path(
    "/Users/wyattsun/Desktop/财务数据/CNY这是现在用的金蝶里边的数据从23.07-26.07月！"
    "/明细分类账-泰国子公司; 开始期间-主账簿-2024年1期至2026年7期.xlsx"
)
DEFAULT_DB = CODE_ROOT / "0_项目基本信息" / "project_name_match_db.json"

DASH_CHARS = ("－", "—", "–", "﹣", "−")
NOISE_PREFIX_RE = re.compile(
    r"^(内部|银行|订正|转账|未知|资金|手续费|利息|税费|工资|社保)"
)
DATE_PREFIX_RE = re.compile(r"^\d{1,2}([./月]\d{1,2})")
# 摘要把业务动作写进了前缀（未正确用 - 分隔）
ACTION_IN_PREFIX_RE = re.compile(
    r"(支付|采购|工资|保证金|退回|外售|客户名|合同款|投标|退押金|存入|调拨|转账手续费)"
)
CJK_OR_LETTER_RE = re.compile(r"[\u4e00-\u9fffA-Za-z]")


def is_empty_project(val) -> bool:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return True
    text = str(val).strip()
    return (not text) or text.lower() == "nan"


def extract_summary_prefix(summary) -> str:
    """提取摘要中第一个 '-' 之前的内容。"""
    if summary is None or (isinstance(summary, float) and pd.isna(summary)):
        return ""
    text = str(summary).strip()
    if not text:
        return ""
    for ch in DASH_CHARS:
        text = text.replace(ch, "-")
    if "-" not in text:
        return ""
    return text.split("-", 1)[0].strip()


def strip_project_suffix(name: str) -> str:
    name = name.strip()
    if name.endswith("项目") and len(name) > 2:
        return name[:-2].strip()
    return name


def is_plausible_project_name(prefix: str) -> bool:
    """启发式：像项目简称，而非日期/长业务描述。"""
    if not prefix:
        return False
    # 过短/过长不像项目名
    if len(prefix) < 2 or len(prefix) > 16:
        return False
    if DATE_PREFIX_RE.match(prefix):
        return False
    if NOISE_PREFIX_RE.search(prefix):
        return False
    if ACTION_IN_PREFIX_RE.search(prefix):
        return False
    if not CJK_OR_LETTER_RE.search(prefix):
        return False
    # 含过多数字/括号的描述
    if prefix.count("(") + prefix.count("（") >= 1 and len(prefix) > 10:
        return False
    digit_ratio = sum(ch.isdigit() for ch in prefix) / max(len(prefix), 1)
    if digit_ratio > 0.4:
        return False
    return True


def build_known_names(series: pd.Series) -> set[str]:
    names: set[str] = set()
    for val in series.dropna().astype(str):
        text = val.strip()
        if not text or text.lower() == "nan":
            continue
        names.add(text)
        names.add(strip_project_suffix(text))
        if not text.endswith("项目"):
            names.add(text + "项目")
    return names


def recognize_project_name(
    prefix: str,
    *,
    db: dict,
    known_names: set[str],
) -> str | None:
    """
    识别摘要前缀是否为项目名称。
    返回应写入「项目名称」的字符串；无法识别返回 None。
    """
    if not prefix:
        return None

    candidates = [prefix]
    stripped = strip_project_suffix(prefix)
    if stripped != prefix:
        candidates.append(stripped)

    # 1) 本文件已有项目名称
    for c in candidates:
        if c in known_names:
            # 写入时与文件既有风格对齐：优先无「项目」后缀的短名（若库中有）
            if stripped in known_names:
                return stripped
            return c

    # 2) 项目匹配库
    for c in candidates:
        hit = match_project(c, db, fuzzy=True)
        if hit and hit.get("project_name"):
            # 台账项目名称列习惯用摘要简称，保留识别到的前缀短名
            return stripped if is_plausible_project_name(stripped) else c

    # 3) 启发式
    if is_plausible_project_name(prefix):
        return stripped if is_plausible_project_name(stripped) else prefix
    if stripped != prefix and is_plausible_project_name(stripped):
        return stripped

    return None


def process_file(
    input_path: Path,
    output_path: Path | None,
    db_path: Path,
    *,
    overwrite: bool = False,
) -> Path:
    input_path = input_path.expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"文件不存在: {input_path}")

    if output_path is None:
        if overwrite:
            output_path = input_path
        else:
            output_path = input_path.with_name(
                f"{input_path.stem}_项目名称已补全{input_path.suffix}"
            )
    else:
        output_path = output_path.expanduser().resolve()

    print(f"读取: {input_path}")
    df = pd.read_excel(input_path)
    if "摘要" not in df.columns or "项目名称" not in df.columns:
        raise SystemExit(
            f"缺少必要列。当前列: {list(df.columns)}"
        )

    db = load_or_init_db(str(db_path))
    known_names = build_known_names(df["项目名称"])

    empty_mask = df["项目名称"].map(is_empty_project)
    empty_count = int(empty_mask.sum())
    print(f"项目名称为空: {empty_count} 行 / 共 {len(df)} 行")

    filled = 0
    no_dash = 0
    unrecognized = 0
    fill_counter: Counter[str] = Counter()
    unrec_samples: Counter[str] = Counter()

    df = df.copy()
    for idx in df.index[empty_mask]:
        summary = df.at[idx, "摘要"]
        prefix = extract_summary_prefix(summary)
        if not prefix:
            no_dash += 1
            continue
        name = recognize_project_name(prefix, db=db, known_names=known_names)
        if not name:
            unrecognized += 1
            unrec_samples[prefix] += 1
            continue
        df.at[idx, "项目名称"] = name
        filled += 1
        fill_counter[name] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)
    print(f"\n已保存: {output_path}")
    print(f"补全成功: {filled} 行")
    print(f"摘要无「-」前缀: {no_dash} 行（仍为空）")
    print(f"有前缀但未识别: {unrecognized} 行（仍为空）")

    if fill_counter:
        print("\n补全项目名称 Top:")
        for name, cnt in fill_counter.most_common(25):
            print(f"  {cnt:4d}  {name}")

    if unrec_samples:
        print("\n未识别前缀 Top（供核对）:")
        for name, cnt in unrec_samples.most_common(20):
            print(f"  {cnt:4d}  {name}")

    still_empty = int(df["项目名称"].map(is_empty_project).sum())
    print(f"\n处理后项目名称仍为空: {still_empty} 行")
    return output_path


def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(description="从摘要「-」前补全明细分类账项目名称")
    p.add_argument("input", nargs="?", default=str(DEFAULT_INPUT), help="明细分类账 xlsx")
    p.add_argument("-o", "--output", default=None, help="输出路径（默认同目录 *_项目名称已补全.xlsx）")
    p.add_argument("--db", default=str(DEFAULT_DB), help="项目名称匹配库 json")
    p.add_argument(
        "--overwrite",
        action="store_true",
        help="直接覆盖原文件（默认写新文件）",
    )
    return p.parse_args()


def main() -> None:
    args = parse_args()
    out = Path(args.output) if args.output else None
    process_file(
        Path(args.input),
        out,
        Path(args.db),
        overwrite=args.overwrite,
    )


if __name__ == "__main__":
    main()
