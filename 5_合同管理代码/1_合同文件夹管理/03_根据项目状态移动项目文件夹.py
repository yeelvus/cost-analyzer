#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""根据“项目基本信息表.xlsx”的项目状态，同步合同管理中的项目文件夹位置。

目录结构：合同类型 / 项目状态 / 项目文件夹

默认仅预览，不移动任何文件。执行移动必须明确加 --apply，并再次确认。
目标目录已经存在时始终跳过，绝不覆盖、合并或删除文件。
"""
from __future__ import annotations

import argparse
import json
import os
import re
import shutil
from collections import Counter
from datetime import datetime
from pathlib import Path

import pandas as pd


DEFAULT_EXCEL = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx")
DEFAULT_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑")
CODE_COLUMN = "项目编码"
STATUS_COLUMN = "项目状态"
CONTRACT_TYPE_KEYWORDS = ("总包合同", "分包合同", "总包签证", "分包签证")
LOG_FOLDER = "日志"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="按项目状态移动合同管理项目文件夹（默认仅预览）")
    parser.add_argument("--excel", default=str(DEFAULT_EXCEL), help="项目基本信息表.xlsx")
    parser.add_argument("--root", default=str(DEFAULT_ROOT), help="合同管理根目录")
    parser.add_argument("--apply", action="store_true", help="执行移动；省略时仅生成预览")
    parser.add_argument("--yes", action="store_true", help="与 --apply 一起使用时跳过二次确认")
    parser.add_argument("--no-log", action="store_true", help="不写入 JSON 日志（仅用于只读预览/测试）")
    return parser.parse_args()


def clean_text(value: object) -> str:
    if value is None or pd.isna(value):
        return ""
    return str(value).strip()


def clean_status(value: object) -> str:
    text = clean_text(value)
    # 与现有建目录工具保持一致：状态中空格不参与目录匹配。
    return re.sub(r"\s+", "", text)


def load_status_map(excel_path: Path) -> tuple[dict[str, str], list[str]]:
    if not excel_path.is_file():
        raise FileNotFoundError(f"项目基本信息表不存在：{excel_path}")
    df = pd.read_excel(excel_path)
    missing = [name for name in (CODE_COLUMN, STATUS_COLUMN) if name not in df.columns]
    if missing:
        raise ValueError(f"Excel 缺少必要列：{missing}")

    result: dict[str, str] = {}
    duplicates: list[str] = []
    for _, row in df.iterrows():
        code = clean_text(row[CODE_COLUMN])
        status = clean_status(row[STATUS_COLUMN])
        if not code or not status:
            continue
        if code in result and result[code] != status:
            duplicates.append(code)
            continue
        result[code] = status
    return result, sorted(set(duplicates))


def find_contract_type_dirs(root: Path) -> list[Path]:
    result: list[Path] = []
    for item in sorted(root.iterdir(), key=lambda path: path.name):
        if not item.is_dir() or item.name.startswith("."):
            continue
        if any(keyword in item.name for keyword in CONTRACT_TYPE_KEYWORDS):
            result.append(item)
    return result


def project_code(folder_name: str) -> str:
    """项目目录以“项目编码_项目名称”为准；无下划线的目录不参与自动移动。"""
    if "_" not in folder_name:
        return ""
    return folder_name.split("_", 1)[0].strip()


def build_plan(root: Path, status_map: dict[str, str]) -> tuple[list[dict], list[dict]]:
    moves: list[dict] = []
    skipped: list[dict] = []
    for contract_dir in find_contract_type_dirs(root):
        for current_status_dir in sorted(contract_dir.iterdir(), key=lambda path: path.name):
            if not current_status_dir.is_dir() or current_status_dir.name.startswith("."):
                continue
            # 项目目录只能位于“合同类型/状态”这一级，避免递归移动合同子文件夹。
            for project_dir in sorted(current_status_dir.iterdir(), key=lambda path: path.name):
                if not project_dir.is_dir() or project_dir.name.startswith("."):
                    continue
                code = project_code(project_dir.name)
                if not code:
                    skipped.append({"reason": "目录名不含项目编码前缀", "source": str(project_dir)})
                    continue
                desired_status = status_map.get(code)
                if not desired_status:
                    skipped.append({"reason": "Excel 未找到项目编码或项目状态为空", "source": str(project_dir), "code": code})
                    continue
                if current_status_dir.name == desired_status:
                    continue
                target = contract_dir / desired_status / project_dir.name
                if target.exists():
                    skipped.append({
                        "reason": "目标项目文件夹已存在（不合并）",
                        "source": str(project_dir), "target": str(target), "code": code,
                    })
                    continue
                moves.append({
                    "code": code,
                    "contract_type_dir": contract_dir.name,
                    "from_status": current_status_dir.name,
                    "to_status": desired_status,
                    "source": str(project_dir),
                    "target": str(target),
                })
    return moves, skipped


def print_preview(moves: list[dict], skipped: list[dict]) -> None:
    print("\n" + "=" * 88)
    print(f"待移动项目文件夹：{len(moves)} 个；跳过/待处理：{len(skipped)} 个")
    print("=" * 88)
    for index, item in enumerate(moves, start=1):
        print(f"{index:>3}. [{item['contract_type_dir']}] {item['code']}")
        print(f"     {item['from_status']}  →  {item['to_status']}")
        print(f"     {item['source']}")
        print(f"     → {item['target']}")
    if skipped:
        print("\n跳过/待处理项目：")
        for item in skipped:
            print(f" - {item['reason']}：{item['source']}")


def execute_plan(moves: list[dict]) -> tuple[list[dict], list[dict]]:
    moved: list[dict] = []
    failed: list[dict] = []
    for item in moves:
        source = Path(item["source"])
        target = Path(item["target"])
        try:
            if not source.is_dir():
                raise FileNotFoundError("源项目文件夹已不存在")
            if target.exists():
                raise FileExistsError("目标项目文件夹已存在")
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(source), str(target))
            moved.append(item)
            print(f"✅ 已移动：{item['code']} · {item['from_status']} → {item['to_status']}")
        except Exception as exc:
            failed.append({**item, "error": str(exc)})
            print(f"❌ 移动失败：{item['source']}（{exc}）")
    return moved, failed


def save_log(root: Path, *, excel_path: Path, moves: list[dict], skipped: list[dict], moved: list[dict], failed: list[dict], dry_run: bool, duplicates: list[str]) -> Path:
    log_dir = root / LOG_FOLDER
    log_dir.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"项目状态同步_{'预览' if dry_run else '执行'}_{stamp}.json"
    payload = {
        "timestamp": datetime.now().isoformat(timespec="seconds"),
        "operation": "按项目基本信息表的项目状态同步合同项目文件夹",
        "dry_run": dry_run,
        "excel_path": str(excel_path),
        "planned_move_count": len(moves),
        "moved_count": len(moved),
        "skipped_count": len(skipped),
        "failed_count": len(failed),
        "duplicate_status_codes": duplicates,
        "planned_moves": moves,
        "skipped": skipped,
        "moved": moved,
        "failed": failed,
    }
    log_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return log_path


def main() -> None:
    args = parse_args()
    root = Path(args.root).expanduser().resolve()
    excel_path = Path(args.excel).expanduser().resolve()
    if not root.is_dir():
        raise SystemExit(f"合同管理根目录不存在：{root}")
    status_map, duplicates = load_status_map(excel_path)
    moves, skipped = build_plan(root, status_map)
    print(f"Excel 有效项目状态：{len(status_map)} 个")
    if duplicates:
        print(f"⚠️ 有 {len(duplicates)} 个项目编码存在互相冲突的状态，已跳过这些重复记录。")
    print_preview(moves, skipped)

    if not args.apply:
        log = None if args.no_log else save_log(root, excel_path=excel_path, moves=moves, skipped=skipped, moved=[], failed=[], dry_run=True, duplicates=duplicates)
        print(f"\n🧪 当前为预览模式，未移动任何文件夹。" + (f"日志：{log}" if log else ""))
        print("如确认无误，请执行：python 本脚本.py --apply")
        return
    if not args.yes:
        answer = input("\n确认执行以上移动？目标存在的项目将跳过，不会合并。(y/n): ").strip().lower()
        if answer not in {"y", "yes", "是"}:
            print("已取消，未移动任何文件夹。")
            return
    moved, failed = execute_plan(moves)
    log = None if args.no_log else save_log(root, excel_path=excel_path, moves=moves, skipped=skipped, moved=moved, failed=failed, dry_run=False, duplicates=duplicates)
    print(f"\n完成：成功移动 {len(moved)} 个，失败 {len(failed)} 个，跳过 {len(skipped)} 个。")
    if log:
        print(f"日志：{log}")


if __name__ == "__main__":
    main()
