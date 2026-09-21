#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""根据“项目基本信息表.xlsx”的项目状态，同步合同管理中的项目文件夹位置。

目录结构：合同类型 / 项目状态 / 项目文件夹

默认仅预览，不移动任何文件。执行必须明确加 --apply，并再次确认。

移动是 os.rename / shutil.move，成功后若源目录仍在（iCloud 复制残留），会删除源目录。
目标已存在时不再整夹跳过：
  - 源为空 → 删除旧文件夹
  - 目标为空 → 删空目标后再移动
  - 两边都有文件 → 把源里目标没有的文件拷入目标（不覆盖），再删除旧文件夹
"""
from __future__ import annotations

import argparse
import json
import re
import shutil
from datetime import datetime
from pathlib import Path

import pandas as pd


DEFAULT_EXCEL = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx")
DEFAULT_ROOT = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理📑")
CODE_COLUMN = "项目编码"
STATUS_COLUMN = "项目状态"
CONTRACT_TYPE_KEYWORDS = ("总包合同", "分包合同", "总包签证", "分包签证")
LOG_FOLDER = "日志"
IGNORE_NAMES = {".DS_Store", "Thumbs.db", ".localized"}


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


def iter_real_files(folder: Path):
    for path in folder.rglob("*"):
        if path.is_file() and path.name not in IGNORE_NAMES:
            yield path


def is_effectively_empty(folder: Path) -> bool:
    if not folder.is_dir():
        return True
    return next(iter_real_files(folder), None) is None


def remove_tree(path: Path) -> None:
    if path.is_dir():
        shutil.rmtree(path)
    elif path.exists():
        path.unlink()


def robust_move(source: Path, target: Path) -> None:
    """移动目录；若底层变成复制（iCloud / 跨卷），结束后删掉源目录。"""
    target.parent.mkdir(parents=True, exist_ok=True)
    if target.exists():
        raise FileExistsError("目标项目文件夹已存在")
    shutil.move(str(source), str(target))
    if source.exists():
        remove_tree(source)


def merge_missing_files(source: Path, target: Path) -> dict:
    """把源中目标没有的文件拷入目标，不覆盖已有文件。"""
    copied = 0
    skipped_existing = 0
    conflicts = []
    for src_file in iter_real_files(source):
        rel = src_file.relative_to(source)
        dest_file = target / rel
        if dest_file.exists():
            skipped_existing += 1
            if dest_file.stat().st_size != src_file.stat().st_size:
                conflicts.append({"rel": str(rel), "source": str(src_file), "target": str(dest_file)})
            continue
        dest_file.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src_file, dest_file)
        copied += 1
    return {"copied": copied, "skipped_existing": skipped_existing, "conflicts": conflicts}


def classify_action(source: Path, target: Path) -> str:
    source_empty = is_effectively_empty(source)
    target_exists = target.exists()
    target_empty = is_effectively_empty(target) if target_exists else True
    if not target_exists:
        return "move"
    if source_empty:
        return "delete_empty_source"
    if target_empty:
        return "replace_empty_target"
    return "merge_then_delete_source"


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
                action = classify_action(project_dir, target)
                moves.append({
                    "code": code,
                    "contract_type_dir": contract_dir.name,
                    "from_status": current_status_dir.name,
                    "to_status": desired_status,
                    "source": str(project_dir),
                    "target": str(target),
                    "action": action,
                })
    return moves, skipped


ACTION_LABELS = {
    "move": "移动（完成后删除源文件夹）",
    "delete_empty_source": "旧文件夹为空，删除",
    "replace_empty_target": "目标为空：先删空目标再移动，并删除源",
    "merge_then_delete_source": "合并到已有目标（不覆盖同名文件），再删除旧文件夹",
}


def print_preview(moves: list[dict], skipped: list[dict]) -> None:
    print("\n" + "=" * 88)
    print(f"待处理项目文件夹：{len(moves)} 个；跳过/待处理：{len(skipped)} 个")
    counts: dict[str, int] = {}
    for item in moves:
        counts[item["action"]] = counts.get(item["action"], 0) + 1
    if counts:
        print("其中：" + "；".join(f"{ACTION_LABELS[key]} {counts[key]} 个" for key in ACTION_LABELS if key in counts))
    print("=" * 88)
    for index, item in enumerate(moves, start=1):
        print(f"{index:>3}. [{item['contract_type_dir']}] {item['code']}  ·  {ACTION_LABELS[item['action']]}")
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
        action = item["action"]
        try:
            if not source.is_dir():
                raise FileNotFoundError("源项目文件夹已不存在")
            extra = {}
            if action == "move":
                robust_move(source, target)
            elif action == "delete_empty_source":
                if not is_effectively_empty(source):
                    raise RuntimeError("源文件夹已不再为空，已中止删除")
                remove_tree(source)
            elif action == "replace_empty_target":
                if target.exists() and not is_effectively_empty(target):
                    raise RuntimeError("目标已有文件，已中止覆盖")
                if target.exists():
                    remove_tree(target)
                robust_move(source, target)
            elif action == "merge_then_delete_source":
                if not target.exists():
                    robust_move(source, target)
                else:
                    extra = merge_missing_files(source, target)
                    remove_tree(source)
            else:
                raise RuntimeError(f"未知动作：{action}")
            if source.exists():
                raise RuntimeError(f"源文件夹仍在，未完成去重：{source}")
            record = {**item, **extra}
            moved.append(record)
            print(f"✅ {ACTION_LABELS[action]}：{item['code']} · {item['from_status']} → {item['to_status']}")
            if extra.get("copied"):
                print(f"     补入目标 {extra['copied']} 个文件；已有同名 {extra.get('skipped_existing', 0)} 个")
            if extra.get("conflicts"):
                print(f"     ⚠️ {len(extra['conflicts'])} 个同名文件大小不同，已保留目标中的版本")
        except Exception as exc:
            failed.append({**item, "error": str(exc)})
            print(f"❌ 失败：{item['source']}（{exc}）")
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
        print(f"\n🧪 当前为预览模式，未移动或删除任何文件夹。" + (f"日志：{log}" if log else ""))
        print("如确认无误，请执行：python 本脚本.py --apply")
        return
    if not args.yes:
        answer = input("\n确认执行以上操作？旧状态目录在移动/合并后会被删除。(y/n): ").strip().lower()
        if answer not in {"y", "yes", "是"}:
            print("已取消，未移动任何文件夹。")
            return
    moved, failed = execute_plan(moves)
    log = None if args.no_log else save_log(root, excel_path=excel_path, moves=moves, skipped=skipped, moved=moved, failed=failed, dry_run=False, duplicates=duplicates)
    print(f"\n完成：成功 {len(moved)} 个，失败 {len(failed)} 个，跳过 {len(skipped)} 个。")
    if log:
        print(f"日志：{log}")


if __name__ == "__main__":
    main()
