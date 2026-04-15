from __future__ import annotations

from pathlib import Path
import re


ROOT_DIR = Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_分包对账与付款/A_分包对账与付款_原始数据")
TARGET_SUFFIX = "分包管控表.xlsx"
NAME_PATTERN = re.compile(r"^(?P<prefix>.+?)分包管控表\.xlsx$")


def rename_files(root_dir: Path) -> None:
    if not root_dir.exists() or not root_dir.is_dir():
        print(f"目录不存在或不是文件夹: {root_dir}")
        return

    matched_files = [
        p for p in root_dir.rglob(f"*{TARGET_SUFFIX}") if p.is_file() and NAME_PATTERN.match(p.name)
    ]
    if not matched_files:
        print("未找到需要重命名的文件。")
        return

    renamed_count = 0
    skipped_count = 0

    for old_path in matched_files:
        match = NAME_PATTERN.match(old_path.name)
        if not match:
            skipped_count += 1
            continue

        prefix = match.group("prefix").rstrip("_ ")
        if not prefix:
            skipped_count += 1
            continue

        new_name = f"{prefix}_分包管控表.xlsx"
        new_path = old_path.with_name(new_name)

        if new_path == old_path:
            print(f"跳过（已是目标格式）: {old_path}")
            skipped_count += 1
            continue

        if new_path.exists():
            print(f"跳过（目标文件已存在）: {old_path} -> {new_path}")
            skipped_count += 1
            continue

        old_path.rename(new_path)
        print(f"已重命名: {old_path} -> {new_path}")
        renamed_count += 1

    print("\n处理完成")
    print(f"匹配到文件数: {len(matched_files)}")
    print(f"成功重命名: {renamed_count}")
    print(f"跳过数量: {skipped_count}")


if __name__ == "__main__":
    rename_files(ROOT_DIR)
