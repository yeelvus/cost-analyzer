#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""在每个项目文件夹下创建空的项目编码.json（已存在则跳过，不覆盖）。

文件夹名：THCC-2024-002_北新建材……
生成文件：THCC-2024-002.json
仅自动写入项目编码、项目名称，其余字段留空自行填写。
"""
from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

from config import DEFAULT_SETTLEMENT_ROOT, JSON_NAME_TMPL, PROJECT_CODE_RE
from schema import JSON_FIELDS, empty_record, ordered_record

CODE_RE = re.compile(PROJECT_CODE_RE)


def parse_folder_name(name: str) -> tuple[str, str]:
    if "_" in name:
        code, title = name.split("_", 1)
        return code.strip(), title.strip()
    return name.strip(), ""


def main() -> int:
    parser = argparse.ArgumentParser(description="为项目文件夹创建空 JSON")
    parser.add_argument("--root", default=str(DEFAULT_SETTLEMENT_ROOT))
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    root = Path(args.root)
    if not root.exists():
        print(f"目录不存在（网络盘未挂载？）: {root}")
        return 1

    created = skipped = 0
    for folder in sorted(p for p in root.iterdir() if p.is_dir() and not p.name.startswith(".")):
        code, name = parse_folder_name(folder.name)
        if not CODE_RE.match(code):
            continue
        json_path = folder / JSON_NAME_TMPL.format(code=code)
        if json_path.exists():
            try:
                with json_path.open("r", encoding="utf-8") as f:
                    data = json.load(f)
            except Exception:
                data = {}
            if not isinstance(data, dict):
                data = {}
            blank = empty_record(code, name)
            missing = [k for k in JSON_FIELDS if k not in data]
            if missing:
                for k in missing:
                    data[k] = blank.get(k, "")
                print(f"[补字段] {code}  增加 {', '.join(missing)}")
                created += 1
                if not args.dry_run:
                    with json_path.open("w", encoding="utf-8") as f:
                        json.dump(ordered_record(data), f, ensure_ascii=False, indent=2)
                        f.write("\n")
            else:
                print(f"[跳过] {code}  已有 {json_path.name}")
                skipped += 1
            continue
        rec = empty_record(code, name)
        print(f"[创建] {code}  {name}")
        created += 1
        if not args.dry_run:
            with json_path.open("w", encoding="utf-8") as f:
                json.dump(ordered_record(rec), f, ensure_ascii=False, indent=2)
                f.write("\n")

    print(f"完成 创建={created} 跳过={skipped}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
