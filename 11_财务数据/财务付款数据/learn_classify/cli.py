# -*- coding: utf-8 -*-
"""
learn_classify 兼容入口（薄封装）。

推荐统一使用：
  python 财务付款一体化.py

本模块保留 `python -m learn_classify …` 方便脚本调用，命令会转发到 payment_hub。
"""

from __future__ import annotations

import sys
from pathlib import Path

# 保证仓库根在 path 上
_ROOT = Path(__file__).resolve().parent.parent
if str(_ROOT) not in sys.path:
    sys.path.insert(0, str(_ROOT))


# 旧 learn_classify 子命令 → 一体化 argv
_MAP = {
    "export-review": "review",
    "compact": "compact-memory",
    "seed-vendor": None,  # 仅 learn_classify 独有，内联处理
}


def main(argv=None) -> int:
    argv = list(sys.argv[1:] if argv is None else argv)

    # 抽出 --data-dir 供转发
    data_dir = None
    rest: list[str] = []
    i = 0
    while i < len(argv):
        if argv[i] == "--data-dir" and i + 1 < len(argv):
            data_dir = argv[i + 1]
            i += 2
            continue
        if argv[i].startswith("--data-dir="):
            data_dir = argv[i].split("=", 1)[1]
            i += 1
            continue
        rest.append(argv[i])
        i += 1

    if not rest:
        print(
            "用法: python -m learn_classify <命令>\n"
            "推荐: python 财务付款一体化.py  （交互菜单）\n"
            "兼容命令: classify / reclassify / compact / health / learn / …"
        )
        return 0

    cmd = rest[0]
    # seed-vendor 仅在此包保留
    if cmd == "seed-vendor":
        return _cmd_seed_vendor(rest[1:], data_dir)

    mapped = _MAP.get(cmd, cmd)
    hub_argv = [mapped] + rest[1:]
    if data_dir:
        hub_argv = ["--data-dir", data_dir] + hub_argv

    from payment_hub.cli import main as hub_main

    return hub_main(hub_argv)


def _cmd_seed_vendor(argv: list[str], data_dir: str | None) -> int:
    import argparse

    from .engine import ClassifyEngine
    from .learn import rebuild_vendor_stats_from_std
    from .store import VendorStats

    p = argparse.ArgumentParser(prog="seed-vendor")
    p.add_argument("--std-dir", required=True)
    p.add_argument("--conf", type=float, default=0.75)
    args = p.parse_args(argv)
    data = Path(data_dir) if data_dir else Path(__file__).resolve().parent / "data"
    n = rebuild_vendor_stats_from_std(
        Path(args.std_dir),
        VendorStats(data / "vendor_stats.json"),
        ClassifyEngine.from_data_dir(data),
        min_conf=args.conf,
    )
    print(f"✅ 供应商画像观察次数: {n}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
