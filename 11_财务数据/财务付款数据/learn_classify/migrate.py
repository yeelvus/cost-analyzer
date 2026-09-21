# -*- coding: utf-8 -*-
"""
分类体系迁移 / 快速转移。

迁移文件示例 migrations/20260717_split_other.json:
{
  "from_version": "1.0.0",
  "to_version": "1.1.0",
  "description": "拆分其他→检验试验/危废等",
  "id_map": {
    "L2_mgmt_other": "L2_misc_other"
  },
  "rename": {
    "L2_mat_generic": "辅材杂项"
  },
  "reparent": {},
  "add_level1": [],
  "add_level2": [
    {
      "id": "L2_test_fee",
      "name": "检验试验费",
      "parent": "L1_mgmt",
      "order": 50
    }
  ],
  "deactivate": [],
  "pending_l2_id": "L2_pending"
}

流程：
1. 备份 data/
2. 按 migration 更新 taxonomy.json
3. 按 id_map 重写 memory / corrections / vendor_stats / rules 中的 l2_id
4. 可选：重跑 std 显示名
"""

from __future__ import annotations

import json
import shutil
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

from .engine import ClassifyEngine
from .store import CorrectionStore, MemoryStore, VendorStats, load_rules, save_rules
from .taxonomy import Taxonomy, apply_migration_to_taxonomy_file


def load_migration(path: Path) -> Dict[str, Any]:
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def backup_data_dir(data_dir: Path) -> Path:
    data_dir = Path(data_dir)
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    bak = data_dir.parent / f"data_backup_{stamp}"
    shutil.copytree(data_dir, bak)
    return bak


def apply_id_map_to_rules(rules: list, id_map: Dict[str, str]) -> int:
    changed = 0
    for r in rules:
        old = str(r.get("l2_id") or r.get("target") or "")
        if old in id_map:
            r["l2_id"] = id_map[old]
            if "target" in r:
                r["target"] = id_map[old]
            changed += 1
    return changed


def apply_id_map_to_subject_map(path: Path, id_map: Dict[str, str]) -> int:
    if not path.exists():
        return 0
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)
    changed = 0
    for k, v in list(data.items()):
        if str(v) in id_map:
            data[k] = id_map[str(v)]
            changed += 1
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return changed


def run_migration(
    data_dir: Path,
    migration_path: Path,
    *,
    do_backup: bool = True,
    reclassify_std: Optional[Path] = None,
) -> Dict[str, Any]:
    data_dir = Path(data_dir)
    migration = load_migration(migration_path)
    report: Dict[str, Any] = {"migration": str(migration_path)}

    if do_backup:
        bak = backup_data_dir(data_dir)
        report["backup"] = str(bak)

    tax_path = data_dir / "taxonomy.json"
    taxonomy = Taxonomy.load(tax_path)
    report["from_version"] = taxonomy.version

    # optional version guard
    if migration.get("from_version") and migration["from_version"] != taxonomy.version:
        report["warning"] = (
            f"from_version 不匹配: migration={migration['from_version']} "
            f"taxonomy={taxonomy.version}（仍继续执行）"
        )

    new_tax = apply_migration_to_taxonomy_file(taxonomy, migration)
    new_tax.save(tax_path)
    report["to_version"] = new_tax.version

    id_map = {str(k): str(v) for k, v in (migration.get("id_map") or {}).items()}
    report["id_map_size"] = len(id_map)

    mem = MemoryStore(data_dir / "memory.jsonl")
    report["memory_changed"] = mem.apply_id_map(id_map) if id_map else 0

    corr = CorrectionStore(data_dir / "corrections.jsonl")
    report["corrections_changed"] = corr.apply_id_map(id_map) if id_map else 0

    vendor = VendorStats(data_dir / "vendor_stats.json")
    report["vendor_changed"] = vendor.apply_id_map(id_map) if id_map else 0

    rules = load_rules(data_dir / "rules.json")
    rc = apply_id_map_to_rules(rules, id_map) if id_map else 0
    if rc or migration.get("to_version"):
        save_rules(data_dir / "rules.json", rules, meta={"version": new_tax.version})
    report["rules_changed"] = rc

    report["subject_map_changed"] = apply_id_map_to_subject_map(
        data_dir / "subject_map.json", id_map
    )

    # 保存迁移记录
    hist_dir = data_dir / "migrations_applied"
    hist_dir.mkdir(parents=True, exist_ok=True)
    hist_file = hist_dir / f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{Path(migration_path).name}"
    with open(hist_file, "w", encoding="utf-8") as f:
        json.dump({"migration": migration, "report": report}, f, ensure_ascii=False, indent=2)
    report["history_file"] = str(hist_file)

    if reclassify_std:
        from .review import reclassify_std_dir

        engine = ClassifyEngine.from_data_dir(data_dir)
        stats = reclassify_std_dir(Path(reclassify_std), engine)
        report["reclassify"] = stats

    return report
