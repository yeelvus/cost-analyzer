"""
通用修正（Corrections）基础设施

设计目标：
- 把所有“手动知识”（项目别名、物品名称映射、单位修正、分类调整、异常值修正等）
  从原始文件里剥离出来，变成显式、可版本化、可重放的结构。
- 提供统一的加载 + 应用接口，让各标准化脚本都能方便地“应用历史修正”。
- 参考并兼容现有的 common.project 项目匹配模式。

推荐用法：
1. 项目别名继续用 common.project（已很成熟）。
2. 物品/单位/分类等用本模块 + 一个中央 corrections/ 目录。
3. 每个业务月处理时，加载“截至该月有效的所有修正”后应用。

目录建议（在 01成本统计 下）：
corrections/
  project_aliases.json          # 复用现有
  item_name_mappings.xlsx       # 物品标准化规则（推荐有 effective_from）
  manual_corrections_log.xlsx   # 通用 ad-hoc 修正日志（最重要）
"""

from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd


@dataclass
class Correction:
    business_key: str                 # 业务键，例如 "项目编码:机具名称:日期" 或 "原始物品名"
    field: str                        # 要修正的字段
    corrected_value: Any
    reason: str = ""
    effective_from: Optional[date] = None
    source: str = ""                  # 哪个 raw 文件或批次产生的


class CorrectionsRegistry:
    """
    统一修正注册表。
    支持：
    - 从 JSON（简单映射）
    - 从 Excel log（推荐用于复杂 ad-hoc 修正）
    加载，并按 effective_from 过滤。
    """

    def __init__(self):
        self._alias_maps: Dict[str, Dict[str, str]] = {}   # type -> {raw: standard}
        self._corrections: List[Correction] = []

    # ---------------------- 加载接口 ----------------------

    def load_project_aliases(self, db_path: Path | str):
        """兼容现有 common.project 的 JSON 结构"""
        db_path = Path(db_path)
        if not db_path.exists():
            return
        with open(db_path, encoding="utf-8") as f:
            db = json.load(f)

        alias_map = {}
        projects = db.get("projects", {}) or {}
        for std_name, meta in projects.items():
            std_name = str(std_name).strip()
            if not std_name:
                continue
            alias_map[std_name] = std_name
            for alias in meta.get("aliases", []) or []:
                if alias:
                    alias_map[str(alias).strip()] = std_name

        self._alias_maps["project"] = alias_map

    def load_item_mappings(self, excel_path: Path | str, sheet: str = "映射表"):
        """加载物品名称/规格统一映射（推荐列：raw_name, standardized_name, category, unit, effective_from）"""
        p = Path(excel_path)
        if not p.exists():
            return
        df = pd.read_excel(p, sheet_name=sheet)
        df.columns = [str(c).strip() for c in df.columns]

        for _, row in df.iterrows():
            raw = str(row.get("raw_name") or row.get("原始名称") or "").strip()
            std = str(row.get("standardized_name") or row.get("标准化名称") or "").strip()
            if raw and std:
                self._alias_maps.setdefault("item", {})[raw] = std

    def load_manual_log(self, excel_path: Path | str, sheet: str = None):
        """
        加载通用手动修正日志。
        推荐列（不强制）：
        business_month, source_type, source_file, business_key,
        field, corrected_value, reason, corrected_at, effective_from
        """
        p = Path(excel_path)
        if not p.exists():
            return

        if sheet:
            df = pd.read_excel(p, sheet_name=sheet)
        else:
            df = pd.read_excel(p)

        df.columns = [str(c).strip().lower() for c in df.columns]

        for _, row in df.iterrows():
            key = str(row.get("business_key") or row.get("key") or "").strip()
            field = str(row.get("field") or row.get("字段") or "").strip()
            value = row.get("corrected_value") or row.get("修正值")
            reason = str(row.get("reason") or row.get("原因") or "")
            eff = row.get("effective_from") or row.get("生效日期")

            if key and field and value is not None:
                eff_date = None
                if pd.notna(eff):
                    try:
                        eff_date = pd.to_datetime(eff).date()
                    except Exception:
                        pass

                self._corrections.append(Correction(
                    business_key=key,
                    field=field,
                    corrected_value=value,
                    reason=reason,
                    effective_from=eff_date,
                    source=str(row.get("source_file") or row.get("来源文件") or "")
                ))

    # ---------------------- 应用接口 ----------------------

    def apply_alias(self, value: str, alias_type: str = "project") -> str:
        """简单别名映射（项目名、物品名等）"""
        if not value:
            return value
        mapping = self._alias_maps.get(alias_type, {})
        return mapping.get(str(value).strip(), value)

    def apply_corrections(self, df: pd.DataFrame, business_key_col: str,
                          target_fields: List[str] | None = None,
                          as_of: Optional[date] = None) -> pd.DataFrame:
        """
        把修正应用到 DataFrame。
        - business_key_col: DataFrame 里能唯一标识一行的业务键列名
        - target_fields: 只修正这些列（None 表示全部）
        - as_of: 只应用 effective_from <= as_of 的修正
        """
        if df.empty or not self._corrections:
            return df

        df = df.copy()
        key_col = business_key_col

        corrections_to_apply = self._corrections
        if as_of:
            corrections_to_apply = [c for c in corrections_to_apply
                                    if c.effective_from is None or c.effective_from <= as_of]

        # 按 business_key + field 建立快速查找
        corr_index: Dict[Tuple[str, str], Any] = {}
        for c in corrections_to_apply:
            corr_index[(c.business_key, c.field)] = c.corrected_value

        for idx, row in df.iterrows():
            key = str(row.get(key_col, "")).strip()
            if not key:
                continue
            for col in (target_fields or df.columns):
                if col == key_col:
                    continue
                val = corr_index.get((key, str(col)))
                if val is not None:
                    df.at[idx, col] = val

        return df

    # ---------------------- 辅助 ----------------------

    def get_effective_corrections(self, as_of: Optional[date] = None) -> List[Correction]:
        if not as_of:
            return self._corrections
        return [c for c in self._corrections
                if c.effective_from is None or c.effective_from <= as_of]

    def summary(self) -> Dict[str, int]:
        return {
            "alias_types": len(self._alias_maps),
            "total_corrections": len(self._corrections),
        }


# ==================== 便捷工厂函数 ====================

def load_all_corrections(corrections_root: Path,
                         project_aliases_file: str = "project_aliases.json",
                         item_mappings_file: Optional[str] = None,
                         manual_log_file: str = "manual_corrections_log.xlsx") -> CorrectionsRegistry:
    """
    一键加载 corrections_root 下的所有常用修正文件。
    """
    reg = CorrectionsRegistry()
    root = Path(corrections_root)

    # 项目别名（优先）
    pa = root / project_aliases_file
    if pa.exists():
        reg.load_project_aliases(pa)

    # 物品映射（可选）
    if item_mappings_file:
        im = root / item_mappings_file
        if im.exists():
            reg.load_item_mappings(im)

    # 通用手动日志（最推荐）
    ml = root / manual_log_file
    if ml.exists():
        reg.load_manual_log(ml)

    return reg


if __name__ == "__main__":
    # 示例用法
    print("CorrectionsRegistry 示例（请根据实际 corrections 目录调整路径）")
    # reg = load_all_corrections(Path("/Users/wyattsun/Desktop/邯郸中材/01成本统计/00_数据治理/corrections"))
    # print(reg.summary())
