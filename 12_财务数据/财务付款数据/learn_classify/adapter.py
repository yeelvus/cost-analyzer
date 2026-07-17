# -*- coding: utf-8 -*-
"""
与旧 SubjectMapper 兼容的适配器，供 3_财务付款数据汇总台账.py 无插拔。

用法：
  from learn_classify.adapter import SubjectMapper
  mapper = SubjectMapper()  # 使用默认 data 目录
  mapper.classify(subject, summary)
"""

from __future__ import annotations

from pathlib import Path
from typing import Optional, Tuple

from .engine import ClassifyEngine
from .review import update_excel_classification_columns

DEFAULT_DATA_DIR = Path(__file__).resolve().parent / "data"


class SubjectMapper:
    """兼容旧接口：classify / process_excel / update_excel_subject_columns。"""

    def __init__(self, rules_json_path: str | None = None, data_dir: str | Path | None = None):
        """
        rules_json_path: 兼容旧调用；若指向 learn_classify/data 外的旧 rules，
        仍优先使用 data_dir 下的新体系。
        """
        if data_dir is not None:
            self.data_dir = Path(data_dir)
        elif rules_json_path:
            p = Path(rules_json_path)
            # 若传入 .../learn_classify/data/rules.json 或 同级 data
            if p.name == "rules.json" and (p.parent / "taxonomy.json").exists():
                self.data_dir = p.parent
            elif (p.parent / "learn_classify" / "data" / "taxonomy.json").exists():
                self.data_dir = p.parent / "learn_classify" / "data"
            else:
                self.data_dir = DEFAULT_DATA_DIR
        else:
            self.data_dir = DEFAULT_DATA_DIR

        self.engine = ClassifyEngine.from_data_dir(self.data_dir)
        # 兼容：旧代码读取 mapper.rules
        self.rules = []
        for r in self.engine.rules:
            l2_id = str(r.get("l2_id") or "")
            l1_name, l2_name = self.engine.taxonomy.display(l2_id) if l2_id else ("", "")
            self.rules.append(
                {
                    "二级科目": l2_name or r.get("id", ""),
                    "一级科目": l1_name,
                    "含有": r.get("含有", ""),
                    "不包含": r.get("不包含", ""),
                    "l2_id": l2_id,
                    "priority": r.get("priority", 100),
                }
            )
        self.subject_to_level1 = {}
        for name, l2_id in self.engine.subject_map.items():
            l1_name, _ = self.engine.taxonomy.display(l2_id)
            if l1_name:
                self.subject_to_level1[name] = l1_name

    def classify(self, subject, summary) -> Tuple[str, str]:
        r = self.engine.classify(subject, summary, None)
        return r.l1_name, r.l2_name

    def classify_full(self, subject, summary, vendor=None):
        return self.engine.classify(subject, summary, vendor)

    def process_excel(self, excel_path: str, output_path: Optional[str] = None):
        import pandas as pd

        df = pd.read_excel(excel_path, sheet_name=0)
        results = [
            self.engine.classify(row.get("科目"), row.get("摘要"), row.get("供应商名称"))
            for _, row in df.iterrows()
        ]
        df["新一级科目"] = [r.l1_name for r in results]
        df["新二级科目"] = [r.l2_name for r in results]
        df["分类ID"] = [r.l2_id for r in results]
        df["命中层"] = [r.layer for r in results]
        df["置信度"] = [r.confidence for r in results]
        out = Path(output_path) if output_path else Path(excel_path).with_name(
            f"{Path(excel_path).stem}_已分类.xlsx"
        )
        out.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(out, index=False)
        return df

    def update_excel_subject_columns(
        self,
        excel_path: str,
        sheet_names=None,
        subject_col: str = "科目",
        summary_col: str = "摘要",
        level1_col: str = "一级科目",
        level2_col: str = "二级科目",
    ) -> int:
        return update_excel_classification_columns(
            Path(excel_path),
            self.engine,
            sheet_names=sheet_names,
            subject_col=subject_col,
            summary_col=summary_col,
            level1_col=level1_col,
            level2_col=level2_col,
            extra_cols=True,
        )
