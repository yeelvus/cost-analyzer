# subject_mapper.py
# 最终修正版 - 输出文件强制保存到桌面

import pandas as pd
import json
import re
from typing import Tuple, Optional
from pathlib import Path
from openpyxl import load_workbook


class SubjectMapper:
    """科目摘要智能分类器（简化JSON版）"""
    
    def __init__(self, rules_json_path: str):
        try:
            with open(rules_json_path, 'r', encoding='utf-8') as f:
                self.rules_dict = json.load(f)
        except json.JSONDecodeError as e:
            raise ValueError(f"rules.json 格式错误：{e}") from e

        self.rules = []
        self.subject_to_level1 = {}

        # 首选新结构（两段式）：
        # {
        #   "科目映射": {
        #      "0_分包": {"二级科目": "分包款;分包借款"},
        #      "1_人工费": {"二级科目": "保安工资;工资薪金"}
        #   },
        #   "关键词规则": {"二级科目": {"含有": "...", "不包含": "..."}, ...}
        # }
        subject_map = self.rules_dict.get("科目映射") if isinstance(self.rules_dict, dict) else None
        keyword_rules = self.rules_dict.get("关键词规则") if isinstance(self.rules_dict, dict) else None
        if isinstance(subject_map, dict) and isinstance(keyword_rules, dict):
            # 统一展开为：{"分包款": "0_分包", "分包借款": "0_分包"}
            expanded_map = {}
            for k, v in subject_map.items():
                key_text = str(k).strip()

                # 结构A: "一级": {"二级科目": "a;b;c"}
                if isinstance(v, dict):
                    if re.match(r"^\d+_", key_text):
                        items_text = str(
                            v.get("二级科目")
                            or v.get("科目")
                            or v.get("包含")
                            or ""
                        ).strip()
                        for level2 in [x.strip() for x in items_text.split(';') if x.strip()]:
                            expanded_map[level2] = key_text
                        continue

                # 结构B: "一级": "a;b;c"
                if isinstance(v, str) and re.match(r"^\d+_", key_text) and ";" in v:
                    for level2 in [x.strip() for x in v.split(';') if x.strip()]:
                        expanded_map[level2] = key_text
                    continue

                # 结构C: "二级": "一级"
                val_text = str(v).strip()
                if key_text:
                    expanded_map[key_text] = val_text

            self.subject_to_level1 = {str(k).strip(): str(v).strip() for k, v in expanded_map.items() if str(k).strip()}

            for level2, info in keyword_rules.items():
                if not isinstance(info, dict):
                    continue
                self.rules.append({
                    "二级科目": level2,
                    "一级科目": str(expanded_map.get(level2, "")).strip(),
                    "含有": info.get("含有", ""),
                    "不包含": info.get("不包含", "")
                })

            print(f"✅ 两段式规则加载成功！共 {len(self.rules)} 条规则")
            return

        # 兼容两种规则结构：
        # 1) 旧结构：{二级科目: {一级科目, 含有, 不包含}}
        # 2) 新结构：{一级科目: {二级科目: {含有, 不包含}}}
        for k1, v1 in self.rules_dict.items():
            if isinstance(v1, dict) and "一级科目" in v1:
                if str(k1).strip() and str(v1.get("一级科目", "")).strip():
                    self.subject_to_level1[str(k1).strip()] = str(v1.get("一级科目", "")).strip()
                self.rules.append({
                    "二级科目": k1,
                    "一级科目": v1.get("一级科目", ""),
                    "含有": v1.get("含有", ""),
                    "不包含": v1.get("不包含", "")
                })
                continue

            if isinstance(v1, dict):
                for level2, info in v1.items():
                    if not isinstance(info, dict):
                        continue
                    if str(level2).strip() and str(k1).strip():
                        self.subject_to_level1[str(level2).strip()] = str(k1).strip()
                    self.rules.append({
                        "二级科目": level2,
                        "一级科目": k1,
                        "含有": info.get("含有", ""),
                        "不包含": info.get("不包含", "")
                    })

        print(f"✅ 简化规则加载成功！共 {len(self.rules)} 条规则")

    def _parse_contains(self, contains_str: str):
        if not contains_str or not contains_str.strip():
            return []
        or_groups = [g.strip() for g in contains_str.split(';') if g.strip()]
        parsed = []
        for group in or_groups:
            and_keywords = [k.strip() for k in group.split('&') if k.strip()]
            if and_keywords:
                parsed.append(and_keywords)
        return parsed

    def _match_rule(self, summary: str, rule: dict) -> bool:
        if not summary or not isinstance(summary, str):
            return False
        s = summary.lower().strip()
        
        and_groups = self._parse_contains(rule.get("含有", ""))
        # 没有“含有”关键词的规则一律不参与自动匹配，避免误命中全部摘要。
        if not and_groups:
            return False

        matched = any(all(kw.lower() in s for kw in group) for group in and_groups)
        if not matched:
            return False
        
        excludes = rule.get("不包含", "")
        if excludes and excludes.strip():
            exclude_list = [k.strip().lower() for k in excludes.split(';') if k.strip()]
            if any(kw in s for kw in exclude_list):
                return False
        return True

    def classify(self, subject: str, summary: str) -> Tuple[str, str]:
        if pd.isna(subject):
            return "", ""

        txt = str(subject).strip()

        # 摘要为空时，优先按原始科目映射一级科目。
        if pd.isna(summary) or str(summary).strip() == "":
            return self.subject_to_level1.get(txt, ""), txt
        
        for rule in self.rules:
            if self._match_rule(summary, rule):
                return rule["一级科目"], rule["二级科目"]
        
        # 关键词未命中时，优先按原始科目映射一级科目。
        return self.subject_to_level1.get(txt, ""), txt

    def process_excel(self, excel_path: str, output_path: Optional[str] = None):
        print(f"📊 正在读取: {excel_path}")
        df = pd.read_excel(excel_path, sheet_name="Sheet1")
        print(f"✅ 读取成功！共 {len(df)} 行")
        
        results = df.apply(lambda row: self.classify(row.get('科目'), row.get('摘要')), axis=1)
        
        df['新一级科目'] = [r[0] for r in results]
        df['新二级科目'] = [r[1] for r in results]
        
        cols = list(df.columns)
        new_order = [c for c in cols if c not in ['新一级科目', '新二级科目']]
        new_order += ['新一级科目', '新二级科目']
        df = df[new_order]
        
        # 未传 output_path 时，默认保存到输入文件同目录，文件名加“_已分类”后缀。
        if output_path:
            out_path = Path(output_path)
        else:
            src = Path(excel_path)
            out_path = src.with_name(f"{src.stem}_已分类.xlsx")

        if out_path.suffix.lower() != '.xlsx':
            out_path = out_path.with_suffix('.xlsx')

        out_path.parent.mkdir(parents=True, exist_ok=True)
        df.to_excel(out_path, index=False)
        print(f"💾 文件已保存到：{out_path.resolve()}")
        print(f"📌 文件存在性检查：{'存在' if out_path.exists() else '不存在'}")
        
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
        """仅更新现有Excel中的一级/二级科目列，不改其他内容。返回更新行数。"""
        wb = load_workbook(excel_path)
        target_sheets = sheet_names or wb.sheetnames
        updated_rows = 0

        for sheet_name in target_sheets:
            if sheet_name not in wb.sheetnames:
                continue

            ws = wb[sheet_name]
            header_to_col = {}
            for c in range(1, ws.max_column + 1):
                header = ws.cell(row=1, column=c).value
                if header is None:
                    continue
                header_to_col[str(header).strip()] = c

            subj_idx = header_to_col.get(subject_col)
            summ_idx = header_to_col.get(summary_col)
            if subj_idx is None:
                continue

            lvl1_idx = header_to_col.get(level1_col)
            lvl2_idx = header_to_col.get(level2_col)

            if lvl1_idx is None:
                lvl1_idx = ws.max_column + 1
                ws.cell(row=1, column=lvl1_idx, value=level1_col)
            if lvl2_idx is None:
                lvl2_idx = ws.max_column + 1
                ws.cell(row=1, column=lvl2_idx, value=level2_col)

            for r in range(2, ws.max_row + 1):
                subject_value = ws.cell(row=r, column=subj_idx).value
                summary_value = ws.cell(row=r, column=summ_idx).value if summ_idx is not None else ""

                if subject_value is None and summary_value in (None, ""):
                    continue

                level1, level2 = self.classify(subject_value, summary_value)
                if not level1:
                    level1 = "6_管理费"
                if not level2:
                    level2 = str(subject_value).strip() if subject_value is not None else "(空白)"

                ws.cell(row=r, column=lvl1_idx, value=level1)
                ws.cell(row=r, column=lvl2_idx, value=level2)
                updated_rows += 1

        wb.save(excel_path)
        return updated_rows
