#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
配置驱动的成本分析报告构建器

功能：
1. 读取JSON配置文件
2. 根据配置动态生成报告结构
3. 整合多数据源（成本统计、合同、采购等）
4. 应用配置的格式化规则
"""


# --- path bootstrap (S_数据分析模块) ---
import sys
from pathlib import Path as _PathForBootstrap
try:
    from S_数据分析模块._bootstrap import ensure_code_root, ensure_script_dir
    ensure_code_root(_PathForBootstrap(__file__))
    ensure_script_dir(__file__)
except Exception:
    _p = _PathForBootstrap(__file__).resolve()
    for _c in [_p.parent, *_p.parents]:
        if (_c / "settings.py").exists():
            if str(_c) not in sys.path:
                sys.path.insert(0, str(_c))
            break
    if str(_p.parent) not in sys.path:
        sys.path.insert(0, str(_p.parent))
# --- end bootstrap ---

import json
import pandas as pd
from pathlib import Path
from typing import Dict, List, Any, Optional, Tuple
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment


class ConfigDrivenReportBuilder:
    """配置驱动的报告构建器"""
    
    # 类级别的配置缓存（单例模式）
    _cached_config = None
    _config_path_cached = None
    
    def __init__(self, config_path: str = None):
        """
        初始化报告构建器
        
        Args:
            config_path: 配置文件路径，默认使用同目录下的配置文件
        """
        if config_path is None:
            from settings import REPORT_CONFIG_PATH

            config_path = REPORT_CONFIG_PATH
        
        self.config_path = Path(config_path)
        self.config = self._load_config()
        
    def _load_config(self) -> Dict:
        """加载配置文件（使用缓存）"""
        if not self.config_path.exists():
            raise FileNotFoundError(f"配置文件不存在: {self.config_path}")
        
        # 使用类级别缓存
        if (ConfigDrivenReportBuilder._cached_config is not None and 
            ConfigDrivenReportBuilder._config_path_cached == str(self.config_path)):
            return ConfigDrivenReportBuilder._cached_config
        
        with open(self.config_path, "r", encoding="utf-8") as f:
            config = json.load(f)
        
        # 更新缓存
        ConfigDrivenReportBuilder._cached_config = config
        ConfigDrivenReportBuilder._config_path_cached = str(self.config_path)
        
        print(f"✓ 已加载配置: {self.config_path.name} (v{config['meta']['version']})")
        return config
    
    def build_report_rows(
        self,
        project_code: str,
        project_name: str,
        cost_stats_row: Optional[pd.Series],
        summary_row: Dict,
        main_rows: pd.DataFrame,
        sub_rows: pd.DataFrame,
        concrete_rows: pd.DataFrame,
        equipment_rows: pd.DataFrame = None,
        labor_rows: pd.DataFrame = None,
        finance_category_rows: pd.DataFrame = None,
        procurement_rows: pd.DataFrame = None,
        responsibility_cost_data: Dict = None,
        data_source_info: Dict = None,
    ) -> List[List]:
        """
        构建报告行数据
        
        Args:
            project_code: 项目编码
            project_name: 项目名称
            cost_stats_row: 成本统计数据行
            summary_row: 汇总数据字典
            main_rows: 总包合同数据
            sub_rows: 分包合同数据
            concrete_rows: 商混数据
            equipment_rows: 机具数据
            labor_rows: 人工数据
            responsibility_cost_data: 责任成本数据 {sheet_name: dataframe}
            data_source_info: 数据来源信息（文件名和统计日期）
            
        Returns:
            报告行列表
        """
        # 保存 data_source_info 到实例变量，供后续方法使用
        self.data_source_info = data_source_info if data_source_info else {}
        
        rows = []
        
        # 标题行
        title = f"{project_name or project_code}项目成本分析报告"
        rows.append([title, "", "", "", ""])
        
        # 遍历配置的报告结构
        for section in self.config["report_structure"]:
            if not section.get("enabled", True):
                continue
            
            # 添加大类标题行
            rows.append([section["section_title"], "", "", "", ""])
            
            # 添加字段表头
            rows.append(["项目/科目", "明细", "币种", "金额", "计算逻辑/备注"])
            
            # 处理该大类的所有字段
            section_rows = self._build_section_rows(
                section,
                cost_stats_row,
                summary_row,
                main_rows,
                sub_rows,
                concrete_rows,
                equipment_rows,
                labor_rows,
                finance_category_rows,
                procurement_rows,
                responsibility_cost_data,
            )
            rows.extend(section_rows)
            
            # 添加自定义字段
            for custom_field in section.get("custom_fields", []):
                rows.append([
                    custom_field.get("display_name", ""),
                    custom_field.get("detail", ""),
                    custom_field.get("currency", ""),
                    custom_field.get("amount", ""),
                    custom_field.get("note", ""),
                ])
            
            # 添加大类备注
            for note in section.get("notes", []):
                rows.append(["备注", "", "", "", note])
            
            # 添加空行分隔
            rows.append(["", "", "", "", ""])
        
        return rows
    
    def _build_section_rows(
        self,
        section: Dict,
        cost_stats_row: Optional[pd.Series],
        summary_row: Dict,
        main_rows: pd.DataFrame,
        sub_rows: pd.DataFrame,
        concrete_rows: pd.DataFrame,
        equipment_rows: pd.DataFrame,
        labor_rows: pd.DataFrame,
        finance_category_rows: pd.DataFrame = None,
        procurement_rows: pd.DataFrame = None,
        responsibility_cost_data: Dict = None,
    ) -> List[List]:
        """
        构建某个大类的所有行数据
        
        Args:
            section: 大类配置
            其他参数同 build_report_rows
            finance_category_rows: 财务分包付款明细数据
            responsibility_cost_data: 责任成本数据
            
        Returns:
            该大类的行列表
        """
        section_id = section["section_id"]
        rows = []
        
        # 特殊处理：业主合同和分包合同需要显示合同明细
        if section_id == "业主合同":
            rows.extend(self._build_contract_rows(main_rows, "业主合同明细"))
        elif section_id == "分包合同":
            # 计算签订分包合同金额汇总
            sub_total = 0.0
            if sub_rows is not None and not sub_rows.empty:
                sub_total = sub_rows["合同金额"].apply(self._coerce_float).sum() if "合同金额" in sub_rows.columns else 0.0
            rows.append(["签订分包合同金额汇总", "", "THB", sub_total, ""])
            # 工程部已签订分包合同明细（按序号排序）
            rows.extend(self._build_contract_rows(sub_rows, "工程部已签订分包合同", sort_by_prefix=True))
        elif section_id == "商混站":
            # 特殊处理商混站的方量
            volume = self._sum_concrete_volume(concrete_rows)
            rows.append(["方量", "", "", volume, ""])
        elif section_id == "人工费" and labor_rows is not None and not labor_rows.empty:
            # 特殊处理人工费：添加详细分项
            # 列名可能是"中方正式工铢"或"求和项:中方正式工铢"
            labor_detail_mapping = {
                "中方正式工": ["中方正式工铢", "求和项:中方正式工铢"],
                "中方劳务工": ["中方劳务工铢", "求和项:中方劳务工铢"],
                "子公司属地工": ["子公司属地工铢", "求和项:子公司属地工铢"],
                "劳务公司": ["劳务公司铢", "求和项:劳务公司铢"],
                "签证费": ["签证费", "求和项:签证费"],
            }
            for label, possible_cols in labor_detail_mapping.items():
                # 找到存在的列名
                actual_col = None
                for col in possible_cols:
                    if col in labor_rows.columns:
                        actual_col = col
                        break
                
                if actual_col:
                    value = labor_rows[actual_col].sum() if len(labor_rows) > 0 else 0
                    rows.append([label, "", "THB", value, "从人工费汇总sheet累计"])
        
        # 处理配置的字段
        for field_config in section["fields"]:
            field_name = field_config["field_name"]
            display_name = field_config["display_name"]
            currency = field_config["currency"]
            source = field_config.get("source", "成本统计")
            
            # 特殊处理：分包已付款从财务分包付款明细获取
            if field_name == "分包已付款" and finance_category_rows is not None and not finance_category_rows.empty:
                # 从财务分包付款明细的"付款汇总（泰铢）"列求和
                if "付款汇总（泰铢）" in finance_category_rows.columns:
                    value = finance_category_rows["付款汇总（泰铢）"].sum()
                else:
                    value = self._get_field_value(
                        field_name,
                        source,
                        cost_stats_row,
                        summary_row,
                        concrete_rows,
                    )
            # 特殊处理：采购部采购金额从材料采购明细获取
            elif field_name == "采购部采购金额" and procurement_rows is not None and not procurement_rows.empty:
                # 从材料采购明细的"含税金额"列求和
                if "含税金额" in procurement_rows.columns:
                    value = procurement_rows["含税金额"].sum()
                else:
                    value = self._get_field_value(
                        field_name,
                        source,
                        cost_stats_row,
                        summary_row,
                        concrete_rows,
                    )
            # 特殊处理：商混总方量和总金额从商混明细获取
            elif field_name in ["商混总方量", "商混总金额"] and concrete_rows is not None and not concrete_rows.empty:
                if field_name == "商混总方量":
                    # 从商混明细的"方量"或"总方量"列求和
                    volume_col = None
                    for col in ["方量", "总方量", "数量", "立方"]:
                        if col in concrete_rows.columns:
                            volume_col = col
                            break
                    if volume_col:
                        value = concrete_rows[volume_col].apply(self._coerce_float).sum()
                    else:
                        value = 0
                else:  # 商混总金额
                    # 从商混明细的"金额"或"计算总金额"或"金额（泰铢）"列求和
                    amount_col = None
                    for col in ["金额_THB", "计算总金额", "金额（泰铢）", "金额（万泰铢）", "金额"]:
                        if col in concrete_rows.columns:
                            amount_col = col
                            break
                    if amount_col:
                        amounts = concrete_rows[amount_col].apply(self._coerce_float)
                        # 如果是万泰铢，需要乘以10000
                        if "万泰铢" in amount_col or amount_col == "金额（万泰铢）":
                            value = (amounts * 10000).sum()
                        else:
                            value = amounts.sum()
                    else:
                        # 从汇总中获取
                        value = summary_row.get("商混站金额", 0)
            # 特殊处理：责任成本汇总金额从责任成本sheet获取
            elif field_name == "责任成本汇总金额" and responsibility_cost_data is not None and responsibility_cost_data:
                # 导入必要的函数来计算责任成本汇总
                from contract_report_generator import calculate_responsibility_cost_summary
                value, cost_currency = calculate_responsibility_cost_summary(responsibility_cost_data)
                if cost_currency and cost_currency != "THB":
                    currency = cost_currency
            else:
                # 获取字段值
                value = self._get_field_value(
                    field_name,
                    source,
                    cost_stats_row,
                    summary_row,
                    concrete_rows,
            )
            
            # 格式化值和计算说明
            if field_config.get("format") == "percentage" and value is not None:
                # 百分比保持原值，Excel格式化时处理
                pass
            
            # 生成计算逻辑说明
            calculation_note = self._generate_calculation_note(field_name, source, field_config)
            
            rows.append([display_name, "", currency, value, calculation_note])
            
            # 分包已付款后新增付款比例行
            if field_name == "分包已付款":
                # 获取分包结算额作为分母
                settlement_value = self._get_field_value("分包结算额", "成本统计", cost_stats_row, summary_row, concrete_rows)
                ratio = None
                if value and settlement_value and self._coerce_float(settlement_value):
                    try:
                        ratio = float(value) / float(settlement_value)
                    except (ValueError, TypeError, ZeroDivisionError):
                        ratio = None
                rows.append(["分包已付款比例", "", "", ratio, "分包已付款比例 = 分包已付款金额 / 分包结算额"])
        
        return rows
    
    def _get_field_value(
        self,
        field_name: str,
        source: str,
        cost_stats_row: Optional[pd.Series],
        summary_row: Dict,
        concrete_rows: pd.DataFrame,
    ) -> Any:
        """
        获取字段值
        
        Args:
            field_name: 字段名称
            source: 数据源
            cost_stats_row: 成本统计数据
            summary_row: 汇总数据
            concrete_rows: 商混数据
            
        Returns:
            字段值
        """
        # 从成本统计获取
        if source == "成本统计":
            if cost_stats_row is None:
                return None
            val = cost_stats_row.get(field_name)
            return None if pd.isna(val) else val
        
        # 从汇总数据获取
        if source == "汇总":
            val = summary_row.get(field_name)
            return None if pd.isna(val) else self._coerce_float(val)
        
        # 从商混数据计算
        if source == "商混" and field_name == "金额":
            val = summary_row.get("商混站金额")
            return None if pd.isna(val) else self._coerce_float(val)
        
        return None
    
    def _build_contract_rows(
        self,
        contract_df: pd.DataFrame,
        title: str,
        max_items: int = None,
        sort_by_prefix: bool = False
    ) -> List[List]:
        """
        构建合同明细行
        
        Args:
            contract_df: 合同数据
            title: 行标题（如"业主合同明细"或"工程部已签订分包合同"）
            max_items: 最多显示几个合同（None表示显示全部）
            sort_by_prefix: 是否按合同子文件夹名的序号前缀排序
            
        Returns:
            合同明细行列表
        """
        import re as _re
        rows = []
        
        if contract_df is None or contract_df.empty:
            rows.append([title, "", "", "", ""])
            rows.append(["", "", "", "", ""])
            return rows
        
        # 判断是否为分包合同（需要显示工程范围和对方公司名称）
        is_subcontract = "分包" in title or "工程部" in title
        
        # 按序号前缀排序（提取合同子文件夹名开头的数字）
        if sort_by_prefix and "合同子文件夹" in contract_df.columns:
            def _extract_prefix_num(name):
                m = _re.match(r'^(\d+)', str(name).strip())
                return int(m.group(1)) if m else 9999
            contract_df = contract_df.copy()
            contract_df["_sort_key"] = contract_df["合同子文件夹"].apply(_extract_prefix_num)
            contract_df = contract_df.sort_values("_sort_key").drop(columns=["_sort_key"])
        
        # 确定要显示的数据行数
        df_to_show = contract_df if max_items is None else contract_df.head(max_items)
        
        # 第一行：标题行
        first_row = True
        for idx, row in df_to_show.iterrows():
            if is_subcontract:
                # 分包合同：显示 合同子文件夹、工程范围、合同金额、对方公司名称
                subfolder = str(row.get("合同子文件夹", "")).strip() or ""
                scope = str(row.get("工程范围", "")).strip() or ""
                amount = row.get("合同金额", 0.0)
                company = str(row.get("对方公司名称", "") or row.get("对方单位公司名称", "")).strip() or ""
                
                if first_row:
                    rows.append([title, subfolder, scope, amount, company])
                    first_row = False
                else:
                    rows.append(["", subfolder, scope, amount, company])
            else:
                # 业主合同：显示 合同子文件夹/合同类型、币种、合同金额
                name = str(row.get("合同子文件夹", "") or row.get("合同类型", "") or "合同").strip()
                currency = str(row.get("币种", "")).strip()
                amount = row.get("合同金额", 0.0)
                
                if first_row:
                    rows.append([title, name, currency, amount, ""])
                    first_row = False
                else:
                    rows.append(["", name, currency, amount, ""])
        
        # 如果没有任何数据，添加空行
        if first_row:
            rows.append([title, "", "", "", ""])
            rows.append(["", "", "", "", ""])
        
        return rows
    
    def _sum_concrete_volume(self, df: pd.DataFrame) -> float:
        """计算商混方量"""
        if df is None or df.empty:
            return 0.0
        
        for col in ["方量", "数量", "方量(立方)"]:
            if col in df.columns:
                return df[col].apply(self._coerce_float).sum()
    def _generate_calculation_note(self, field_name: str, source: str, field_config: Dict) -> str:
        """
        生成计算逻辑说明
        
        Args:
            field_name: 字段名称
            source: 数据源
            field_config: 字段配置
            
        Returns:
            计算逻辑说明文本
        """
        # 定义字段与数据源的映射关系
        field_to_source_mapping = {
            "预计总成本": "项目成本统计汇总表",
            "收入": "项目成本统计汇总表",
            "业主预计结算额": "项目成本统计汇总表",
            "分包结算额": "项目成本统计汇总表",
            "分包已付款": "财务分包付款明细",
            "材料采购实际已发生": "项目成本统计汇总表",
            "材料采购后续预估发生": "项目成本统计汇总表",
            "采购部采购金额": "材料采购明细",
            "项目部采购金额": "项目成本统计汇总表",
            "工资支出": "项目成本统计汇总表",
            "管理人员工资": "项目成本统计汇总表",
            "机具费支出": "项目成本统计汇总表",
            "机具费后续预估": "项目成本统计汇总表",
            "材料租赁费用": "项目成本统计汇总表",
            "材料租赁后续预估": "项目成本统计汇总表",
            "人员调遣支出": "项目成本统计汇总表",
            "人员调遣后续预估": "项目成本统计汇总表",
            "其他费用支出": "项目成本统计汇总表",
            "其他费用后续预估": "项目成本统计汇总表",
            "临建费用支出": "项目成本统计汇总表",
            "临建费用后续预估": "项目成本统计汇总表",
            "商混总方量": "商混站明细",
            "商混总金额": "商混站明细",
            "机具使用费": "机具使用费明细",
            "责任成本汇总金额": "责任成本",
        }

        # 获取数据源信息（如果可用）
        data_source_info = getattr(self, 'data_source_info', {})

        def build_source_suffix(source_key: str) -> str:
            if not source_key or source_key not in data_source_info:
                return ""
            source_info = data_source_info[source_key]
            file_name = source_info.get("文件名", "____")
            if source_key == "责任成本":
                date_field = "编制日期"
                date_label = "编制日期"
                date_value = source_info.get(date_field, "____年__月")
            else:
                date_field = "统计截止日期"
                date_label = "统计截止日期"
                date_value = source_info.get(date_field, "____年__月")
            return f"（数据来源文件：{file_name}，{date_label}：{date_value}）"

        # 如果配置中已有备注，也追加动态来源时间，确保随数据自动更新。
        if field_config.get("note"):
            source_key = field_to_source_mapping.get(field_name)
            return f"{field_config['note']}{build_source_suffix(source_key)}"
        
        # 对于计算字段，不需要数据源信息
        if field_name == "毛利":
            return "毛利 = 收入 - 预计总成本"
        elif field_name == "毛利率":
            return "毛利率 = 毛利 / 收入 × 100%"
        
        # 获取该字段对应的数据源key
        source_key = field_to_source_mapping.get(field_name)
        
        if source_key and source_key in data_source_info:
            source_info = data_source_info[source_key]
            file_name = source_info.get("文件名", "____")
            
            # 根据不同的数据源使用不同的日期字段名
            if source_key == "责任成本":
                date_field = "编制日期"
            else:
                date_field = "统计截止日期"
            
            date_value = source_info.get(date_field, "____年__月")
            
            # 根据字段生成说明文本
            if source_key == "项目成本统计汇总表":
                return f"项目提供数据（统计截止日期：{date_value}）"
            elif source_key == "财务分包付款明细":
                return f"引用财务分包付款明细（付款汇总泰铢）（数据来源文件：{file_name}，统计截止日期：{date_value}）"
            elif source_key == "材料采购明细":
                return f"引用材料采购明细sheet页（数据来源文件：{file_name}，统计截止日期：{date_value}）"
            elif source_key == "商混站明细":
                return f"引用商混站明细sheet页（数据来源文件：{file_name}，统计截止日期：{date_value}）"
            elif source_key == "机具使用费明细":
                return f"引用机具使用费明细sheet页（数据来源文件：{file_name}，统计截止日期：{date_value}）"
            elif source_key == "责任成本":
                return f"引用责任成本sheet页（数据来源文件：{file_name}，编制日期：{date_value}）"
        
        # 如果没有找到对应的数据源信息，使用默认的占位符格式
        calculations = {
            "预计总成本": "项目提供数据（统计截止日期：____年__月）",
            "收入": "项目提供数据（统计截止日期：____年__月）",
            "业主预计结算额": "项目提供数据（统计截止日期：____年__月）",
            "分包结算额": "项目提供数据（统计截止日期：____年__月）",
            "分包已付款": "引用财务分包付款明细（付款汇总泰铢）（数据来源文件：____，统计截止日期：____年__月）",
            "材料采购实际已发生": "项目提供数据（统计截止日期：____年__月）",
            "材料采购后续预估发生": "项目提供数据（统计截止日期：____年__月）",
            "采购部采购金额": "引用材料采购明细sheet页（数据来源文件：____，统计截止日期：____年__月）",
            "项目部采购金额": "项目提供数据（统计截止日期：____年__月）",
            "工资支出": "项目提供数据（统计截止日期：____年__月）",
            "管理人员工资": "项目提供数据（统计截止日期：____年__月）",
            "机具费支出": "项目提供数据（统计截止日期：____年__月）",
            "机具费后续预估": "项目提供数据（统计截止日期：____年__月）",
            "材料租赁费用": "项目提供数据（统计截止日期：____年__月）",
            "材料租赁后续预估": "项目提供数据（统计截止日期：____年__月）",
            "人员调遣支出": "项目提供数据（统计截止日期：____年__月）",
            "人员调遣后续预估": "项目提供数据（统计截止日期：____年__月）",
            "其他费用支出": "项目提供数据（统计截止日期：____年__月）",
            "其他费用后续预估": "项目提供数据（统计截止日期：____年__月）",
            "临建费用支出": "项目提供数据（统计截止日期：____年__月）",
            "临建费用后续预估": "项目提供数据（统计截止日期：____年__月）",
            "商混总方量": "引用商混站明细sheet页（数据来源文件：____，统计截止日期：____年__月）",
            "商混总金额": "引用商混站明细sheet页（数据来源文件：____，统计截止日期：____年__月）",
            "机具使用费": "引用机具使用费明细sheet页（数据来源文件：____，统计截止日期：____年__月）",
            "责任成本汇总金额": "引用责任成本sheet页（数据来源文件：____，编制日期：____年__月）",
        }
        
        return calculations.get(field_name, f"来源: {source}")
    
    @staticmethod
    def _coerce_float(val) -> Optional[float]:
        """转换为浮点数"""
        if pd.isna(val):
            return None
        try:
            return float(val)
        except (ValueError, TypeError):
            return None
    
    def style_report_sheet(self, ws):
        """
        应用配置的格式化样式（专业版 - DCF风格）
        
        Args:
            ws: openpyxl worksheet对象
        """
        import re
        
        # 专业配色方案 - 类似DCF估值表
        color_scheme = {
            "main_title": {"bg": "1F4E78", "font": "FFFFFF"},  # 深蓝色主标题
            "section_title": {"bg": "4472C4", "font": "FFFFFF"},  # 蓝色大类标题
            "sub_header": {"bg": "8EA9DB", "font": "000000"},  # 浅蓝色子标题
            "data_row_odd": {"bg": "D9E2F3", "font": "000000"},  # 浅蓝色数据行（奇数）
            "data_row_even": {"bg": "FFFFFF", "font": "000000"},  # 白色数据行（偶数）
            "calculation": {"bg": "FFF2CC", "font": "000000"},  # 淡黄色计算说明
            "key_metric": {"bg": "C6E0B4", "font": "000000"},  # 淡绿色关键指标
        }
        
        # 字体配置
        fonts = {
            "main_title": Font(name="微软雅黑", size=14, bold=True, color=color_scheme["main_title"]["font"]),
            "section": Font(name="微软雅黑", size=12, bold=True, color=color_scheme["section_title"]["font"]),
            "sub_header": Font(name="微软雅黑", size=11, bold=True, color=color_scheme["sub_header"]["font"]),
            "data": Font(name="微软雅黑", size=10, color="000000"),
            "calculation": Font(name="微软雅黑", size=9, italic=True, color="7F7F7F"),
        }
        
        # 边框样式
        thin_border = Border(
            left=Side(style="thin", color="B4C7E7"),
            right=Side(style="thin", color="B4C7E7"),
            top=Side(style="thin", color="B4C7E7"),
            bottom=Side(style="thin", color="B4C7E7")
        )
        
        medium_border = Border(
            left=Side(style="medium", color="4472C4"),
            right=Side(style="medium", color="4472C4"),
            top=Side(style="medium", color="4472C4"),
            bottom=Side(style="medium", color="4472C4")
        )
        
        # 1. 主标题样式（第1行）
        ws.row_dimensions[1].height = 30
        for cell in ws[1]:
            cell.font = fonts["main_title"]
            cell.fill = PatternFill(start_color=color_scheme["main_title"]["bg"], 
                                   end_color=color_scheme["main_title"]["bg"], 
                                   fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = medium_border
        ws.merge_cells("A1:E1")
        
        # 2. 遍历所有行应用样式
        current_section = 0
        data_row_counter = 0
        section_has_header = {}  # 追踪各个section是否已经出现过表头
        
        for row_idx, row in enumerate(ws.iter_rows(min_row=2, max_row=ws.max_row), start=2):
            cell_value = row[0].value
            cell_value_str = str(cell_value).strip() if cell_value is not None else ""
            
            # 大类标题（一、二、三...）
            if cell_value and isinstance(cell_value, str) and re.match(r"^[一二三四五六七八九十]、", cell_value):
                current_section += 1
                data_row_counter = 0
                section_has_header[current_section] = False  # 重置该section的表头标记
                ws.row_dimensions[row_idx].height = 25
                for cell in row:
                    cell.font = fonts["section"]
                    cell.fill = PatternFill(start_color=color_scheme["section_title"]["bg"],
                                          end_color=color_scheme["section_title"]["bg"],
                                          fill_type="solid")
                    cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                    cell.border = medium_border
                ws.merge_cells(f"A{row_idx}:E{row_idx}")
            
            # 字段表头（项目/科目）
            elif cell_value == "项目/科目":
                section_has_header[current_section] = True  # 标记该section已有表头
                ws.row_dimensions[row_idx].height = 22
                for cell in row:
                    cell.font = fonts["sub_header"]
                    cell.fill = PatternFill(start_color=color_scheme["sub_header"]["bg"],
                                          end_color=color_scheme["sub_header"]["bg"],
                                          fill_type="solid")
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                    cell.border = thin_border
            
            # 备注行（行首为"备注"或行尾为"备注"）
            elif cell_value_str == "备注" or (len(row) > 4 and str(row[4].value).strip() == "备注"):
                for cell in row:
                    cell.font = fonts["calculation"]
                    cell.fill = PatternFill(start_color=color_scheme["calculation"]["bg"],
                                          end_color=color_scheme["calculation"]["bg"],
                                          fill_type="solid")
                    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
                    cell.border = thin_border
            
            # 空行（所有列都为空）
            elif not cell_value_str and all(str(cell.value).strip() == "" if cell.value is not None else True for cell in row):
                ws.row_dimensions[row_idx].height = 8
                for cell in row:
                    cell.border = None
            
            # 合同明细行和数据行（包括汇总行）
            else:
                # 检查是否为分包/业主合同明细行或汇总行
                # 这些行的特点是：第一列通常为空或为分包/业主标签，后续列有数据
                is_detail_row = (
                    cell_value_str == "" or  # 明细行首列为空
                    "合同" in cell_value_str or  # 包含"合同"关键字
                    "小计" in cell_value_str or  # 小计/汇总行
                    "合计" in cell_value_str or  # 合计行
                    "汇总" in cell_value_str or  # 汇总行
                    any(str(cell.value).strip() for cell in row[1:])  # 或后续列有非空内容
                )
                
                if is_detail_row or section_has_header.get(current_section, False):
                    # 这是一个数据行，需要应用交替颜色
                    data_row_counter += 1
                    ws.row_dimensions[row_idx].height = 20
                    
                    # 交替行颜色（项目提供数据来源行用淡绿色）
                    note_cell_value = str(row[4].value).strip() if len(row) > 4 and row[4].value is not None else ""
                    if note_cell_value.startswith("项目提供数据"):
                        bg_color = color_scheme["key_metric"]["bg"]  # 淡绿色
                    else:
                        bg_color = color_scheme["data_row_odd"]["bg"] if data_row_counter % 2 == 1 else color_scheme["data_row_even"]["bg"]
                    
                    for col_idx, cell in enumerate(row):
                        cell.font = fonts["data"]
                        cell.fill = PatternFill(start_color=bg_color, end_color=bg_color, fill_type="solid")
                        cell.border = thin_border
                        
                        # 对齐方式
                        if col_idx == 0:  # 项目/科目列
                            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1 if cell_value_str else 0)
                        elif col_idx in [1, 4]:  # 明细、备注列
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        elif col_idx == 2:  # 币种列
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        elif col_idx == 3:  # 金额列
                            cell.alignment = Alignment(horizontal="right", vertical="center")
                else:
                    # 无法识别的行，应用基础样式
                    ws.row_dimensions[row_idx].height = 20
                    for col_idx, cell in enumerate(row):
                        cell.font = fonts["data"]
                        cell.fill = PatternFill(start_color=color_scheme["data_row_even"]["bg"], 
                                              end_color=color_scheme["data_row_even"]["bg"], 
                                              fill_type="solid")
                        cell.border = thin_border
                        
                        if col_idx == 0:
                            cell.alignment = Alignment(horizontal="left", vertical="center", indent=1)
                        elif col_idx in [1, 4]:
                            cell.alignment = Alignment(horizontal="left", vertical="center")
                        elif col_idx == 2:
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                        elif col_idx == 3:
                            cell.alignment = Alignment(horizontal="right", vertical="center")
        
        # 3. 设置列宽（更宽松的布局）
        ws.column_dimensions["A"].width = 28  # 项目/科目
        ws.column_dimensions["B"].width = 35  # 明细
        ws.column_dimensions["C"].width = 10  # 币种
        ws.column_dimensions["D"].width = 20  # 金额
        ws.column_dimensions["E"].width = 45  # 备注/计算说明
        
        # 4. 应用数字格式（使用完整金额，不使用M简写）
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
            # D列是金额列
            if row[3].value is not None:
                try:
                    val = float(row[3].value)
                    # 检查是否为百分比格式（通过行首列内容或计算说明判断）
                    is_percentage = False
                    if row[0].value:
                        first_col_str = str(row[0].value).strip()
                        # 如果是毛利率或付款比例，则判定为百分比
                        if '毛利率' in first_col_str or '比例' in first_col_str:
                            is_percentage = True
                    
                    # 应用相应的格式
                    if is_percentage:
                        row[3].number_format = '0.00%'  # 百分比格式
                    else:
                        row[3].number_format = '#,##0.00'  # 千分位格式
                except (ValueError, TypeError):
                    pass
        
        # 5. 冻结窗格
        ws.freeze_panes = "A2"
        
        # 6. 设置打印选项
        ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
    
    def reload_config(self):
        """重新加载配置文件（用于配置更新后）"""
        # 清除缓存
        ConfigDrivenReportBuilder._cached_config = None
        ConfigDrivenReportBuilder._config_path_cached = None
        
        self.config = self._load_config()
        print("✓ 配置已重新加载")


# 便捷函数，供外部调用
def create_report_builder(config_path: str = None) -> ConfigDrivenReportBuilder:
    """
    创建报告构建器实例
    
    Args:
        config_path: 配置文件路径
        
    Returns:
        ConfigDrivenReportBuilder实例
    """
    return ConfigDrivenReportBuilder(config_path)


if __name__ == "__main__":
    # 测试配置加载
    builder = create_report_builder()
    print(f"配置版本: {builder.config['meta']['version']}")
    print(f"报告大类数: {len(builder.config['report_structure'])}")
    
    for section in builder.config["report_structure"]:
        print(f"  - {section['section_title']}: {len(section['fields'])} 个字段")
