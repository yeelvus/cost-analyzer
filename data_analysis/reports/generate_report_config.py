#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
成本分析报告配置自动生成工具

功能：
1. 从成本统计数据自动识别所有字段
2. 基于字段名称和业务逻辑自动分组归类
3. 生成可编辑的JSON配置文件
4. 支持用户自定义调整
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

import pandas as pd
import json
import re
from pathlib import Path
from typing import Dict, List, Any
from datetime import datetime


class ReportConfigGenerator:
    """报告配置生成器"""
    
    # 字段关键词分类规则
    CATEGORY_RULES = {
        "业主合同": ["业主", "结算额"],
        "分包合同": ["分包"],
        "材料费": ["材料", "采购"],
        "商混站": ["商混"],
        "人工费": ["工资", "人工"],
        "机具费": ["机具"],
        "材料租赁": ["租赁"],
        "其他费用": ["调遣", "其他费用", "临建"],
        "综合指标": ["预计总成本", "收入", "毛利", "毛利率"],
    }
    
    # 字段显示顺序优先级（数字越小越靠前）
    FIELD_PRIORITY = {
        "预计总成本": 1,
        "收入": 2,
        "毛利": 3,
        "毛利率": 4,
        "业主预计结算额": 10,
        "分包结算额": 20,
        "分包已付款": 21,
        "材料采购实际已发生": 30,
        "材料采购后续预估发生": 31,
        "采购部采购金额": 32,
        "项目部采购金额": 33,
        "工资支出": 40,
        "管理人员工资": 41,
        "机具费支出": 50,
        "机具费后续预估": 51,
        "材料租赁费用": 60,
        "材料租赁后续预估": 61,
        "人员调遣支出": 70,
        "人员调遣后续预估": 71,
        "其他费用支出": 72,
        "其他费用后续预估": 73,
        "临建费用支出": 74,
        "临建费用后续预估": 75,
    }
    
    def __init__(self, cost_stats_file: str):
        """
        初始化配置生成器
        
        Args:
            cost_stats_file: 成本统计数据文件路径
        """
        self.cost_stats_file = Path(cost_stats_file)
        self.fields = []
        self.config = {}
        
    def load_fields(self):
        """加载成本统计数据的所有字段"""
        xl = pd.ExcelFile(self.cost_stats_file)
        df = pd.read_excel(self.cost_stats_file, sheet_name=xl.sheet_names[0])
        
        # 排除基本信息字段
        exclude_fields = ["项目名称", "标准项目编码", "标准项目名称", 
                         "项目名称（合同管理）", "匹配度", "完整文件名"]
        
        self.fields = [col for col in df.columns if col not in exclude_fields]
        print(f"✓ 已加载 {len(self.fields)} 个成本统计字段")
        
    def classify_field(self, field_name: str) -> str:
        """
        根据字段名称自动分类
        
        Args:
            field_name: 字段名称
            
        Returns:
            分类名称
        """
        for category, keywords in self.CATEGORY_RULES.items():
            if any(kw in field_name for kw in keywords):
                return category
        return "其他费用"  # 默认分类
    
    def get_field_priority(self, field_name: str) -> int:
        """获取字段优先级"""
        return self.FIELD_PRIORITY.get(field_name, 999)
    
    def infer_display_name(self, field_name: str) -> str:
        """
        推断字段显示名称（更简洁的表述）
        
        Args:
            field_name: 原始字段名
            
        Returns:
            显示名称
        """
        # 某些字段需要保持原名
        keep_as_is = ["预计总成本", "收入", "毛利", "毛利率"]
        if field_name in keep_as_is:
            return field_name
            
        # 简化名称
        simplifications = {
            "业主预计结算额": "预计结算额",
            "分包结算额": "结算额",
            "分包已付款": "已付款金额",
            "材料采购实际已发生": "实际已发生",
            "材料采购后续预估发生": "后续预估发生",
            "采购部采购金额": "采购部采购",
            "项目部采购金额": "项目部采购",
            "工资支出": "工资支出",
            "管理人员工资": "管理人员工资",
            "机具费支出": "机具费支出",
            "机具费后续预估": "机具费后续预估",
            "材料租赁费用": "租赁费用",
            "材料租赁后续预估": "租赁后续预估",
            "人员调遣支出": "人员调遣支出",
            "人员调遣后续预估": "人员调遣后续预估",
            "其他费用支出": "其他费用支出",
            "其他费用后续预估": "其他费用后续预估",
            "临建费用支出": "临建费用支出",
            "临建费用后续预估": "临建费用后续预估",
        }
        
        return simplifications.get(field_name, field_name)
    
    def infer_currency(self, field_name: str) -> str:
        """
        推断字段币种
        
        Args:
            field_name: 字段名称
            
        Returns:
            币种代码
        """
        # 比率字段无币种
        if "率" in field_name:
            return ""
        # 默认THB
        return "THB"
    
    def generate_config(self):
        """生成报告配置"""
        # 按分类组织字段
        categories = {}
        for field in self.fields:
            category = self.classify_field(field)
            if category not in categories:
                categories[category] = []
            
            field_config = {
                "field_name": field,
                "display_name": self.infer_display_name(field),
                "currency": self.infer_currency(field),
                "source": "成本统计",
                "priority": self.get_field_priority(field),
                "format": "percentage" if "率" in field else "amount",
                "note": ""  # 用户可手动添加备注说明
            }
            categories[category].append(field_config)
        
        # 排序每个分类内的字段
        for category in categories:
            categories[category].sort(key=lambda x: x["priority"])
        
        # 构建完整配置
        self.config = {
            "meta": {
                "version": "1.0",
                "generated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "source_file": str(self.cost_stats_file),
                "description": "成本分析报告自动生成配置"
            },
            "report_structure": self._build_report_structure(categories),
            "formatting": self._build_formatting_config(),
            "validation": self._build_validation_config()
        }
        
        print(f"✓ 已生成 {len(self.config['report_structure'])} 个大类的报告配置")
        
    def _build_report_structure(self, categories: Dict[str, List[Dict]]) -> List[Dict]:
        """
        构建报告结构配置
        
        Args:
            categories: 分类后的字段
            
        Returns:
            报告结构配置列表
        """
        # 定义大类顺序和显示名称
        category_order = [
            ("综合指标", "项目综合指标"),
            ("业主合同", "业主合同额"),
            ("分包合同", "分包合同金额"),
            ("材料费", "材料采购费"),
            ("商混站", "商混站"),
            ("人工费", "人工费"),
            ("机具费", "机具费"),
            ("材料租赁", "材料租赁费"),
            ("其他费用", "其他费用"),
        ]
        
        structure = []
        for idx, (category_key, category_display) in enumerate(category_order, 1):
            if category_key not in categories:
                continue
                
            # 中文数字
            section_number = ["一", "二", "三", "四", "五", "六", "七", "八", "九"][idx - 1]
            
            section_config = {
                "section_number": section_number,
                "section_id": category_key,
                "section_title": f"{section_number}、{category_display}",
                "enabled": True,
                "fields": categories[category_key],
                "custom_fields": [],  # 用户可添加自定义字段
                "notes": []  # 该大类的备注说明
            }
            structure.append(section_config)
        
        return structure
    
    def _build_formatting_config(self) -> Dict:
        """构建格式化配置"""
        return {
            "title": {
                "font_size": 14,
                "font_bold": True,
                "font_color": "FFFFFF",
                "fill_color": "4F81BD",
                "alignment": "center"
            },
            "section_header": {
                "font_size": 12,
                "font_bold": True,
                "font_color": "000000",
                "fill_color": "D9E1F2",
                "alignment": "left"
            },
            "field_header": {
                "font_size": 11,
                "font_bold": True,
                "font_color": "000000",
                "fill_color": "E7E6E6",
                "alignment": "center"
            },
            "data_row": {
                "font_size": 11,
                "font_bold": False,
                "alignment": "left"
            },
            "note_row": {
                "font_size": 10,
                "fill_color": "FFF2CC",
                "font_italic": True
            },
            "column_widths": {
                "项目/科目": 20,
                "明细": 25,
                "币种": 8,
                "金额": 15,
                "备注": 30
            },
            "number_format": {
                "amount": "#,##0.00",
                "percentage": "0.00%"
            }
        }
    
    def _build_validation_config(self) -> Dict:
        """构建验证配置"""
        return {
            "required_sections": ["综合指标", "业主合同", "分包合同"],
            "min_fields_per_section": 1,
            "allow_empty_values": True,
            "currency_validation": ["THB", "CNY", "USD", ""],
        }
    
    def save_config(self, output_path: str = None):
        """
        保存配置到JSON文件
        
        Args:
            output_path: 输出文件路径，默认为当前目录
        """
        if output_path is None:
            from settings import REPORT_CONFIG_PATH

            output_path = REPORT_CONFIG_PATH
        else:
            output_path = Path(output_path)
        
        with open(output_path, "w", encoding="utf-8") as f:
            json.dump(self.config, f, ensure_ascii=False, indent=2)
        
        print(f"✓ 配置已保存到: {output_path}")
        return output_path
    
    def print_summary(self):
        """打印配置摘要"""
        print("\n" + "="*60)
        print("报告配置摘要")
        print("="*60)
        
        for section in self.config["report_structure"]:
            print(f"\n{section['section_title']}")
            print(f"  字段数: {len(section['fields'])}")
            for field in section["fields"]:
                currency = f"[{field['currency']}]" if field['currency'] else ""
                print(f"    - {field['display_name']:20s} {currency:6s} <- {field['field_name']}")
        
        print("\n" + "="*60)
        print(f"总计: {sum(len(s['fields']) for s in self.config['report_structure'])} 个字段")
        print("="*60)


def main():
    """主函数"""
    # 成本统计数据文件路径
    cost_stats_file = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/2_成本统计收集excel情况汇总/2025年12月/项目统计数据汇总.xlsx"
    
    print("成本分析报告配置自动生成工具")
    print("="*60)
    
    # 创建生成器
    generator = ReportConfigGenerator(cost_stats_file)
    
    # 加载字段
    generator.load_fields()
    
    # 生成配置
    generator.generate_config()
    
    # 打印摘要
    generator.print_summary()
    
    # 保存配置
    output_file = generator.save_config()
    
    print(f"\n✓ 配置生成完成！")
    print(f"\n下一步操作：")
    print(f"  1. 查看并编辑配置文件: {output_file}")
    print(f"  2. 可以调整字段顺序、显示名称、添加备注等")
    print(f"  3. 运行报告生成程序时会自动读取此配置")


if __name__ == "__main__":
    main()
