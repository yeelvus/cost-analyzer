# -*- coding: utf-8 -*-
import pandas as pd
import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill
from pathlib import Path
import re
import os
import sys
from datetime import datetime
import json

BASE_CODE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
MATCHER_DIR_CANDIDATES = [
    os.path.join(BASE_CODE_DIR, '0_项目基本信息'),
    os.path.join(BASE_CODE_DIR, '0_项目基本信息情况'),
]

for matcher_dir in MATCHER_DIR_CANDIDATES:
    if os.path.exists(os.path.join(matcher_dir, 'project_name_matcher.py')):
        if matcher_dir not in sys.path:
            sys.path.insert(0, matcher_dir)
        break

from project_name_matcher import (
    build_standard_projects_from_dataframe,
    is_unmatched_alias,
    load_or_init_db,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
)

# ==================== 配置区 ====================
# 标准化文件目录（自动选择最新文件）
STANDARDIZED_INPUT_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_std"

# 如需临时强制指定某个文件，可填写绝对路径；默认留空自动选择最新
EXCEL_FILE_OVERRIDE = ""

# 映射表文件名（JSON）
PROJECT_MAPPING_JSON = "/Users/wyattsun/Desktop/code/10_机具费汇总代码/区域映射表.json"      # ←←← 项目区域映射文件(JSON)
MACHINE_MAPPING_JSON = "/Users/wyattsun/Desktop/code/10_机具费汇总代码/机具映射表.json"      # ←←← 机具种类和型号映射文件(JSON)

# 兼容旧版TXT映射表（仅用于首次自动迁移）
PROJECT_MAPPING_TXT_LEGACY = "/Users/wyattsun/Desktop/code/10_机具费汇总代码/区域映射表.txt"
MACHINE_MAPPING_TXT_LEGACY = "/Users/wyattsun/Desktop/code/10_机具费汇总代码/机具映射表.txt"

# 项目编码名称汇总文件（统一标准项目映射）
PROJECT_CODE_NAME_MAPPING_FILE = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"  # ←←← 统一的项目编码映射表

# 汇总文件输出文件夹路径
OUTPUT_FOLDER = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_sum"                  # ←←← 汇总文件保存位置

# 汇总输出文件名前缀（精简）
OUTPUT_FILENAME_PREFIX = "机具费汇总"
OUTPUT_FILE_SUFFIX = "汇总"

PROJECT_MATCH_DB_PATH = os.path.join("/Users/wyattsun/Desktop/code/0_项目基本信息", 'project_name_match_db.json')

# 汇总后新 sheet 的名字
SUMMARY_SHEET_NAME = "汇总"
PROJECT_SUMMARY_SHEET_NAME = "项目汇总机具费"
PROJECT_MONTHLY_CHANGE_SHEET_NAME = "项目月度金额增量"

# ================================================

LATEST_FILE_MONTH_RE = re.compile(r"(20\d{2})\s*年\s*(1[0-2]|0?[1-9])\s*月")


def extract_year_month_from_filename(file_path):
    """从文件名提取年月，用于“最新月份”判定。"""
    match = LATEST_FILE_MONTH_RE.search(file_path.name)
    if not match:
        return None
    return int(match.group(1)), int(match.group(2))


def pick_latest_standardized_file(input_dir):
    """自动选择标准化目录中最新文件：优先按文件名年月，其次按修改时间。"""
    input_path = Path(input_dir)
    if not input_path.exists():
        return None

    candidates = [
        p for p in input_path.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    if not candidates:
        return None

    dated_candidates = []
    undated_candidates = []
    for file_path in candidates:
        ym = extract_year_month_from_filename(file_path)
        mtime = file_path.stat().st_mtime
        if ym is not None:
            dated_candidates.append((ym, mtime, file_path))
        else:
            undated_candidates.append((mtime, file_path))

    if dated_candidates:
        dated_candidates.sort(key=lambda item: (item[0], item[1]))
        return dated_candidates[-1][2]

    undated_candidates.sort(key=lambda item: item[0])
    return undated_candidates[-1][1]


def list_all_standardized_files(input_dir):
    """列出标准化目录下全部待汇总文件（优先 *_标准化.xlsx）。"""
    input_path = Path(input_dir)
    if not input_path.exists():
        return []

    preferred = [
        p for p in input_path.glob("*_标准化.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    if preferred:
        return sorted(preferred, key=lambda p: p.name)

    fallback = [
        p for p in input_path.glob("*.xlsx")
        if p.is_file() and not p.name.startswith("~$")
    ]
    return sorted(fallback, key=lambda p: p.name)


def build_summary_output_basename(excel_path: Path) -> str:
    """输出名使用原台账名，去掉“_标准化”后再拼接时间后缀。"""
    stem = excel_path.stem
    if stem.endswith("_标准化"):
        stem = stem[:-4]
    return stem

def load_project_config():
    """加载项目配置文件（从合同管理系统）"""
    config_path = os.path.join(
        os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
        '5_合同管理', '0_合同管理代码', 'migration_config.json'
    )
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"警告：无法加载项目配置文件 {config_path}: {e}")
        return []


def load_unified_project_mapping():
    """从 项目编码名称汇总.xlsx 中加载统一的项目编码映射表
    包含Excel工作簿名称用于Sheet名称匹配
    返回格式: [{'project_code': 'THCC-2024-007', 'project_name': '双塔项目', 'sheet_name': 'THCC-2024-007_双塔项目', 'status': '成本梳理'}, ...]
    """
    projects = []
    
    if not os.path.exists(PROJECT_CODE_NAME_MAPPING_FILE):
        print(f"警告：统一项目映射文件不存在: {PROJECT_CODE_NAME_MAPPING_FILE}")
        return projects
    
    try:
        # 读取「汇总」工作簿
        df = pd.read_excel(PROJECT_CODE_NAME_MAPPING_FILE, sheet_name="汇总")
        
        # 期望的列名
        code_col = "项目编码"
        name_col = "项目名称（合同管理）"  # 优先用合同管理的名称
        status_col = "项目状态"
        
        # 查找Excel工作簿名称列（可能是"Unnamed: 5"或用户自定义的列名）
        excel_sheet_col = None
        for col in df.columns:
            if "sheet" in col.lower() or "工作簿" in col or "Unnamed" in col:
                excel_sheet_col = col
                break
        
        # 如果没找到，使用Unnamed: 5（如果存在）
        if not excel_sheet_col and "Unnamed: 5" in df.columns:
            excel_sheet_col = "Unnamed: 5"
        
        # 遍历所有行
        for idx, row in df.iterrows():
            code = str(row[code_col]).strip() if pd.notna(row[code_col]) else ""
            name = str(row[name_col]).strip() if pd.notna(row[name_col]) else ""
            status = str(row[status_col]).strip() if pd.notna(row[status_col]) else ""
            
            # 如果合同管理名称为空，使用成本梳理名称
            if not name and pd.notna(row.get("项目名称（成本梳理）")):
                name = str(row["项目名称（成本梳理）"]).strip()
            
            # 获取Excel工作簿名称（如果有）
            sheet_name = ""
            if excel_sheet_col and pd.notna(row.get(excel_sheet_col)):
                sheet_name = str(row[excel_sheet_col]).strip()
            
            if code and name:
                projects.append({
                    "project_code": code,
                    "project_name": name,
                    "sheet_name": sheet_name,
                    "status": status
                })
        
        print(f"✓ 已从统一映射表加载 {len(projects)} 个项目编码映射")
        return projects
    
    except Exception as e:
        print(f"警告：读取统一项目映射表失败: {e}")
        import traceback
        traceback.print_exc()
        return projects

def extract_project_name_from_sheet(sheet_name):
    """从sheet名称提取项目名称"""
    # 尝试匹配 "XX项目" 模式
    match = re.match(r'^(.+?项目)', sheet_name)
    if match:
        return match.group(1)
    
    # 如果有下划线或空格，取第一部分
    if '_' in sheet_name:
        return sheet_name.split('_')[0]
    if ' ' in sheet_name:
        return sheet_name.split(' ')[0]
    
    return sheet_name  # 否则使用整个sheet名

def parse_project_mapping_txt(txt_path):
    """解析旧版TXT：标准区域: 别名1, 别名2... -> {别名: 标准区域}。"""
    project_to_region = {}
    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#") or ":" not in line:
                continue
            standard, raw_list = line.split(":", 1)
            standard = standard.strip()
            for raw in raw_list.split(","):
                raw = raw.strip()
                if raw:
                    project_to_region[raw] = standard
    return project_to_region


def load_project_mapping(json_path, legacy_txt_path=None):
    """加载项目→区域映射。优先JSON，不存在时自动从旧TXT迁移。

    JSON支持两种结构：
    1) 推荐分组：{"region_to_projects": {"春武里府": ["欣旺达", "樱花"]}}
    2) 兼容扁平：{"project_to_region": {"欣旺达": "春武里府"}}
    返回统一为 {项目别名: 区域}。
    """
    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            mapping = {}
            if isinstance(data, dict) and "region_to_projects" in data:
                region_to_projects = data.get("region_to_projects", {})
                if isinstance(region_to_projects, dict):
                    for region, projects in region_to_projects.items():
                        region = str(region).strip()
                        if not region:
                            continue
                        if isinstance(projects, list):
                            for project in projects:
                                project = str(project).strip()
                                if project:
                                    mapping[project] = region
                        elif isinstance(projects, str):
                            for project in re.split(r'[，,]', projects):
                                project = project.strip()
                                if project:
                                    mapping[project] = region
            elif isinstance(data, dict) and "project_to_region" in data:
                flat_mapping = data.get("project_to_region", {})
                if isinstance(flat_mapping, dict):
                    mapping = {
                        str(project).strip(): str(region).strip()
                        for project, region in flat_mapping.items()
                        if str(project).strip()
                    }
            elif isinstance(data, dict):
                # 兼容直接以 {raw: std} 存储的JSON
                mapping = {
                    str(project).strip(): str(region).strip()
                    for project, region in data.items()
                    if str(project).strip()
                }

            return mapping
        except Exception as e:
            print(f"警告：读取项目映射JSON失败 {json_path}: {e}")

    if legacy_txt_path and os.path.exists(legacy_txt_path):
        mapping = parse_project_mapping_txt(legacy_txt_path)
        save_project_mapping(json_path, mapping)
        print(f"已将旧TXT项目映射迁移到JSON：{json_path}")
        return mapping

    save_project_mapping(json_path, {})
    print(f"已创建项目区域映射JSON模板：{json_path}")
    return {}

def apply_smart_mapping_rules(raw_name):
    """智能映射规则：根据关键词自动识别机具种类"""
    raw_lower = raw_name.lower().replace(" ", "")
    
    # 定义基础映射规则（按优先级排序）
    rules = [
        # 挖掘机类
        (["挖掘机", "挖机", "excavator", "pc200", "pc300", "cat320", "sk200", "sy215"], "挖掘机"),
        (["小挖", "pc30", "pc35", "pc45", "pc60", "pc75"], "小型挖掘机"),
        (["挖铲", "jcb", "反铲"], "挖铲"),
        
        # 装载机类
        (["装载机", "铲车", "loader", "zl50", "zw140"], "装载机"),
        
        # 推土机类
        (["推土机", "bulldozer"], "推土机"),
        
        # 压路机类
        (["压路机", "压实机", "震动压路机", "roller"], "压路机"),
        
        # 起重机械类
        (["塔吊", "塔机", "tower", "c7030", "mc310"], "塔吊"),
        (["汽车吊", "汽吊", "吊车", "25吨", "50吨", "crane"], "汽车吊"),
        (["履带吊"], "履带吊"),
        (["随车吊", "自卸吊"], "随车吊"),
        
        # 混凝土设备类
        (["泵车", "混凝土泵", "汽车泵", "pump"], "混凝土泵车"),
        (["地泵"], "地泵"),
        
        # 运输车辆类
        (["拉土车", "渣土车", "自卸车", "翻斗车", "dump"], "自卸车"),
        (["平板车", "板车", "拖板车", "拖车"], "平板车"),
        (["水车", "洒水车"], "洒水车"),
        (["叉车", "forklift"], "叉车"),
        (["高栏车"], "高栏车"),
        
        # 升降设备类
        (["升降车", "曲臂", "曲柄"], "升降车"),
        
        # 发电机类
        (["发电机", "generator"], "发电机"),
        
        # 其他设备
        (["平地机", "grader"], "平地机"),
        (["切断机", "剪断机"], "钢筋切断机"),
        (["破锤", "炮机"], "破碎锤"),
        (["抽粪车"], "抽粪车"),
    ]
    
    # 按规则匹配
    for keywords, standard_name in rules:
        for keyword in keywords:
            if keyword in raw_lower:
                return standard_name
    
    # 费用类项目统一归类
    if any(x in raw_lower for x in ["运费", "检测费", "安装", "拆卸", "进场", "出场", "搬迁", "移", "预付", "押金", "延时", "超时", "服务费"]):
        return "费用类"
    
    # 未匹配到返回None
    return None


def normalize_special_vehicle_kind(raw_name):
    """将“拉土车/随车吊+车牌号”等形式统一为基础名称。

    示例：
    - 拉土车82-2515 -> 拉土车
    - 拉土车 82-7041 -> 拉土车
    - 随车吊80-8816 -> 随车吊
    """
    if raw_name is None:
        return ""

    value = str(raw_name).strip()
    if not value or value == "nan":
        return value

    compact = re.sub(r'\s+', '', value)
    for base_name in ["拉土车", "随车吊"]:
        if compact.startswith(base_name):
            suffix = compact[len(base_name):]
            # 只要后缀出现数字（常见车牌/编号），就统一归一到基础名称
            if suffix and re.search(r'\d', suffix):
                return base_name
            # 后缀仅有分隔符/括号，也归一
            if suffix and re.fullmatch(r'[-_/()（）\[\]【】]+', suffix):
                return base_name
            if not suffix:
                return base_name

    return value

def parse_machine_mapping_txt(txt_path):
    """解析旧版TXT机具映射。"""
    machine_kind_map = {}
    model_map = {}
    current_kind = None

    with open(txt_path, "r", encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith("#"):
                continue

            if line.startswith("[") and line.endswith("]"):
                current_kind = line[1:-1].strip()
                continue

            if ":" not in line:
                continue

            standard, raw_list = line.split(":", 1)
            standard = standard.strip()
            raws = [r.strip() for r in raw_list.split(",") if r.strip()]

            if current_kind is None:
                for raw in raws:
                    machine_kind_map[raw] = standard
            else:
                for raw in raws:
                    model_map[(current_kind, raw)] = standard

    return machine_kind_map, model_map


def load_machine_mapping(json_path, legacy_txt_path=None):
    """加载机具种类和型号规格映射。优先JSON，不存在时自动从旧TXT迁移。

    JSON支持两种结构：
    1) 推荐分组：
       - kind_to_aliases: {标准机具种类: [原始机具别名...]}
       - model_to_aliases: {标准机具种类: {标准型号: [原始型号别名...]}}
    2) 兼容旧版：
       - machine_kind_map: {原始机具别名: 标准机具种类}
       - model_map: {标准机具种类: {原始型号别名: 标准型号}}
    返回统一为：
       machine_kind_map = {原始机具别名: 标准机具种类}
       model_map = {(标准机具种类, 原始型号别名): 标准型号}
    """
    machine_kind_map = {}
    model_map = {}

    if os.path.exists(json_path):
        try:
            with open(json_path, "r", encoding="utf-8") as f:
                data = json.load(f)

            # 机具种类：优先读取推荐分组结构
            grouped_kind_map = data.get("kind_to_aliases", {}) if isinstance(data, dict) else {}
            legacy_kind_map = data.get("machine_kind_map", {}) if isinstance(data, dict) else {}

            if isinstance(grouped_kind_map, dict) and grouped_kind_map:
                for std_kind, aliases in grouped_kind_map.items():
                    std_kind = str(std_kind).strip()
                    if not std_kind:
                        continue
                    if isinstance(aliases, list):
                        for alias in aliases:
                            alias = str(alias).strip()
                            if alias:
                                machine_kind_map[alias] = std_kind
                    elif isinstance(aliases, str):
                        for alias in re.split(r'[，,]', aliases):
                            alias = alias.strip()
                            if alias:
                                machine_kind_map[alias] = std_kind
            elif isinstance(legacy_kind_map, dict):
                machine_kind_map = {
                    str(raw).strip(): str(std).strip()
                    for raw, std in legacy_kind_map.items()
                    if str(raw).strip()
                }

            # 型号映射：优先读取推荐分组结构
            grouped_model_map = data.get("model_to_aliases", {}) if isinstance(data, dict) else {}
            legacy_model_map = data.get("model_map", {}) if isinstance(data, dict) else {}

            if isinstance(grouped_model_map, dict) and grouped_model_map:
                for kind, std_model_dict in grouped_model_map.items():
                    kind = str(kind).strip()
                    if not kind or not isinstance(std_model_dict, dict):
                        continue
                    for std_model, aliases in std_model_dict.items():
                        std_model = str(std_model).strip()
                        if not std_model:
                            continue
                        if isinstance(aliases, list):
                            for alias in aliases:
                                alias = str(alias).strip()
                                if alias:
                                    model_map[(kind, alias)] = std_model
                        elif isinstance(aliases, str):
                            for alias in re.split(r'[，,]', aliases):
                                alias = alias.strip()
                                if alias:
                                    model_map[(kind, alias)] = std_model
            elif isinstance(legacy_model_map, dict):
                for kind, model_dict in legacy_model_map.items():
                    kind = str(kind).strip()
                    if not kind or not isinstance(model_dict, dict):
                        continue
                    for raw_model, std_model in model_dict.items():
                        raw_model = str(raw_model).strip()
                        std_model = str(std_model).strip()
                        if raw_model:
                            model_map[(kind, raw_model)] = std_model

            return machine_kind_map, model_map
        except Exception as e:
            print(f"警告：读取机具映射JSON失败 {json_path}: {e}")

    if legacy_txt_path and os.path.exists(legacy_txt_path):
        machine_kind_map, model_map = parse_machine_mapping_txt(legacy_txt_path)
        save_machine_mapping(json_path, machine_kind_map, model_map)
        print(f"已将旧TXT机具映射迁移到JSON：{json_path}")
        return machine_kind_map, model_map

    save_machine_mapping(json_path, machine_kind_map, model_map)
    print(f"已创建机具映射JSON模板：{json_path}")
    return machine_kind_map, model_map

def save_project_mapping(json_path, project_mapping):
    """保存项目区域映射JSON（按区域分组）。"""
    try:
        normalized = {
            str(raw).strip(): str(std).strip()
            for raw, std in project_mapping.items()
            if str(raw).strip()
        }

        region_to_projects = {}
        for project, region in normalized.items():
            if not region:
                continue
            region_to_projects.setdefault(region, []).append(project)

        region_to_projects = {
            region: sorted(projects)
            for region, projects in sorted(region_to_projects.items(), key=lambda kv: kv[0])
        }

        payload = {
            "_comment": "项目→区域映射，推荐结构: {区域: [项目别名1, 项目别名2, ...]}",
            "region_to_projects": region_to_projects,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)
        print(f"已保存区域映射JSON到：{json_path}")
    except Exception as e:
        print(f"保存区域映射JSON失败：{e}")

def save_machine_mapping(json_path, kind_mapping, model_mapping):
    """保存机具种类和型号规格映射JSON（按标准项分组）。"""
    try:
        normalized_kind_map = {
            str(raw).strip(): str(std).strip()
            for raw, std in kind_mapping.items()
            if str(raw).strip()
        }

        kind_to_aliases = {}
        for raw_alias, std_kind in normalized_kind_map.items():
            if not std_kind:
                continue
            kind_to_aliases.setdefault(std_kind, []).append(raw_alias)

        kind_to_aliases = {
            std_kind: sorted(aliases)
            for std_kind, aliases in sorted(kind_to_aliases.items(), key=lambda kv: kv[0])
        }

        model_to_aliases = {}
        for (kind, raw_model), std_model in model_mapping.items():
            kind = str(kind).strip()
            raw_model = str(raw_model).strip()
            std_model = str(std_model).strip()
            if not kind or not raw_model:
                continue
            model_to_aliases.setdefault(kind, {}).setdefault(std_model, []).append(raw_model)

        model_to_aliases = {
            kind: {
                std_model: sorted(alias_list)
                for std_model, alias_list in sorted(std_model_dict.items(), key=lambda kv: kv[0])
            }
            for kind, std_model_dict in sorted(model_to_aliases.items(), key=lambda kv: kv[0])
        }

        payload = {
            "_comment": "机具映射，推荐结构：kind_to_aliases为{标准机具种类:[原始机具别名...]}; model_to_aliases为{标准机具种类:{标准型号:[原始型号别名...]}}",
            "kind_to_aliases": kind_to_aliases,
            "model_to_aliases": model_to_aliases,
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2)

        print(f"已保存机具映射JSON到：{json_path}")
    except Exception as e:
        print(f"保存机具映射JSON失败：{e}")


def guess_region_from_sheet_name(sheet_name, mapping):
    """从 sheet 名找最长的匹配项目作为区域"""
    for raw_project in sorted(mapping.keys(), key=len, reverse=True):
        if raw_project in sheet_name:
            return mapping[raw_project]
    return "未知区域"   # 没匹配到就给个默认值

def normalize_column_name(col_name):
    """标准化列名：去除空格、换行符、全角符号等"""
    if not col_name:
        return ""
    col_name = str(col_name)
    # 去除所有空白字符（空格、换行、制表符等）
    col_name = re.sub(r'\s+', '', col_name)
    # 替换全角符号为半角
    replacements = {
        '（': '(', '）': ')', '：': ':', '％': '%',
        '，': ',', '＋': '+', '／': '/'
    }
    for old, new in replacements.items():
        col_name = col_name.replace(old, new)
    return col_name

def find_matching_column(df_columns, target_patterns):
    """
    在DataFrame列中查找匹配目标模式的列名
    target_patterns: 列表，包含可能的列名关键词
    返回：匹配到的列名，如果没找到返回None
    """
    normalized_cols = {normalize_column_name(col): col for col in df_columns}
    
    for pattern in target_patterns:
        pattern_norm = normalize_column_name(pattern)
        # 精确匹配
        if pattern_norm in normalized_cols:
            return normalized_cols[pattern_norm]
        # 包含匹配
        for norm_col, orig_col in normalized_cols.items():
            if pattern_norm in norm_col or norm_col in pattern_norm:
                return orig_col
    return None

def find_header_row(ws):
    """在工作表中查找包含目标列名的行"""
    for row_idx in range(1, min(20, ws.max_row + 1)):  # 搜索前20行
        row_values = []
        for cell in ws[row_idx]:
            val = str(cell.value) if cell.value else ""
            # 清理换行符和多余空格
            val = re.sub(r'\s+', '', val)
            row_values.append(val)
        
        # 检查是否包含关键列名
        if any("序号" in val for val in row_values) and any("机具种类" in val for val in row_values):
            return row_idx
    return None

def safe_to_float(value):
    """安全地将值转换为浮点数"""
    if value is None or value == "" or str(value).strip() == "":
        return 0.0
    try:
        # 去除逗号和空格
        val_str = str(value).replace(",", "").replace(" ", "").strip()
        if val_str == "" or val_str == "nan":
            return 0.0
        return float(val_str)
    except:
        return 0.0


def validate_overtime_fee(row):
    """验证加班费是否正确
    正确公式: 加班费 = 台班费_泰铢 / 8 * 加班小时
    返回: (是否正确, 实际值, 计算值)
    """
    actual = safe_to_float(row.get('加班费', 0))
    unit_price = safe_to_float(row.get('台班费_泰铢', 0))
    overtime_hours = safe_to_float(row.get('加班小时', 0))
    
    # 如果没有加班小时，加班费应该为0
    if overtime_hours == 0:
        calculated = 0.0
    else:
        calculated = unit_price / 8 * overtime_hours
    
    # 允许0.01的误差（因为四舍五入）
    is_correct = abs(actual - calculated) < 0.01
    
    return is_correct, actual, calculated

def validate_settlement_amount(row):
    """验证结算费用是否正确
    正确公式（两种方式之一即可）：
    方式1: 结算费用_泰铢 = 台班费_泰铢 * 本次结算台班 + 加班费 + 进出场费运费 + 移动费拆装费
    方式2: 结算费用_泰铢 = 台班费_泰铢 + 台班费_泰铢 / 8 * 加班小时 + 进出场费运费 + 移动费拆装费
    返回: (是否正确, 实际值, 计算值1, 计算值2)
    """
    actual = safe_to_float(row.get('结算费用_泰铢', 0))
    unit_price = safe_to_float(row.get('台班费_泰铢', 0))
    quantity = safe_to_float(row.get('本次结算台班', 0))
    overtime_fee = safe_to_float(row.get('加班费', 0))
    overtime_hours = safe_to_float(row.get('加班小时', 0))
    entry_exit_fee = safe_to_float(row.get('进出场费运费', 0))
    move_install_fee = safe_to_float(row.get('移动费拆装费', 0))
    
    # 方式1: 台班费 * 结算台班 + 加班费 + 进出场费 + 移动费
    calculated1 = unit_price * quantity + overtime_fee + entry_exit_fee + move_install_fee
    
    # 方式2: 台班费 + (台班费/8 * 加班小时) + 进出场费 + 移动费
    calculated2 = unit_price + (unit_price / 8 * overtime_hours) + entry_exit_fee + move_install_fee
    
    # 允许0.01的误差（因为四舍五入），两种方式有一种正确即可
    is_correct = abs(actual - calculated1) < 0.01 or abs(actual - calculated2) < 0.01
    
    return is_correct, actual, calculated1, calculated2

def validate_total_amount(row):
    """验证本次费用合计是否正确
    正确公式: 本次费用合计 = 结算费用_泰铢 + 增值税7% - 机具租赁预扣税3% - 机具租赁预扣税5%
    返回: (是否正确, 实际值, 计算值)
    """
    actual = safe_to_float(row.get('本次费用合计', 0))
    settlement = safe_to_float(row.get('结算费用_泰铢', 0))
    vat = safe_to_float(row.get('增值税7%', 0))
    tax3 = safe_to_float(row.get('机具租赁预扣税3%', 0))
    tax5 = safe_to_float(row.get('机具租赁预扣税5%', 0))
    
    calculated = settlement + vat - tax3 - tax5
    
    # 允许0.01的误差（因为四舍五入）
    is_correct = abs(actual - calculated) < 0.01
    
    return is_correct, actual, calculated

def detect_and_fix_swapped_values(df):
    """检测并修正台班费和结算台班填反的情况
    判断逻辑：
    1. 台班费通常 > 100（单价），结算台班通常 < 100（数量）
    2. 如果台班费 < 100 且 结算台班 > 1000，很可能填反了
    返回: (修正后的df, 被修正的行索引列表)
    """
    swapped_indices = []
    
    for idx in range(len(df)):
        unit_price = safe_to_float(df.iloc[idx].get('台班费_泰铢', 0))
        quantity = safe_to_float(df.iloc[idx].get('本次结算台班', 0))
        
        # 检测填反的情况：
        # 1. 台班费很小（< 50）且结算台班很大（> 500）
        # 2. 或者台班费在合理台班数范围（1-100）且结算台班像价格（> 1000）
        is_swapped = False
        
        if unit_price > 0 and quantity > 0:
            # 情况1：台班费 < 50 且 结算台班 > 500
            if unit_price < 50 and quantity > 500:
                is_swapped = True
            # 情况2：台班费像数量（1-100）且结算台班像价格（> 1000）
            elif unit_price <= 100 and quantity > 1000:
                is_swapped = True
        
        if is_swapped:
            # 交换两列的值
            df.at[idx, '台班费_泰铢'] = quantity
            df.at[idx, '本次结算台班'] = unit_price
            swapped_indices.append(idx)
            print(f"    ⚠️  第 {idx+2} 行数据已自动交换：台班费 {unit_price:.2f} ↔ 结算台班 {quantity:.2f}")
    
    return df, swapped_indices

def detect_and_fix_supplier_month_swap(df):
    """检测并修正出租单位和月份填反的情况
    判断逻辑：
    1. 月份通常是数字格式（如"1月"、"2月"、"202501"、"8月份"等）
    2. 出租单位通常是公司名称（包含中文、英文、特殊字符等）
    3. 如果出租单位像月份，月份像公司名，则可能填反
    返回: (修正后的df, 被修正的行索引列表)
    """
    swapped_month_indices = []
    
    # 月份模式：纯数字、包含"月"字、YYYYMM格式等
    month_patterns = [
        r'^\d{1,2}$',              # 1-12
        r'^\d{1,2}月$',            # 1月-12月
        r'^\d{1,2}月份$',          # 1月份-12月份 ← 新增
        r'^\d{4}年\d{1,2}月$',     # 2024年1月
        r'^\d{4}年\d{1,2}月份$',   # 2024年1月份 ← 新增
        r'^\d{6}$',                # 202401
        r'^\d{4}-\d{1,2}$',        # 2024-01
        r'^\d{4}/\d{1,2}$',        # 2024/01
        r'^[一二三四五六七八九十]{1,2}月$',      # 一月-十二月（中文数字）
        r'^[一二三四五六七八九十]{1,2}月份$',    # 一月份-十二月份
    ]
    
    for idx in range(len(df)):
        supplier = str(df.iloc[idx].get('出租单位', '')).strip()
        month = str(df.iloc[idx].get('月份', '')).strip()
        
        if not supplier or supplier == 'nan':
            continue
            
        # 检查出租单位是否像月份
        is_supplier_like_month = False
        for pattern in month_patterns:
            if re.match(pattern, supplier):
                is_supplier_like_month = True
                break
        
        # 如果出租单位像月份，且月份不像月份（或为空），则可能填反
        if is_supplier_like_month:
            # 检查月份是否不像月份（更像公司名）
            is_month_like_supplier = False
            if month and month != 'nan':
                # 月份包含中文、英文字母（除了"月"、"份"、"年"）等，可能是公司名
                month_chars = set('月份年')
                if any(c.isalpha() and c not in month_chars for c in month) or len(month) > 15:
                    is_month_like_supplier = True
            
            # 如果确定填反了，或月份为空
            if is_month_like_supplier or not month or month == 'nan':
                # 交换两列的值
                df.at[idx, '出租单位'], df.at[idx, '月份'] = month if month and month != 'nan' else '', supplier
                swapped_month_indices.append(idx)
                print(f"    ⚠️  第 {idx+2} 行数据已自动交换：出租单位「{supplier}」↔ 月份「{month}」")
    
    return df, swapped_month_indices

def extract_project_info(ws):
    """从第二行提取项目名称和开始月份"""
    project_name = ""
    start_month = ""
    
    # 搜索前10行中包含"项目名称"和"开始月份"的单元格
    for row_idx in range(1, min(11, ws.max_row + 1)):
        for cell in ws[row_idx]:
            cell_value = str(cell.value).strip() if cell.value else ""
            
            # 查找"项目名称：XXX"
            if "项目名称" in cell_value and "：" in cell_value:
                project_name = cell_value.split("：", 1)[1].strip()
            elif "项目名称" in cell_value:
                # 如果冒号在相邻单元格
                next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                if next_cell.value:
                    project_name = str(next_cell.value).strip()
            
            # 查找"开始月份：XXXX"
            if "开始月份" in cell_value and "：" in cell_value:
                start_month = cell_value.split("：", 1)[1].strip()
            elif "开始月份" in cell_value:
                # 如果冒号在相邻单元格
                next_cell = ws.cell(row=cell.row, column=cell.column + 1)
                if next_cell.value:
                    start_month = str(next_cell.value).strip()
    
    return project_name, start_month

def main(excel_path_override=None):
    if excel_path_override:
        excel_path = Path(excel_path_override).expanduser().resolve()
        print(f"批量模式-当前处理文件: {excel_path.name}")
    elif EXCEL_FILE_OVERRIDE and str(EXCEL_FILE_OVERRIDE).strip():
        excel_path = Path(EXCEL_FILE_OVERRIDE).expanduser().resolve()
        print(f"使用手工指定文件: {excel_path}")
    else:
        picked = pick_latest_standardized_file(STANDARDIZED_INPUT_DIR)
        if not picked:
            print(f"找不到标准化文件：{STANDARDIZED_INPUT_DIR}")
            return
        excel_path = picked.resolve()
        print(f"自动选择最新标准化文件: {excel_path.name}")

    if not excel_path.exists():
        print(f"找不到文件：{excel_path}")
        return

    # 1. 加载映射表和项目配置
    project_mapping = load_project_mapping(PROJECT_MAPPING_JSON, PROJECT_MAPPING_TXT_LEGACY)          # 项目→区域
    kind_mapping, model_mapping = load_machine_mapping(MACHINE_MAPPING_JSON, MACHINE_MAPPING_TXT_LEGACY)  # 机具种类、型号
    
    # 加载项目配置：优先使用统一的项目编码映射表
    project_config = load_unified_project_mapping()
    
    # 如果统一映射表为空，回退到旧配置
    if not project_config:
        print("统一映射表加载失败或为空，回退到旧项目配置")
        project_config = load_project_config()

    # 初始化统一项目匹配库（project_name_matcher + project_name_match_db.json）
    matcher_info_df = pd.DataFrame({
        '项目名称': [p.get('project_name', '') for p in project_config],
        '项目编码': [p.get('project_code', '') for p in project_config],
        '备注': [p.get('status', p.get('target_status', '需手动核对')) for p in project_config],
    })
    standard_projects = build_standard_projects_from_dataframe(matcher_info_df)
    match_db = load_or_init_db(PROJECT_MATCH_DB_PATH, standard_projects)
    print(
        f"✅ 共享匹配库已加载：标准项目 {len(standard_projects)} 个，"
        f"别名 {len(match_db.get('alias_to_standard', {}))} 个，"
        f"未匹配留痕 {len(match_db.get('unmatched_aliases', []))} 个"
    )
    
    print("项目匹配已切换为统一匹配库模式（不再使用本地缓存/重新匹配弹窗）")

    # 用于收集新发现的项目、机具种类、型号
    new_projects = set()
    new_kinds = set()
    new_models = set()
    
    all_data = []
    report_data = []  # 用于生成报告
    
    # 第一遍扫描：收集所有需要匹配的项目
    sheet_match_results = {}
    unmatched_aliases = []
    alias_to_sheet_names = {}
    # 2. 读取所有 sheet
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    print("\n=== 第一遍扫描：提取项目名称与匹配 ===")
    for sheet_name in wb.sheetnames:
        if sheet_name == SUMMARY_SHEET_NAME:
            continue
        
        ws = wb[sheet_name]
        project_name_from_content, _ = extract_project_info(ws)
        
        # 如果没有从sheet内容提取到项目名，尝试从sheet名称提取
        if not project_name_from_content:
            project_name_from_content = extract_project_name_from_sheet(sheet_name)
        
        print(f"  Sheet: {sheet_name} (内容提取: '{project_name_from_content}')")
        
        # 使用统一matcher：先尝试Sheet名，再尝试内容提取的项目名
        candidate_aliases = [sheet_name]
        if project_name_from_content and project_name_from_content != sheet_name:
            candidate_aliases.append(project_name_from_content)

        matched_project = None
        for alias in candidate_aliases:
            hit = match_project(str(alias).strip(), match_db)
            if hit:
                matched_project = {
                    'project_code': hit.get('project_code', ''),
                    'project_name': hit.get('project_name', ''),
                    'status': hit.get('target_status', ''),
                }
                break

        if matched_project:
            # 保存到本轮匹配结果（由共享匹配库驱动）
            sheet_match_results[sheet_name] = {
                'project_code': matched_project['project_code'],
                'project_name': matched_project['project_name'],
                'status': matched_project.get('status', '')
            }
            print(f"    ✓ 匹配成功: {matched_project['project_code']} - {matched_project['project_name']}")
        else:
            # 未匹配，收集alias用于统一matcher弹窗
            for alias in candidate_aliases:
                alias = str(alias).strip()
                if not alias:
                    continue
                alias_to_sheet_names.setdefault(alias, set()).add(sheet_name)
                if is_unmatched_alias(match_db, alias):
                    continue
                if alias not in unmatched_aliases:
                    unmatched_aliases.append(alias)
            print(f"    ✗ 未找到匹配")
    
    # 如果有未匹配别名，使用统一matcher弹窗
    if unmatched_aliases:
        print(f"\n发现 {len(unmatched_aliases)} 个未匹配项目别名，打开统一项目匹配窗口...")
        prompt_unmatched_project_mappings(unmatched_aliases, match_db, window_title='机具使用费-手动项目匹配')
        save_db(PROJECT_MATCH_DB_PATH, match_db)

        print("\n=== 统一匹配结果回填 ===")
        for alias, sheet_names in alias_to_sheet_names.items():
            hit = match_project(alias, match_db)
            if not hit:
                continue
            for sheet_name in sheet_names:
                sheet_match_results[sheet_name] = {
                    'project_code': hit.get('project_code', ''),
                    'project_name': hit.get('project_name', ''),
                    'status': hit.get('target_status', ''),
                }
                print(f"  '{sheet_name}' → [{hit.get('project_code', '')}] {hit.get('project_name', '')}")

    print(f"统一匹配结果数量：{len(sheet_match_results)}")
    
    # 重新打开工作簿开始第二遍处理
    wb.close()
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    
    for sheet_name in wb.sheetnames:
        if sheet_name == SUMMARY_SHEET_NAME:
            continue
        
        print(f"\n正在处理 sheet：{sheet_name}")
        ws = wb[sheet_name]
        
        # 提取项目信息
        project_name, start_month = extract_project_info(ws)
        print(f"  项目名称: {project_name if project_name else '未找到'}")
        print(f"  开始月份: {start_month if start_month else '未找到'}")
        
        # 查找表头行
        header_row = find_header_row(ws)
        if not header_row:
            print(f"  ⚠️  未找到表头行，跳过此sheet")
            report_data.append({
                "sheet名称": sheet_name,
                "项目名称": sheet_name,
                "项目编码": '',
                "开始月份": start_month,
                "状态": "失败：未找到表头",
                "数据行数": 0,
                "加班费错误行数": 0,
                "加班费错误率(%)": 0,
                "加班费差异金额(泰铢)": 0,
                "结算费用错误行数": 0,
                "结算费用错误率(%)": 0,
                "结算费用差异金额(泰铢)": 0,
                "本次费用合计错误行数": 0,
                "本次费用合计错误率(%)": 0,
                "本次费用合计差异金额(泰铢)": 0,
                "费用合计(泰铢)": 0
            })
            continue
        
        print(f"  表头位于第 {header_row} 行")
        
        # 使用pandas读取，从表头行开始
        # 先不指定dtype让pandas读取原始数据类型，这样有助于公式计算
        df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header_row-1, engine='openpyxl')
        
        # 清理列名：去除换行符和多余空格
        df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
        
        # 定义标准列名及其可能的变体（用于模糊匹配）
        column_mapping = {
            "序号": ["序号"],
            "机具种类": ["机具种类"],
            "型号规格": ["型号规格"],
            "台班费_泰铢": ["台班费", "台班费泰铢", "台班费(泰铢)", "台班费 泰铢"],
            "本次结算台班": ["本次结算台班", "结算台班", "本次结算台班（小时、天、月)", "本次结算台班(小时、天、月)", "本次结算台班 （小时、天、月)"],
            "加班小时": ["加班", "加班(小时)", "加班（小时）", "加班 (小时)"],
            "加班费": ["加班费", "加班费（小时）", "加班费(小时)", "加班费 （小时）", "加班费 "],
            "进出场费运费": ["进出场费", "运费", "进出场费/运费", "运费＋拆装费", "运费+拆装费"],
            "移动费拆装费": ["移动费", "拆装费", "移动费＋拆装费", "移动费+拆装费"],
            "结算费用_泰铢": ["结算费用", "本次结算费用", "结算费用(泰铢)", "本次结算费用(泰铢)", "结算费用 (泰铢)", "本次结算费用 (泰铢)"],
            "增值税7%": ["增值税", "增值税7%", "增值税7％"],
            "机具租赁预扣税3%": ["预扣税", "机具租赁预扣税", "机具租赁预扣税3%", "机具租赁预扣税3％", "机具租赁预 扣税"],
            "机具租赁预扣税5%": ["机具租赁预扣税5%", "机具租赁预扣税5％"],
            "本次费用合计": ["本次费用合计", "本次费用合计（泰铢）", "本次费用合计(泰铢)", "本次费用合计 （泰铢）", "费用合计", "费用 合计"],
            "使用日期": ["使用日期", "使用时间"],
            "出租单位": ["出租单位"],
            "月份": ["月份"],
            "备注": ["备注", "项目说明"]
        }
        
        # 创建新的DataFrame，使用标准列名
        new_df = pd.DataFrame()
        unmatched_cols = []
        
        for std_col, patterns in column_mapping.items():
            matched_col = find_matching_column(df.columns, patterns)
            if matched_col:
                new_df[std_col] = df[matched_col]
                print(f"    '{matched_col}' -> '{std_col}'")
            else:
                new_df[std_col] = ""
                unmatched_cols.append(std_col)
        
        if unmatched_cols:
            print(f"  ⚠️  未匹配到的列: {', '.join(unmatched_cols)}")
        
        df = new_df
        
        # 重新计算费用相关列（处理公式未缓存的情况）
        # Excel公式结构：
        # 1. 结算费用 = 台班费 × 台班数 + 加班费 + 运费
        # 2. 增值税7% = 结算费用 × 0.07
        # 3. 预扣税3% = 结算费用 × 0.03
        # 4. 本次费用合计 = 结算费用 + 增值税 - 预扣税3% - 预扣税5%
        print(f"  正在检查并重新计算费用...")
        
        settlement_recalc = 0
        vat_recalc = 0
        tax_recalc = 0
        fee_recalc = 0
        
        for idx in range(len(df)):
            # 第一步：重新计算"结算费用"
            settlement = safe_to_float(df.iloc[idx]["结算费用_泰铢"])
            
            if settlement == 0:
                # 尝试从基础数据计算
                unit_price = safe_to_float(df.iloc[idx]["台班费_泰铢"])
                quantity = safe_to_float(df.iloc[idx]["本次结算台班"])
                overtime_fee = safe_to_float(df.iloc[idx]["加班费"])
                entry_fee = safe_to_float(df.iloc[idx]["进出场费运费"])
                move_fee = safe_to_float(df.iloc[idx]["移动费拆装费"])
                
                # 计算结算费用 = 台班费×台班数 + 加班费 + 各种运费
                calc_settlement = unit_price * quantity + overtime_fee + entry_fee + move_fee
                
                if calc_settlement != 0:
                    settlement = calc_settlement
                    df.at[idx, "结算费用_泰铢"] = str(settlement)
                    settlement_recalc += 1
            
            # 第二步：重新计算"增值税7%"
            vat = safe_to_float(df.iloc[idx]["增值税7%"])
            if vat == 0 and settlement != 0:
                calc_vat = settlement * 0.07
                df.at[idx, "增值税7%"] = str(calc_vat)
                vat = calc_vat
                vat_recalc += 1
            
            # 第三步：重新计算"预扣税3%"
            tax3 = safe_to_float(df.iloc[idx]["机具租赁预扣税3%"])
            if tax3 == 0 and settlement != 0:
                calc_tax3 = settlement * 0.03
                df.at[idx, "机具租赁预扣税3%"] = str(calc_tax3)
                tax3 = calc_tax3
                tax_recalc += 1
            
            # 第四步：重新计算"本次费用合计"
            current_fee = safe_to_float(df.iloc[idx]["本次费用合计"])
            if current_fee == 0 and settlement != 0:
                tax5 = safe_to_float(df.iloc[idx]["机具租赁预扣税5%"])
                calculated_fee = settlement + vat - tax3 - tax5
                df.at[idx, "本次费用合计"] = str(calculated_fee)
                fee_recalc += 1
        
        if settlement_recalc > 0 or vat_recalc > 0 or tax_recalc > 0 or fee_recalc > 0:
            print(f"  ✓ 重新计算: 结算费用{settlement_recalc}行, 增值税{vat_recalc}行, 预扣税{tax_recalc}行, 费用合计{fee_recalc}行")
        
        # 过滤空行：保留序号有效的行，或保留序号为空但有关键费用数据的行
        key_cols = ["台班费_泰铢", "本次结算台班", "加班费", "结算费用_泰铢", "本次费用合计"]
        
        # 先保留序号有效的行
        has_seq = df["序号"].notna() & (df["序号"].astype(str).str.strip() != "")
        
        # 对于序号为空的行，检查是否有关键列有非空数据
        has_data = pd.Series(False, index=df.index)
        for col in key_cols:
            if col in df.columns:
                has_data |= (df[col].notna() & (df[col].astype(str).str.strip() != ""))
        
        # 保留有序号或有数据的行
        df = df[has_seq | has_data].copy()  # 使用.copy()避免SettingWithCopyWarning
        
        if len(df) == 0:
            print(f"  ⚠️  没有有效数据行")
            report_data.append({
                "sheet名称": sheet_name,
                "项目名称": sheet_name,
                "项目编码": '',
                "开始月份": start_month,
                "状态": "失败：无有效数据",
                "数据行数": 0,
                "加班费错误行数": 0,
                "加班费错误率(%)": 0,
                "加班费差异金额(泰铢)": 0,
                "结算费用错误行数": 0,
                "结算费用错误率(%)": 0,
                "结算费用差异金额(泰铢)": 0,
                "本次费用合计错误行数": 0,
                "本次费用合计错误率(%)": 0,
                "本次费用合计差异金额(泰铢)": 0,
                "费用合计(泰铢)": 0
            })
            continue
        
        print(f"  有效数据行数: {len(df)}")
        
        # 从统一匹配结果获取项目编码和项目名称
        matched_info = sheet_match_results.get(sheet_name)
        if matched_info and isinstance(matched_info, dict):
            project_code = matched_info.get('project_code', '')
            matched_project_name = matched_info.get('project_name', '')
            project_status = matched_info.get('status', '')
            print(f"  匹配结果: [{project_code}] {matched_project_name} ({project_status})")
        else:
            project_code = ''
            matched_project_name = ''
            project_status = ''
            print(f"  ⚠️  未找到匹配 (Sheet: '{sheet_name}')")
        
        # 汇总统一使用匹配后的项目名称；未匹配则回退到原始sheet名称
        final_project_name = matched_project_name if matched_project_name else sheet_name
        
        # 基础字段只保留：sheet名称 / 项目编码 / 项目名称
        df["sheet名称"] = sheet_name
        df["项目编码"] = project_code
        df["项目名称"] = final_project_name
        df["开始月份"] = start_month
        
        # 增加区域并收集新项目
        region = guess_region_from_sheet_name(sheet_name, project_mapping)
        if region == "未知区域":
            new_projects.add(sheet_name)
        df["区域"] = region
        
        # 映射标准机具种类并收集新种类
        def map_kind(x):
            # 先做专项归一：拉土车/随车吊 + 车牌号 => 基础名称
            val = normalize_special_vehicle_kind(x)
            if val and val != "nan" and val not in kind_mapping:
                new_kinds.add(val)
            return kind_mapping.get(val, val)
        
        df["标准机具种类"] = df["机具种类"].apply(map_kind)
        
        # 映射标准型号规格并收集新型号（基于标准机具种类）
        def map_model(row):
            std_kind = row["标准机具种类"]
            raw_model = str(row["型号规格"]).strip()
            
            if not raw_model or raw_model == "nan":
                return raw_model
            
            # 使用(标准机具种类, 原始型号)作为键
            key = (std_kind, raw_model)
            if key in model_mapping:
                return model_mapping[key]
            else:
                # 收集新型号（记录在哪个机具种类下）
                new_models.add((std_kind, raw_model))
                return raw_model
        
        df["标准型号规格"] = df.apply(map_model, axis=1)
        
        # 计算本sheet的费用合计
        sheet_total = df["本次费用合计"].apply(safe_to_float).sum()
        
        all_data.append(df)
        
        # 计算各类费用
        overtime_fee_total = df["加班费"].apply(safe_to_float).sum()
        settlement_total = df["结算费用_泰铢"].apply(safe_to_float).sum()
        
        # 统计有数据的行数
        overtime_count = (df["加班费"].apply(safe_to_float) != 0).sum()
        settlement_count = (df["结算费用_泰铢"].apply(safe_to_float) != 0).sum()
        total_count = (df["本次费用合计"].apply(safe_to_float) != 0).sum()
        
        # 统计验算错误的行数和金额差异
        overtime_error_rows = 0
        settlement_error_rows = 0
        total_error_rows = 0
        overtime_diff_sum = 0.0
        settlement_diff_sum = 0.0
        total_diff_sum = 0.0
        
        for idx in range(len(df)):
            row = df.iloc[idx]
            
            # 检查加班费错误
            is_correct, actual, calculated = validate_overtime_fee(row)
            if not is_correct and actual != 0:
                overtime_error_rows += 1
                overtime_diff_sum += abs(actual - calculated)
            
            # 检查结算费用错误
            is_correct, actual, calc1, calc2 = validate_settlement_amount(row)
            if not is_correct and actual != 0:
                settlement_error_rows += 1
                # 选择差异较小的计算方式
                diff = min(abs(actual - calc1), abs(actual - calc2))
                settlement_diff_sum += diff
            
            # 检查本次费用合计错误
            is_correct, actual, calculated = validate_total_amount(row)
            if not is_correct:
                total_error_rows += 1
                total_diff_sum += abs(actual - calculated)
        
        # 计算错误率
        data_rows = len(df)
        overtime_error_rate = (overtime_error_rows / data_rows * 100) if data_rows > 0 else 0
        settlement_error_rate = (settlement_error_rows / data_rows * 100) if data_rows > 0 else 0
        total_error_rate = (total_error_rows / data_rows * 100) if data_rows > 0 else 0
        
        # 添加报告数据
        report_data.append({
            "sheet名称": sheet_name,
            "项目名称": final_project_name,
            "项目编码": project_code,
            "开始月份": start_month,
            "状态": "成功",
            "数据行数": len(df),
            "区域": region,
            "机具种类数": df["标准机具种类"].nunique(),
            "型号规格数": df["标准型号规格"].nunique(),
            "加班费错误行数": overtime_error_rows,
            "加班费错误率(%)": round(overtime_error_rate, 2),
            "加班费差异金额(泰铢)": round(overtime_diff_sum, 2),
            "结算费用错误行数": settlement_error_rows,
            "结算费用错误率(%)": round(settlement_error_rate, 2),
            "结算费用差异金额(泰铢)": round(settlement_diff_sum, 2),
            "本次费用合计错误行数": total_error_rows,
            "本次费用合计错误率(%)": round(total_error_rate, 2),
            "本次费用合计差异金额(泰铢)": round(total_diff_sum, 2),
            "加班费行数": overtime_count,
            "加班费合计(泰铢)": round(overtime_fee_total, 2),
            "结算费用行数": settlement_count,
            "结算费用合计(泰铢)": round(settlement_total, 2),
            "本次费用行数": total_count,
            "费用合计(泰铢)": round(sheet_total, 2)
        })
    
    wb.close()

    # 3. 处理新发现的项目、机具种类、型号
    if new_projects or new_kinds or new_models:
        print(f"\n发现新数据：")
        if new_projects:
            print(f"  新项目（{len(new_projects)}个）：") 
            for p in sorted(new_projects):
                print(f"    {p}")
                # 自动添加到映射表（标准名=原始名）
                project_mapping[p] = p
        if new_kinds:
            print(f"  新机具种类（{len(new_kinds)}个）：")
            # 应用智能映射规则
            smart_mapped = 0
            for k in sorted(new_kinds):
                smart_result = apply_smart_mapping_rules(k)
                if smart_result:
                    kind_mapping[k] = smart_result
                    print(f"    {k} -> {smart_result} (智能映射)")
                    smart_mapped += 1
                else:
                    kind_mapping[k] = k
                    print(f"    {k} (未匹配规则，保持原名)")
            if smart_mapped > 0:
                print(f"  ✅ 智能映射成功 {smart_mapped}/{len(new_kinds)} 个机具种类")
                print(f"  💡 已自动更新机具映射表，建议人工复核调整")
        if new_models:
            print(f"  新型号规格（{len(new_models)}个）：")
            # 按机具种类分组显示
            models_by_kind = {}
            for std_kind, raw_model in sorted(new_models):
                if std_kind not in models_by_kind:
                    models_by_kind[std_kind] = []
                models_by_kind[std_kind].append(raw_model)
            
            for kind in sorted(models_by_kind.keys()):
                print(f"    [{kind}]")
                for model in models_by_kind[kind]:
                    print(f"      {model}")
        
        # 保存更新后的映射表到两个独立文件
        if new_projects:
            save_project_mapping(PROJECT_MAPPING_JSON, project_mapping)
        if new_kinds or new_models:
            # 将新型号添加到model_mapping
            for std_kind, raw_model in new_models:
                model_mapping[(std_kind, raw_model)] = raw_model
            save_machine_mapping(MACHINE_MAPPING_JSON, kind_mapping, model_mapping)
        print(f"已自动补充新数据到映射表，请人工检查/修改！\n")
    
    # 4. 合并所有 sheet
    if not all_data:
        print("\n没有读取到任何数据！")
        return
    
    print(f"\n合并 {len(all_data)} 个sheet的数据...")
    # 打印第一个DataFrame的列，用于调试
    if len(all_data) > 0:
        print(f"第一个sheet的列: {list(all_data[0].columns)}")
        
    summary_df = pd.concat(all_data, ignore_index=True)
    print(f"合并后的列: {list(summary_df.columns)}")
    
    # 清理数据格式
    print(f"\n清理数据格式...")
    
    # 1. 清理型号规格：删除换行符、制表符等空白字符
    def clean_model_spec(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""
        # 将所有换行符、制表符等替换为空格，然后合并多余空格
        cleaned = re.sub(r'[\n\r\t]+', ' ', str(value))
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        return cleaned
    
    summary_df["型号规格"] = summary_df["型号规格"].apply(clean_model_spec)
    summary_df["标准型号规格"] = summary_df["标准型号规格"].apply(clean_model_spec)
    print(f"  ✅ 已清理型号规格列的换行符")
    
    # 2. 清理备注：将连续空格（3个以上）替换为分号
    def clean_remark(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""
        # 将3个以上的连续空格替换为分号
        cleaned = re.sub(r' {3,}', '; ', str(value))
        # 清理换行符为分号
        cleaned = re.sub(r'[\n\r]+', '; ', cleaned)
        # 合并多个分号
        cleaned = re.sub(r';\s*;+', ';', cleaned)
        return cleaned.strip()
    
    summary_df["备注"] = summary_df["备注"].apply(clean_remark)
    print(f"  ✅ 已清理备注列的空格")

    # 2.1 清理使用日期：将连续3个及以上空白替换为分号，避免多个日期黏连
    def clean_use_date(value):
        if not value or str(value).strip() == "" or str(value) == "nan":
            return ""

        cleaned = str(value)
        # 先将换行、制表符标准化为空格
        cleaned = re.sub(r'[\n\r\t]+', ' ', cleaned)
        # 连续3个及以上空白替换为中文分号
        cleaned = re.sub(r'\s{3,}', '；', cleaned)
        # 规整分号周围空白并去重
        cleaned = re.sub(r'\s*；\s*', '；', cleaned)
        cleaned = re.sub(r'；{2,}', '；', cleaned)
        return cleaned.strip('； ').strip()

    summary_df["使用日期"] = summary_df["使用日期"].apply(clean_use_date)
    print(f"  ✅ 已清理使用日期列的连续空白")
    
    # 3. 从备注列智能解析月租金额和加班费单价
    def extract_monthly_rent(remark):
        """从备注中提取月租金额（单位：泰铢）"""
        if not remark or str(remark).strip() == "" or str(remark) == "nan":
            return ""
        
        remark_str = str(remark)
        
        # 匹配模式（按优先级）
        patterns = [
            # 模式1: 月租 250000/月
            r'月租[\s]*([\d,.]+)[\s]*/[\s]*月',
            # 模式2: 月租250000/月含油
            r'月租[\s]*([\d,.]+)[\s]*/[\s]*月',
            # 模式3: 月租/70000
            r'月租[\s]*/[\s]*([\d,.]+)',
            # 模式4: 月租125000
            r'月租[\s]*([\d,.]+)(?![\d/天小时])',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, remark_str, re.IGNORECASE)
            if match:
                amount_str = match.group(1)
                # 移除逗号
                amount_str = amount_str.replace(',', '')
                try:
                    amount = float(amount_str)
                    # 验证金额合理性（通常月租在几千到几百万泰铢之间）
                    if 1000 <= amount <= 10000000:
                        return amount
                except ValueError:
                    continue
        
        return ""
    
    def extract_overtime_rate(remark):
        """从备注中提取加班费单价（单位：泰铢/小时）"""
        if not remark or str(remark).strip() == "" or str(remark) == "nan":
            return ""
        
        remark_str = str(remark)
        
        # 匹配模式（按优先级）
        patterns = [
            # 模式1: 加班费1250/小时
            r'加班费[\s]*([\d,.]+)[\s]*/[\s]*小时',
            # 模式2: 加班 833.33/小时
            r'加班[\s]+([\d,.]+)[\s]*/[\s]*小时',
            # 模式3: 加班费 1041.66/小时已退租
            r'加班费[\s]+([\d,.]+)[\s]*/[\s]*小时',
            # 模式4: 加班 1小时/1200泰铢
            r'加班[\s]+[\d]+[\s]*小时[\s]*/[\s]*([\d,.]+)[\s]*泰铢',
            r'加班[\s]+[\d]+[\s]*小时[\s]*/[\s]*([\d,.]+)',
        ]
        
        for pattern in patterns:
            match = re.search(pattern, remark_str, re.IGNORECASE)
            if match:
                rate_str = match.group(1)
                # 移除逗号
                rate_str = rate_str.replace(',', '')
                try:
                    rate = float(rate_str)
                    # 验证单价合理性（通常加班费单价在几百到几千泰铢/小时）
                    if 100 <= rate <= 50000:
                        return rate
                except ValueError:
                    continue
        
        return ""
    
    summary_df["月租金额_泰铢"] = summary_df["备注"].apply(extract_monthly_rent)
    summary_df["备注加班费单价"] = summary_df["备注"].apply(extract_overtime_rate)
    
    # 统计提取成功率
    monthly_rent_count = (summary_df["月租金额_泰铢"] != "").sum()
    overtime_rate_count = (summary_df["备注加班费单价"] != "").sum()
    print(f"  ✅ 已从备注中提取月租金额：{monthly_rent_count}/{len(summary_df)} 行")
    print(f"  ✅ 已从备注中提取加班费单价：{overtime_rate_count}/{len(summary_df)} 行")
    
    # 4. 生成年月列：根据使用日期智能推断
    def extract_year_month(row):
        use_date = str(row.get('使用日期', '')).strip()
        month_field = str(row.get('月份', '')).strip()
        
        if not use_date or use_date == 'nan' or use_date == '':
            return ""
        
        # 尝试多种日期格式解析
        try:
            # 情况1：Excel日期序列号（纯数字且大于1000）
            if use_date.replace('.', '').replace('-', '').isdigit():
                date_num = float(use_date)
                if date_num > 1000:  # Excel日期序列号
                    from datetime import datetime, timedelta
                    # Excel日期基准：1900-01-01（注意Excel bug：1900不是闰年但Excel认为是）
                    base_date = datetime(1899, 12, 30)
                    actual_date = base_date + timedelta(days=date_num)
                    return actual_date.strftime('%Y-%m')
            
            # 情况2：标准日期格式 YYYY-MM-DD, YYYY/MM/DD, YYYY.MM.DD
            date_patterns = [
                r'(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})',  # YYYY-MM-DD
                r'(\d{1,2})[-/\.](\d{1,2})[-/\.]?(\d{4})',  # DD-MM-YYYY
                r'(\d{4})(\d{2})(\d{2})',  # YYYYMMDD
            ]
            
            for pattern in date_patterns:
                match = re.search(pattern, use_date)
                if match:
                    groups = match.groups()
                    # 判断哪个是年份（4位数字）
                    if len(groups[0]) == 4:
                        year, month = groups[0], groups[1]
                    elif len(groups[2]) == 4:
                        year, month = groups[2], groups[1]
                    else:
                        continue
                    
                    # 验证月份范围
                    month_int = int(month)
                    if 1 <= month_int <= 12:
                        return f"{year}-{month.zfill(2)}"
            
            # 情况3：从月份字段辅助推断年份
            if month_field and month_field != 'nan':
                # 提取月份中的年份信息
                year_match = re.search(r'(20\d{2})', month_field)
                if year_match:
                    year = year_match.group(1)
                    # 提取月份数字
                    month_match = re.search(r'(\d{1,2})月', month_field)
                    if month_match:
                        month = month_match.group(1).zfill(2)
                        return f"{year}-{month}"
        
        except Exception as e:
            pass
        
        return ""  # 无法解析则返回空
    
    summary_df["年月"] = summary_df.apply(extract_year_month, axis=1)
    year_month_count = (summary_df["年月"] != "").sum()
    print(f"  ✅ 已生成年月列，成功解析 {year_month_count}/{len(summary_df)} 行")
    
    # 检测并修正台班费和结算台班填反的情况
    print(f"\n检测数据填反情况...")
    summary_df, swapped_indices = detect_and_fix_swapped_values(summary_df)
    if swapped_indices:
        print(f"✅ 已自动修正 {len(swapped_indices)} 行台班费/结算台班填反的数据，将用浅蓝色标注")
    else:
        print(f"✅ 未发现台班费/结算台班填反的数据")
    
    # 检测并修正出租单位和月份填反的情况
    print(f"\n检测出租单位和月份填反情况...")
    summary_df, swapped_month_indices = detect_and_fix_supplier_month_swap(summary_df)
    if swapped_month_indices:
        print(f"✅ 已自动修正 {len(swapped_month_indices)} 行出租单位/月份填反的数据，将用黄色标注")
    else:
        print(f"✅ 未发现出租单位/月份填反的数据")
    
    # 调整列顺序（基础字段统一为：sheet名称、项目编码、项目名称）
    col_order = [
        "sheet名称", "项目编码", "项目名称", "开始月份", "区域", "标准机具种类", "标准型号规格",
        "序号", "机具种类", "型号规格", "台班费_泰铢", "本次结算台班",
        "加班小时", "加班费", "进出场费运费", "移动费拆装费", "结算费用_泰铢",
        "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%", "本次费用合计",
        "使用日期", "年月", "出租单位", "月份", "备注", "月租金额_泰铢", "备注加班费单价"
    ]
    
    # 检查是否所有列都存在
    missing_cols = [col for col in col_order if col not in summary_df.columns]
    if missing_cols:
        print(f"\n⚠️  警告：以下列在数据中不存在: {missing_cols}")
        print(f"实际存在的列: {list(summary_df.columns)}")
    
    # 只选择存在的列进行重新排序
    existing_cols = [col for col in col_order if col in summary_df.columns]
    summary_df = summary_df[existing_cols]
    print(f"重新排序后的列: {list(summary_df.columns)}")

    # 5. 计算汇总金额并验证
    summary_total = summary_df["本次费用合计"].apply(safe_to_float).sum()
    report_total = sum(item.get("费用合计(泰铢)", 0) for item in report_data if item["状态"] == "成功")
    
    print(f"\n{'='*60}")
    print(f"金额验证：")
    print(f"  各sheet费用合计: {report_total:,.2f} 泰铢")
    print(f"  汇总表费用合计: {summary_total:,.2f} 泰铢")
    if abs(summary_total - report_total) < 0.01:
        print(f"  ✅ 金额验证通过，差异: {abs(summary_total - report_total):.2f} 泰铢")
    else:
        print(f"  ⚠️  金额存在差异: {abs(summary_total - report_total):.2f} 泰铢")
    print(f"{'='*60}\n")
    
    # 添加汇总行到报告
    total_data_rows = sum(item.get("数据行数", 0) for item in report_data)
    total_overtime_errors = sum(item.get("加班费错误行数", 0) for item in report_data)
    total_settlement_errors = sum(item.get("结算费用错误行数", 0) for item in report_data)
    total_amount_errors = sum(item.get("本次费用合计错误行数", 0) for item in report_data)
    
    report_data.append({
        "sheet名称": "【合计】",
        "项目名称": "",
        "项目编码": "",
        "开始月份": "",
        "状态": "",
        "数据行数": total_data_rows,
        "区域": "",
        "机具种类数": "",
        "型号规格数": "",
        "加班费错误行数": total_overtime_errors,
        "加班费错误率(%)": round((total_overtime_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "加班费差异金额(泰铢)": round(sum(item.get("加班费差异金额(泰铢)", 0) for item in report_data), 2),
        "结算费用错误行数": total_settlement_errors,
        "结算费用错误率(%)": round((total_settlement_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "结算费用差异金额(泰铢)": round(sum(item.get("结算费用差异金额(泰铢)", 0) for item in report_data), 2),
        "本次费用合计错误行数": total_amount_errors,
        "本次费用合计错误率(%)": round((total_amount_errors / total_data_rows * 100) if total_data_rows > 0 else 0, 2),
        "本次费用合计差异金额(泰铢)": round(sum(item.get("本次费用合计差异金额(泰铢)", 0) for item in report_data), 2),
        "加班费行数": sum(item.get("加班费行数", 0) for item in report_data),
        "加班费合计(泰铢)": round(sum(item.get("加班费合计(泰铢)", 0) for item in report_data), 2),
        "结算费用行数": sum(item.get("结算费用行数", 0) for item in report_data),
        "结算费用合计(泰铢)": round(sum(item.get("结算费用合计(泰铢)", 0) for item in report_data), 2),
        "本次费用行数": sum(item.get("本次费用行数", 0) for item in report_data),
        "费用合计(泰铢)": round(report_total, 2)
    })
    
    # 6. 生成处理报告
    report_df = pd.DataFrame(report_data)
    
    # 7. 生成按项目的机具种类费用汇总
    print(f"\n生成详细统计报告...")
    
    # 按项目和月份汇总
    project_month_stats = summary_df.groupby(['项目名称', '项目编码', '月份']).agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count',
        '标准机具种类': 'nunique'
    }).round(2)
    project_month_stats.columns = ['费用合计(泰铢)', '使用次数', '机具种类数']
    project_month_stats = project_month_stats.reset_index()
    
    # 按区域和机具种类汇总
    region_kind_stats = summary_df.groupby(['区域', '标准机具种类']).agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count'
    }).round(2)
    region_kind_stats.columns = ['费用合计(泰铢)', '使用次数']
    region_kind_stats = region_kind_stats.reset_index()
    
    # 按出租单位汇总
    supplier_stats = summary_df.groupby('出租单位').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count',
        '标准机具种类': 'nunique',
        '项目名称': lambda x: ', '.join(x.unique()[:3]),  # 显示前3个项目
        '项目编码': lambda x: ', '.join([str(c) for c in x.unique()[:3] if c])
    }).round(2)
    supplier_stats.columns = ['费用合计(泰铢)', '使用次数', '机具种类数', '关联项目', '项目编码']
    supplier_stats = supplier_stats.sort_values('费用合计(泰铢)', ascending=False).reset_index()
    
    # 月度趋势分析
    month_trend = summary_df.groupby('月份').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count',
        '标准机具种类': 'nunique',
        '项目名称': 'nunique'
    }).round(2)
    month_trend.columns = ['费用合计(泰铢)', '使用次数', '机具种类数', '项目数']
    month_trend = month_trend.reset_index()
    
    # 8. 处理数值列（转换为浮点数，保留原始数值类型）
    numeric_columns = [
        "台班费_泰铢", "本次结算台班", "加班小时", "加班费",
        "进出场费运费", "移动费拆装费", "结算费用_泰铢",
        "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%", "本次费用合计",
        "月租金额_泰铢", "备注加班费单价"
    ]
    
    # 注意：年月列不进行数值转换，保持文本格式
    
    # 定义需要货币格式的列
    currency_columns = [
        "台班费_泰铢", "加班费", "进出场费运费", "移动费拆装费",
        "结算费用_泰铢", "增值税7%", "机具租赁预扣税3%", "机具租赁预扣税5%",
        "月租金额_泰铢", "备注加班费单价"
    ]
    
    for col in numeric_columns:
        if col in summary_df.columns:
            summary_df[col] = summary_df[col].apply(safe_to_float)

    # 生成项目月度机具费汇总表：纵向项目，横向使用“月份”列
    monthly_report_source = summary_df.copy()

    monthly_report_source["月度项目编码"] = monthly_report_source["项目编码"].fillna("").astype(str).str.strip()
    monthly_report_source["月度项目名称"] = monthly_report_source["项目名称"].fillna("未匹配项目").astype(str).str.strip()
    monthly_report_source["月度报告月份"] = monthly_report_source["月份"].fillna("").astype(str).str.strip()
    monthly_report_source = monthly_report_source[
        (monthly_report_source["月度报告月份"] != "") &
        (monthly_report_source["月度报告月份"] != "nan")
    ].copy()

    project_monthly_fee_report = monthly_report_source.pivot_table(
        index=["月度项目编码", "月度项目名称"],
        columns="月度报告月份",
        values="本次费用合计",
        aggfunc="sum",
        fill_value=0
    )

    month_columns = list(project_monthly_fee_report.columns)

    def sort_monthly_report_column(col_name):
        col_str = str(col_name)
        if re.match(r'^\d{4}-\d{2}$', col_str):
            return (0, col_str)
        return (1, col_str)

    project_monthly_fee_report = project_monthly_fee_report.reindex(
        columns=sorted(month_columns, key=sort_monthly_report_column)
    )
    project_monthly_fee_report = project_monthly_fee_report.reset_index().round(2)
    project_monthly_fee_report = project_monthly_fee_report.rename(columns={
        "月度项目编码": "项目编码",
        "月度项目名称": "项目名称"
    })

    # 生成项目汇总机具费表
    project_summary_df = summary_df.groupby(['项目编码', '项目名称']).agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count',
        '标准机具种类': 'nunique',
        '出租单位': 'nunique',
        '月份': 'nunique'
    }).round(2).reset_index()
    project_summary_df.columns = ['项目编码', '项目名称', '费用合计(泰铢)', '使用次数', '机具种类数', '出租单位数', '覆盖月份数']
    project_summary_df = project_summary_df.sort_values('费用合计(泰铢)', ascending=False).reset_index(drop=True)

    # 生成项目月度金额增量表：按项目展示每月金额汇总
    project_monthly_change_report = project_monthly_fee_report.copy().round(2)
    
    # 8. 保存到新Excel（输出到汇总文件夹，带时间戳）
    # 确保输出文件夹存在
    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    
    output_basename = build_summary_output_basename(excel_path)
    output_filename = f"{output_basename}_{OUTPUT_FILE_SUFFIX}.xlsx"
    output_path = Path(OUTPUT_FOLDER) / output_filename

    # 先收集旧汇总文件，待新文件输出成功后再删除
    old_summary_files = []
    old_summary_patterns = [
        f"{output_basename}_*.xlsx",
        f"{excel_path.stem}_已汇总*.xlsx",
    ]
    for pattern in old_summary_patterns:
        for old_file in Path(OUTPUT_FOLDER).glob(pattern):
            if old_file != output_path and old_file not in old_summary_files:
                old_summary_files.append(old_file)

    if old_summary_files:
        print(f"检测到 {len(old_summary_files)} 个旧汇总文件，输出成功后将自动删除")
    
    # 定义填充颜色
    pink_fill = PatternFill(start_color="FFC0CB", end_color="FFC0CB", fill_type="solid")  # 粉红色：本次费用合计错误
    light_blue_fill = PatternFill(start_color="ADD8E6", end_color="ADD8E6", fill_type="solid")  # 浅蓝色：台班费/结算台班已交换
    yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # 黄色：出租单位/月份已交换
    orange_fill = PatternFill(start_color="FFD700", end_color="FFD700", fill_type="solid")  # 金色：加班费错误
    light_green_fill = PatternFill(start_color="90EE90", end_color="90EE90", fill_type="solid")  # 浅绿色：结算费用错误
    
    # 验证数据并记录错误
    overtime_error_count = 0
    settlement_error_count = 0
    total_error_count = 0
    
    print(f"\n验证费用计算...")
    
    # 验证加班费
    print(f"\n1. 验证加班费...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calculated = validate_overtime_fee(row)
        if not is_correct and actual != 0:  # 只报告非零的错误
            overtime_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行加班费错误：实际={actual:.2f}，应为={calculated:.2f}，差异={abs(actual-calculated):.2f}")
    
    if overtime_error_count > 0:
        print(f"发现 {overtime_error_count} 行加班费计算错误，将用金色标注")
    else:
        print(f"✅ 所有加班费验证通过")
    
    # 验证结算费用
    print(f"\n2. 验证结算费用...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calc1, calc2 = validate_settlement_amount(row)
        if not is_correct and actual != 0:  # 只报告非零的错误
            settlement_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行结算费用错误：实际={actual:.2f}，方式1={calc1:.2f}，方式2={calc2:.2f}")
    
    if settlement_error_count > 0:
        print(f"发现 {settlement_error_count} 行结算费用计算错误，将用浅绿色标注")
    else:
        print(f"✅ 所有结算费用验证通过")
    
    # 验证本次费用合计
    print(f"\n3. 验证本次费用合计...")
    for idx in range(len(summary_df)):
        row = summary_df.iloc[idx]
        is_correct, actual, calculated = validate_total_amount(row)
        if not is_correct:
            total_error_count += 1
            print(f"  ⚠️  第 {idx+2} 行本次费用合计错误：实际={actual:.2f}，应为={calculated:.2f}，差异={abs(actual-calculated):.2f}")
    
    if total_error_count > 0:
        print(f"发现 {total_error_count} 行本次费用合计计算错误，将用粉红色标注")
    else:
        print(f"✅ 所有本次费用合计验证通过")
    
    # 总结
    total_errors = overtime_error_count + settlement_error_count + total_error_count
    if total_errors > 0:
        print(f"\n⚠️  共发现 {total_errors} 处计算错误")
    else:
        print(f"\n✅ 所有数据验证通过，无计算错误")

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        # 先写入重点sheet（放在最前）
        project_summary_df.to_excel(writer, sheet_name=PROJECT_SUMMARY_SHEET_NAME, index=False)
        project_monthly_change_report.to_excel(writer, sheet_name=PROJECT_MONTHLY_CHANGE_SHEET_NAME, index=False)

        # 写入报告表
        report_df.to_excel(writer, sheet_name="处理报告", index=False)

        # 写入汇总表
        summary_df.to_excel(writer, sheet_name=SUMMARY_SHEET_NAME, index=False)
        
        # 写入统计报表
        project_month_stats.to_excel(writer, sheet_name="项目_月份统计", index=False)
        region_kind_stats.to_excel(writer, sheet_name="区域_机具种类统计", index=False)
        supplier_stats.to_excel(writer, sheet_name="出租单位统计", index=False)
        month_trend.to_excel(writer, sheet_name="月度趋势", index=False)

        from openpyxl.styles import Alignment, Font

        # 定义加粗字体
        bold_font = Font(bold=True)

        def format_summary_like_sheet(ws, display_df, left_align_use_date=False):
            """按汇总页规则格式化工作表；可选将使用日期列左对齐"""
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

            for col_idx, col_name in enumerate(display_df.columns, 1):
                max_length = len(str(col_name))
                for row_idx in range(2, len(display_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell_value = cell.value
                    df_row_idx = row_idx - 2

                    # 为货币列设置货币格式（2位小数，千分位）
                    if col_name in currency_columns:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'
                    # 为其他数值列设置普通数值格式
                    elif col_name in numeric_columns:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'
                    # 为使用日期列设置日期格式
                    elif col_name == "使用日期":
                        if isinstance(cell_value, (int, float)) and cell_value > 0:
                            cell.number_format = 'YYYY-MM-DD'
                        elif isinstance(cell_value, str) and cell_value.strip():
                            try:
                                if cell_value.strip().replace('.', '').replace('-', '').replace('/', '').isdigit():
                                    date_value = float(cell_value)
                                    if date_value > 0:
                                        cell.value = date_value
                                        cell.number_format = 'YYYY-MM-DD'
                            except:
                                pass

                        if left_align_use_date:
                            cell.alignment = Alignment(horizontal='left')
                    # 年月列设置居中
                    elif col_name == "年月":
                        cell.alignment = Alignment(horizontal='center')

                    # 为交换过的行标注浅蓝色（台班费和结算台班列）
                    if df_row_idx in swapped_indices and col_name in ['台班费_泰铢', '本次结算台班']:
                        cell.fill = light_blue_fill

                    # 为交换过的行标注黄色（出租单位和月份列）
                    if df_row_idx in swapped_month_indices and col_name in ['出租单位', '月份']:
                        cell.fill = yellow_fill

                    # 验证错误着色（基于summary_df同索引行）
                    if df_row_idx >= 0 and df_row_idx < len(summary_df):
                        source_row = summary_df.iloc[df_row_idx]

                        if col_name == "加班费":
                            is_correct, actual, calculated = validate_overtime_fee(source_row)
                            if not is_correct and actual != 0:
                                cell.fill = orange_fill

                        if col_name == "结算费用_泰铢":
                            is_correct, actual, calc1, calc2 = validate_settlement_amount(source_row)
                            if not is_correct and actual != 0:
                                cell.fill = light_green_fill

                        if col_name == "本次费用合计":
                            cell.font = bold_font
                            is_correct, actual, calculated = validate_total_amount(source_row)
                            if not is_correct:
                                cell.fill = pink_fill

                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))

                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            ws.freeze_panes = 'A2'
            ws.auto_filter.ref = ws.dimensions

        def format_monthly_report_sheet(ws, display_df):
            """格式化项目月度机具费汇总表"""
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')

            for col_idx, col_name in enumerate(display_df.columns, 1):
                max_length = len(str(col_name))
                is_amount_col = col_name not in ["项目编码", "项目名称"]

                for row_idx in range(2, len(display_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell_value = cell.value

                    if is_amount_col:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'

                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))

                if col_name == "项目编码":
                    adjusted_width = min(max(max_length + 2, 16), 24)
                elif col_name == "项目名称":
                    adjusted_width = min(max(max_length + 2, 18), 40)
                else:
                    adjusted_width = max(14, min(max_length + 2, 18))
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width

            ws.freeze_panes = 'C2'
            ws.auto_filter.ref = ws.dimensions
        
        # 美化汇总表
        ws_summary = writer.sheets[SUMMARY_SHEET_NAME]
        format_summary_like_sheet(ws_summary, summary_df, left_align_use_date=False)

        # 美化项目汇总机具费表
        ws_project_summary = writer.sheets[PROJECT_SUMMARY_SHEET_NAME]
        ws_project_summary.row_dimensions[1].height = 40
        for cell in ws_project_summary[1]:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        for col_idx, col_name in enumerate(project_summary_df.columns, 1):
            max_length = len(str(col_name))
            for row_idx in range(2, len(project_summary_df) + 2):
                cell = ws_project_summary.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                if col_name == '费用合计(泰铢)':
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    cell.number_format = '#,##0.00'
                elif col_name in ['使用次数', '机具种类数', '出租单位数', '覆盖月份数']:
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    cell.number_format = '0'
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            adjusted_width = min(max_length + 2, 40)
            ws_project_summary.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        ws_project_summary.freeze_panes = 'A2'
        ws_project_summary.auto_filter.ref = ws_project_summary.dimensions

        # 美化项目月度金额增量表
        ws_project_monthly_change = writer.sheets[PROJECT_MONTHLY_CHANGE_SHEET_NAME]
        format_monthly_report_sheet(ws_project_monthly_change, project_monthly_change_report)
        
        # 美化报告表
        ws_report = writer.sheets["处理报告"]
        
        # 设置首行格式：高度40，自动换行
        ws_report.row_dimensions[1].height = 40
        for cell in ws_report[1]:
            cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
        
        for col_idx, col_name in enumerate(report_df.columns, 1):
            max_length = len(str(col_name))
            for row_idx in range(2, len(report_df) + 2):
                cell = ws_report.cell(row=row_idx, column=col_idx)
                cell_value = cell.value
                
                # 为费用合计列设置货币格式（带泰铢符号）
                if col_name == "费用合计(泰铢)":
                    # 转换字符串为数字
                    if isinstance(cell_value, str):
                        try:
                            cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                        except:
                            cell.value = 0.0
                    # 应用货币格式
                    cell.number_format = '#,##0.00'
                
                if cell_value:
                    max_length = max(max_length, len(str(cell_value)))
            adjusted_width = min(max_length + 2, 40)
            ws_report.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
        
        # 冻结首行
        ws_report.freeze_panes = 'A2'
        
        # 添加自动筛选
        ws_report.auto_filter.ref = ws_report.dimensions
        
        # 美化统计报表
        stats_sheets = [
            ("项目_月份统计", project_month_stats),
            ("区域_机具种类统计", region_kind_stats),
            ("出租单位统计", supplier_stats),
            ("月度趋势", month_trend)
        ]
        
        for sheet_name, stats_df in stats_sheets:
            ws = writer.sheets[sheet_name]
            
            # 设置首行格式
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.alignment = Alignment(wrap_text=True, vertical='center', horizontal='center')
            
            # 设置列宽和格式
            for col_idx, col_name in enumerate(stats_df.columns, 1):
                max_length = len(str(col_name))
                for row_idx in range(2, len(stats_df) + 2):
                    cell = ws.cell(row=row_idx, column=col_idx)
                    cell_value = cell.value
                    
                    # 为费用合计列和台班费列设置货币格式
                    if col_name in ["费用合计(泰铢)", "台班费_泰铢"]:
                        if isinstance(cell_value, str):
                            try:
                                cell.value = float(cell_value.replace(',', '').strip()) if cell_value.strip() else 0.0
                            except:
                                cell.value = 0.0
                        cell.number_format = '#,##0.00'
                    
                    if cell_value:
                        max_length = max(max_length, len(str(cell_value)))
                
                adjusted_width = min(max_length + 2, 50)
                ws.column_dimensions[get_column_letter(col_idx)].width = adjusted_width
            
            # 冻结首行
            ws.freeze_panes = 'A2'
            
            # 添加自动筛选
            ws.auto_filter.ref = ws.dimensions

    # 新文件写入成功后，删除旧汇总文件
    deleted_old_files = 0
    for old_file in old_summary_files:
        try:
            old_file.unlink()
            deleted_old_files += 1
            print(f"已删除旧汇总文件: {old_file.name}")
        except Exception as e:
            print(f"⚠️ 删除旧汇总文件失败: {old_file.name}，原因: {e}")

    print(f"\n{'='*60}")
    print(f"汇总完成！文件已保存为：")
    print(f"  {output_path}")
    if deleted_old_files > 0:
        print(f"已删除 {deleted_old_files} 个旧汇总文件")
    print(f"{'='*60}")
    print(f"共处理 {len(report_data)-1} 个 sheet")  # -1因为包含了合计行
    print(f"成功汇总 {len(all_data)} 个 sheet，共 {len(summary_df)} 行数据")
    print(f"\n处理报告：")
    for item in report_data[:-1]:  # 不显示合计行
        status_icon = "✅" if item["状态"] == "成功" else "❌"
        amount_str = f"{item.get('费用合计(泰铢)', 0):,.2f}" if item["状态"] == "成功" else "0.00"
        print(f"  {status_icon} {item['sheet名称']}: {item['状态']} - {item.get('数据行数', 0)} 行 - {amount_str} 泰铢")
    print(f"\n  💰 总计: {report_total:,.2f} 泰铢")
    print(f"\n已生成8个工作表：")
    print(f"  1. '{PROJECT_SUMMARY_SHEET_NAME}' - 项目汇总机具使用费")
    print(f"  2. '{PROJECT_MONTHLY_CHANGE_SHEET_NAME}' - 项目每月机具费用金额汇总")
    print(f"  3. '处理报告' - 各项目处理情况统计 (包含错误率、差异金额)")
    print(f"  4. '{SUMMARY_SHEET_NAME}' - 所有数据汇总 (共 {len(summary_df)} 行)")
    print(f"  5. '项目_月份统计' - 按项目和月份分组汇总")
    print(f"  6. '区域_机具种类统计' - 按区域和机具种类分组汇总")
    print(f"  7. '出租单位统计' - 按出租单位汇总")
    print(f"  8. '月度趋势' - 按月份汇总趋势分析")
    
    # 9. 生成筛选分析
    print(f"\n{'='*60}")
    print(f"数据分析：")
    print(f"{'='*60}")
    
    # 按区域统计
    print(f"\n📍 按区域统计：")
    region_stats = summary_df.groupby('区域').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count'
    }).round(2)
    region_stats.columns = ['费用合计(泰铢)', '数据行数']
    region_stats = region_stats.sort_values('费用合计(泰铢)', ascending=False)
    for region, row in region_stats.iterrows():
        print(f"  {region}: {row['费用合计(泰铢)']:,.2f} 泰铢 ({int(row['数据行数'])} 行)")
    
    # 按标准机具种类统计
    print(f"\n🔧 按机具种类统计（Top 10）：")
    kind_stats = summary_df.groupby('标准机具种类').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '序号': 'count'
    }).round(2)
    kind_stats.columns = ['费用合计(泰铢)', '使用次数']
    kind_stats = kind_stats.sort_values('费用合计(泰铢)', ascending=False).head(10)
    for idx, (kind, row) in enumerate(kind_stats.iterrows(), 1):
        print(f"  {idx}. {kind}: {row['费用合计(泰铢)']:,.2f} 泰铢 ({int(row['使用次数'])} 次)")
    
    # 按项目统计
    print(f"\n📊 按项目统计：")
    project_stats = summary_df.groupby('项目名称').agg({
        '本次费用合计': lambda x: x.apply(safe_to_float).sum(),
        '标准机具种类': 'nunique'
    }).round(2)
    project_stats.columns = ['费用合计(泰铢)', '机具种类数']
    project_stats = project_stats.sort_values('费用合计(泰铢)', ascending=False)
    for project, row in project_stats.iterrows():
        if project:  # 排除空项目名
            print(f"  {project}: {row['费用合计(泰铢)']:,.2f} 泰铢 ({int(row['机具种类数'])} 种机具)")
    
    print(f"\n{'='*60}")

if __name__ == "__main__":
    if EXCEL_FILE_OVERRIDE and str(EXCEL_FILE_OVERRIDE).strip():
        main()
    else:
        all_input_files = list_all_standardized_files(STANDARDIZED_INPUT_DIR)
        if not all_input_files:
            print(f"找不到标准化文件：{STANDARDIZED_INPUT_DIR}")
        else:
            print(f"自动发现标准化文件 {len(all_input_files)} 个，开始逐个汇总...")
            for idx, input_file in enumerate(all_input_files, start=1):
                print(f"\n[{idx}/{len(all_input_files)}] {input_file.name}")
                main(input_file)