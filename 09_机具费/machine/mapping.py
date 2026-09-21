# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import os
import re
from pathlib import Path

import pandas as pd

from machine.config import PROJECT_INFO_XLSX
from machine.textutil import normalize_lookup_text

PROJECT_CODE_NAME_MAPPING_FILE = str(PROJECT_INFO_XLSX)


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
                            alias = normalize_lookup_text(alias)
                            if alias:
                                machine_kind_map[alias] = std_kind
                    elif isinstance(aliases, str):
                        for alias in re.split(r'[，,]', aliases):
                            alias = normalize_lookup_text(alias)
                            if alias:
                                machine_kind_map[alias] = std_kind
            elif isinstance(legacy_kind_map, dict):
                machine_kind_map = {
                    normalize_lookup_text(raw): str(std).strip()
                    for raw, std in legacy_kind_map.items()
                    if normalize_lookup_text(raw)
                }

            # 型号映射：优先读取推荐分组结构
            grouped_model_map = data.get("model_to_aliases", {}) if isinstance(data, dict) else {}
            legacy_model_map = data.get("model_map", {}) if isinstance(data, dict) else {}

            if isinstance(grouped_model_map, dict) and grouped_model_map:
                for kind, std_model_dict in grouped_model_map.items():
                    kind = normalize_lookup_text(kind)
                    if not kind or not isinstance(std_model_dict, dict):
                        continue
                    for std_model, aliases in std_model_dict.items():
                        std_model = str(std_model).strip()
                        if not std_model:
                            continue
                        if isinstance(aliases, list):
                            for alias in aliases:
                                alias = normalize_lookup_text(alias)
                                if alias:
                                    model_map[(kind, alias)] = std_model
                        elif isinstance(aliases, str):
                            for alias in re.split(r'[，,]', aliases):
                                alias = normalize_lookup_text(alias)
                                if alias:
                                    model_map[(kind, alias)] = std_model
            elif isinstance(legacy_model_map, dict):
                for kind, model_dict in legacy_model_map.items():
                    kind = normalize_lookup_text(kind)
                    if not kind or not isinstance(model_dict, dict):
                        continue
                    for raw_model, std_model in model_dict.items():
                        raw_model = normalize_lookup_text(raw_model)
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
            normalize_lookup_text(raw): str(std).strip()
            for raw, std in kind_mapping.items()
            if normalize_lookup_text(raw)
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
            kind = normalize_lookup_text(kind)
            raw_model = normalize_lookup_text(raw_model)
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


def guess_region_from_project_or_sheet(project_name, sheet_name, mapping):
    """优先按项目名匹配区域，失败后回退按sheet名匹配。"""
    project_name = str(project_name).strip() if project_name is not None else ""
    if project_name and project_name != "nan":
        for raw_project in sorted(mapping.keys(), key=len, reverse=True):
            if raw_project and raw_project in project_name:
                return mapping[raw_project]
    return guess_region_from_sheet_name(sheet_name, mapping)


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

