#!/usr/bin/env python3
"""
合并 A_机具费_raw 下的3个Excel文件：
1. 统一列名（03月表与04/05月表结构有差异）
2. 纵向合并
3. 输出为 clean_机具费_raw合并.xlsx
"""

import pandas as pd
import os
import re
import sys
import importlib.util
import inspect
from pathlib import Path
from datetime import datetime

# ========== 路径 ==========
RAW_DIR = os.path.dirname(os.path.abspath(__file__))
RAW_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/A_机具费_raw"
OUT_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/10_机具费🚜/B_机具费_std"
os.makedirs(OUT_DIR, exist_ok=True)

def load_project_matcher(matcher_py_path: str):
    """尝试从指定的 .py 文件加载匹配函数，返回一个可调用的匹配函数或 None。"""
    try:
        spec = importlib.util.spec_from_file_location("project_name_matcher", matcher_py_path)
        mod = importlib.util.module_from_spec(spec)
        spec.loader.exec_module(mod)
        # 常见函数名尝试（优先返回单参数可调用包装器）
        for fname in ("match_project_name", "match_project", "find_standard_project", "get_project_code"):
            if hasattr(mod, fname):
                fn = getattr(mod, fname)
                sig = None
                try:
                    sig = inspect.signature(fn)
                except Exception:
                    sig = None

                # 如果函数需要 db 参数（如 match_project(raw_name, db)），尝试初始化 DB 并返回包装器
                if sig is not None and len(sig.parameters) >= 2:
                    # 默认 db 路径（与 project_db_manager 中一致）
                    default_db_path = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"
                    if hasattr(mod, 'load_or_init_db'):
                        try:
                            db = mod.load_or_init_db(default_db_path)
                            return lambda raw: fn(raw, db)
                        except Exception as e:
                            print(f"⚠️ 无法初始化匹配库: {e}")
                    # 若无法初始化，返回原函数 (调用处需传 db)
                    return fn
                return fn
        # 如果模块导出的是一个可调用对象
        if hasattr(mod, "PROJECT_MATCHER") and callable(mod.PROJECT_MATCHER):
            return mod.PROJECT_MATCHER
        print(f"⚠️ project_name_matcher 未找到已知的匹配函数（{matcher_py_path}）。")
    except Exception as e:
        print(f"⚠️ 导入 project_name_matcher 失败: {e}")
    return None
files = sorted([f for f in os.listdir(RAW_DIR) if f.endswith('.xlsx') and not f.startswith('~')])
print(f"找到 {len(files)} 个文件: {files}")

# ========== 标准化映射函数 ==========
def process_file(filepath, filename):
    """读取并标准化一个Excel文件"""

    # 读取全部数据（不设header，手动处理）
    raw = pd.read_excel(filepath, header=None)
    print(f"  原始 shape: {raw.shape}")

    # 判断格式：03月是19列，04/05月有32列（后13列空白）
    # 只取前19列
    raw = raw.iloc[:, :19]

    # 第一行是标题行
    headers_raw = list(raw.iloc[0])
    data = raw.iloc[1:].copy().reset_index(drop=True)
    print(f"  数据行数: {len(data)}")

    # 标准化列名
    new_headers = []
    for c in headers_raw:
        c_stripped = str(c).strip() if pd.notna(c) else ''

        if re.match(r'^序号$', c_stripped):
            new_headers.append('序号')
        elif re.match(r'^机具种类$', c_stripped):
            new_headers.append('机具种类')
        elif re.match(r'^型号规格$', c_stripped):
            new_headers.append('型号规格')
        elif '台班费' in c_stripped:
            new_headers.append('台班费（泰铢）')
        elif '结算台班' in c_stripped or '本次结算台班' in c_stripped:
            new_headers.append('本次结算台班（小时、天、月）')
        elif c_stripped == '加班费（小时）':
            # 03月特有：加班费单价 → 映射到加班费/超时费
            new_headers.append('加班费/超时费')
        elif '加班费' in c_stripped and '超时' in c_stripped:
            new_headers.append('加班费/超时费')
        elif '加班' in c_stripped and ('小时' in c_stripped or '(' in c_stripped or '（' in c_stripped):
            new_headers.append('加班（小时）')
        elif c_stripped == '加班费':
            # 03月：第7列是"加班费"→ 实际是加班费总额，跳过（与超时费合并）
            new_headers.append('__drop__')
        elif '进出场' in c_stripped or c_stripped == '运费＋拆装费' or c_stripped == '进出场费/运费':
            new_headers.append('进出场费/运费')
        elif '移动费' in c_stripped or '拆装费' in c_stripped or '接管费' in c_stripped:
            new_headers.append('移动费＋拆装费+接管费')
        elif '结算费用' in c_stripped or '本次结算费用' in c_stripped:
            new_headers.append('结算费用（泰铢）')
        elif c_stripped == '增值税7%':
            new_headers.append('增值税7%')
        elif '预扣税' in c_stripped:
            if '5%' in c_stripped:
                new_headers.append('机具租赁预扣税5%')
            else:
                new_headers.append('机具租赁预扣税3%')
        elif '费用合计' in c_stripped:
            new_headers.append('本次费用合计（泰铢）')
        elif '使用时间' in c_stripped or '使用日期' in c_stripped:
            new_headers.append('使用日期')
        elif '出租单位' in c_stripped:
            new_headers.append('出租单位')
        elif '使用项目' in c_stripped:
            new_headers.append('使用项目')
        elif '备注' in c_stripped:
            new_headers.append('备注')
        elif '月份' in c_stripped:
            new_headers.append('月份')
        else:
            print(f"  ⚠️  无法映射列: [{c_stripped}]，跳过")
            new_headers.append('__drop__')

    # 删除标记为drop的列
    drop_indices = [i for i, h in enumerate(new_headers) if h == '__drop__']
    keep_indices = [i for i, h in enumerate(new_headers) if h != '__drop__']
    final_headers = [h for h in new_headers if h != '__drop__']

    data = data.iloc[:, keep_indices].copy()
    data.columns = final_headers

    # 处理03月的"加班费"列：如果原列是"加班费（小时）"（第6列）和"加班费"（第7列）
    # "加班费（小时）"实际上是"加班费/超时费"，"加班费"是另一列（多余的已丢弃）
    # 但这里"加班费（小时）"映射到了"加班费/超时费"
    # 03月缺少"移动费＋拆装费+接管费"列，补上
    if '移动费＋拆装费+接管费' not in data.columns:
        data['移动费＋拆装费+接管费'] = None

    # 确保标准列顺序
    std_order = [
        '序号', '机具种类', '型号规格', '台班费（泰铢）',
        '本次结算台班（小时、天、月）', '加班（小时）', '加班费/超时费',
        '进出场费/运费', '移动费＋拆装费+接管费', '结算费用（泰铢）',
        '增值税7%', '机具租赁预扣税3%', '机具租赁预扣税5%',
        '本次费用合计（泰铢）', '使用日期', '出租单位', '使用项目', '备注', '月份'
    ]
    for col in std_order:
        if col not in data.columns:
            data[col] = None

    data = data[std_order]

    # 月份标准化：03月格式是"2024-06"，04/05月格式是 datetime → 统一
    if '月份' in data.columns:
        data['月份'] = data['月份'].apply(lambda x: str(x) if pd.notna(x) else None)

    print(f"  处理后列数: {len(data.columns)}")
    return data


# ========== 逐个处理 ==========
all_dfs = []
for f in files:
    path = os.path.join(RAW_DIR, f)
    print(f"\n📄 处理: {f}")
    df = process_file(path, f)
    print(f"  ✅ 行数: {len(df)}")
    all_dfs.append(df)

# ========== 合并 ==========
if len(all_dfs) == 0:
    print("\n❌ 没有成功读取任何文件")
    exit(1)

merged = pd.concat(all_dfs, ignore_index=True)

print(f"\n{'='*60}")
print(f"✅ 合并完成！")
print(f"  总行数: {len(merged)}")
print(f"  总列数: {len(merged.columns)}")
print(f"  列名: {list(merged.columns)}")

# ====== 项目名称匹配：从 external matcher 添加两列 ======
# 优先使用 JSON DB 直接匹配，绕过 project_name_matcher.py 的函数签名差异
DB_PATH = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"
import json

def load_db(db_path=DB_PATH):
    with open(db_path, 'r', encoding='utf-8') as f:
        return json.load(f)

def build_alias_and_code_maps(db):
    alias_map = {}
    code_map = {}
    for std, meta in (db.get("projects", {}) or {}).items():
        std_s = str(std).strip()
        if not std_s:
            continue
        alias_map[std_s] = std_s
        for a in meta.get("aliases", []) or []:
            a_s = str(a).strip()
            if a_s:
                alias_map[a_s] = std_s
        code_map[std_s] = meta.get("project_code")
    return alias_map, code_map

def match_by_db(raw_name, alias_map, code_map):
    if raw_name is None:
        return (None, None)
    s = str(raw_name).strip()
    if not s:
        return (None, None)
    std = alias_map.get(s)
    if std:
        return (std, code_map.get(std))
    key = re.sub(r"\s+", "", s).lower()
    for k, v in alias_map.items():
        if re.sub(r"\s+", "", str(k)).lower() == key:
            return (v, code_map.get(v))
    return (None, None)

# 初始化 DB 映射
try:
    _db = load_db(DB_PATH)
    _alias_map, _code_map = build_alias_and_code_maps(_db)
    matcher_fn = lambda raw: match_by_db(raw, _alias_map, _code_map)
except Exception as e:
    print(f"⚠️ 无法加载 JSON DB，回退到 module matcher: {e}")
    matcher_py = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_matcher.py"
    matcher_fn = load_project_matcher(matcher_py)

if '使用项目' in merged.columns:
    def _apply_match(x):
        if pd.isna(x) or str(x).strip() == '':
            return (None, None)
        if matcher_fn:
            try:
                res = matcher_fn(str(x))
                if res is None:
                    return (None, None)
                if isinstance(res, tuple):
                    return (res[0], res[1] if len(res) > 1 else None)
                if isinstance(res, dict):
                    name = res.get('项目名称') or res.get('标准项目名称') or res.get('name') or res.get('standard_name')
                    code = res.get('项目编码') or res.get('code') or res.get('id')
                    return (name, code)
                if isinstance(res, str):
                    return (res, None)
            except Exception as e:
                print(f"⚠️ 匹配函数调用错误: {e}")
                return (None, None)
        return (None, None)

    matches = merged['使用项目'].apply(_apply_match)
    merged['项目名称'] = [m[0] for m in matches]
    merged['项目编码'] = [m[1] for m in matches]
    print("  ✅ 已添加列: 项目名称, 项目编码")

# ========== 按月份排序 ==========
if '月份' in merged.columns:
    merged = merged.sort_values('月份', na_position='last').reset_index(drop=True)
    print(f"  已按月份排序")

# ========== 输出到 B_机具费_sum ==========
out_path = os.path.join(OUT_DIR, "clean_机具费_raw合并.xlsx")
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    merged.to_excel(writer, sheet_name='机具费合并', index=False)
    ws = writer.sheets['机具费合并']
    for col_idx, col_name in enumerate(merged.columns, 1):
        max_len = len(str(col_name)) * 2 + 2
        sample_vals = merged.iloc[:100, col_idx-1].dropna()
        for val in sample_vals:
            max_len = max(max_len, len(str(val)))
        ws.column_dimensions[ws.cell(row=1, column=col_idx).column_letter].width = min(max_len + 2, 45)

print(f"\n📁 输出: {out_path}")

# ========== 摘要 ==========
print(f"\n{'='*60}")
print("📊 数据摘要")
print(f"{'='*60}")

if '月份' in merged.columns:
    print(f"\n◆ 各月份数据量:")
    print(merged['月份'].value_counts().sort_index().to_string())

if '机具种类' in merged.columns:
    print(f"\n◆ 机具种类分布（前20）:")
    print(merged['机具种类'].value_counts().head(20).to_string())

if '使用项目' in merged.columns:
    print(f"\n◆ 使用项目分布:")
    print(merged['使用项目'].value_counts().to_string())

# 金额汇总
money_cols = [c for c in merged.columns if '泰铢' in c or '费用' in c or '增值税' in c]
if money_cols:
    print(f"\n◆ 金额列汇总（泰铢）:")
    for c in money_cols:
        merged[c] = pd.to_numeric(merged[c], errors='coerce')
        total = merged[c].sum()
        print(f"  {c}: {total:,.2f}")
