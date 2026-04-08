# -*- coding: utf-8 -*-
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from datetime import datetime, timedelta
import numpy as np
import os
import re

# ================== 自动加载系统中文字体 ==================
plt.rcParams['font.sans-serif'] = [
    'PingFang SC', 'Microsoft YaHei', 'SimHei', 'Arial Unicode MS', 'DejaVu Sans'
]
plt.rcParams['axes.unicode_minus'] = False

# ================== 【全局字体大小控制区】只需修改这里即可批量调整所有文字大小 ==================
# 推荐范围：8~14（正常查看）、15~20（汇报打印大屏）、6~9（项目超多时防重叠）
GLOBAL_FONT_SCALE = 20      # <--- 主开关：修改这个数字即可全局控制字体大小

# 各部分字体大小（基于主开关自动缩放，比例经过精心调整，保持视觉协调）
FONT_TITLE         = int(GLOBAL_FONT_SCALE * 1.91)    # 标题最大
FONT_PROJECT_NAME  = int(GLOBAL_FONT_SCALE * 0.96)    # Y轴项目名
FONT_DATE_LABEL    = int(GLOBAL_FONT_SCALE * 0.83)    # 开始/结束日期标签
FONT_STATUS        = int(GLOBAL_FONT_SCALE * 0.96)    # 右侧“剩余/超期”状态文字
FONT_TODAY_LABEL   = int(GLOBAL_FONT_SCALE * 1.13)    # “今日”文字
FONT_XTICK         = int(GLOBAL_FONT_SCALE * 0.96)    # X轴年月标签
FONT_LEGEND        = int(GLOBAL_FONT_SCALE * 1.00)    # 图例文字
FONT_TICK_LABEL    = int(GLOBAL_FONT_SCALE * 0.96)    # Y轴刻度标签（项目名已单独控制）

# ================== 配置 ==================
file_path = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx"
sheet_name = "在建项目"
today = datetime.today().replace(hour=0, minute=0, second=0, microsecond=0)  # 动态今日

INCLUDE_REMARKS = ["B_新开项目", "A_在建项目","AS_收尾项目"]
SORT_MODE = 'code_asc'

# ================== 高明度配色方案 ==================
COLOR_PLAN = "#A0D8F1"      # 浅蓝灰（计划）
COLOR_ON_TRACK = "#66BB6A"  # 柔和绿（按期）
COLOR_DELAYED = "#FF7043"   # 珊瑚橙（超期）
COLOR_OVERDUE = "#E53935"   # 深红（超期部分）
COLOR_TODAY = "#333333"     # 今日线

# ================== 日期解析 ==================
def parse_date(value):
    if pd.isna(value) or value is None: return None
    s = str(value).strip()
    if not s or s == 'nan': return None
    s_clean = s.split('.')[0]
    if s_clean.isdigit() and len(s_clean) == 8:
        try: return datetime.strptime(s_clean, "%Y%m%d")
        except: pass
    try:
        num = float(s)
        if 40000 < num < 50000:
            return datetime(1899, 12, 30) + timedelta(days=int(num))
    except: pass
    for fmt in ("%Y/%m/%d", "%Y-%m-%d", "%Y.%m.%d", "%Y年%m月%d日", "%m/%d/%Y"):
        try: return datetime.strptime(s, fmt)
        except: continue
    try:
        dt = pd.to_datetime(s, errors='coerce')
        if pd.notna(dt): return dt.to_pydatetime()
    except: pass
    return None

def extract_days(value):
    if pd.isna(value) or value is None: return None
    s = str(value).strip()
    if not s or s.lower() in {'nan', 'nat', 'none'}:
        return None

    # 避免把日期字符串（如 2026-07-13）误识别为天数
    if re.search(r'\d{4}\s*[-/年\.]\s*\d{1,2}\s*[-/月\.]\s*\d{1,2}', s) or re.fullmatch(r'\d{8}', s):
        return None

    # 提取第一个数值，支持 180.0 / 180天 / 工期180
    match = re.search(r'(\d+(?:\.\d+)?)', s)
    if match:
        try:
            days = float(match.group(1))
            if days > 0:
                return int(round(days))
        except:
            pass
    return None

# ================== 读取 + 筛选 ==================
df_raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
header_row = 0
keywords = ['项目名称', '合同签订日期', '开工日期', '合同工期', '历天', '项目状态', '项目编码']
for i in range(min(20, len(df_raw))):
    row_str = ''.join(str(v).strip() for v in df_raw.iloc[i].tolist() if pd.notna(v))
    if any(k in row_str for k in keywords):
        header_row = i
        break

df = pd.read_excel(file_path, sheet_name=sheet_name, skiprows=header_row)
df = df.dropna(how='all').reset_index(drop=True)
df.columns = df.columns.str.strip()

required = ['项目编码', '项目名称', '合同签订日期', '开工日期', '合同工期', '历天', '项目状态']
for col in required:
    if col not in df.columns: df[col] = ''
df[required] = df[required].fillna('').astype(str)
df['项目编码'] = df['项目编码'].str.strip()
df['项目状态'] = df['项目状态'].str.strip()

df_filtered = df[df['项目状态'].isin(INCLUDE_REMARKS)].copy()
if df_filtered.empty:
    print("警告：没有符合筛选条件的项目！")
    exit()

# ================== 数据处理 ==================
records = []
for _, row in df_filtered.iterrows():
    name = row['项目名称'].strip()
    code = row['项目编码'].strip()
    if not name: continue

    start = parse_date(row['开工日期']) or parse_date(row['合同签订日期'])
    if not start: continue

    end = parse_date(row['合同工期'])
    days_from_litian = extract_days(row['历天'])
    days_from_contract = extract_days(row['合同工期'])

    # 结束日期优先使用有效的“合同工期日期”；否则按“开工日期 + 历天”兜底
    if not end or (end - start).days <= 0:
        fallback_days = days_from_litian or days_from_contract
        end = start + timedelta(days=fallback_days) if fallback_days else None
    if not end or (end - start).days <= 0: continue

    actual_days = max((today - start).days, 0)
    overdue_days = max((today - end).days, 0)
    days_left = (end - today).days
    status = f"剩余 {days_left}天" if days_left >= 0 else f"超期 {overdue_days}天"
    color = COLOR_ON_TRACK if days_left >= 0 else COLOR_DELAYED

    records.append({
        'code': code, 'name': name, 'start': start, 'end': end,
        'duration': (end - start).days, 'actual_days': actual_days,
        'overdue_days': overdue_days, 'progress_color': color, 'status': status,
        'start_text': start.strftime("%Y-%m-%d"),
        'end_text': end.strftime("%Y-%m-%d")
    })

# ================== 排序 ==================
if SORT_MODE == 'code_asc':
    records.sort(key=lambda x: x['code'] if x['code'] else 'ZZZ')
elif SORT_MODE == 'code_desc':
    records.sort(key=lambda x: x['code'] if x['code'] else '', reverse=True)
else:
    records.sort(key=lambda x: x['start'])

# ================== 绘图（终极美化 + 防遮挡）==================
if not records:
    print("没有可绘制的数据")
else:
    n = len(records)
    bar_height = min(0.75, 20.0 / n)
    fig, ax = plt.subplots(figsize=(36, max(n * bar_height * 1.4, 10)))

    # X轴范围
    all_dates = [r['start'] for r in records] + [r['end'] for r in records]
    if any(r['overdue_days'] > 0 for r in records):
        all_dates += [r['end'] + timedelta(days=r['overdue_days']) for r in records]
    min_date = min(all_dates) - timedelta(days=60)
    max_date = max(all_dates) + timedelta(days=150)

    y_pos = np.arange(n)
    for i, r in enumerate(records):
        y = y_pos[i]

        # 计划工期
        ax.barh(y, r['duration'], left=r['start'], color=COLOR_PLAN, height=bar_height, alpha=0.9)
        # 实际进度
        ax.barh(y, r['actual_days'], left=r['start'], color=r['progress_color'], height=bar_height, alpha=0.95)
        # 超期部分
        if r['overdue_days'] > 0:
            ax.barh(y, r['overdue_days'], left=r['end'], color=COLOR_OVERDUE, height=bar_height, alpha=0.95)

        # 开始日期：左端（防遮挡）
        ax.text(
            r['start'] - timedelta(days=1.5), y, r['start_text'],
            va='center', ha='right', fontsize=FONT_DATE_LABEL,
            fontweight='bold', color='white',
            bbox=dict(boxstyle="round,pad=0.35", facecolor='#2E2E2E', alpha=0.9, edgecolor='none')
        )

        # 结束日期：右端（防遮挡）
        ax.text(
            r['end'] + timedelta(days=1.5), y, r['end_text'],
            va='center', ha='left', fontsize=FONT_DATE_LABEL,
            fontweight='bold', color='white',
            bbox=dict(boxstyle="round,pad=0.35", facecolor='#2E2E2E', alpha=0.9, edgecolor='none')
        )

    # 状态（右侧）
    status_x = max_date - timedelta(days=40)
    for i, r in enumerate(records):
        ax.text(status_x, y_pos[i], r['status'],
                va='center', ha='left', fontsize=FONT_STATUS,
                fontweight='bold', color=r['progress_color'])

    # Y轴：项目名 + 编码
    max_len = max(len(r['name']) for r in records)
    y_labels = [f"{r['name']}{' ' * (max_len - len(r['name']) + 3)}{r['code']}" for r in records]
    ax.set_yticks(y_pos)
    ax.set_yticklabels(y_labels, fontsize=FONT_PROJECT_NAME, fontweight='bold')
    ax.invert_yaxis()

    # 今日线
    ax.axvline(today, color=COLOR_TODAY, linestyle='--', linewidth=1.8, alpha=0.9)
    ax.text(today, -0.8, f'今日 {today.strftime("%Y-%m-%d")}', color=COLOR_TODAY, fontsize=FONT_TODAY_LABEL, fontweight='bold',
            ha='center', va='top',
            bbox=dict(boxstyle="round,pad=0.4", facecolor='#FFFFCC', alpha=0.7, edgecolor='none'))

    # X轴：每3个月
    ax.xaxis_date()
    ax.xaxis.set_major_locator(mdates.MonthLocator(interval=3))
    ax.xaxis.set_minor_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
    plt.setp(ax.get_xticklabels(), rotation=45, ha='right', fontsize=FONT_XTICK, fontweight='bold')

    ax.set_xlim(min_date, max_date)
    ax.set_ylim(n - 0.5, -0.5)
    ax.set_title("项目进度情况", fontsize=FONT_TITLE, pad=30, fontweight='bold')

    # 图例
    legend_elements = [
        plt.Rectangle((0,0),1,1, facecolor=COLOR_PLAN, label='计划工期'),
        plt.Rectangle((0,0),1,1, facecolor=COLOR_ON_TRACK, label='按期进行'),
        plt.Rectangle((0,0),1,1, facecolor=COLOR_DELAYED, label='已超期（进度）'),
        plt.Rectangle((0,0),1,1, facecolor=COLOR_OVERDUE, label='超期部分')
    ]
    leg = ax.legend(handles=legend_elements, loc='upper right',
                    bbox_to_anchor=(0.88, 0.98), fontsize=FONT_LEGEND,
                    frameon=True, fancybox=True, shadow=False)
    leg.get_frame().set_facecolor('#F8F9FA')
    leg.get_frame().set_edgecolor('#DDDDDD')

    # 网格
    ax.grid(True, axis='x', which='major', linestyle='--', alpha=0.6, linewidth=0.8)
    ax.grid(True, axis='x', which='minor', linestyle=':', alpha=0.3, linewidth=0.5)
    for spine in ax.spines.values():
        spine.set_visible(False)

    # ========== 动态防遮挡 + 精确布局 ==========
    date_box_width = 8
    max_label_width = max(len(label) for label in y_labels) * 0.8
    left_padding_days = date_box_width + max_label_width * 2
    dynamic_left = max(0.20, (left_padding_days / (max_date - min_date).days) * 0.8)

    plt.subplots_adjust(
        left=dynamic_left,
        bottom=0.13,
        right=0.944,
        top=0.871
    )

    # 保存
    output_dir = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/0_项目基本信息情况"
    os.makedirs(output_dir, exist_ok=True)
    date_tag = today.strftime("%Y%m%d")
    output_path = os.path.join(output_dir, f"在建项目甘特图_{date_tag}.png")

    fig.savefig(output_path, dpi=350, bbox_inches='tight', facecolor='white')
    print(f"高颜值甘特图已保存：{output_path}")