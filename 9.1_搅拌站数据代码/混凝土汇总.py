import pandas as pd
import numpy as np
import os
import sys
import importlib.util
from datetime import datetime

# ===== 项目配置与客户映射 =====
CODE_ROOT = os.path.dirname(os.path.abspath(__file__))
DEFAULT_DATA_WORKSPACE = '/Users/wyattsun/Desktop/邯郸中材/01成本统计'
DEFAULT_RAW_DATA_DIR = '/Users/wyattsun/Desktop/邯郸中材/01成本统计/9.1_商混站🚛/A_商混站_raw'
DEFAULT_STD_OUTPUT_DIR = '/Users/wyattsun/Desktop/邯郸中材/01成本统计/9.1_商混站🚛/B_商混站_std'


def _data_workspace():
    """数据总目录，可通过环境变量覆盖。"""
    return os.environ.get('CONCRETE_DATA_WORKSPACE', DEFAULT_DATA_WORKSPACE)


def _resolve_data_dir():
    """解析混凝土数据目录，优先环境变量，其次自动探测。"""
    explicit = os.environ.get('CONCRETE_DATA_DIR', '').strip()
    if explicit and os.path.isdir(explicit):
        return explicit

    # 默认使用商混站原始数据目录
    if os.path.isdir(DEFAULT_RAW_DATA_DIR):
        return DEFAULT_RAW_DATA_DIR

    ws = _data_workspace()
    candidates = [
        os.path.join(ws, '9.1_商混站混凝土'),
        os.path.join(ws, '9.1_商混凝站混凝土'),
        os.path.join(ws, '9.1_商混站混凝土', 'A_商混凝站混凝土数据_按月统计数据'),
    ]
    for p in candidates:
        if os.path.isdir(p):
            return p

    if os.path.isdir(ws):
        for name in os.listdir(ws):
            p = os.path.join(ws, name)
            if not os.path.isdir(p):
                continue
            if name.startswith('9.1_') and ('混凝土' in name or '商混' in name):
                return p

    # 兼容旧布局：数据与脚本放在同目录
    return CODE_ROOT


def _output_dir(data_dir):
    # 优先环境变量，其次使用固定标准化目录
    explicit = os.environ.get('CONCRETE_OUTPUT_DIR', '').strip()
    if explicit:
        out = explicit
    elif os.path.isdir(os.path.dirname(DEFAULT_STD_OUTPUT_DIR)):
        out = DEFAULT_STD_OUTPUT_DIR
    else:
        out = os.path.join(data_dir, 'B_标准化数据')
    os.makedirs(out, exist_ok=True)
    return out


MATCHER_DIR = os.path.join(os.path.dirname(CODE_ROOT), '0_项目基本信息')


def _load_module_from_path(module_name, file_path):
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    if spec is None or spec.loader is None:
        raise ImportError(f"无法加载模块: {file_path}")
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


_project_name_matcher = _load_module_from_path(
    'project_name_matcher',
    os.path.join(MATCHER_DIR, 'project_name_matcher.py')
)
build_standard_projects_from_dataframe = _project_name_matcher.build_standard_projects_from_dataframe
is_unmatched_alias = _project_name_matcher.is_unmatched_alias
load_or_init_db = _project_name_matcher.load_or_init_db
match_project = _project_name_matcher.match_project
prompt_unmatched_project_mappings = _project_name_matcher.prompt_unmatched_project_mappings
save_db = _project_name_matcher.save_db

PROJECT_MATCH_DB_PATH = os.path.join(MATCHER_DIR, 'project_name_match_db.json')

def _norm(s):
    return (str(s).strip() if s is not None else '').upper()

def _read_detail_excel(detail_path):
    """优先读取“数据源”工作表，失败则读取首个工作表。"""
    try:
        return pd.read_excel(detail_path, sheet_name='数据源')
    except Exception:
        return pd.read_excel(detail_path)

def _read_detail_with_header_detection(detail_path):
    """自动检测表头行，遍历所有工作表寻找包含客户名等关键列的表。"""
    customer_candidates = {
        "客户名",
        "客户名称",
        "客户",
        "项目名称",
        "工程名称",
        "领用项目",
    }

    # 确定尝试顺序：优先 "数据源"，其余按原始顺序
    try:
        all_sheets = pd.ExcelFile(detail_path).sheet_names
    except Exception:
        all_sheets = []

    ordered_sheets = []
    if '数据源' in all_sheets:
        ordered_sheets.append('数据源')
    for s in all_sheets:
        if s != '数据源':
            ordered_sheets.append(s)
    if not ordered_sheets:
        ordered_sheets = [0]  # fallback: 首个 sheet

    for sheet in ordered_sheets:
        try:
            raw_df = pd.read_excel(detail_path, sheet_name=sheet, header=None)
        except Exception:
            continue
        for i in range(min(20, len(raw_df))):
            row_vals = raw_df.iloc[i].astype(str).str.strip().tolist()
            if any(v in customer_candidates for v in row_vals):
                header = raw_df.iloc[i].astype(str).str.strip().tolist()
                df = raw_df.iloc[i + 1:].copy()
                df.columns = header
                return df.reset_index(drop=True)

    # 所有 sheet 都未检测到，退回默认读取
    return _read_detail_excel(detail_path)


def _make_unique_columns(columns):
    """将重复列名改为唯一列名，保留首个原名，其余追加后缀 _2/_3。"""
    seen = {}
    unique_cols = []
    for col in columns:
        name = str(col).strip()
        if name in seen:
            seen[name] += 1
            unique_cols.append(f"{name}_{seen[name]}")
        else:
            seen[name] = 1
            unique_cols.append(name)
    return unique_cols

def _normalize_detail_df(detail_df):
    # 清洗列名空格，防止匹配失败
    detail_df.columns = _make_unique_columns(detail_df.columns)

    def pick_col(candidates):
        return next((c for c in candidates if c in detail_df.columns), None)

    # 适配日期列（后期数据）或年/月（前期数据）
    date_col = pick_col(["日期", "生产日期", "开盘日期"])
    if date_col:
        detail_df["日期"] = pd.to_datetime(detail_df[date_col], errors="coerce")
    else:
        year_col = pick_col(["年", "年份"])
        month_col = pick_col(["月", "月份"])
        if year_col and month_col:
            years = detail_df[year_col].astype(str).str.replace(r"\D+", "", regex=True)
            months = detail_df[month_col].astype(str).str.replace(r"\D+", "", regex=True)
            ym = years.str.zfill(4) + "-" + months.str.zfill(2) + "-01"
            detail_df["日期"] = pd.to_datetime(ym, format="%Y-%m-%d", errors="coerce")

    # 适配“客户名”列名变更
    customer_candidates = [
        "客户名",
        "客户名称",
        "客户",
        "项目名称",
        "工程名称",
        "领用项目",
    ]
    customer_col = pick_col(customer_candidates)
    if customer_col is None:
        raise RuntimeError(f"未找到客户名列，请检查明细表列名。已检测列：{list(detail_df.columns)}")
    if customer_col != "客户名":
        detail_df["客户名"] = detail_df[customer_col]
    detail_df["客户名"] = detail_df["客户名"].astype(str).str.strip()

    # 适配“总方量”列名变更：将可能的列名统一映射为“总方量”
    volume_candidates = [
        "总方量",
        "方量",
        "数量",
        "方量(立方)",
        "方量（立方）",
        "立方",
        "方量m3",
        "方量(M3)",
        "方量（m3）",
    ]
    volume_col = pick_col(volume_candidates)
    if volume_col is None:
        raise RuntimeError(f"未找到方量列，请检查明细表列名。已检测列：{list(detail_df.columns)}")
    if volume_col != "总方量":
        detail_df["总方量"] = detail_df[volume_col]
    detail_df["总方量"] = pd.to_numeric(detail_df["总方量"], errors="coerce").fillna(0)

    # 适配“塌落度/标准型号”列名变更：标准型号作为塌落度的替代列
    slump_candidates = ["塌落度", "坍落度", "标准型号"]
    slump_col = pick_col(slump_candidates)
    if slump_col and slump_col not in {"塌落度", "坍落度"}:
        detail_df["塌落度"] = detail_df[slump_col]

    # 可选字段：总车次、建筑部位、抗渗等级
    trip_col = pick_col(["总车次", "车次", "车数"])
    if trip_col and trip_col != "总车次":
        detail_df["总车次"] = detail_df[trip_col]
    if "总车次" in detail_df.columns:
        detail_df["总车次"] = pd.to_numeric(detail_df["总车次"], errors="coerce")

    part_col = pick_col(["建筑部位", "部位", "浇筑部位", "施工部位"])
    if part_col and part_col != "建筑部位":
        detail_df["建筑部位"] = detail_df[part_col]

    imperme_col = pick_col(["抗渗", "抗渗等级", "抗渗标号"])
    if imperme_col and imperme_col != "抗渗":
        detail_df["抗渗"] = detail_df[imperme_col]

    # 删除原材料用量及成本明细列（不参与汇总）
    drop_patterns = [
        r"^[0-9]*#",        # 1#砂, 2#石, 1#水 泥 ...
        r"(?i)sika",         # SIKA-D2323
        r"(?i)mighty",       # MIGHTY 150HB-9
        r"(?i)shk",          # SHK-G703CW
        r"(?i)hf-",          # HF-703G (G703)
        r"(?i)g703",
        r"^成本单价$",
    ]
    import re as _re
    drop_cols = [
        c for c in detail_df.columns
        if any(_re.search(p, str(c).strip()) for p in drop_patterns)
    ]
    if drop_cols:
        detail_df = detail_df.drop(columns=drop_cols)

    return detail_df

def _find_price_file(base_dir):
    """查找报价单：优先当前目录，其次上级目录（兼容 A_raw + 上级报价单 布局）。"""
    search_dirs = [base_dir]
    parent_dir = os.path.dirname(base_dir)
    if parent_dir and parent_dir not in search_dirs:
        search_dirs.append(parent_dir)

    for d in search_dirs:
        direct = os.path.join(d, '混凝土报价单.xlsx')
        if os.path.exists(direct):
            return direct

    for d in search_dirs:
        if not os.path.isdir(d):
            continue
        candidates = []
        for f in os.listdir(d):
            if f.startswith('~$'):
                continue
            if not f.lower().endswith(('.xlsx', '.xlsm')):
                continue
            if '报价' in f:
                candidates.append(os.path.join(d, f))
        if candidates:
            # 优先 xlsx，再按文件名排序保证稳定
            candidates = sorted(candidates, key=lambda p: (0 if p.lower().endswith('.xlsx') else 1, os.path.basename(p)))
            return candidates[0]

    return None

def _pick_detail_files(folder_path):
    """返回文件夹下所有符合条件的明细 Excel 路径列表。"""
    excel_files = []
    for f in os.listdir(folder_path):
        if f.startswith('~$'):
            continue
        if not f.lower().endswith(('.xlsx', '.xlsm')):
            continue
        if '报价' in f or '汇总报表' in f or '客户总额汇总' in f or '混凝土汇总报表' in f:
            continue
        excel_files.append(os.path.join(folder_path, f))
    return excel_files

def _iter_detail_files(data_dir):
    detail_items = []
    output_dir = _output_dir(data_dir)

    # 兼容明细直接放在数据根目录
    root_detail_paths = _pick_detail_files(data_dir)
    if root_detail_paths:
        detail_items.append({
            'folder': os.path.basename(data_dir),
            'detail_paths': root_detail_paths,
            'output_dir': output_dir,
        })

    for name in os.listdir(data_dir):
        if name.startswith('.') or name == '__pycache__':
            continue
        folder_path = os.path.join(data_dir, name)
        if not os.path.isdir(folder_path):
            continue
        detail_paths = _pick_detail_files(folder_path)
        if detail_paths:
            detail_items.append({'folder': name, 'detail_paths': detail_paths, 'output_dir': output_dir})
    return detail_items

def _write_outputs(merged_df, detail_summary, customer_summary, monthly_volume_pivot, monthly_amount_pivot, output_path):
    # 在导出前确保包含“金额（万泰铢）”列（写入后会用公式填充）
    if '金额（万泰铢）' not in merged_df.columns:
        merged_df['金额（万泰铢）'] = ""

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        customer_summary.to_excel(writer, sheet_name='客户总额汇总', index=False)
        monthly_volume_pivot.to_excel(writer, sheet_name='月度方量汇总', index=False)
        monthly_amount_pivot.to_excel(writer, sheet_name='月度金额汇总', index=False)
        merged_df.to_excel(writer, sheet_name='计算明细(全)', index=False)
        detail_summary.to_excel(writer, sheet_name='客户强度明细汇总', index=False)

    from openpyxl import load_workbook
    wb = load_workbook(output_path)

    for sheet_name in ['计算明细(全)', '客户强度明细汇总', '客户总额汇总', '月度方量汇总', '月度金额汇总']:
        if sheet_name in wb.sheetnames:
            wsx = wb[sheet_name]
            wsx.row_dimensions[1].height = 40
            wsx.freeze_panes = 'A2'
            wsx.auto_filter.ref = wsx.dimensions
            # 日期列格式化（去掉时间部分）
            from openpyxl.utils import get_column_letter as _gcl
            for c in range(1, wsx.max_column + 1):
                if wsx.cell(row=1, column=c).value == '日期':
                    for r in range(2, wsx.max_row + 1):
                        wsx.cell(row=r, column=c).number_format = 'YYYY-MM-DD'
                    break

    wb.save(output_path)
    wb.close()

def find_project_for_customer(customer_name, match_db):
    """统一使用 project_name_matcher 做客户名到标准项目映射。"""
    if customer_name is None:
        return {'project_code': '', 'project_name': ''}
    hit = match_project(str(customer_name).strip(), match_db) if match_db else None
    if hit:
        return {
            'project_code': hit.get('project_code', ''),
            'project_name': hit.get('project_name', ''),
        }
    return {'project_code': '', 'project_name': ''}

def process_concrete_data():
    try:
        data_dir = _resolve_data_dir()
        print(f"数据目录: {data_dir}")

        price_path = _find_price_file(data_dir)
        if not price_path:
            raise RuntimeError(
                "未找到报价单文件，请在以下目录放置含‘报价’的 Excel："
                f"{data_dir} 或 {os.path.dirname(data_dir)}"
            )
        price_df = pd.read_excel(price_path)
        price_df.columns = _make_unique_columns(price_df.columns)

        detail_items = _iter_detail_files(data_dir)
        if not detail_items:
            raise RuntimeError(f"未在子文件夹中找到明细文件：{data_dir}")

        # 初始化统一项目匹配库（仅使用 project_name_match_db.json）
        match_db = load_or_init_db(PROJECT_MATCH_DB_PATH, [])
        print(
            "✅ 共享匹配库已加载（不再读取 migration_config.json）："
            f"标准项目 {len(match_db.get('projects', {}))} 个，"
            f"别名 {len(match_db.get('alias_to_standard', {}))} 个，"
            f"未匹配留痕 {len(match_db.get('unmatched_aliases', []))} 个"
        )

        # 自动识别报价单中的价格列名（优先顺序）
        price_candidates = ['价格', '单价', '报价', 'Price', 'Unit Price']
        price_col = None
        for c in price_candidates:
            if c in price_df.columns:
                price_col = c
                break
        if price_col is None:
            # 如果未找到，尝试剩余数值列（按数值类型列）
            num_cols = price_df.select_dtypes(include='number').columns.tolist()
            price_col = num_cols[0] if len(num_cols) > 0 else None

        if price_col is None:
            raise RuntimeError('未在报价表中识别到价格列，请检查列名')

        # 归一化关键字段以提高匹配成功率
        import re

        def get_series(df, col_name, fallback_cols=None):
            """安全获取列为Series；若不存在则返回空Series（与df长度一致）。"""
            if col_name in df.columns:
                return df[col_name]
            if fallback_cols:
                for alt in fallback_cols:
                    if alt in df.columns:
                        return df[alt]
            return pd.Series([''] * len(df), index=df.index)

        def norm_strength(x):
            if pd.isna(x):
                return ''
            s = str(x)
            s = re.sub(r"\s+", "", s)
            m = re.match(r"^(\d+)\.0$", s)
            if m:
                s = m.group(1)
            return s.upper()

        def norm_slump(x):
            if pd.isna(x):
                return ''
            s = str(x)
            s = re.sub(r"\s+", "", s)
            return s

        def detect_price_date_col(df):
            date_candidates = [
                "生效日期",
                "月份",
                "日期",
                "报价日期",
                "价格日期",
                "执行日期",
                "开始日期",
            ]
            for c in date_candidates:
                if c in df.columns:
                    return c
            return None

        price_date_col = detect_price_date_col(price_df)
        if price_date_col:
            price_df["_price_date"] = pd.to_datetime(price_df[price_date_col], errors="coerce")
        else:
            price_df["_price_date"] = pd.NaT

        price_df['_s_key'] = get_series(price_df, '砼强度').apply(norm_strength)
        price_df['_t_key'] = get_series(price_df, '塌落度', ['坍落度', '标准型号']).apply(norm_slump)

        def build_price_lookup(df, price_column):
            by_key = {}
            by_strength = {}
            valid = df.dropna(subset=["_price_date"]).copy()
            if valid.empty:
                return by_key, by_strength
            valid = valid.sort_values("_price_date")
            for (s_key, t_key), grp in valid.groupby(["_s_key", "_t_key"], dropna=False):
                dates = grp["_price_date"].values.astype('datetime64[ns]')
                prices = grp[price_column].to_numpy()
                by_key[(s_key, t_key)] = (dates, prices)
            for s_key, grp in valid.groupby(["_s_key"], dropna=False):
                dates = grp["_price_date"].values.astype('datetime64[ns]')
                prices = grp[price_column].to_numpy()
                by_strength[s_key] = (dates, prices)
            return by_key, by_strength

        price_by_key, price_by_strength = build_price_lookup(price_df, price_col)

        # 预加载明细并收集客户
        pending = []
        detail_frames = []
        for item in detail_items:
            sub_dfs = []
            for detail_path in item['detail_paths']:
                try:
                    df = _read_detail_with_header_detection(detail_path)
                    df = _normalize_detail_df(df)
                    df['_源文件'] = os.path.basename(detail_path)
                    sub_dfs.append(df)
                    print(f"  读取: {os.path.basename(detail_path)} ({len(df)} 行)")
                except Exception as e:
                    print(f"  跳过: {os.path.basename(detail_path)} - {e}")
            if not sub_dfs:
                print(f"警告：{item['folder']} 下无有效明细文件，跳过")
                continue
            detail_df = pd.concat(sub_dfs, ignore_index=True)
            print(f"  {item['folder']} 合并后共 {len(detail_df)} 行")
            item['detail_df'] = detail_df
            detail_frames.append(item)
            unique_customers = sorted(set(detail_df['客户名'].dropna().astype(str).str.strip()))
            for name in unique_customers:
                if is_unmatched_alias(match_db, name):
                    continue
                if match_project(name, match_db):
                    continue
                if name not in pending:
                    pending.append(name)

        if pending:
            print(f"发现 {len(pending)} 个客户需确认，打开统一项目匹配窗口...")
            prompt_unmatched_project_mappings(pending, match_db, window_title='混凝土-手动项目匹配')
            save_db(PROJECT_MATCH_DB_PATH, match_db)

        all_merged = []
        total_groups = len(detail_frames)
        for idx_item, item in enumerate(detail_frames, start=1):
            detail_df = item['detail_df']
            folder = item['folder']
            output_dir = item['output_dir']

            detail_df['_s_key'] = get_series(detail_df, '砼强度').apply(norm_strength)
            detail_df['_t_key'] = get_series(detail_df, '塌落度', ['坍落度', '标准型号']).apply(norm_slump)

            use_date_price = price_date_col is not None and "日期" in detail_df.columns
            if use_date_price:
                def match_price(s_key, t_key, dt):
                    if pd.isna(dt):
                        return np.nan
                    dt64 = np.datetime64(pd.Timestamp(dt), 'ns')
                    key = (s_key, t_key)
                    if key in price_by_key:
                        dates, prices = price_by_key[key]
                        idx = np.searchsorted(dates, dt64, side="right") - 1
                        if idx >= 0:
                            return prices[idx]
                    if s_key in price_by_strength:
                        dates, prices = price_by_strength[s_key]
                        idx = np.searchsorted(dates, dt64, side="right") - 1
                        if idx >= 0:
                            return prices[idx]
                    return np.nan

                merged_df = detail_df.copy()
                merged_df["价格"] = merged_df.apply(
                    lambda r: match_price(r.get("_s_key", ""), r.get("_t_key", ""), r.get("日期")),
                    axis=1,
                )

                # 日期匹配失败的行，回退到静态价格匹配（取最新价格）
                still_missing = merged_df['价格'].isna()
                if still_missing.any():
                    latest_price = price_df.sort_values('_price_date', na_position='first').drop_duplicates(
                        subset=['_s_key', '_t_key'], keep='last'
                    ).set_index(['_s_key', '_t_key'])[price_col]
                    for idx in merged_df.index[still_missing]:
                        sk = merged_df.at[idx, '_s_key']
                        tk = merged_df.at[idx, '_t_key']
                        if (sk, tk) in latest_price.index:
                            merged_df.at[idx, '价格'] = latest_price[(sk, tk)]
                        else:
                            # 仅按强度回退
                            fallback_s = price_df.sort_values('_price_date', na_position='first').drop_duplicates(
                                subset=['_s_key'], keep='last'
                            ).set_index('_s_key')[price_col]
                            if sk in fallback_s.index:
                                merged_df.at[idx, '价格'] = fallback_s[sk]
            else:
                merged_df = pd.merge(
                    detail_df,
                    price_df[['_s_key', '_t_key', price_col]].rename(columns={price_col: '价格'}),
                    left_on=['_s_key', '_t_key'],
                    right_on=['_s_key', '_t_key'],
                    how='left'
                )

                missing_mask = merged_df['价格'].isna()
                if missing_mask.any():
                    fallback = price_df.groupby('_s_key')[price_col].first().to_dict()
                    merged_df.loc[missing_mask, '价格'] = merged_df.loc[missing_mask, '_s_key'].map(fallback)

                missing_mask = merged_df['价格'].isna()
                if missing_mask.any():
                    slurry_mask = merged_df['_s_key'].astype(str).str.upper() == '砂浆'
                    target_mask = missing_mask & slurry_mask
                    if target_mask.any():
                        price_df['_t_key'] = price_df['_t_key'].fillna('')
                        slurry_price = price_df.loc[
                            (price_df['_s_key'].astype(str).str.upper() == '砂浆') & (price_df['_t_key'] == ''),
                            price_col,
                        ]
                        if not slurry_price.empty:
                            merged_df.loc[target_mask, '价格'] = slurry_price.iloc[0]

            merged_df['价格'] = pd.to_numeric(merged_df['价格'], errors='coerce').fillna(0)
            merged_df['金额'] = merged_df['总方量'].fillna(0) * merged_df['价格']

            if '总车次' not in merged_df.columns:
                merged_df['总车次'] = 0

            matched = (merged_df['价格'] > 0).sum()
            total = len(merged_df)
            print(f"{folder} 报价匹配：{matched}/{total} 行获得价格")

            def map_code(name):
                return find_project_for_customer(name, match_db).get('project_code', '')
            def map_name(name):
                return find_project_for_customer(name, match_db).get('project_name', '')
            merged_df['项目编码'] = merged_df['客户名'].apply(map_code)
            merged_df['项目名称'] = merged_df['客户名'].apply(map_name)
            merged_df['项目名称'] = merged_df['项目名称'].replace('', pd.NA).fillna(merged_df['客户名'])

            # 生成 年、月列：优先保留原始年/月数据，缺失时从日期补充
            if '年' not in merged_df.columns:
                merged_df['年'] = np.nan
            if '月' not in merged_df.columns:
                merged_df['月'] = np.nan
            merged_df['年'] = pd.to_numeric(merged_df['年'], errors='coerce')
            merged_df['月'] = pd.to_numeric(merged_df['月'], errors='coerce')
            if '日期' in merged_df.columns:
                mask_y = merged_df['年'].isna()
                mask_m = merged_df['月'].isna()
                if mask_y.any():
                    merged_df.loc[mask_y, '年'] = merged_df.loc[mask_y, '日期'].dt.year
                if mask_m.any():
                    merged_df.loc[mask_m, '月'] = merged_df.loc[mask_m, '日期'].dt.month
            # 年/月转为整数显示
            merged_df['年'] = merged_df['年'].astype('Int64')
            merged_df['月'] = merged_df['月'].astype('Int64')
            # 日期列去掉时间部分
            if '日期' in merged_df.columns:
                merged_df['日期'] = merged_df['日期'].dt.date
            merged_df['强度（塌落度）'] = merged_df['_s_key'].astype(str) + '(' + merged_df['_t_key'].astype(str) + ')'
            merged_df['金额（万泰铢）'] = merged_df['金额'] / 10000

            # 只保留指定列
            keep_cols = [
                '项目编码', '项目名称', '砼强度', '建筑部位', '总方量', '总车次', '塌落度', '抗渗',
                '年', '月', '强度（塌落度）', '价格', '金额', '金额（万泰铢）',
            ]
            merged_df = merged_df.reindex(columns=[c for c in keep_cols if c in merged_df.columns])

            customer_summary = merged_df.groupby(['项目编码', '项目名称']).agg({
                '总方量': 'sum',
                '金额': 'sum',
                '总车次': 'sum'
            }).reset_index()
            customer_summary = customer_summary.sort_values(by='项目编码', ascending=False, na_position='last')
            customer_summary['金额（万泰铢）'] = customer_summary['金额'] / 10000
            cs_order = ['项目编码', '项目名称', '总方量', '总车次', '金额', '金额（万泰铢）']
            customer_summary = customer_summary.reindex(columns=[c for c in cs_order if c in customer_summary.columns])

            detail_summary = merged_df.groupby(['项目编码', '项目名称', '砼强度', '塌落度']).agg({
                '总方量': 'sum',
                '价格': 'first',
                '金额': 'sum',
            }).reset_index()
            detail_summary['金额（万泰铢）'] = detail_summary['金额'] / 10000
            ds_order = ['项目编码', '项目名称', '砼强度', '塌落度', '总方量', '价格', '金额', '金额（万泰铢）']
            detail_summary = detail_summary.reindex(columns=[c for c in ds_order if c in detail_summary.columns])

            merged_for_month = merged_df.copy()
            merged_for_month['年月'] = merged_for_month.apply(
                lambda r: f"{int(r['年'])}-{int(r['月']):02d}" if pd.notna(r['年']) and pd.notna(r['月']) else '',
                axis=1
            )
            merged_for_month = merged_for_month[merged_for_month['年月'] != '']

            monthly_volume_pivot = merged_for_month.pivot_table(
                index=['项目编码', '项目名称'],
                columns='年月',
                values='总方量',
                aggfunc='sum',
                fill_value=0,
            )
            monthly_amount_pivot = merged_for_month.pivot_table(
                index=['项目编码', '项目名称'],
                columns='年月',
                values='金额',
                aggfunc='sum',
                fill_value=0,
            )

            def sort_month_cols(cols):
                return sorted(cols, key=lambda s: (int(str(s).split('-')[0]), int(str(s).split('-')[1])))

            if len(monthly_volume_pivot.columns) > 0:
                monthly_volume_pivot = monthly_volume_pivot.reindex(columns=sort_month_cols(monthly_volume_pivot.columns))
            if len(monthly_amount_pivot.columns) > 0:
                monthly_amount_pivot = monthly_amount_pivot.reindex(columns=sort_month_cols(monthly_amount_pivot.columns))

            monthly_volume_pivot = monthly_volume_pivot.reset_index()
            monthly_amount_pivot = monthly_amount_pivot.reset_index()

            monthly_volume_pivot = monthly_volume_pivot.sort_values(by='项目编码', ascending=False, na_position='last')
            monthly_amount_pivot = monthly_amount_pivot.sort_values(by='项目编码', ascending=False, na_position='last')

            ym_suffix = datetime.now().strftime('%Y%m')
            if total_groups > 1:
                output_name = f"搅拌站混凝土汇总报表_{ym_suffix}_{idx_item:02d}.xlsx"
            else:
                output_name = f"搅拌站混凝土汇总报表_{ym_suffix}.xlsx"
            output_path = os.path.join(output_dir, output_name)
            _write_outputs(merged_df, detail_summary, customer_summary, monthly_volume_pivot, monthly_amount_pivot, output_path)

        print("数据处理成功！")
        print(f"文件保存位置: {data_dir}")

    except Exception as e:
        print(f"处理过程中出现错误: {e}")

if __name__ == "__main__":
    process_concrete_data()