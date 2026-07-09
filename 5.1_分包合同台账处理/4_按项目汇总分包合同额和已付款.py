from pathlib import Path
import numpy as np
import pandas as pd


ROOT = Path('/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同')
INPUT_DIR = ROOT / '2_分包合同台账_std'
OUTPUT_DIR = ROOT / '3_分包合同_sum'


def pick_column(columns, candidates):
    """在列名列表中按候选名匹配，并返回实际列名。"""
    col_map = {str(c).strip().replace(' ', ''): c for c in columns}
    for candidate in candidates:
        key = str(candidate).strip().replace(' ', '')
        if key in col_map:
            return col_map[key]
    return None


def find_latest_excel(input_dir: Path) -> Path:
    excel_files = sorted(input_dir.glob('*.xlsx'))
    if not excel_files:
        raise FileNotFoundError(f'未在 {input_dir} 下找到 Excel 文件')
    return max(excel_files, key=lambda p: p.stat().st_mtime)


def build_summary(input_file: Path, output_file: Path):
    df = pd.read_excel(input_file)

    project_code_col = pick_column(df.columns, ['项目编码', '项目代码', 'Project Code'])
    project_name_col = pick_column(df.columns, ['项目名称', 'Project Name'])
    amount_col = pick_column(df.columns, [
        '合同额（泰铢）', '合同额(泰铢)', '合同额', '分包合同额（泰铢）', '分包合同额(泰铢)', '合同金额'
    ])
    paid_col = pick_column(df.columns, ['分包已付款', '公司实付含税', '已付款', '实付含税'])

    if not project_code_col or not project_name_col or not amount_col or not paid_col:
        raise ValueError(
            f'无法识别所需列。项目编码={project_code_col}, 项目名称={project_name_col}, '
            f'合同额={amount_col}, 分包已付款={paid_col}'
        )

    work_df = df[[project_code_col, project_name_col, amount_col, paid_col]].copy()
    work_df.columns = ['项目编码', '项目名称', '合同额（泰铢）', '分包已付款（泰铢）']

    # 清洗并转换为数值
    for col in ['合同额（泰铢）', '分包已付款（泰铢）']:
        work_df[col] = pd.to_numeric(work_df[col], errors='coerce').fillna(0)

    # 去除空项目编码/名称，统一清理
    work_df['项目编码'] = work_df['项目编码'].fillna('').astype(str).str.strip()
    work_df['项目名称'] = work_df['项目名称'].fillna('').astype(str).str.strip()
    work_df = work_df[(work_df['项目编码'] != '') | (work_df['项目名称'] != '')]

    summary = (
        work_df.groupby(['项目编码', '项目名称'], as_index=False)
        .agg({'合同额（泰铢）': 'sum', '分包已付款（泰铢）': 'sum'})
    )

    summary['分包付款比例'] = np.where(
        summary['合同额（泰铢）'] > 0,
        summary['分包已付款（泰铢）'] / summary['合同额（泰铢）'],
        0
    )

    summary_thb = summary[['项目编码', '项目名称', '合同额（泰铢）', '分包已付款（泰铢）', '分包付款比例']].copy()
    summary_thb['分包付款比例'] = summary_thb['分包付款比例'].fillna(0)

    summary_wan = summary_thb.copy()
    summary_wan.rename(columns={
        '合同额（泰铢）': '合同额（万泰铢）',
        '分包已付款（泰铢）': '分包已付款（万泰铢）',
    }, inplace=True)
    summary_wan['合同额（万泰铢）'] = summary_wan['合同额（万泰铢）'] / 10000
    summary_wan['分包已付款（万泰铢）'] = summary_wan['分包已付款（万泰铢）'] / 10000

    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        summary_thb.to_excel(writer, sheet_name='泰铢', index=False)
        summary_wan.to_excel(writer, sheet_name='万泰铢', index=False)

        # 设置格式
        for sheet_name in ['泰铢', '万泰铢']:
            ws = writer.book[sheet_name]
            # 自动列宽
            for column in ws.columns:
                max_length = 0
                column = list(column)
                for cell in column:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                adjusted_width = min(max_length + 2, 40)
                ws.column_dimensions[column[0].column_letter].width = adjusted_width

            # 数值列格式
            if sheet_name == '泰铢':
                value_columns = ['C', 'D']
            else:
                value_columns = ['C', 'D']

            for col_letter in value_columns:
                for row in range(2, len(summary_thb) + 2):
                    ws[f'{col_letter}{row}'].number_format = '0.00'

            ratio_col = 'E'
            for row in range(2, len(summary_thb) + 2):
                ws[f'{ratio_col}{row}'].number_format = '0.00%'

    print(f'已生成汇总文件: {output_file}')


if __name__ == '__main__':
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    input_file = find_latest_excel(INPUT_DIR)
    output_file = OUTPUT_DIR / f'{input_file.stem}_项目汇总.xlsx'
    build_summary(input_file, output_file)
