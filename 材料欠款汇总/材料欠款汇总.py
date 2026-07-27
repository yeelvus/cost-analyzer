import sys
import os
from datetime import datetime
import pandas as pd

# ====================== 模块路径修复 ======================
matcher_dir = "/Users/wyattsun/Desktop/code/0_项目基本信息"
if matcher_dir not in sys.path:
    sys.path.insert(0, matcher_dir)

# excel_formatting_utils.py 所在目录（根据你的实际位置调整）
utils_dir = "/Users/wyattsun/Desktop/code"
if utils_dir not in sys.path:
    sys.path.insert(0, utils_dir)

from project_name_matcher import (
    load_or_init_db,
    match_project,
    prompt_unmatched_project_mappings,
    save_db,
)

from excel_formatting_utils import apply_common_excel_layout

# ====================== 配置区 ======================
excel_path = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9.1_材料欠款/A_材料欠款_raw/泰国子公司各项目采购欠款汇总7.9.xlsx"
db_path = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"

# 输出目录
output_dir = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/Z_report"
os.makedirs(output_dir, exist_ok=True)

sheet_name = "汇总表"

# 输出文件名（不再带时间戳，按要求固定名称）
output_path = os.path.join(output_dir, "泰国子公司材料采购欠款汇总_标准化.xlsx")

# ====================== 主程序 ======================
print("🚀 开始处理泰国子公司材料采购欠款汇总...\n")

# 1. 加载项目名称匹配数据库
db = load_or_init_db(db_path)
print(f"✅ 已加载匹配数据库，当前标准项目数量: {len(db.get('projects', {}))}")

# 2. 读取 Excel 文件
print(f"正在读取: {os.path.basename(excel_path)} → Sheet: {sheet_name}")

df = None
for header in [0, 1, 2]:
    try:
        temp_df = pd.read_excel(excel_path, sheet_name=sheet_name, header=header)
        col_str = [str(c).strip() for c in temp_df.columns]
        if any("项目名称" in c for c in col_str):
            df = temp_df
            print(f"✓ 使用第 {header+1} 行作为表头")
            break
    except:
        continue

if df is None:
    df = pd.read_excel(excel_path, sheet_name=sheet_name, header=0)
    print("⚠️ 使用默认第1行作为表头")

# 清理列名
df.columns = [str(col).strip() for col in df.columns]

# 关键列
project_col = "项目名称"
currency_col = "货币"
amount_col = "欠款金额（泰铢）"

# 检查列是否存在
for col in [project_col, currency_col, amount_col]:
    if col not in df.columns:
        print(f"❌ 错误：找不到列 '{col}'")
        print(f"当前可用列: {list(df.columns)}")
        sys.exit(1)

# 3. 数据清洗
df = df[~df[project_col].astype(str).str.contains("小计", na=False)]
df = df[df[project_col].notna() & (df[project_col].astype(str).str.strip() != "")]

print(f"过滤后有效记录数: {len(df)} 条")

# 清理金额
def clean_amount(val):
    if pd.isna(val):
        return 0.0
    if isinstance(val, str):
        val = val.replace(",", "").replace(" ", "").replace("，", "").strip()
    try:
        return float(val)
    except:
        return 0.0

df[amount_col] = df[amount_col].apply(clean_amount)

# 4. 项目名称匹配
print("\n🔍 正在进行项目名称匹配...")

records = []
unmatched_list = []

for _, row in df.iterrows():
    raw_name = str(row[project_col]).strip()
    if not raw_name:
        continue

    match_result = match_project(raw_name, db)
    
    if match_result:
        std_name = match_result["project_name"]
        project_code = match_result.get("project_code", "需手动核对")
    else:
        std_name = raw_name
        project_code = "需手动核对"
        unmatched_list.append(raw_name)

    records.append({
        "项目编码": project_code,
        "项目名称": std_name,           # 按要求使用“项目名称”
        "货币": str(row.get(currency_col, "THB")).strip(),
        "欠款金额（泰铢）": row[amount_col],
    })

# 5. 未匹配项目 → 手动匹配（自动保存到 db.json）
if unmatched_list:
    unique_unmatched = sorted(set(unmatched_list))
    print(f"\n⚠️ 发现 {len(unique_unmatched)} 个未自动匹配的项目，正在打开手动匹配窗口...")
    
    # 传入 db_path，让手动匹配后立即自动保存
    prompt_unmatched_project_mappings(unique_unmatched, db, db_path=db_path)
    
    # 重新匹配一次，确保使用最新映射
    for rec in records:
        if rec["项目名称"] == rec.get("原始项目名称", "") or rec["项目编码"] == "需手动核对":
            new_match = match_project(rec.get("原始项目名称", rec["项目名称"]), db)
            if new_match:
                rec["项目名称"] = new_match["project_name"]
                rec["项目编码"] = new_match.get("project_code", "需手动核对")

print(f"✅ 项目名称匹配完成！共处理 {len(records)} 条记录")

# 6. 按 项目编码 → 项目名称 → 货币 排序并汇总
final_df = pd.DataFrame(records)

# 先按项目编码和项目名称排序（确保总计在最后）
summary_df = final_df.groupby(["项目编码", "项目名称", "货币"], as_index=False)["欠款金额（泰铢）"].sum()

# 排序：项目编码为空的（总计）排在最后
summary_df = summary_df.sort_values(
    by=["项目编码", "项目名称", "货币"],
    key=lambda x: x.str.lower() if x.name in ["项目名称", "货币"] else x  # 忽略编码排序时的字符串问题
)

# 添加总计行
total_amount = summary_df["欠款金额（泰铢）"].sum()
total_row = pd.DataFrame([{
    "项目编码": "",
    "项目名称": "【总计】",
    "货币": "THB",
    "欠款金额（泰铢）": total_amount
}])
summary_df = pd.concat([summary_df, total_row], ignore_index=True)

# 7. 保存并美化 Excel
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    summary_df.to_excel(writer, index=False, sheet_name="欠款汇总")
    
    worksheet = writer.sheets["欠款汇总"]
    
    # 应用美化格式
    apply_common_excel_layout(
        worksheet,
        numeric_start_col=4,                    # 第4列开始是金额
        amount_headers=["欠款金额（泰铢）"],
        header_height=35,
        freeze_panes="A2",
        thousand_number_format="#,##0.00"
    )

# 保存更新后的匹配数据库
save_db(db_path, db)

print("\n" + "="*95)
print("🎉 处理完成！")
print(f"✅ 汇总文件已保存至：")
print(f"   {output_path}")
print(f"\n总欠款金额：{total_amount:,.2f} 泰铢")
print(f"涉及项目数量：{len(summary_df) - 1} 个")
print(f"项目匹配数据库（db.json）已更新并保存")
print("="*95)

print("\n💡 提示：下次运行时，之前手动匹配的结果会被自动记住，无需重复操作。")