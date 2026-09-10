import re
import sys
import os
from datetime import datetime
import pandas as pd

# ====================== 模块路径修复 ======================
matcher_dir = "/Users/wyattsun/Desktop/code/0_项目基本信息"
if matcher_dir not in sys.path:
    sys.path.insert(0, matcher_dir)

# excel_formatting_utils.py 与匹配库同目录
utils_dir = "/Users/wyattsun/Desktop/code/0_项目基本信息"
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
_CODE_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _CODE_ROOT not in sys.path:
    sys.path.insert(0, _CODE_ROOT)
from common.paths import iter_excel_files, raw_data_path  # noqa: E402

_raw_dir = raw_data_path(
    "材料欠款",
    legacy="/Users/wyattsun/Desktop/邯郸中材/01成本统计/9.1_材料欠款/A_材料欠款_raw",
)
_raw_files = iter_excel_files(_raw_dir)
excel_path = (
    str(_raw_files[0])
    if _raw_files
    else os.path.join(str(_raw_dir), "泰国子公司各项目采购欠款汇总7.31.xlsx")
)
db_path = "/Users/wyattsun/Desktop/code/0_项目基本信息/project_name_match_db.json"

# 输出目录（模块自己的 std，不写 Z_report）
output_dir = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/9.1_材料欠款/B_材料欠款_std"
os.makedirs(output_dir, exist_ok=True)

sheet_name = "汇总表"

# 输出文件名（不再带时间戳，按要求固定名称）
output_path = os.path.join(output_dir, "泰国子公司材料采购欠款汇总_标准化.xlsx")

# 非交互模式：存在 --non-interactive / -y 时不弹窗
INTERACTIVE = not any(a in ("--non-interactive", "-y", "--yes") for a in sys.argv[1:])

# 汇总表中的分组标题 / 小计 / 总计行（不应作为项目明细）
# 匹配：含「小计/合计/总计」、或以「：/:」结尾的分段标题（如 在建项目：、已完工项目：）
SECTION_OR_TOTAL_RE = re.compile(
    r"(小计|合计|总计|汇总)|(?:项目)?[：:]\s*$"
)


def is_section_or_total_row(name) -> bool:
    """判断是否为分组标题、小计、总计等非项目明细行。"""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return True
    text = str(name).strip()
    if not text or text.lower() == "nan":
        return True
    if SECTION_OR_TOTAL_RE.search(text):
        return True
    # 纯分段名（无冒号时的兜底）
    if text in {"在建项目", "已完工项目", "新开工项目", "完工项目"}:
        return True
    return False


def clean_amount(val) -> float:
    if pd.isna(val):
        return 0.0
    if isinstance(val, str):
        val = val.replace(",", "").replace(" ", "").replace("，", "").strip()
    try:
        return float(val)
    except (TypeError, ValueError):
        return 0.0


def normalize_project_code(code) -> str:
    if code is None or (isinstance(code, float) and pd.isna(code)):
        return "需手动核对"
    text = str(code).strip()
    return text if text else "需手动核对"


def normalize_currency(val, default: str = "泰铢") -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    text = str(val).strip()
    if not text or text.lower() == "nan":
        return default
    # 统一展示
    if text.upper() in {"THB", "BAHT"}:
        return "泰铢"
    return text


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
            print(f"✓ 使用第 {header + 1} 行作为表头")
            break
    except Exception:
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

# 3. 数据清洗：去掉空行、分组标题、小计/总计（避免金额重复计入）
before_n = len(df)
df = df[df[project_col].notna()].copy()
df[project_col] = df[project_col].astype(str).str.strip()
df = df[df[project_col] != ""]
df = df[df[project_col].str.lower() != "nan"]

dropped_sections = df[df[project_col].map(is_section_or_total_row)]
if len(dropped_sections) > 0:
    print("\n🧹 已排除分组/小计/总计行（避免重复汇总）：")
    for _, r in dropped_sections.iterrows():
        amt = clean_amount(r.get(amount_col))
        print(f"   - {r[project_col]!r}  金额={amt:,.2f}")

df = df[~df[project_col].map(is_section_or_total_row)].copy()
print(f"\n过滤后有效项目记录数: {len(df)} 条（原始 {before_n} 行）")

# 清理金额
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
        project_code = normalize_project_code(match_result.get("project_code"))
        # 库内存在但编码仍为「需手动核对」/空：仍记入待补清单
        if project_code == "需手动核对":
            unmatched_list.append(raw_name)
    else:
        std_name = raw_name
        project_code = "需手动核对"
        unmatched_list.append(raw_name)

    records.append({
        "原始项目名称": raw_name,
        "项目编码": project_code,
        "项目名称": std_name,
        "货币": normalize_currency(row.get(currency_col, "泰铢")),
        "欠款金额（泰铢）": row[amount_col],
    })

# 5. 未匹配项目 → 手动匹配（自动保存到 db.json）
if unmatched_list:
    unique_unmatched = sorted(set(unmatched_list))
    print(f"\n⚠️ 发现 {len(unique_unmatched)} 个需补全编码/映射的项目：{unique_unmatched}")
    if INTERACTIVE:
        print("正在打开手动匹配窗口...")
        prompt_unmatched_project_mappings(unique_unmatched, db, db_path=db_path)
    else:
        print("（非交互模式：跳过弹窗，保留「需手动核对」）")
        prompt_unmatched_project_mappings(
            unique_unmatched, db, db_path=db_path, interactive=False
        )

    # 重新匹配一次，使用最新映射（必须用原始名称）
    for rec in records:
        raw = rec.get("原始项目名称") or rec["项目名称"]
        new_match = match_project(raw, db)
        if new_match:
            rec["项目名称"] = new_match["project_name"]
            rec["项目编码"] = normalize_project_code(new_match.get("project_code"))

print(f"✅ 项目名称匹配完成！共处理 {len(records)} 条记录")

# 6. 按 项目编码 → 项目名称 → 货币 汇总
final_df = pd.DataFrame(records)
if final_df.empty:
    print("❌ 没有可汇总的有效记录")
    sys.exit(1)

summary_df = (
    final_df.groupby(["项目编码", "项目名称", "货币"], as_index=False)["欠款金额（泰铢）"]
    .sum()
)

# 排序：正式编码在前，「需手动核对」在后；同组内按项目名称
def _sort_key_frame(frame: pd.DataFrame) -> pd.DataFrame:
    out = frame.copy()
    out["_code_rank"] = out["项目编码"].map(
        lambda c: 1 if str(c).strip() in {"", "需手动核对", "nan"} else 0
    )
    out["_code"] = out["项目编码"].astype(str).str.lower()
    out["_name"] = out["项目名称"].astype(str).str.lower()
    out["_cur"] = out["货币"].astype(str).str.lower()
    out = out.sort_values(by=["_code_rank", "_code", "_name", "_cur"]).drop(
        columns=["_code_rank", "_code", "_name", "_cur"]
    )
    return out.reset_index(drop=True)


summary_df = _sort_key_frame(summary_df)

# 添加总计行（仅项目明细合计，不含源表分组小计）
total_amount = float(summary_df["欠款金额（泰铢）"].sum())
# 若源表存在「总计」行，可与之对照校验
try:
    raw_check = pd.read_excel(excel_path, sheet_name=sheet_name, header=1)
    raw_check.columns = [str(c).strip() for c in raw_check.columns]
    if amount_col in raw_check.columns and project_col in raw_check.columns:
        total_rows = raw_check[
            raw_check[project_col].astype(str).str.contains("总计", na=False)
        ]
        if len(total_rows) > 0:
            source_total = clean_amount(total_rows.iloc[0][amount_col])
            diff = abs(total_amount - source_total)
            if diff > 0.02:
                print(
                    f"\n⚠️ 汇总金额与源表「总计」不一致："
                    f"本表 {total_amount:,.2f} vs 源表 {source_total:,.2f} "
                    f"(差额 {total_amount - source_total:,.2f})"
                )
            else:
                print(f"\n✓ 汇总金额与源表「总计」一致：{total_amount:,.2f} 泰铢")
except Exception as exc:
    print(f"\n（跳过与源表总计核对：{exc}）")

total_row = pd.DataFrame([{
    "项目编码": "",
    "项目名称": "【总计】",
    "货币": "泰铢",
    "欠款金额（泰铢）": total_amount,
}])
summary_df = pd.concat([summary_df, total_row], ignore_index=True)

# 输出明细时去掉内部字段「原始项目名称」
export_df = summary_df[["项目编码", "项目名称", "货币", "欠款金额（泰铢）"]].copy()

# 7. 保存并美化 Excel
with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
    export_df.to_excel(writer, index=False, sheet_name="欠款汇总")

    worksheet = writer.sheets["欠款汇总"]

    apply_common_excel_layout(
        worksheet,
        numeric_start_col=4,
        amount_headers=["欠款金额（泰铢）"],
        header_height=35,
        freeze_panes="A2",
        thousand_number_format="#,##0.00",
    )

# 保存更新后的匹配数据库
save_db(db_path, db)

project_count = len(export_df) - 1  # 不含总计行
print("\n" + "=" * 95)
print("🎉 处理完成！")
print(f"✅ 汇总文件已保存至：")
print(f"   {output_path}")
print(f"\n总欠款金额：{total_amount:,.2f} 泰铢")
print(f"涉及项目数量：{project_count} 个")
print(f"项目匹配数据库（db.json）已更新并保存")
print("=" * 95)

print("\n💡 提示：下次运行时，之前手动匹配的结果会被自动记住，无需重复操作。")
print("   非交互重跑可加参数: python 材料欠款汇总.py --non-interactive")
