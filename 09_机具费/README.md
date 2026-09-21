# 机具费处理

## 数据目录（按 A→F 顺序）

```text
01_Raw/10_机具费/                      原始台账（输入）
02_Std/10_机具费/
├── B_机具费_raw_sum/                  机具费_raw合并.xlsx（过程合并 + 项目匹配）
├── C_机具费_raw_split/                原始拆分 {编码}_{项目}_机具费.xlsx
└── E_机具费_std_split/                标准拆分（标准汇总按项目）
03_Sum/10_机具费/                      机具费_汇总.xlsx
04_html/机具费分析/                    分析报告 HTML
```

## 流水线

```text
A_机具费_raw
    │ merge-raw（仅此步做项目名称匹配）
    ▼
02_Std/10_机具费/B_机具费_raw_sum/机具费_raw合并.xlsx
    │ 自动
    ▼
02_Std/10_机具费/C_机具费_raw_split/{编码}_{项目}_机具费.xlsx
    │ summarize（不弹窗匹配；汇总 sheet 置首）
    ▼
03_Sum/10_机具费/机具费_汇总.xlsx
    │ 自动
    ▼
02_Std/10_机具费/E_机具费_std_split/{编码}_{项目}_机具费.xlsx
    │ report
    ▼
04_html/机具费分析/
```

## 用法

```bash
cd "/Users/wyattsun/Desktop/code/10_机具费"

python3 00_机具费处理入口.py paths      # 查看/创建 A→F 目录，迁移旧文件
python3 00_机具费处理入口.py            # 全流程
python3 00_机具费处理入口.py merge-raw  # A → B + C
python3 00_机具费处理入口.py summarize  # C → D + E
python3 00_机具费处理入口.py report     # D → F
python3 00_机具费处理入口.py split-ledger  # 仅 B→C
python3 00_机具费处理入口.py split-sum     # 仅 D→E
```
