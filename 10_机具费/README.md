# 机具费处理

## 数据目录（按 A→F 顺序）

```text
10_机具费🚜/
├── A_机具费_raw/              原始台账（输入）
├── B_机具费_raw_sum/          机具费_raw合并.xlsx（过程合并 + 项目匹配）
├── C_机具费_raw_split/        原始拆分 {编码}_{项目}_机具费.xlsx
├── D_机具费_std_sum/          仅 机具费汇总_全量标准化_汇总.xlsx
├── E_机具费_std_split/        标准拆分（标准汇总按项目）
└── F_机具费_report/           分析报告 HTML / Excel
```

## 流水线

```text
A_机具费_raw
    │ merge-raw（仅此步做项目名称匹配）
    ▼
B_机具费_raw_sum/机具费_raw合并.xlsx
    │ 自动
    ▼
C_机具费_raw_split/{编码}_{项目}_机具费.xlsx
    │ summarize（不弹窗匹配；汇总 sheet 置首）
    ▼
D_机具费_std_sum/机具费汇总_全量标准化_汇总.xlsx
    │ 自动
    ▼
E_机具费_std_split/{编码}_{项目}_机具费.xlsx
    │ report
    ▼
F_机具费_report/
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
