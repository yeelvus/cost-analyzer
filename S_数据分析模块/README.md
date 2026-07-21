# S_.据分析模块

汇总「看数据 / 出报告 / 出看板」类脚本。各业务领域的 **raw→std 标准化流水线** 仍留在编号目录（`3_预算`、`5_合同`、`9_采购`…），本模块只做 **聚合后的分析与呈现**。

## 目录

```text
S_数据分析模块/
├── 00_分析入口.py          # 推荐入口
├── _bootstrap.py           # 自动挂载 code 根到 sys.path
├── dashboards/             # HTML 看板
│   ├── build_full_lifecycle_dashboard.py   # 全过程成本看板（推荐）
│   ├── build_unified_cost_dashboard.py     # 综合成本看板（单入口文件夹）
│   ├── build_cost_kanban.py                # 卡片看板
│   ├── build_comprehensive_cost_analysis_html.py  # 多层级钻取综合报告
│   ├── build_company_cost_monthly_html.py  # 报公司表 → 月度进度 HTML
│   └── build_budget_vs_process.py          # 预算 vs 过程 人材机对比
├── analysis/               # Excel/文本分析
│   ├── 成本_数据分析.py
│   ├── 预计毛利率对比_数据统计与环比.py
│   └── 3_分包付款汇总分析_按日期与月增量.py
├── reports/                # 配置驱动合同/成本报告
│   ├── config_driven_report_builder.py
│   ├── contract_report_*.py
│   ├── generate_report_config.py
│   └── cost_analysis_report_config.json
└── docs/                   # 设计说明
```

## 用法

```bash
cd "/Users/wyattsun/Desktop/code"

# 列出命令
python3 S_数据分析模块/00_分析入口.py --list

# 全过程看板
python3 S_数据分析模块/00_分析入口.py dashboard lifecycle

# 综合成本看板
python3 S_数据分析模块/00_分析入口.py dashboard unified

# 预算 vs 过程
python3 S_数据分析模块/00_分析入口.py dashboard budget-vs-process

# 成本差异分析 / 毛利率环比
python3 S_数据分析模块/00_分析入口.py analysis cost
python3 S_数据分析模块/00_分析入口.py analysis margin

# 旧路径仍可用（根目录 shim 转发）
python3 build_unified_cost_dashboard.py
python3 2_成本代码/成本_数据分析.py
```

输出一般在数据目录 `01成本统计/Z_report/`（见根目录 `settings.py`）。

## 与流水线的关系

```text
领域 ETL（编号目录）          中央聚合（2_成本代码）         本模块
raw → std → sum      →   成本_汇总按月收集…master  →  dashboards / analysis / reports
合同/预算/采购/…              0_一键运行_成本流程.py         00_分析入口.py
```

一键成本流程仍会调用本模块中的「成本分析 + 综合 HTML」。
