# 9_采购数据 — 代码结构

## 目录

```text
9_采购数据/
├── 0_run_采购全流程.py     ★ 推荐入口
├── _采购路径.py            路径与公共工具
│
├── 流水线/                 ETL：raw → extract → std → sum
│   ├── 1_rename_excel_month_filenames.py
│   ├── 2_采购台账月度去重.py
│   ├── 3_采购部excel格式标准化.py
│   ├── 4_采购部数据汇总.py
│   └── 5_项目采购分类汇总.py
│
├── 规格标准化/             物品名称/规格规则（桩、钢筋、网片、砖…）
├── 单位转换/               数量单位换算（kg / m² / m³）
├── 内容标准化/             步骤4 调用：单位纠错 + 规格流水线
│
├── 分析看板/               钢筋价格、桩成本 HTML/Excel
└── 工具/                   查询、精准归类、金额增量等旁路脚本
```

## 数据盘（`邯郸中材/01成本统计/9_采购台账📦`）

| 目录 | 含义 |
|------|------|
| A_…_raw | 月度原始台账 |
| A_采购台账提取 | 去重后每项目 1 文件 |
| B_…_std | 标准化明细 `*_processed.xlsx` |
| D_…_sum | 汇总 `*_std_merged.xlsx` |
| Z_report/C_采购数据分析 | 分析输出 |

## 用法

```bash
cd /Users/wyattsun/Desktop/code/9_采购数据

# 查看结构与路径
PYTHONPATH=.. python3 0_run_采购全流程.py --tree
PYTHONPATH=.. python3 0_run_采购全流程.py --show-paths

# 全流程 1→4
PYTHONPATH=.. python3 0_run_采购全流程.py

# 仅汇总标准化
PYTHONPATH=.. python3 0_run_采购全流程.py --only 4

# 含分类汇总 + 分析看板
PYTHONPATH=.. python3 0_run_采购全流程.py --with-5 --with-analysis
```

也可直接运行子模块，例如：

```bash
PYTHONPATH=..:. python3 流水线/4_采购部数据汇总.py
PYTHONPATH=..:. python3 分析看板/0_汇总分析批量运行.py
```

## 模块关系

```text
流水线/1→2→3→4
              └─ 内容标准化/采购内容标准化_utils
                    ├─ 规格标准化/*（process_worksheet）
                    └─ 单位转换/*（try_convert）
分析看板/*  ← 读 D_sum
工具/*      ← 独立旁路
```
