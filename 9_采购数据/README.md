# 9_采购数据 — 代码结构

## 目录

```text
9_采购数据/
├── 0_run_采购全流程.py     ★ 推荐入口
├── _采购路径.py            路径与公共工具
│
├── 流水线/                 ETL
│   ├── 1_rename_…py        ★ 手动
│   ├── 2_采购台账月度去重.py ★ 手动
│   ├── 3_采购部excel格式标准化.py ★ 手动
│   ├── 4_采购部数据汇总.py        ← 默认只跑这一步
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
| C_…_raw_sum | 纯合并 `采购部采购台账_merged.xlsx` |
| D_…_sum | 标准化汇总 `*_std_merged.xlsx`、`项目采购分类汇总.xlsx`（含按月「采购总金额」） |
| Z_report/C_采购数据分析 | 分析输出 |

## 用法

```bash
cd /Users/wyattsun/Desktop/code/9_采购数据

# 查看结构与路径
PYTHONPATH=.. python3 0_run_采购全流程.py --tree
PYTHONPATH=.. python3 0_run_采购全流程.py --show-paths

# 日常默认：仅 4 汇总（不含 1/2/3）
PYTHONPATH=.. python3 0_run_采购全流程.py

# 仅汇总（同上）
PYTHONPATH=.. python3 0_run_采购全流程.py --only 4

# 手动步骤（需要时单独跑）
PYTHONPATH=.. python3 0_run_采购全流程.py --only 1   # 重命名
PYTHONPATH=.. python3 0_run_采购全流程.py --only 2   # 月度去重
PYTHONPATH=.. python3 0_run_采购全流程.py --only 3   # 格式标准化→std
# 或: python3 流水线/1_….py / 2_….py / 3_….py

# 一次跑完准备+汇总：1→2→3→4
PYTHONPATH=.. python3 0_run_采购全流程.py --with-prep

# 含分类汇总 + 分析看板
PYTHONPATH=.. python3 0_run_采购全流程.py --with-5 --with-analysis
```

### 项目匹配弹窗（步骤3）

| 行为 | 说明 |
|------|------|
| 默认 | 仅对**新出现**且无法自动/模糊匹配的名称弹窗一次 |
| 选「保持原名称」 | 写入匹配库 `unmatched_aliases`，**下次不再弹** |
| `--no-prompt` | 全程不弹窗，新名直接保留原名 |
| `--force-prompt` | 强制对当前无法匹配的名再弹一次 |

匹配顺序：精确别名 → 去空格 → 清洗（去「泰国/项目」）→ 唯一包含 → 高相似度。

也可直接运行子模块，例如：

```bash
PYTHONPATH=..:. python3 流水线/4_采购部数据汇总.py
PYTHONPATH=..:. python3 分析看板/0_汇总分析批量运行.py
```

## 模块关系

```text
【手动，按需】1 重命名 → 2 去重 → EXTRACT → 3 格式标准化 → STD
【默认日常】  流水线/4 汇总（读 STD → SUM）
                    └─ 内容标准化
                          ├─ 规格标准化/*
                          └─ 单位转换/*
分析看板/*  ← 读 D_sum
工具/*      ← 独立旁路
```
