# 机具费处理（重构版）

## 结构

```text
10_机具费/
├── 00_机具费处理入口.py     # 推荐入口
├── machine/                 # 核心包
│   ├── config.py            # 路径与常量
│   ├── merge_raw.py         # 原 1_：raw 合并
│   ├── summarize.py         # 原 2_：std→sum 汇总
│   ├── split.py             # 原两个 4_：拆分
│   ├── report.py            # 原 html 报告
│   ├── mapping.py           # 机具/区域映射
│   ├── validate.py          # 金额校验与纠错
│   ├── columns.py / textutil / discover
│   └── __main__.py          # CLI
├── 机具映射表.json
├── 1_merge_raw.py           # 兼容 shim
├── 2_机具使用汇总.py
├── 4_*.py / 机具费html报告.py
└── ARCHITECTURE.md
```

## 流水线

```text
A_机具费_raw/*.xlsx
        │ merge-raw
        ▼
B_机具费_std/
  · 机具费_raw合并.xlsx              ← 过程文件（唯一做项目名称匹配）
  · {编码}_{项目名}_机具费.xlsx      ← 按项目拆分，仅「明细」sheet（无按月子表）
        │ summarize（不弹窗匹配；用文件内/文件名项目字段）
        ▼
B_机具费_sum/  机具费汇总_*_汇总.xlsx
        │ report
        ▼
C_机具租赁分析报告/
```

> 不再生成 `clean_机具费_raw合并.xlsx`。  
> 项目匹配只在 `merge-raw`；`summarize` 不再弹出手动匹配。

## 用法

```bash
cd "/Users/wyattsun/Desktop/code/10_机具费"

python3 00_机具费处理入口.py paths
python3 00_机具费处理入口.py              # all
python3 00_机具费处理入口.py merge-raw
python3 00_机具费处理入口.py summarize
python3 00_机具费处理入口.py report
python3 00_机具费处理入口.py split-sum
python3 00_机具费处理入口.py split-ledger

# 等价
python3 -m machine summarize
```

环境变量：

```bash
export COST_DATA_ROOT=".../01成本统计"
export MACHINE_RUN_MODE=all          # 或 monthly_pair
export MACHINE_EXCEL_OVERRIDE=       # 可选：强制汇总某个文件
```

## 配置与映射

| 文件 | 用途 |
|------|------|
| `machine/config.py` | 数据路径、sheet 名、运行模式 |
| `机具映射表.json` | 机具种类/型号别名 |
| `区域映射表.json` | 项目→区域（若有） |
| `../0_项目基本信息/project_name_match_db.json` | 项目别名 |

## 兼容

旧命令仍可用（shim 转发到 `machine.*`）。
