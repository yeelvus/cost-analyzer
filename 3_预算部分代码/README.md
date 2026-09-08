# 预算处理模块（budget）

将原分散的几个 Python 脚本重构为可扩展包，统一入口、配置与单位换算。

## 结构

```text
3_预算部分代码/
├── budget/                      # 核心包（新功能加在这里）
│   ├── __init__.py
│   ├── __main__.py              # python -m budget
│   ├── config.py                # 路径与业务常量
│   ├── units.py                 # 单位规范化（吨/kg/m³…）
│   ├── glodon/
│   │   ├── __init__.py
│   │   └── pipeline.py          # 广联达 raw→std + 总汇总
│   ├── dashboard.py             # HTML 看板
│   └── responsibility.py        # 责任成本
├── 00_预算处理入口.py            # 推荐入口
├── 广联达预算人材机文件统计.py    # 兼容旧入口
├── 人材机总汇总看板生成.py
├── process_responsibility_costs_complete.py
├── tests/
│   └── test_units_rebar.py
└── README.md
```

## 数据流

```text
1_广联达导出人材料机_raw
        │  budget glodon
        ▼
2_人材机_std/*_processed.xlsx
        + 【总汇总】所有标段.xlsx
        │  budget dashboard
        ▼
总汇总项目看板.html

3.1_责任成本/1_责任成本_raw
        │  budget responsibility
        ▼
3.1_责任成本/2_责任成本_std + 完整处理报告.xlsx
```

默认 `all` 会按顺序跑完：**广联达 → 看板 → 责任成本**。

## 用法

```bash
cd "/Users/wyattsun/Desktop/code/3_预算部分代码"

# 默认全流程：广联达 + 看板 + 责任成本
python3 00_预算处理入口.py
# 或
python3 -m budget all

# 仅广联达
python3 -m budget glodon
python3 -m budget glodon -i /path/raw -o /path/std

# 仅看板
python3 -m budget dashboard
python3 -m budget dashboard --summary-file /path/to/总汇总.xlsx

# 仅责任成本
python3 -m budget responsibility
# 或兼容旧脚本
python3 process_responsibility_costs_complete.py

# 单位换算自检
python3 -m budget units-test
python3 tests/test_units_rebar.py
```

## 配置

编辑 `budget/config.py`，或设置环境变量：

```bash
export COST_DATA_ROOT="/Users/wyattsun/Desktop/邯郸中材/01成本统计"
```

主要路径：

| 配置 | 含义 |
|------|------|
| `GLODON_RAW_DIR` | 广联达原始导出 |
| `GLODON_STD_DIR` | 标准化 + 总汇总输出 |
| `RESP_RAW_DIR` / `RESP_STD_DIR` | 责任成本 |

## 扩展新功能

1. 在 `budget/` 下新建模块，例如 `budget/compare_versions.py`
2. 在 `budget/__main__.py` 的 `build_parser()` 里增加子命令
3. 需要配置时写入 `budget/config.py`
4. 与数量相关的单位逻辑 **只** 调用 `budget.units`

示例子命令骨架：

```python
def cmd_my_feature(args):
    from budget import my_feature
    my_feature.run(...)

# in build_parser:
p = sub.add_parser("my-feature", help="...")
p.set_defaults(func=cmd_my_feature)
```

## 钢筋单位

- 基准单位：**kg**
- `1 吨 / t = 1000 kg`
- 未知单位不进入合计，写入总汇总 **「单位异常明细」** Sheet

## 依赖

- pandas, openpyxl
- 处理旧版 `.xls` 时：`pip install 'xlrd<2'`
