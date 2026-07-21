# 机具费架构（1.0 重构）

## 设计目标

1. **单配置**：所有路径走 `machine.config` + `settings.DATA_ROOT`
2. **单入口**：`00_机具费处理入口.py` / `python -m machine`
3. **职责分离**：merge / summarize / split / report 独立模块
4. **业务保真**：汇总核心逻辑从原 `2_机具使用汇总.py` 迁入 `summarize.py`，保留映射、校验、着色导出
5. **兼容 shim**：根目录 1/2/4/html 仍可直接运行

## 模块职责

| 模块 | 职责 |
|------|------|
| `config` | RAW/STD/SUM/REPORT、映射路径、RUN_MODE |
| `merge_raw` | 多月 raw 列标准化、合并、匹配库补项目编码 |
| `discover` | 找最新 std、按项目/月份分组文件 |
| `mapping` | 区域/机具 JSON 加载保存、智能种类规则 |
| `columns` | 表头探测、列模糊匹配、sheet 项目信息 |
| `validate` | 加班费/结算/合计校验，台班与供应商月份纠错 |
| `textutil` | 文本归一、金额解析、年月归一 |
| `summarize` | 两遍扫描匹配 → 明细合并 → 报告/宽表 → Excel |
| `split` | sum 按项目拆；raw 合并按项目拆 |
| `report` | Plotly HTML + Excel 统计 |

## 数据约定

| 阶段 | 目录 | 文件 |
|------|------|------|
| raw | `A_机具费_raw` | 业务原始 xlsx |
| 合并 | `B_机具费_std` | `机具费_raw合并.xlsx` |
| 项目 std | `B_机具费_std` | `{YYYYMM}_{项目}机具费.xlsx` |
| 汇总 | `B_机具费_sum` | `机具费汇总_*_汇总.xlsx` |
| 报告 | `C_机具租赁分析报告` | 带时间戳 html/xlsx |

## 后续可优化

- 为 `summarize.main` 再拆 `transform` / `export` 子模块
- 增加金额回归测试（固定样例 std → sum 合计）
- 区域映射 JSON 若缺失时自动从项目主数据生成
