# Desktop/code 代码目录说明

> 整理日期：2026-07-21  
> 代码根：`/Users/wyattsun/Desktop/code`  
> 数据根：`settings.DATA_ROOT` → 默认 `Desktop/邯郸中材/01成本统计`

## 一、总览心智模型

```text
0_Raw（数据根下）   全部原始收集；目录名可带 0_/1_ 序号
0_项目基本信息     主数据：项目编码/别名匹配/单位映射
编号业务目录       std / sum（ETL 产物；raw 已集中到 0_Raw）
2_成本代码         中央聚合 master 宽表 + 一键流程
S_数据分析模块     看板 / 分析 / 报告（消费 std/sum/master）
common / settings  公共路径与修正工具
```

**原则**：业务标准化脚本留在编号目录；「读结果做分析、出 HTML/Excel 报告」放进 `S_数据分析模块`。

---

## 二、目录职责一览

| 目录 | 类型 | 作用 |
|------|------|------|
| `0_项目基本信息/` | 主数据 | 项目名匹配库、单位映射、Excel 格式化、进度图 |
| `2_成本代码/` | ETL 聚合 | 月度成本模板刮取 + 多源 join → master；一键运行 |
| `3_预算部分代码/` | 领域 ETL | 广联达人材机、看板、责任成本（`budget` 包 + `00_预算处理入口`） |
| `5_合同管理代码/` | 领域 ETL | 合同文件夹、明细/金额汇总、签证 |
| `5.1_分包合同台账处理/` | 领域 ETL | 分包台账 1+2+4 全量合并 |
| `9_采购数据/` | 领域 ETL | 采购台账清洗、规格标准化、分类汇总、专项分析 |
| `9.1_搅拌站/` | 领域 ETL | 混凝土汇总 |
| `10_机具费/` | 领域 ETL | 机具台账合并、使用汇总、按项目拆分、HTML |
| `11_人工费/` | 领域 ETL | 人工费/机票汇总 |
| `12_财务数据/` | 领域 ETL | 财务付款、分包付款台账 |
| `13_结算整理/` | 领域 ETL | 业主结算额空 JSON + 汇总 Excel（汇率折泰铢） |
| `S_数据分析模块/` | **分析** | 全过程/综合看板、成本分析、配置驱动报告 |
| `common/` | 公共库 | paths、corrections、excel、project |
| `templates/` | 资源 | 看板 HTML 模板与 chart.js |
| `settings.py` | 配置 | `CODE_ROOT` / `DATA_ROOT` / `REPORT_ROOT`(Excel) / `0_html`(看板) |
| `WORKFLOW.md` | 文档 | 月度增量与治理目标架构 |

---

## 三、各模块关键脚本

### 0_项目基本信息
- `project_name_matcher.py` + `project_name_match_db.json`：别名 → 标准名/编码
- `unit_map_utils.py`：单位规范化
- `excel_formatting_utils.py`：Excel 样式工具
- `项目进度图绘制.py`：进度可视化

### 2_成本代码（聚合，非纯分析）
- `成本_汇总按月收集成本统计.py`：**中央 master** 刮取与 join
- `0_一键运行_成本流程.py`：汇总 → 分析 → 毛利率 → 综合 HTML
- `成本_数据分析.py` / `预计毛利率…`：**shim**，实体在 `S_数据分析模块/analysis/`

### 3_预算部分代码
- `00_预算处理入口.py` / `budget/`：广联达 + 看板 + 责任成本（默认 `all` 全跑）

### 5 / 5.1 合同
- `2_合同汇总输出/合同汇总_全量合并输出.py`：默认先跑 5.1 台账再合并总包分包
- `5.1/00_全量合并汇总.py`：分包台账 1+2+4

### 9 / 10 / 11 / 12
- 各自 raw→std→sum 流水线；领域 HTML（如机具、人工）仍可在本目录，跨域综合看板在 S 模块

### S_数据分析模块（本次整理重点）

| 子目录 | 脚本 | 作用 |
|--------|------|------|
| dashboards | `build_full_lifecycle_dashboard.py` | 全过程多项目看板（报公司表优先） |
| dashboards | `build_unified_cost_dashboard.py` | 综合看板单入口文件夹 |
| dashboards | `build_cost_kanban.py` | 卡片式成本看板 |
| dashboards | `build_comprehensive_cost_analysis_html.py` | master 多层级钻取 HTML |
| dashboards | `build_company_cost_monthly_html.py` | 报公司月度进度页 |
| dashboards | `build_budget_vs_process.py` | 预算 vs 过程人材机 |
| analysis | `成本_数据分析.py` | 项目差异文本/汇总报告 |
| analysis | `预计毛利率对比_数据统计与环比.py` | 毛利率环比 |
| analysis | `3_分包付款汇总分析_…` | 分包付款按日/月增量 |
| reports | `config_driven_*` / `contract_report_*` | 配置驱动成本/合同报告 |

根目录同名 `build_*.py` 等为 **兼容 shim**，自动转发到 S 模块。

---

## 四、推荐运行顺序（月度）

1. 各领域更新 raw 后跑领域脚本（采购/机具/人工/财务/合同/预算）  
2. `2_成本代码/0_一键运行_成本流程.py`（master + 分析 + 综合 HTML）  
3. 需要管理层决策面时：  
   `python S_数据分析模块/00_分析入口.py dashboard lifecycle`  
   或 `dashboard unified`

---

## 五、未迁入 S 的说明

| 保留位置 | 原因 |
|----------|------|
| `2_成本代码/成本_汇总按月收集…` | 数据工程/ETL，不是分析呈现 |
| `9_采购数据/采购部数据分析代码/` | 强依赖采购 std 路径，领域内专项 |
| `10_机具费/*html*`、`11_人工费/*` | 领域专项报告，输入输出紧耦合领域 sum |
| `3_预算…/budget/dashboard.py` | 预算流水线内嵌步骤，由预算入口调用 |

若后续要把领域专项 HTML 也统一入口，可在 `00_分析入口.py` 增加 `domain` 分组做转发。

---

## 六、兼容性

- 旧命令 `python build_xxx.py` 仍可用（shim）  
- 旧路径 `2_成本代码/成本_数据分析.py` 仍可用（shim）  
- 新推荐：`S_数据分析模块/00_分析入口.py`
