# 项目基本信息 · 共用工具

本目录保留与「项目基本信息表」相关的 Excel 样式、单位换算和同步脚本。

项目名称匹配已迁到 `code/common/`，全仓库统一从那里调用。

## 本目录文件

| 文件 | 作用 |
|------|------|
| `excel_formatting_utils.py` | Excel 通用样式 |
| `unit_map_utils.py` / `unit_map.json` | 单位换算 |
| `copy_latest_cost_excel.py` | 同步最新成本 Excel |
| `项目进度图绘制.py` | 进度图 |
| `project_db_manager.py` | 从「项目基本信息表」刷新匹配库（薄入口） |

## 项目名称匹配（common）

```python
from common.project import ProjectMatcher

m = ProjectMatcher()                 # 自动加载 JSON，有主数据表则合并
hit = m.match("永强")                # dict 或 None
name, code = m.resolve("永强")       # ("永强车业项目", "THCC-2025-005")

results = m.match_batch(["永强", "新项目A"], prompt=True)
df = m.enrich_dataframe(df, col="项目名称", prompt=True)
m.save()
```

更短：

```python
from common.project import load_matcher
m = load_matcher()
print(m.resolve("樱花"))
```

图形界面 / 命令行：

```bash
cd /Users/wyattsun/Desktop/code
python -m common.project_name_matcher
python -m common.project_name_matcher gui
python -m common.project_name_matcher_gui
python -m common.project_name_matcher match 永强
python -m common.project_name_matcher refresh
python -m common.project_name_matcher stats
python -m common.project_name_matcher excel 台账.xlsx --process-summary --prompt
```

匹配库路径：`code/common/project_name_match_db.json`  
主数据：`邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx`
