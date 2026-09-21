# 项目基本信息 · 共用工具

全仓库（成本 / 财务 / 人工 / 机具 / 搅拌站…）共用的**项目名称 → 标准名 + 编码**匹配。

## 文件

| 文件 | 作用 |
|------|------|
| `project_name_matcher.py` | **核心模块**（匹配 + 弹窗 + JSON） |
| `project_name_match_db.json` | 标准项目 + 别名库 |
| `project_db_manager.py` | 从「项目基本信息表」刷新库的入口 |
| `excel_formatting_utils.py` | Excel 通用样式 |
| `unit_map_utils.py` | 单位换算 |

## 推荐用法（新代码）

```python
import sys
sys.path.insert(0, "/Users/wyattsun/Desktop/code/0_项目基本信息")

from project_name_matcher import ProjectMatcher

m = ProjectMatcher()                 # 自动加载 JSON，有主数据表则合并
hit = m.match("永强")                # dict 或 None
name, code = m.resolve("永强")       # ("永强车业项目", "THCC-2025-005")

# 批量 + 未匹配弹窗
results = m.match_batch(["永强", "新项目A"], prompt=True)

# DataFrame 一列搞定
df = m.enrich_dataframe(df, col="项目名称", prompt=True)

m.save()                             # 写回 JSON
```

更短：

```python
from project_name_matcher import load_matcher
m = load_matcher()
print(m.resolve("樱花"))
```

## 图形界面（选 Excel / 列 / 填写）

```bash
cd /Users/wyattsun/Desktop/code/0_项目基本信息
python project_name_matcher.py          # 无参数即打开界面
python project_name_matcher.py gui
python project_name_matcher_gui.py
```

在界面里选择 Excel、待匹配列、标准名称输出列、项目编码输出列；可选「处理摘要列」。
只改这些列的单元格值，列宽、样式、日期、筛选等原格式保持不变。

摘要处理规则：若摘要以待匹配列的项目名称开头，则删去该前缀；删完后若以 `-`（或同类连接符）开头，再删掉连接符。

```bash
python project_name_matcher.py excel 台账.xlsx --process-summary
python project_name_matcher.py excel 台账.xlsx --match 项目名称 --name-out 项目标准名称 --code-out 项目编码 --process-summary --prompt
```

## 命令行

```bash
cd /Users/wyattsun/Desktop/code/0_项目基本信息
python project_name_matcher.py match 永强
python project_name_matcher.py refresh      # 从项目基本信息表更新 JSON
python project_name_matcher.py stats
python project_name_matcher.py demo 永强 樱花
```

## 旧函数式 API（仍可用）

```python
from project_name_matcher import (
    load_or_init_db, match_project, save_db,
    prompt_unmatched_project_mappings,
    build_standard_projects_from_dataframe,
)
db = load_or_init_db(db_path, standard_projects)
hit = match_project(raw, db)
save_db(db_path, db)
```

## 默认路径

- 匹配库：`0_项目基本信息/project_name_match_db.json`
- 主数据：`邯郸中材/01成本统计/0_项目基本信息情况/项目基本信息表.xlsx`

## 匹配顺序

1. 精确别名 / 标准名  
2. 去空格归一化  
3. 软匹配（清洗全等 / 唯一子串 / 高相似度）  
4. 曾「保留原名」→ 自动批处理静默原名；**手动运行 / 勾选弹窗补录时仍会弹出对话框**  

软匹配成功会自动写入别名，下次精确命中。
