# class_master 拆分 PR 计划（函数级迁移表）

目标：把 `learn_classify/class_master.py`（约 1080 行）拆成可测、可独立演进的小模块，**行为零变更**。

原则：

1. 每个 PR 可单独合并、可单独回滚  
2. 对外 API 保持：`export_class_master` / `apply_class_master` / `merge_pending_fixes_into_master` / `apply_dropdowns_to_pending_report`  
3. 先「搬文件 + re-export」，再「瘦内部」；不在同一 PR 改业务语义  

---

## 目标目录结构

```
learn_classify/
  class_master.py          # 薄 facade：from .master_*. import *  （过渡期保留）
  master/
    __init__.py            # 对外 re-export
    constants.py           # MASTER_COLUMNS / STATUS_* / _SKIP_L2
    collect.py             # 扫 B_std、去重聚合
    labels.py              # 从总表读已有标签
    export.py              # export_class_master + Excel 格式/下拉
    apply.py               # apply_class_master + 学习记录
    pending.py             # 待人工报告下拉 + merge pending
    report.py              # _export_learn_report
```

可选后续：把 Excel 下拉依赖从 `payment_hub.category_dropdowns` 抽到 `excel_ui/`，切断 `learn_classify → payment_hub`。

---

## 函数级迁移表

| 现函数 / 符号 | 行号约 | 目标模块 | 公开？ | 备注 |
|---------------|--------|----------|--------|------|
| `MASTER_SHEET` / `README_SHEET` | 33–34 | `constants.py` | 包内 | |
| `MASTER_COLUMNS` | 38–50 | `constants.py` | 是 | 测试可 import |
| `STATUS_CONFIRMED` / `STATUS_PENDING` | 52–53 | `constants.py` | 是 | |
| `_CONFIRM_VALUES` / `_SKIP_L2` | 54–56 | `constants.py` | 包内 | |
| `_safe_str` | 59–65 | `collect.py` 或 `util` | 否 | 可与 normalize 合并 |
| `_pick_sheet` | 67–77 | `collect.py` | 否 | |
| `_collect_raw_lines` | 79–138 | `collect.py` | 否 | IO 重 |
| `_is_confirmed` | 140–143 | `labels.py` | 否 | |
| `_is_learnable_l2` | 145–148 | `labels.py` | 否 | |
| `resolve_leaf_from_row` | 150–175 | `labels.py` / `apply.py` | 是 | apply 核心 |
| `_is_learnable_leaf` | 177–187 | `labels.py` | 否 | |
| `load_labels_from_master` | 189–232 | `labels.py` | 是 | export 保留标签 |
| `build_deduped_master` | 234–343 | `collect.py` | 是 | 纯聚合，单测重点 |
| `_apply_master_format` | 345–509 | `export.py` | 否 | openpyxl 重 |
| `export_class_master` | 511–547 | `export.py` | **是** | pipeline 入口 |
| `_export_learn_report` | 550–626 | `report.py` | 否 | |
| `apply_class_master` | 629–820 | `apply.py` | **是** | 学习主路径 |
| `apply_dropdowns_to_pending_report` | 824–950 | `pending.py` | **是** | CLI 用 |
| `merge_pending_fixes_into_master` | 952–1080 | `pending.py` | **是** | CLI 用 |

行号会随编辑漂移；以**符号名**为准。

---

## PR 序列（推荐 5 个小 PR）

### PR-A · 建包 + 搬 constants（无行为变化）

1. 新建 `learn_classify/master/`  
2. 迁 `constants.py`  
3. `class_master.py` 改为 `from .master.constants import *`  
4. 冒烟：`python -c "from learn_classify.class_master import MASTER_COLUMNS"`  

**验收**：import 路径旧代码仍可用。

### PR-B · collect + labels

1. 迁 `_collect_raw_lines` / `build_deduped_master` → `collect.py`  
2. 迁标签解析函数 → `labels.py`  
3. `class_master` re-export  
4. 单测：`build_deduped_master` 用 2 行 fixture std  

**验收**：`export-class` dry 逻辑与拆前一致（行数、确认状态分布）。

### PR-C · export

1. 迁 `_apply_master_format` + `export_class_master`  
2. 保持 `pipeline.step_export_class_master` 调用点不变  

**验收**：导出总表列序、下拉、preserve 标签。

### PR-D · apply + report

1. 迁 `apply_class_master` + `_export_learn_report`  
2. 单测：叶子未变跳过、pending 跳过、变更写入 memory（临时 data dir）  

**验收**：对同一总表两次 apply，第二次 `learned=0`。

### PR-E · pending + 删空壳

1. 迁 pending 两函数  
2. `class_master.py` 仅 re-export 或删除并改 import 到 `learn_classify.master`  
3. 更新 `pipeline` / `cli` 的 import  

**验收**：`apply-pending-fixes` / `pending-dropdowns` 冒烟。

---

## 依赖方向（拆完后）

```
pipeline / cli
    → master.export / master.apply / master.pending
        → engine / store / taxonomy / learn.learn_one
        → （可选）excel_ui.dropdowns   # 勿反向依赖 payment_hub 业务
```

---

## 单测优先级（随 PR 补）

| 用例 | 模块 |
|------|------|
| 相同摘要多项目 → 一行 + 出现次数 | collect |
| 有效二级保留 / 待复核不覆盖已有有效 | labels + export |
| resolve_leaf：有 L3 用 L3，否则 L2 | labels |
| apply：叶子相同不写 memory | apply |
| apply：only_confirmed 过滤 | apply |

---

## 不做（本轮）

- 改学习语义（何时写入 memory）  
- 换存储（SQLite）— 可在 compact 稳定后再做  
- 合并 review/learn 与 export-class 产品路径 — 文档层已收敛，代码可后置  

---

## 与已落地优化的关系

| 已做 | 对本拆分的影响 |
|------|----------------|
| `engine.classify_many` | collect/export 扫描时应用（可选 follow-up） |
| `compact-memory` | 不依赖 class_master 结构 |
| 文档默认 apply-class | 无冲突 |

拆分时可在 `collect._collect_raw_lines` 内对唯一摘要批量分类，进一步加速 `export-class`。
