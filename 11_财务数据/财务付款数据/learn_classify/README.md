# 财务付款分类 · 可持续学习闭环

路径：`12_财务数据/财务付款数据/learn_classify/`

> **推荐日常使用一体化入口**（已融合 3/4/5 + 本模块）：  
> `python 财务付款一体化.py help-guide`  
> 详见上级目录 `使用说明_财务付款一体化.md`。

## 解决什么问题

1. **可解释分层分类**（记忆 → 高优规则 → 科目映射 → 规则 → 供应商众数 → **待复核**）
2. **人工纠错后自动学习**（精确记忆 + 供应商画像 + 规则建议）
3. **一二级调整后快速转移**（稳定 `L2_xxx` ID + migration `id_map` + 一键重分类）

---

## 架构

```
新付款明细
    │
    ▼
ClassifyEngine（可解释多层）
    │
    ├─ 高置信 → 写入 std（一级/二级/分类ID/命中层/置信度）
    │
    └─ 低置信/待复核 → export-review → 待复核池.xlsx
                              │
                              │ 人工填「正确二级科目」
                              ▼
                           learn
                              │
              ┌───────────────┼───────────────┐
              ▼               ▼               ▼
         memory.jsonl   corrections.jsonl  vendor_stats
              │               │
              │          suggest-rules → 规则建议.xlsx
              │               │ status=approved
              │               ▼
              │         approve-rules → rules.json
              │
              └── 下次分类自动命中
```

**分类转移（taxonomy 变更）**

```
编辑 migrations/xxx.json（id_map / rename / reparent / add_level2）
        │
        ▼
     migrate --map ... [--reclassify-std B_财务付款_std]
        │
        ├─ 备份 data/
        ├─ 更新 taxonomy.json
        ├─ 重写 memory / corrections / vendor / rules / subject_map 中的 ID
        └─ 可选：整目录重刷显示名
```

记忆与规则存的是 **稳定 ID**，不是「钢筋」「材料采购」这种会改的名字，所以改名/合并/拆分只需 `id_map`。

---

## 目录

```
learn_classify/
  engine.py taxonomy.py store.py learn.py review.py migrate.py
  adapter.py cli.py
  data/
    taxonomy.json      # 分类树（version + L1/L2 稳定 ID）
    rules.json         # 关键词规则 → l2_id + priority
    subject_map.json   # 原始财务科目名 → l2_id
    memory.jsonl       # 精确记忆（学习产物）
    corrections.jsonl  # 纠错流水
    vendor_stats.json  # 供应商众数
    migrations/        # 迁移方案（人写）
    migrations_applied/# 已执行记录（自动）
```

---

## 常用命令

在 `财务付款数据` 目录下：

```bash
cd /Users/wyattsun/Desktop/code/12_财务数据/财务付款数据

# 单条调试
python -m learn_classify classify --summary "预扣税" --subject "税款"

# 导出待复核（从 B_财务付款_std）
python -m learn_classify export-review -o ./learn_classify/data/待复核池.xlsx

# 人工填写「正确二级科目」后学习
python -m learn_classify learn -i ./learn_classify/data/待复核池.xlsx

# 从纠错生成规则建议 → 改 status=approved → 合并
python -m learn_classify suggest-rules -o ./learn_classify/data/规则建议.xlsx
python -m learn_classify approve-rules -i ./learn_classify/data/规则建议.xlsx

# 分类体系迁移（快速转移）
python -m learn_classify migrate \
  --map ./learn_classify/data/migrations/example_v1_to_v1.1.json \
  --reclassify-std "/Users/wyattsun/Desktop/邯郸中材/01成本统计/12-财务数据💰/A_财务付款数据/B_财务付款_std"

# 仅重分类 std（不改 taxonomy）
python -m learn_classify reclassify

# 健康度
python -m learn_classify health

# 供应商画像热身（可选）
python -m learn_classify seed-vendor
```

跑原有汇总台账即可自动走新引擎（`subject_mapper.py` 已接入）：

```bash
python 3_财务付款数据汇总台账.py
```

---

## 迁移文件怎么写

`data/migrations/your_change.json`：

```json
{
  "from_version": "1.0.0",
  "to_version": "1.1.0",
  "description": "其他拆出危废清运；材料采购改名",
  "id_map": {
    "L2_mgmt_other": "L2_mgmt_other"
  },
  "rename": {
    "L2_mat_generic": "辅材杂项"
  },
  "add_level2": [
    {
      "id": "L2_waste",
      "name": "危废清运",
      "parent": "L1_mgmt",
      "order": 45,
      "aliases": ["危废处理"]
    }
  ],
  "reparent": {
    "L2_log_transport": "L1_logistics"
  },
  "deactivate": [],
  "pending_l2_id": "L2_pending"
}
```

| 字段 | 作用 |
|------|------|
| `id_map` | 旧二级 ID → 新二级 ID（记忆/规则/画像全部跟着转） |
| `rename` | 只改显示名，ID 不变 |
| `reparent` | 二级换一级（显示一级跟着变） |
| `add_level2` | 新增叶子类 |
| `deactivate` | 停用旧类（引擎会落到待复核） |

合并两类：`"L2_a": "L2_b"`  
拆分：先 `add_level2`，再对记忆做人工 learn，或分批 `id_map` 无法表达的一对多用纠正表处理。

---

## 待复核表列说明

| 列 | 说明 |
|----|------|
| 一级/二级科目 | 当前自动结果 |
| 分类ID | 稳定 ID（迁移用） |
| 命中层 / 置信度 / 原因 | 可解释信息 |
| **正确一级科目** | 可选 |
| **正确二级科目** | **必填才能学习**（填显示名或 ID） |
| 备注 | 可选 |

---

## Python API

```python
from learn_classify import ClassifyEngine, DEFAULT_DATA_DIR

engine = ClassifyEngine.from_data_dir(DEFAULT_DATA_DIR)
r = engine.classify("材料采购", "角钢、工字钢", "某某公司")
print(r.l1_name, r.l2_name, r.l2_id, r.layer, r.confidence)
```

---

## 建议运营节奏

| 频率 | 动作 |
|------|------|
| 每次出账后 | `export-review` → 标 20～50 条 → `learn` |
| 每周 | `suggest-rules` → 审核 → `approve-rules` |
| 改分类树时 | 写 migration → `migrate --reclassify-std ...` |
| 每月 | `health` 看待复核率 / 其他占比 |

目标参考：待复核 &lt; 10%，二级「其他」持续下降。
