import json
import os
from collections import OrderedDict


def update_contract_json_logic(root_dir):
    # 1. 定义最终的标准字段顺序及默认值
    # 已将 "对方单位公司名称" 修改为 "对方公司名称"
    template = OrderedDict([
        ("项目名称", ""),
        ("合同编号", ""),
        ("合同类型", ""),
        ("对方公司名称", ""),  # 统一后的字段名
        ("合同签订日期", ""),
        ("生效日期", ""),
        ("结束日期", ""),
        ("工期/交货期", ""),
        ("币种", ""),
        ("汇率", 1.0),
        ("合同金额", 0.0),
        ("含税金额", 0.0),
        ("不含税金额", 0.0),
        ("税率", ""),
        ("付款方式", ""),
        ("履约保证金", ""),
        ("合同状态", ""),
        ("关联主合同编号", ""),
        ("联系人及电话", ""),
        ("PDF文件路径", ""),
        ("收付款记录", []),
        ("备注", "")
    ])

    def _fmt(v):
        try:
            return json.dumps(v, ensure_ascii=False)
        except Exception:
            return str(v)

    def _format_date_val(v):
        """把类似 20241008 或 2024/10/08 转为 2024-10-08，其他保持原样。"""
        if v is None:
            return ""
        # 如果已经是标准格式，直接返回
        try:
            s = str(v).strip()
        except Exception:
            return v
        if s == "":
            return ""
        if "-" in s and len(s) >= 8:
            return s
        # 提取数字
        digits = ''.join(ch for ch in s if ch.isdigit())
        if len(digits) == 8:
            return f"{digits[0:4]}-{digits[4:6]}-{digits[6:8]}"
        # 处理类似 2024/10/08
        if "/" in s:
            parts = [p for p in s.replace('/', '-').split('-') if p]
            if len(parts) == 3 and all(p.isdigit() for p in parts):
                y, m, d = parts
                if len(y) == 4:
                    m = m.zfill(2)
                    d = d.zfill(2)
                    return f"{y}-{m}-{d}"
        return v

    # 首先扫描，收集所有需要将汇率改为 5 的文件（仅当币种为 CNY 且汇率为 1/空/0/None 时）
    candidates = []
    for root, dirs, files in os.walk(root_dir):
        for file in files:
            if file != "合同信息.json":
                continue
            file_path = os.path.join(root, file)

            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    orig_full = json.load(f)
            except Exception as e:
                print(f"读取失败 {file_path}: {e}")
                continue

            # 使用 orig 作为可变副本来处理字段迁移，但用 orig_full 保存原始值以便比较
            orig = dict(orig_full)

            # 获取可能存在的旧字段内容（保留到 final_company_name）
            val1 = orig.pop("对方单位公司名称", None) if "对方单位公司名称" in orig else None
            val2 = orig.pop("对方公司名称", None) if "对方公司名称" in orig else None
            final_company_name = val1 if val1 else (val2 if val2 else "")

            current_currency = orig_full.get("币种", "")
            curr_str = str(current_currency).upper()
            is_cny = ("CNY" in curr_str) or ("人民币" in curr_str)

            current_rate = orig_full.get("汇率")
            # 判断汇率是否需要改为默认 5
            def _rate_needs_change(rate):
                if rate is None:
                    return True
                s = str(rate).strip()
                if s == "":
                    return True
                try:
                    v = float(s)
                    if v == 0 or v == 1:
                        return True
                    return False
                except Exception:
                    return False

            # 构建 new_data（仅作预览，暂不写入）
            final_currency = "CNY" if is_cny else orig_full.get("币种", "")
            new_data = OrderedDict()
            # 将常见的日期字段统一格式化（包含更多可能的命名变体）
            date_keys = {
                "合同签订日期", "生效日期", "结束日期",
                "签署时间", "签署日期", "签订时间", "签订日期",
                "交货日期", "交货时间", "交付日期"
            }
            for key, default_value in template.items():
                if key == "对方公司名称":
                    new_data[key] = final_company_name
                elif key == "汇率":
                    # 仅当币种为 CNY 且需要更改时才设置为 5.0，否则保留原值或默认
                    if is_cny and _rate_needs_change(current_rate):
                        new_data[key] = 5.0
                    else:
                        new_data[key] = orig_full.get(key, default_value)
                elif key == "币种":
                    new_data[key] = final_currency
                else:
                    val = orig.get(key, default_value)
                    if key in date_keys:
                        new_data[key] = _format_date_val(val)
                    else:
                        new_data[key] = val
            # 保留额外字段（原始 JSON 中存在但不在 template 的字段）
            for k in orig_full:
                if k not in new_data:
                    new_data[k] = orig_full[k]

            # 记录所有将要发生的变更（对比 orig_full 与 new_data）
            changes = []
            all_keys = list(new_data.keys())
            for k in all_keys:
                old_val = orig_full.get(k, template.get(k, ""))
                new_val = new_data.get(k)
                # 日期字段比较：保留原始旧值，和格式化后的新值做对比
                if k in date_keys:
                    old_cmp = old_val
                    new_cmp = _format_date_val(new_val)
                else:
                    old_cmp = old_val
                    new_cmp = new_val
                if old_cmp != new_cmp:
                    changes.append((k, old_cmp, new_cmp))

            # 只有在确实有变更时才加入候选列表
            if changes:
                candidates.append({
                    "path": file_path,
                    "orig": orig_full,
                    "new_data": new_data,
                    "changes": changes
                })

    # 只打印符合条件需要修改的文件清单
    if not candidates:
        print("未发现需要将汇率改为 5 的 CNY JSON 文件。")
        return

    print("发现以下需要修改的文件（币种为 CNY 且汇率为 1/空/0）：")
    for c in candidates:
        print(f"- {c['path']}")
        for k, ov, nv in c['changes']:
            print(f"    * {k}: {_fmt(ov)} -> {_fmt(nv)}")

    ans = input("输入 y 确认将上述文件的汇率改为 5 并写回文件，其他任意键取消：")
    if ans.lower() != 'y':
        print("已取消操作。没有修改任何文件。")
        return

    # 写回确认的文件
    for c in candidates:
        path = c['path']
        new_data = c['new_data']
        try:
            with open(path, 'w', encoding='utf-8') as f:
                json.dump(new_data, f, ensure_ascii=False, indent=2)
            print(f"已写回: {path}")
        except Exception as e:
            print(f"写入失败 {path}: {e}")


# --- 配置区 ---
target_path = r"/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理"


if __name__ == "__main__":
    update_contract_json_logic(target_path)
    print("代码已就绪，请在 target_path 填入路径并运行脚本以处理文件。")