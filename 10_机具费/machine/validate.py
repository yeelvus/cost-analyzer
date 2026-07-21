# -*- coding: utf-8 -*-
from __future__ import annotations

import re
from machine.textutil import safe_to_float

def validate_overtime_fee(row):
    """验证加班费是否正确
    正确公式: 加班费 = 台班费_泰铢 / 8 * 加班小时
    返回: (是否正确, 实际值, 计算值)
    """
    actual = safe_to_float(row.get('加班费', 0))
    unit_price = safe_to_float(row.get('台班费_泰铢', 0))
    overtime_hours = safe_to_float(row.get('加班小时', 0))
    
    # 如果没有加班小时，加班费应该为0
    if overtime_hours == 0:
        calculated = 0.0
    else:
        calculated = unit_price / 8 * overtime_hours
    
    # 允许0.01的误差（因为四舍五入）
    is_correct = abs(actual - calculated) < 0.01
    
    return is_correct, actual, calculated


def validate_settlement_amount(row):
    """验证结算费用是否正确
    正确公式（两种方式之一即可）：
    方式1: 结算费用_泰铢 = 台班费_泰铢 * 本次结算台班 + 加班费 + 进出场费运费 + 移动费拆装费
    方式2: 结算费用_泰铢 = 台班费_泰铢 + 台班费_泰铢 / 8 * 加班小时 + 进出场费运费 + 移动费拆装费
    返回: (是否正确, 实际值, 计算值1, 计算值2)
    """
    actual = safe_to_float(row.get('结算费用_泰铢', 0))
    unit_price = safe_to_float(row.get('台班费_泰铢', 0))
    quantity = safe_to_float(row.get('本次结算台班', 0))
    overtime_fee = safe_to_float(row.get('加班费', 0))
    overtime_hours = safe_to_float(row.get('加班小时', 0))
    entry_exit_fee = safe_to_float(row.get('进出场费运费', 0))
    move_install_fee = safe_to_float(row.get('移动费拆装费', 0))
    
    # 方式1: 台班费 * 结算台班 + 加班费 + 进出场费 + 移动费
    calculated1 = unit_price * quantity + overtime_fee + entry_exit_fee + move_install_fee
    
    # 方式2: 台班费 + (台班费/8 * 加班小时) + 进出场费 + 移动费
    calculated2 = unit_price + (unit_price / 8 * overtime_hours) + entry_exit_fee + move_install_fee
    
    # 允许0.01的误差（因为四舍五入），两种方式有一种正确即可
    is_correct = abs(actual - calculated1) < 0.01 or abs(actual - calculated2) < 0.01
    
    return is_correct, actual, calculated1, calculated2


def validate_total_amount(row):
    """验证本次费用合计是否正确
    正确公式: 本次费用合计 = 结算费用_泰铢 + 增值税7% - 机具租赁预扣税3% - 机具租赁预扣税5%
    返回: (是否正确, 实际值, 计算值)
    """
    actual = safe_to_float(row.get('本次费用合计', 0))
    settlement = safe_to_float(row.get('结算费用_泰铢', 0))
    vat = safe_to_float(row.get('增值税7%', 0))
    tax3 = safe_to_float(row.get('机具租赁预扣税3%', 0))
    tax5 = safe_to_float(row.get('机具租赁预扣税5%', 0))
    
    calculated = settlement + vat - tax3 - tax5
    
    # 允许0.01的误差（因为四舍五入）
    is_correct = abs(actual - calculated) < 0.01
    
    return is_correct, actual, calculated


def detect_and_fix_swapped_values(df):
    """检测并修正台班费和结算台班填反的情况
    判断逻辑：
    1. 台班费通常 > 100（单价），结算台班通常 < 100（数量）
    2. 如果台班费 < 100 且 结算台班 > 1000，很可能填反了
    返回: (修正后的df, 被修正的行索引列表)
    """
    swapped_indices = []
    
    for idx in range(len(df)):
        unit_price = safe_to_float(df.iloc[idx].get('台班费_泰铢', 0))
        quantity = safe_to_float(df.iloc[idx].get('本次结算台班', 0))
        
        # 检测填反的情况：
        # 1. 台班费很小（< 50）且结算台班很大（> 500）
        # 2. 或者台班费在合理台班数范围（1-100）且结算台班像价格（> 1000）
        is_swapped = False
        
        if unit_price > 0 and quantity > 0:
            # 情况1：台班费 < 50 且 结算台班 > 500
            if unit_price < 50 and quantity > 500:
                is_swapped = True
            # 情况2：台班费像数量（1-100）且结算台班像价格（> 1000）
            elif unit_price <= 100 and quantity > 1000:
                is_swapped = True
        
        if is_swapped:
            # 交换两列的值
            df.at[idx, '台班费_泰铢'] = quantity
            df.at[idx, '本次结算台班'] = unit_price
            swapped_indices.append(idx)
            print(f"    ⚠️  第 {idx+2} 行数据已自动交换：台班费 {unit_price:.2f} ↔ 结算台班 {quantity:.2f}")
    
    return df, swapped_indices


def detect_and_fix_supplier_month_swap(df):
    """检测并修正出租单位和月份填反的情况
    判断逻辑：
    1. 月份通常是数字格式（如"1月"、"2月"、"202501"、"8月份"等）
    2. 出租单位通常是公司名称（包含中文、英文、特殊字符等）
    3. 如果出租单位像月份，月份像公司名，则可能填反
    返回: (修正后的df, 被修正的行索引列表)
    """
    swapped_month_indices = []
    
    # 月份模式：纯数字、包含"月"字、YYYYMM格式等
    month_patterns = [
        r'^\d{1,2}$',              # 1-12
        r'^\d{1,2}月$',            # 1月-12月
        r'^\d{1,2}月份$',          # 1月份-12月份 ← 新增
        r'^\d{4}年\d{1,2}月$',     # 2024年1月
        r'^\d{4}年\d{1,2}月份$',   # 2024年1月份 ← 新增
        r'^\d{6}$',                # 202401
        r'^\d{4}-\d{1,2}$',        # 2024-01
        r'^\d{4}/\d{1,2}$',        # 2024/01
        r'^[一二三四五六七八九十]{1,2}月$',      # 一月-十二月（中文数字）
        r'^[一二三四五六七八九十]{1,2}月份$',    # 一月份-十二月份
    ]
    
    for idx in range(len(df)):
        supplier = str(df.iloc[idx].get('出租单位', '')).strip()
        month = str(df.iloc[idx].get('月份', '')).strip()
        
        if not supplier or supplier == 'nan':
            continue
            
        # 检查出租单位是否像月份
        is_supplier_like_month = False
        for pattern in month_patterns:
            if re.match(pattern, supplier):
                is_supplier_like_month = True
                break
        
        # 如果出租单位像月份，且月份不像月份（或为空），则可能填反
        if is_supplier_like_month:
            # 检查月份是否不像月份（更像公司名）
            is_month_like_supplier = False
            if month and month != 'nan':
                # 月份包含中文、英文字母（除了"月"、"份"、"年"）等，可能是公司名
                month_chars = set('月份年')
                if any(c.isalpha() and c not in month_chars for c in month) or len(month) > 15:
                    is_month_like_supplier = True
            
            # 如果确定填反了，或月份为空
            if is_month_like_supplier or not month or month == 'nan':
                # 交换两列的值
                df.at[idx, '出租单位'], df.at[idx, '月份'] = month if month and month != 'nan' else '', supplier
                swapped_month_indices.append(idx)
                print(f"    ⚠️  第 {idx+2} 行数据已自动交换：出租单位「{supplier}」↔ 月份「{month}」")
    
    return df, swapped_month_indices
