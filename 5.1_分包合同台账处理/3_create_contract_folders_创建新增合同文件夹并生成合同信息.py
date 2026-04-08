"""
分包合同完整性检查与自动创建工具
功能:
1. 读取合同汇总Excel，按项目编码分组统计合同总额
2. 扫描合同管理目录，读取所有合同信息.json并统计总额
3. 对比差异，找出缺失的合同
4. 列出待创建清单，经用户确认后自动创建文件夹和JSON
"""

import pandas as pd
import os
import json
import glob
import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional


# =========================
# 配置区域
# =========================
SUMMARY_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/2_标准化汇总分包合同台账"
# 兼容历史目录名，按顺序查找可用目录
SUMMARY_DIR_CANDIDATES = [
    SUMMARY_DIR,
    "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5.1_分包合同/2_整理标准化汇总分包签订合同文件",
]
CONTRACT_BASE_DIR = "/Users/wyattsun/Desktop/邯郸中材/01成本统计/5_合同管理"
AMOUNT_TOLERANCE = 10.0  # 金额容差（泰铢）


def resolve_summary_dir(candidates: List[str]) -> str:
    """返回第一个存在的汇总目录；若都不存在则返回首选目录。"""
    for path in candidates:
        if os.path.isdir(path):
            return path
    return candidates[0]


def find_latest_summary_file(summary_dir: str) -> Optional[str]:
    """在汇总目录中查找最新的汇总Excel文件。"""
    if not os.path.isdir(summary_dir):
        return None

    candidates = []
    for fn in os.listdir(summary_dir):
        if fn.startswith('~$'):
            continue
        lower_fn = fn.lower()
        if not lower_fn.endswith(('.xlsx', '.xlsm', '.xls')):
            continue
        if '汇总' not in fn:
            continue
        full_path = os.path.join(summary_dir, fn)
        if os.path.isfile(full_path):
            candidates.append(full_path)

    if not candidates:
        return None

    def extract_ymd_rank(path: str):
        name = os.path.basename(path)
        # 支持 yyyy-mm-dd / yyyy_mm_dd / yyyy年m月d日 / yyyy-mm
        m = re.search(r'(20\d{2})[-_年\.](\d{1,2})(?:[-_月\.](\d{1,2}))?', name)
        if not m:
            return None
        year = int(m.group(1))
        month = int(m.group(2))
        day = int(m.group(3)) if m.group(3) else 0
        if not (1 <= month <= 12 and 0 <= day <= 31):
            return None
        return (year, month, day)

    ranked = []
    for path in candidates:
        rank = extract_ymd_rank(path)
        mtime = os.path.getmtime(path)
        ranked.append((rank, mtime, path))

    # 优先按文件名日期排序；若缺失日期则按修改时间兜底
    with_date = [x for x in ranked if x[0] is not None]
    if with_date:
        return max(with_date, key=lambda x: (x[0], x[1]))[2]

    return max(ranked, key=lambda x: x[1])[2]


def find_project_folder(contract_base_dir: str, project_code: str) -> Tuple[Optional[str], Optional[str]]:
    """
    在合同管理目录下查找项目文件夹
    
    Args:
        contract_base_dir: 合同管理基础目录
        project_code: 项目编码 (如: THCC-2025-002)
    
    Returns:
        (项目文件夹路径, 分包合同文件夹路径) 或 (None, None) 如果未找到
    """
    
    # 搜索所有一级子文件夹（A_在建项目、AS_收尾项目等）
    for category_dir in os.listdir(contract_base_dir):
        category_path = os.path.join(contract_base_dir, category_dir)
        
        # 只处理文件夹
        if not os.path.isdir(category_path) or category_dir.startswith('.'):
            continue
        
        # 搜索该分类下包含项目编码的文件夹
        for item in os.listdir(category_path):
            item_path = os.path.join(category_path, item)
            
            if not os.path.isdir(item_path):
                continue
            
            # 检查文件夹名称是否包含项目编码
            if project_code in item:
                # 找到项目文件夹，现在查找分包合同文件夹
                for subitem in os.listdir(item_path):
                    subitem_path = os.path.join(item_path, subitem)
                    
                    if not os.path.isdir(subitem_path):
                        continue
                    
                    # 检查是否是分包合同文件夹 (F_项目编码_分包合同)
                    if subitem.startswith('F_') and '分包合同' in subitem:
                        return item_path, subitem_path
    
    return None, None


def scan_existing_contracts(subpackage_dir: str) -> Dict[str, Dict]:
    """
    扫描分包合同文件夹，读取所有合同信息.json
    
    Args:
        subpackage_dir: 分包合同文件夹路径
    
    Returns:
        {合同编号: {folder: 文件夹路径, amount: 金额, json_data: JSON数据}}
    """
    contracts = {}
    
    if not os.path.exists(subpackage_dir):
        return contracts
    
    for item in os.listdir(subpackage_dir):
        item_path = os.path.join(subpackage_dir, item)
        
        if not os.path.isdir(item_path) or item.startswith('.'):
            continue
        
        # 查找合同信息.json
        json_file = os.path.join(item_path, '合同信息.json')
        
        if os.path.exists(json_file):
            try:
                with open(json_file, 'r', encoding='utf-8') as f:
                    json_data = json.load(f)
                
                contract_no = json_data.get('合同编号', '')
                amount = json_data.get('合同金额', 0.0)
                
                if contract_no:
                    contracts[contract_no] = {
                        'folder': item_path,
                        'folder_name': item,
                        'amount': float(amount) if amount else 0.0,
                        'json_data': json_data
                    }
            except Exception as e:
                print(f"  警告：读取JSON失败 {json_file}: {e}")
    
    return contracts


def analyze_contract_completeness(excel_file: str, contract_base_dir: str) -> Dict:
    """
    分析合同完整性
    
    Returns:
        {
            'projects': {项目编码: {项目信息}},
            'missing_contracts': [缺失的合同列表],
            'amount_mismatches': [金额不匹配的项目列表]
        }
    """
    print("="*80)
    print("分包合同完整性检查")
    print("="*80)
    
    # 读取Excel
    print(f"\n正在读取Excel: {excel_file}")
    df = pd.read_excel(excel_file, sheet_name='Sheet1')
    df.columns = [str(c).strip() for c in df.columns]
    print(f"  共 {len(df)} 条记录")

    def pick_required_column(candidates: List[str]) -> str:
        for col in candidates:
            if col in df.columns:
                return col
        raise KeyError(f"缺少必要列，候选列: {candidates}，实际列: {list(df.columns)}")

    def pick_optional_column(candidates: List[str]) -> Optional[str]:
        for col in candidates:
            if col in df.columns:
                return col
        return None

    # 兼容新旧版本汇总列名
    project_code_col = pick_required_column(['项目编码', '项目编码(匹配)'])
    project_name_col = pick_optional_column(['项目名称', '标准项目名称(JSON)'])
    amount_col = pick_required_column(['合同额（泰铢）', '合同额(泰铢)', '合同金额'])
    contract_no_col = pick_required_column(['合同编号'])
    work_scope_col = pick_optional_column(['工程范围'])
    company_name_col = pick_optional_column(['协作体名称'])
    
    # 按项目编码分组
    project_codes = df[project_code_col].unique()
    print(f"  涉及 {len(project_codes)} 个项目")
    
    analysis_result = {
        'projects': {},
        'missing_contracts': [],
        'amount_mismatches': []
    }
    
    # 逐项目检查
    print("\n" + "="*80)
    print("开始检查各项目...")
    print("="*80)
    
    for project_code in project_codes:
        if pd.isna(project_code) or str(project_code).strip() == '':
            continue
        
        project_code = str(project_code).strip()
        
        # 获取该项目的所有合同
        project_contracts = df[df[project_code_col] == project_code].copy()
        if project_name_col:
            project_name = project_contracts.iloc[0][project_name_col]
        else:
            project_name = project_code
        
        # 计算Excel中的总金额
        excel_total = project_contracts[amount_col].sum()
        
        print(f"\n【项目】{project_name} ({project_code})")
        print(f"  Excel合同数: {len(project_contracts)}")
        print(f"  Excel总金额: {excel_total:,.2f} 泰铢")
        
        # 查找项目文件夹
        project_folder, subpackage_folder = find_project_folder(contract_base_dir, project_code)
        
        if not subpackage_folder:
            print(f"  ⚠️  警告：找不到分包合同文件夹")
            analysis_result['amount_mismatches'].append({
                'project_code': project_code,
                'project_name': project_name,
                'excel_total': excel_total,
                'json_total': 0,
                'difference': excel_total,
                'reason': '找不到分包合同文件夹'
            })
            
            # 标记所有合同为缺失
            for _, contract_row in project_contracts.iterrows():
                # 跳过合同额为空或0的记录
                contract_amount = contract_row[amount_col]
                if pd.isna(contract_amount) or contract_amount == 0:
                    continue
                
                # 处理工程范围可能为空值的情况
                work_scope = contract_row.get(work_scope_col, '') if work_scope_col else ''
                if pd.isna(work_scope):
                    work_scope = ''
                else:
                    work_scope = str(work_scope).strip()
                
                # 处理协作体名称
                company_name = contract_row.get(company_name_col, '') if company_name_col else ''
                if pd.isna(company_name):
                    company_name = ''
                else:
                    company_name = str(company_name).strip()
                
                analysis_result['missing_contracts'].append({
                    'project_code': project_code,
                    'project_name': project_name,
                    'contract_no': contract_row[contract_no_col],
                    'work_scope': work_scope,
                    'amount': contract_amount,
                    'company_name': company_name,
                    'reason': '分包合同文件夹不存在'
                })
            continue
        
        print(f"  文件夹: {subpackage_folder}")
        
        # 扫描现有合同
        existing_contracts = scan_existing_contracts(subpackage_folder)
        json_total = sum(c['amount'] for c in existing_contracts.values())
        
        print(f"  JSON合同数: {len(existing_contracts)}")
        print(f"  JSON总金额: {json_total:,.2f} 泰铢")
        
        # 比较金额
        diff = abs(excel_total - json_total)
        
        if diff <= AMOUNT_TOLERANCE:
            print(f"  ✓ 金额匹配（差异: {diff:.2f} 泰铢）")
        else:
            print(f"  ✗ 金额不匹配（差异: {diff:.2f} 泰铢）")
            
            # 检查缺失的合同
            for _, contract_row in project_contracts.iterrows():
                contract_no = str(contract_row[contract_no_col]).strip()
                
                # 跳过合同额为空或0的记录
                contract_amount = contract_row[amount_col]
                if pd.isna(contract_amount) or contract_amount == 0:
                    continue
                
                if contract_no not in existing_contracts:
                    print(f"    → 缺失合同: {contract_no}")
                    # 处理工程范围可能为空值的情况
                    work_scope = contract_row.get(work_scope_col, '') if work_scope_col else ''
                    if pd.isna(work_scope):
                        work_scope = ''
                    else:
                        work_scope = str(work_scope).strip()
                    
                    # 处理协作体名称
                    company_name = contract_row.get(company_name_col, '') if company_name_col else ''
                    if pd.isna(company_name):
                        company_name = ''
                    else:
                        company_name = str(company_name).strip()
                    
                    analysis_result['missing_contracts'].append({
                        'project_code': project_code,
                        'project_name': project_name,
                        'contract_no': contract_no,
                        'work_scope': work_scope,
                        'amount': contract_amount,
                        'company_name': company_name,
                        'subpackage_folder': subpackage_folder,
                        'reason': '文件夹中不存在'
                    })
            
            analysis_result['amount_mismatches'].append({
                'project_code': project_code,
                'project_name': project_name,
                'excel_total': excel_total,
                'json_total': json_total,
                'difference': diff,
                'existing_contracts': len(existing_contracts),
                'excel_contracts': len(project_contracts)
            })
        
        analysis_result['projects'][project_code] = {
            'name': project_name,
            'excel_total': excel_total,
            'json_total': json_total,
            'excel_count': len(project_contracts),
            'json_count': len(existing_contracts),
            'subpackage_folder': subpackage_folder
        }
    
    return analysis_result


def get_next_contract_number(subpackage_dir: str) -> int:
    """
    获取下一个合同编号序号
    
    Args:
        subpackage_dir: 分包合同文件夹路径
    
    Returns:
        下一个序号
    """
    max_num = 0
    
    if not os.path.exists(subpackage_dir):
        return 1
    
    for item in os.listdir(subpackage_dir):
        # 匹配 数字_合同编号_工程范围 格式
        match = re.match(r'^(\d+)_', item)
        if match:
            num = int(match.group(1))
            max_num = max(max_num, num)
    
    return max_num + 1


def print_analysis_summary(analysis: Dict):
    """打印分析摘要"""
    print("\n" + "="*80)
    print("检查结果汇总")
    print("="*80)
    
    print(f"\n总项目数: {len(analysis['projects'])}")
    print(f"金额不匹配项目数: {len(analysis['amount_mismatches'])}")
    print(f"缺失合同数: {len(analysis['missing_contracts'])}")
    
    if analysis['amount_mismatches']:
        print("\n" + "-"*80)
        print("金额不匹配的项目：")
        print("-"*80)
        for item in analysis['amount_mismatches']:
            print(f"\n项目: {item['project_name']} ({item['project_code']})")
            print(f"  Excel总额: {item['excel_total']:,.2f} 泰铢 ({item.get('excel_contracts', 0)} 个合同)")
            print(f"  JSON总额:  {item['json_total']:,.2f} 泰铢 ({item.get('existing_contracts', 0)} 个合同)")
            print(f"  差异:     {item['difference']:,.2f} 泰铢")
            if 'reason' in item:
                print(f"  原因: {item['reason']}")
    
    if analysis['missing_contracts']:
        print("\n" + "-"*80)
        print("缺失的合同清单：")
        print("-"*80)
        
        # 按项目分组显示
        by_project = {}
        for contract in analysis['missing_contracts']:
            proj = contract['project_code']
            if proj not in by_project:
                by_project[proj] = []
            by_project[proj].append(contract)
        
        for project_code, contracts in by_project.items():
            print(f"\n【{contracts[0]['project_name']}】({project_code})")
            print(f"  缺失 {len(contracts)} 个合同：")
            for c in contracts:
                print(f"    - {c['contract_no']}")
                print(f"      工程范围: {c['work_scope']}")
                print(f"      合同金额: {c['amount']:,.2f} 泰铢")
                print(f"      协作体: {c.get('company_name', '')}")


def create_contract_info_json(contract_data: Dict) -> Dict:
    """
    创建合同信息JSON字典
    
    Args:
        contract_data: 包含合同基础信息的字典
    
    Returns:
        合同信息JSON字典
    """
    
    return {
        "项目名称": contract_data.get('project_name', ''),
        "合同编号": contract_data.get('contract_no', ''),
        "合同类型": "",
        "对方公司名称": contract_data.get('company_name', ''),
        "合同签订日期": "",
        "生效日期": "",
        "结束日期": "",
        "工期/交货期": "",
        "币种": "",  # 默认空白
        "汇率": 1.0,  # 默认1
        "合同金额": contract_data.get('amount', 0.0),
        "含税金额": 0.0,
        "不含税金额": 0.0,
        "税率": "",
        "付款方式": "",
        "履约保证金": "",
        "合同状态": "",
        "关联主合同编号": "",
        "联系人及电话": "",
        "PDF文件路径": "",
        "收付款记录": [],
        "备注": "",
        "工程范围": contract_data.get('work_scope', '')
    }


def create_missing_contracts(missing_contracts: List[Dict], dry_run: bool = True) -> Dict:
    """
    创建缺失的合同文件夹和JSON
    
    Args:
        missing_contracts: 缺失合同列表
        dry_run: 是否为演练模式（不实际创建）
    
    Returns:
        操作结果
    """
    results = {
        'success': [],
        'failed': [],
        'skipped': [],
        'total': len(missing_contracts)
    }
    
    mode_text = "【演练模式】" if dry_run else "【执行模式】"
    print(f"\n{mode_text} 开始创建缺失的合同...")
    
    for idx, contract in enumerate(missing_contracts, 1):
        try:
            contract_no = contract['contract_no']
            project_code = contract['project_code']
            project_name = contract['project_name']
            work_scope = contract['work_scope']
            amount = contract['amount']
            company_name = contract.get('company_name', '')
            
            print(f"\n[{idx}/{len(missing_contracts)}] {contract_no}")
            print(f"  项目: {project_name}")
            print(f"  工程范围: {work_scope}")
            print(f"  金额: {amount:,.2f} 泰铢")
            
            # 检查是否有分包合同文件夹路径
            if 'subpackage_folder' not in contract:
                error_msg = "找不到分包合同文件夹路径"
                print(f"  ✗ 跳过: {error_msg}")
                results['skipped'].append({
                    'contract_no': contract_no,
                    'reason': error_msg
                })
                continue
            
            subpackage_folder = contract['subpackage_folder']
            
            # 获取下一个合同序号
            next_num = get_next_contract_number(subpackage_folder)
            
            # 创建文件夹名称
            # 清理工程范围中的非法字符
            work_scope_str = str(work_scope) if work_scope else '未命名'
            safe_work_scope = re.sub(r'[<>:"/\\|?*]', '_', work_scope_str)
            folder_name = f"{next_num}_{contract_no}_{safe_work_scope}"
            contract_folder = os.path.join(subpackage_folder, folder_name)
            
            print(f"  文件夹: {folder_name}")
            
            if not dry_run:
                # 创建文件夹
                os.makedirs(contract_folder, exist_ok=True)
                
                # 创建合同信息.json
                contract_info = create_contract_info_json({
                    'project_name': project_name,
                    'contract_no': contract_no,
                    'work_scope': work_scope,
                    'amount': amount,
                    'company_name': company_name
                })
                
                json_file = os.path.join(contract_folder, '合同信息.json')
                with open(json_file, 'w', encoding='utf-8') as f:
                    json.dump(contract_info, f, ensure_ascii=False, indent=2)
                
                print(f"  ✓ 已创建")
                
                results['success'].append({
                    'contract_no': contract_no,
                    'folder': contract_folder,
                    'json_file': json_file
                })
            else:
                print(f"  → 将创建于: {contract_folder}")
                results['success'].append({
                    'contract_no': contract_no,
                    'folder': contract_folder
                })
            
        except Exception as e:
            error_msg = f"处理异常: {str(e)}"
            print(f"  ✗ 错误: {error_msg}")
            results['failed'].append({
                'contract_no': contract.get('contract_no', 'Unknown'),
                'error': error_msg
            })
    
    return results


def main():
    """主函数"""

    summary_dir = resolve_summary_dir(SUMMARY_DIR_CANDIDATES)
    excel_file = find_latest_summary_file(summary_dir)
    
    print("="*80)
    print("分包合同完整性检查与自动创建工具")
    print("="*80)
    print(f"\n汇总目录: {summary_dir}")
    print(f"Excel文件: {excel_file if excel_file else '未找到可用文件'}")
    print(f"合同目录: {CONTRACT_BASE_DIR}")
    print(f"金额容差: {AMOUNT_TOLERANCE} 泰铢\n")
    
    # 检查文件是否存在
    if not excel_file or not os.path.exists(excel_file):
        print(f"错误：未在目录中找到可用汇总Excel: {summary_dir}")
        return
    
    if not os.path.exists(CONTRACT_BASE_DIR):
        print(f"错误：合同目录不存在: {CONTRACT_BASE_DIR}")
        return
    
    # 第一步：分析完整性
    analysis = analyze_contract_completeness(excel_file, CONTRACT_BASE_DIR)
    
    # 第二步：打印摘要
    print_analysis_summary(analysis)
    
    # 第三步：如果有缺失合同，询问是否创建
    if not analysis['missing_contracts']:
        print("\n" + "="*80)
        print("✓ 所有合同已完整，无需创建")
        print("="*80)
        return
    
    print("\n" + "="*80)
    print("待创建合同预览")
    print("="*80)
    
    # 按项目分组显示待创建清单，并为每个合同编号
    by_project = {}
    for contract in analysis['missing_contracts']:
        proj = contract['project_code']
        if proj not in by_project:
            by_project[proj] = []
        by_project[proj].append(contract)
    
    total_to_create = len(analysis['missing_contracts'])
    print(f"\n共 {total_to_create} 个合同需要创建：\n")
    
    # 创建编号映射
    contract_index_map = {}
    current_index = 1
    
    for project_code, contracts in by_project.items():
        print(f"【{contracts[0]['project_name']}】({project_code}) - {len(contracts)} 个合同")
        for c in contracts:
            # 生成预览的文件夹名称
            next_num = get_next_contract_number(c.get('subpackage_folder', '')) if 'subpackage_folder' in c else '?'
            work_scope_str = str(c.get('work_scope', '')) if c.get('work_scope') else '未命名'
            safe_work_scope = re.sub(r'[<>:"/\\|?*]', '_', work_scope_str)
            folder_preview = f"{next_num}_{c['contract_no']}_{safe_work_scope}"
            
            # 显示编号
            print(f"  [{current_index}] {folder_preview}")
            contract_index_map[current_index] = c
            current_index += 1
    
    # 询问用户确认并选择要创建的合同
    print("\n" + "="*80)
    print(f"⚠️  检测到 {total_to_create} 个缺失合同")
    print("="*80)
    print("\n请输入要创建的合同编号：")
    print("  - 输入编号，用逗号分隔（如：1,2,3 或 1-5,8,10）")
    print("  - 输入 'all' 或 'a' 创建所有合同")
    print("  - 直接回车取消操作")
    
    user_input = input("\n请输入: ").strip()
    
    if not user_input:
        print("\n❌ 用户取消操作")
        return
    
    # 解析用户输入
    selected_contracts = []
    
    if user_input.lower() in ['all', 'a']:
        selected_contracts = list(analysis['missing_contracts'])
        print(f"\n✓ 已选择所有 {len(selected_contracts)} 个合同")
    else:
        selected_indices = set()
        
        try:
            # 解析输入（支持 1,2,3 和 1-5 格式）
            parts = user_input.replace('，', ',').split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    # 范围格式 1-5
                    start, end = part.split('-')
                    selected_indices.update(range(int(start.strip()), int(end.strip()) + 1))
                else:
                    # 单个编号
                    selected_indices.add(int(part))
            
            # 根据编号获取合同
            for idx in sorted(selected_indices):
                if idx in contract_index_map:
                    selected_contracts.append(contract_index_map[idx])
                else:
                    print(f"  ⚠️  编号 {idx} 无效，已跳过")
            
            if not selected_contracts:
                print("\n❌ 未选择有效的合同编号，操作取消")
                return
            
            print(f"\n✓ 已选择 {len(selected_contracts)} 个合同")
            print("  选中的合同：")
            for c in selected_contracts:
                print(f"    - {c['contract_no']} ({c['project_name']})")
        
        except ValueError:
            print("\n❌ 输入格式错误，操作取消")
            return
    
    # 最终确认
    final_confirm = input(f"\n确认创建以上 {len(selected_contracts)} 个合同？(输入 y/yes 确认): ").strip().lower()
    
    if final_confirm not in ['y', 'yes']:
        print("\n❌ 用户取消操作")
        return
    
    print("\n✓ 确认创建，开始执行...")
    
    # 第四步：执行创建
    results = create_missing_contracts(selected_contracts, dry_run=False)
    
    # 第五步：显示结果
    print("\n" + "="*80)
    print("创建完成")
    print("="*80)
    print(f"总计: {results['total']}")
    print(f"成功: {len(results['success'])}")
    print(f"失败: {len(results['failed'])}")
    print(f"跳过: {len(results['skipped'])}")
    
    if results['success']:
        print("\n✓ 成功创建的合同:")
        for item in results['success']:
            print(f"  - {item['contract_no']}")
            if 'json_file' in item:
                print(f"    {item['json_file']}")
    
    if results['failed']:
        print("\n✗ 失败的合同:")
        for item in results['failed']:
            print(f"  - {item['contract_no']}")
            print(f"    原因: {item['error']}")
    
    if results['skipped']:
        print("\n⊘ 跳过的合同:")
        for item in results['skipped']:
            print(f"  - {item['contract_no']}")
            print(f"    原因: {item['reason']}")
    
    print("\n" + "="*80)
    print("✅ 所有操作已完成！")
    print("="*80)


if __name__ == '__main__':
    main()
