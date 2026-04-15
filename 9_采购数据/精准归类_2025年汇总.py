#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
精准材料归类系统 + 2025年材料采购汇总报告

基于实际物品名称分析，建立完善的材料归类规则
重点生成2025年度材料采购情况汇总
"""
import pandas as pd
import numpy as np
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import re
import difflib
import os

class PreciseMaterialClassifier:
    """精准材料分类器"""
    
    def _build_super_category_mapping(self):
        """建立超级分类映射规则（50+个超级分类）"""
        return {
            '螺栓螺丝类': ['螺丝', '螺母', '螺栓', '扣件', '山型', '山形', '燕尾', '地脚', '膨胀', '化学螺栓', 
                      '止水螺', '对拉', '通丝', '三段', '自攻', '钻尾', '批头'],
            '五金工具': ['内六角', '卡箍', '压力表', '套筒', '扳手', '钳', '水表', '马鞍', '喉箍', 
                     'U型卡', '卡头', '线管', '网管', '过墙管', '线盒', '转扣', '套管', '热缩管', 
                     '航空插头', '插头', '开孔器', '导轨', '卡轨'],
            '铁丝铁钉类': ['铁丝', '铁钉', '钢钉', '扎丝', '镀锌铁丝', '镀锌扎丝', '钉', '镀锌钉', '铁钉', '钢钉'],
            '井类及沟盖': ['井盖', '篦子', '箅子', '雨水井', '污水井', '检修孔盖', '沟盖', '电缆井', '污水井盖', '雨水井盖'],
            '集装箱及板房': ['集装箱', '板房', '活动房', '预制墙板', '围墙板', '会议室', '厕所箱', '保安室'],
            '电梯及货梯': ['电梯', '货梯', '客梯'],
            '围栏及护栏': ['围栏', '护栏', '走道', '外楼梯'],
            '化粪池及隔油池': ['化粪池', '隔油池', '玻璃钢化粪池'],
            '钢材型材': ['工字钢', '方钢', '槽钢', '镀锌方管', '镀锌钢管', '镀锌扁', '圆钢', '角钢', 
                     '不锈钢板', 'C型钢', '扁钢', '角铁', '钢管', '方管'],
            '灯具及照明': ['灯', 'LED', '路灯', '泛光灯', '工厂灯', '应急灯', '日光灯', '吸顶灯', 
                      '标志灯', '诱导灯', '灯口', '节能灯', '筒灯'],
            '油漆及涂料': ['油漆', '涂料', '底漆', '冷底油', '防锈', '稀释剂', '稀料', '自喷漆', '漆料'],
            '泵及水泵': ['泵', '潜水泵', '污水泵', '水泵', '加压泵'],
            'PE管及配件': ['PE膜', 'PE管', 'PE闸阀', 'PE法兰', 'PE三通', 'PE弯头', 'PE变径', 'PE直接', 
                       'PE承插', 'PE对接', 'PE内丝', 'PE外丝', 'PE-PPR'],
            'PPR管及配件': ['PPR', 'ppr', 'PP-R'],
            'PVC管及配件': ['PVC', 'pvc'],
            '阀门及配件': ['阀门', '阀', '过滤器', '止回阀', '球阀', '闸阀'],
            '法兰及垫片': ['法兰', '垫片', '同心大小头', '盲板', '皮垫'],
            '弯头及三通': ['弯头', '三通', '变径', '卡箍弯头', '电熔焊三通', '电熔焊弯头', '电熔焊变径'],
            '电气柜及箱': ['配电柜', '配电箱', '照明配电箱', '巡检柜', '直流屏', '端子箱', '插座箱', 
                      '开关柜', '出线柜', '进线柜', '充气柜', 'PT柜', '联络柜', '补偿柜', '转换柜', '电控柜'],
            '电缆及电线': ['电缆', '电线', '照明线', '铜线', '铝电缆', 'YJLV', '铜鼻', '铝鼻', '铜排'],
            '变压器及开关': ['变压器', '干式变压器', '开关', '断路器', '漏电断路器', '倒顺开关'],
            '机械设备': ['搅拌', '装载机', '泵站', '抹光机', '划线机', '登高车', '发电机', '电焊机', 
                     '角磨机', '洗车机', '推车', '斗车', '吊带', '打印机', '弯箍机', '切割机', 
                     '制砖机', '地磅', '卸货平台'],
            '钢筋及配件': ['钢筋', '套筒', '钢筋网', '扎钩'],
            '门窗类': ['门', '窗', '卷帘门', '平开门', '滑段门', '伸缩门', '道闸', '人脸识别'],
            '防水保温材料': ['防水', '保温', '挤塑板', '岩棉', 'SBS', 'APP', '丙纶布', '聚氨酯防水', 
                        '防水套管', '止水带', '渗透结晶', '阻燃', '安全网'],
            '砖砌体类': ['加气块', '实心砖', '空心砖', '人行道砖', '水泥砖'],
            '混凝土及砂浆': ['混凝土', '砂浆', '抹灰', '砌筑', '粉刷', '墙缝', '商砼', '振动', '振捣'],
            '水泥': ['水泥', '硫铝酸盐水泥', '白水泥', '普通硅酸盐水泥', '水泥制品'],
            '砂石': ['砂', '沙', '碎石', '炉渣', '砂轮'],
            '管材管件': ['波纹管', 'HDPE管', '雨水管', '软水管', '给水管', '检查口', '隔断'],
            '装饰材料': ['地砖', '瓷砖', '墙砖', '路沿石', '腻子', '腻子粉', '地坪', '固体沥青'],
            '脚手架及支撑': ['脚手架', '盘扣', '横杆', '立杆', '跳板', '支撑', '顶托', '底托', '斜杆', '拉杆'],
            '预制桩及桩基': ['桩', '方桩', '工字桩', '打桩', '引孔', '运费', '预制', '混凝土桩'],
            '测试及检测': ['大应变', '小应变', '试装', '进出场费'],
            '安全及标识': ['安全标识', '马甲', '反光背心', '安全帽', '警示', '条幅', '标语', '宣传栏', '目镜'],
            '办公家具': ['办公桌', '桌椅', '会议桌', '餐椅', '茶桌', '茶几', '沙发', '椅子', '文件柜', '餐桌'],
            '胶及注浆': ['胶', '注浆', '植筋胶', '密封胶', '胶带', '胶布', '绝缘胶', '胶泥', '胶水'],
            '切割工具': ['切割片', '锯片', '锯', '美工刀'],
            '手工具': ['直尺', '刮尺', '步步紧', '水平仪', '卷尺', '水带', '盒尺', '铁锹', '铁铲', 
                    '锤', '手套', '扫把', '灰盆', '灰桶', '放线仪', '万用表', '灰铲', '托灰板', 
                    '墨斗', '墨汁', '剪刀', '编织袋', '工具包', '剥线钳'],
            '电工工具': ['冲击钻', '电锤', '电镐', '洋镐', '电笔', '钻头'],
            '灭火及消防': ['灭火器', '消火栓', '消防', '水桶'],
            '网格布及网': ['网格布', '纤维网', '玻璃纤维', '钢丝网', '电焊网', '遮阳网', '彩条布', '打包带'],
            '预埋件': ['预埋'],
            '土方回填': ['土方', '回填', '强夯'],
            # 新增单独分类
            '木方': ['木方', '方木'],
            '钢板': ['钢板', '花纹板', '花纹钢板'],
            '空心砖': ['空心砖'],
            '实心砖': ['实心砖'],
            '彩钢瓦及钢格栅': ['彩钢瓦', '彩钢', '钢格栅', '格栅', '铁格栅', '钢格板'],
            '其他杂项': []  # 默认类别
        }
    
    def _map_to_super_category(self, item_name):
        """将物品名称映射到45个超级分类之一"""
        if pd.isna(item_name):
            return '其他杂项'
        
        item_name_str = str(item_name).lower()
        
        # 遍历所有超级分类规则
        for super_cat, keywords in self.super_category_rules.items():
            if any(kw.lower() in item_name_str for kw in keywords):
                return super_cat
        
        return '其他杂项'
    
    def __init__(self, excel_path, similarity_threshold=0.85, min_group_size=2):
        self.excel_path = excel_path
        self.df = None
        self.similarity_threshold = similarity_threshold
        self.min_group_size = min_group_size
        self.similarity_groups = []
        self.similarity_mapping = {}
        self.super_category_rules = self._build_super_category_mapping()  # 新增：超级分类规则
        self.classification_rules = self._build_classification_rules()
        self.load_data()
        
    def _build_classification_rules(self):
        """建立完善的材料分类规则"""
        return {            '桩施工': {
                'keywords': ['桩施工', '施工', '打桩', '引孔', '试桩', '运费', '大小应变'],
                'sub_class': ['施工服务', '运输', '测试'],
                'priority': 10
            },            '钢筋类': {
                'keywords': ['钢筋', 'rebar', 'steel bar', '螺纹钢', '钢筋网片', '钢筋网'],
                'sub_class': ['螺纹钢筋', '钢筋网', '钢筋制品'],
                'priority': 10
            },
            '桩基类': {
                'keywords': ['桩', 'pile', '预制桩', '工程桩', '方桩', '工字桩'],
                'exclude': ['桩机', '运费', '引孔', '试桩', '大小应变'],
                'sub_class': ['预制桩', '工程桩', '桩材'],
                'priority': 9
            },
            '混凝土及砂浆': {
                'keywords': ['混凝土', 'concrete', '砂浆', 'mortar', '抹灰', '砌筑', '粉刷', '商砼'],
                'sub_class': ['混凝土', '砂浆', '抹灰材料'],
                'priority': 8
            },
            '空心砖': {
                'keywords': ['空心砖'],
                'sub_class': ['空心砖'],
                'priority': 9
            },
            '实心砖': {
                'keywords': ['实心砖'],
                'sub_class': ['实心砖'],
                'priority': 9
            },
            '加气块': {
                'keywords': ['加气块'],
                'sub_class': ['加气块'],
                'priority': 9
            },
            '砖砌体类': {
                'keywords': ['砖', 'brick', '砌块'],
                'exclude': ['空心砖', '实心砖', '加气块'],
                'sub_class': ['砌块', '砖类'],
                'priority': 7
            },
            '碎石': {
                'keywords': ['碎石', 'stone', '石子'],
                'exclude': ['砂浆', '抹灰'],
                'sub_class': ['碎石', '石子'],
                'priority': 9
            },
            '砂子': {
                'keywords': ['沙子', '砂子', '砂', '黄沙'],
                'exclude': ['砂浆', '抹灰', '碎石'],
                'sub_class': ['细砂', '粗砂', '黄沙'],
                'priority': 9
            },
            '炉渣': {
                'keywords': ['炉渣'],
                'sub_class': ['炉渣', '废渣'],
                'priority': 9
            },
            '水泥管': {
                'keywords': ['水泥管'],
                'sub_class': ['水泥管', '排水管'],
                'priority': 9
            },
            '水泥制品': {
                'keywords': ['水泥', 'cement'],
                'exclude': ['砂浆', '砖', '抹灰', '水泥管'],
                'sub_class': ['水泥', '水泥制品'],
                'priority': 7
            },
            '电气材料': {
                'keywords': ['电缆', 'cable', '配电箱', '配电柜', '开关', '变压器', '变频', '电线', 
                           '电梯', '低压', '高压', '出线柜', '进线柜', '照明', '铜电缆'],
                'sub_class': ['电缆', '配电设备', '开关装置', '照明'],
                'priority': 9
            },
            '钢材类': {
                'keywords': ['钢板', '钢管', '角钢', '槽钢', 'h型钢', '方管', '圆钢', '预埋板', 
                           '预埋钢板', '镀锌', '花纹钢板'],
                'exclude': ['钢筋', '螺纹钢'],
                'sub_class': ['钢板', '型钢', '预埋件'],
                'priority': 8
            },
            '模版': {
                'keywords': ['模板', '木模'],
                'exclude': ['脚手架', '木方'],
                'sub_class': ['模板', '模具'],
                'priority': 9
            },
            '木方': {
                'keywords': ['木方'],
                'exclude': ['模板', '脚手架'],
                'sub_class': ['木方', '方木'],
                'priority': 9
            },
            '脚手架': {
                'keywords': ['脚手架', '支撑'],
                'exclude': ['模板', '木方'],
                'sub_class': ['脚手架', '支撑架'],
                'priority': 9
            },
            '门': {
                'keywords': ['门', '卷帘门', '防火门', '平开门'],
                'exclude': ['窗', '阀门'],
                'sub_class': ['门', '卷帘门', '防火门'],
                'priority': 9
            },
            '窗': {
                'keywords': ['窗', 'window'],
                'exclude': ['门', '阀门'],
                'sub_class': ['窗户', '玻璃窗'],
                'priority': 9
            },
            '防水保温': {
                'keywords': ['防水', '保温', '挤塑板', '岩棉', '泡沫', '卷材'],
                'sub_class': ['防水材料', '保温材料'],
                'priority': 7
            },
            '管材管件': {
                'keywords': ['管', 'pipe', 'pvc', 'hdpe', 'upvc', '管道', '管件'],
                'exclude': ['水泥管', '电缆管', '门', '窗', '阀门'],
                'sub_class': ['塑料管', '金属管', '管件'],
                'priority': 6
            },
            '五金材料': {
                'keywords': ['螺栓', '螺丝', '螺母', '卡扣', '扣件', '连接件', '铁丝', '钉'],
                'sub_class': ['螺栓类', '螺丝螺母', '卡扣件', '铁丝钉'],
                'priority': 5
            },
            '井盖及附属': {
                'keywords': ['井盖', '雨水井', '污水井', '井篦子', '篦子'],
                'sub_class': ['井盖', '井类'],
                'priority': 7
            },
            '机械设备': {
                'keywords': ['搅拌', '机', '设备', '吊', '车', '泵', '装载机', '挖掘'],
                'exclude': ['搅拌站', '制砖机'],
                'sub_class': ['搅拌机', '运输机械', '起重机械'],
                'priority': 4
            },
            '装饰材料': {
                'keywords': ['瓷砖', '地砖', '墙砖', '涂料', '腻子', '石材', '地坪', '路沿石'],
                'sub_class': ['瓷砖', '涂料', '石材'],
                'priority': 6
            },
            '土方及回填': {
                'keywords': ['土方', '回填', '夯', '强夯'],
                'sub_class': ['土方', '回填', '夯实'],
                'priority': 6
            },
            '临建设施': {
                'keywords': ['集装箱', '板房', '活动房', '保安室'],
                'sub_class': ['活动房', '集装箱'],
                'priority': 5
            },
            '其他材料': {
                'keywords': [],
                'sub_class': ['其他'],
                'priority': 0
            }
        }
    
    def load_data(self):
        """加载数据"""
        print(f"正在加载数据: {self.excel_path}")
        self.df = pd.read_excel(self.excel_path, sheet_name='合并数据')
        
        # 确保数值列为数值类型
        numeric_cols = ['数量', '单价', '不含税金额', '含税金额', '汇率']
        for col in numeric_cols:
            if col in self.df.columns:
                self.df[col] = pd.to_numeric(self.df[col], errors='coerce')
        
        print(f"数据加载完成，共 {len(self.df)} 条记录")

    def _normalize_name(self, name):
        """物品名称标准化（用于相似度计算）"""
        if pd.isna(name):
            return ''
        text = str(name).strip().lower()
        # 去除空格与常见符号
        text = re.sub(r"\s+", "", text)
        text = re.sub(r"[\-_/()（）【】\[\]{}<>《》·,，。;；:：!！?？'\"]", "", text)
        return text

    def _build_similarity_groups(self):
        """基于物品名称去重后进行相似归类"""
        print("\n正在进行物品名称相似归类...")
        unique_names = pd.Series(self.df['物品名称'].dropna().unique())
        normalized = unique_names.apply(self._normalize_name)
        name_pairs = list(zip(unique_names.tolist(), normalized.tolist()))

        # 按首字符与长度分桶，减少比较次数
        buckets = {}
        for original, norm in name_pairs:
            if not norm:
                continue
            key = (norm[0], len(norm) // 4)
            buckets.setdefault(key, []).append((original, norm))

        grouped = []
        used = set()

        for bucket_items in buckets.values():
            for i, (name_i, norm_i) in enumerate(bucket_items):
                if name_i in used:
                    continue
                group = [name_i]
                used.add(name_i)
                for name_j, norm_j in bucket_items[i + 1:]:
                    if name_j in used:
                        continue
                    ratio = difflib.SequenceMatcher(None, norm_i, norm_j).ratio()
                    contains = norm_i in norm_j or norm_j in norm_i
                    if ratio >= self.similarity_threshold or (ratio >= 0.65 and contains):
                        group.append(name_j)
                        used.add(name_j)
                if len(group) >= self.min_group_size:
                    grouped.append(group)

        self.similarity_groups = grouped
        print(f"相似归类完成：共 {len(grouped)} 组")

    def _apply_similarity_grouping(self):
        """将相似归类结果映射为标准物品名称"""
        self._build_similarity_groups()

        # 构建映射：选最短且最常见的名称作为标准名
        mapping = {}
        for group in self.similarity_groups:
            freq = self.df['物品名称'].value_counts()
            group_df = pd.DataFrame({'name': group})
            group_df['freq'] = group_df['name'].map(freq).fillna(0)
            group_df['len'] = group_df['name'].astype(str).str.len()
            standard_name = group_df.sort_values(['freq', 'len'], ascending=[False, True]).iloc[0]['name']
            for name in group:
                if name != standard_name:
                    mapping[name] = standard_name

        self.similarity_mapping = mapping

        # 生成标准物品名称列
        self.df['标准物品名称'] = self.df['物品名称'].replace(mapping)
        print(f"标准物品名称生成完成：映射 {len(mapping)} 个名称")
    
    def classify_item(self, item_name, current_category=None):
        """对单个物品进行分类"""
        if pd.isna(item_name):
            return '未知物品', '未知'

        # 若已有分类，直接采用（优先级最高）
        if pd.notna(current_category) and str(current_category).strip() != '':
            category = str(current_category).strip()
            sub_class = self.classification_rules.get(category, {}).get('sub_class', [''])[0]
            return category, sub_class
        
        item_name_lower = str(item_name).lower()
        
        # 按优先级排序的规则
        sorted_rules = sorted(
            self.classification_rules.items(),
            key=lambda x: x[1]['priority'],
            reverse=True
        )
        
        for category, rule in sorted_rules:
            # 检查排除词
            if 'exclude' in rule:
                if any(ex.lower() in item_name_lower for ex in rule['exclude']):
                    continue
            
            # 检查关键词
            keywords = rule.get('keywords', [])
            if any(kw.lower() in item_name_lower for kw in keywords):
                sub_class = rule.get('sub_class', [''])[0]
                return category, sub_class
        
        return '其他材料', '其他'
    
    def apply_classification(self):
        """应用分类到所有数据"""
        print("\n正在进行精准材料分类...")
        
        # 对每一行应用分类（优先使用标准物品名称）
        name_col = '标准物品名称' if '标准物品名称' in self.df.columns else '物品名称'
        
        # 创建分类结果列
        classify_results = self.df.apply(
            lambda row: self.classify_item(row[name_col], row['材料归类']),
            axis=1
        )
        
        # 分离主分类和小类
        self.df['精准材料归类'] = classify_results.apply(lambda x: x[0])
        self.df['材料小类'] = classify_results.apply(lambda x: x[1])
        
        # 统计分类效果
        original_null = self.df['材料归类'].isna().sum()
        new_classified = self.df[self.df['材料归类'].isna() & 
                                 (self.df['精准材料归类'] != '其他材料')].shape[0]
        
        print(f"✓ 分类完成")
        print(f"  原始空归类: {original_null} 条")
        print(f"  新分类成功: {new_classified} 条")
        print(f"  仍为其他: {self.df[self.df['精准材料归类'] == '其他材料'].shape[0]} 条")
        
        # 显示分类分布
        print("\n精准材料归类分布:")
        category_dist = self.df['精准材料归类'].value_counts()
        for cat, count in category_dist.items():
            pct = count / len(self.df) * 100
            print(f"  {cat:15s}: {count:5d} 条 ({pct:5.1f}%)")
    
    def generate_2025_summary(self):
        """生成2025年材料采购汇总"""
        print("\n" + "="*80)
        print("生成2025年材料采购汇总报告")
        print("="*80)
        
        # 筛选2025年数据
        df_2025 = self.df[self.df['年'] == 2025].copy()
        print(f"\n2025年数据: {len(df_2025)} 条记录")
        print(f"2025年总金额: {df_2025['含税金额'].sum()/1e8:.2f} 亿泰铢")
        
        # 1. 按精准材料归类汇总
        summary_by_category = df_2025.groupby('精准材料归类').agg({
            '含税金额': 'sum',
            '数量': 'sum',
            '序号': 'count',
            '标准单位': lambda x: x.mode()[0] if len(x.mode()) > 0 and x.nunique() == 1 else '',
            '材料小类': lambda x: '\n'.join(sorted(set(x.dropna().astype(str))))  # 所有小类用换行符合并
        }).reset_index()
        summary_by_category.columns = ['材料归类', '总金额', '总数量', '记录条数', '单位', '小类']
        summary_by_category['占比%'] = summary_by_category['总金额'] / summary_by_category['总金额'].sum() * 100
        # 调整列顺序
        summary_by_category = summary_by_category[['材料归类', '小类', '总金额', '总数量', '单位', '记录条数', '占比%']]
        summary_by_category = summary_by_category.sort_values('总金额', ascending=False)
        
        # 2. 按材料归类+物品名称汇总
        name_col = '标准物品名称' if '标准物品名称' in df_2025.columns else '物品名称'
        df_2025_detail = df_2025.copy()
        df_2025_detail['标准单位'] = df_2025_detail['标准单位'].fillna('未填')
        detail_by_item = df_2025_detail.groupby(['精准材料归类', name_col, '标准单位'], dropna=False).agg({
            '含税金额': 'sum',
            '数量': 'sum',
            '单价': 'mean',  # 平均单价
            '序号': 'count'
        }).reset_index()
        detail_by_item.columns = ['材料归类', '物品名称', '单位', '总金额', '总数量', '平均单价', '记录条数']
        detail_by_item = detail_by_item.sort_values(['材料归类', '总金额'], ascending=[True, False])
        
        # 3. 按月度+材料归类汇总（保留空月份）
        df_2025_month = df_2025.copy()
        df_2025_month['月'] = df_2025_month['月'].fillna('未填')
        monthly_summary = df_2025_month.groupby(['月', '精准材料归类'], dropna=False).agg({
            '含税金额': 'sum'
        }).reset_index()
        monthly_pivot = monthly_summary.pivot_table(
            values='含税金额',
            index='精准材料归类',
            columns='月',
            aggfunc='sum',
            fill_value=0
        )
        monthly_pivot['全年合计'] = monthly_pivot.sum(axis=1)
        monthly_pivot = monthly_pivot.sort_values('全年合计', ascending=False)
        
        # 4. 全量物品明细（含规格）
        # 先转换规格为字符串避免类型混合
        df_2025_copy = df_2025.copy()
        df_2025_copy['规格'] = df_2025_copy['规格'].astype(str)
        df_2025_copy['标准单位'] = df_2025_copy['标准单位'].fillna('未填').astype(str)
        if '标准物品名称' not in df_2025_copy.columns:
            df_2025_copy['标准物品名称'] = df_2025_copy['物品名称']
        
        top_items = df_2025_copy.groupby(['精准材料归类', '标准物品名称', '规格', '标准单位'], dropna=False).agg({
            '含税金额': 'sum',
            '数量': 'sum',
            '单价': 'mean',
            '供货商名称': lambda x: '、'.join(str(s) for s in x.dropna().unique()[:3])  # 前3个供应商
        }).reset_index()
        top_items.columns = ['材料归类', '物品名称', '规格', '单位', '总金额', '总数量', '平均单价', '主要供货商']
        top_items = top_items.sort_values('总金额', ascending=False)
        
        # 5. 项目维度分析（保留空项目）
        df_2025_project = df_2025.copy()
        df_2025_project['标准项目名称'] = df_2025_project['标准项目名称'].fillna('未填')
        project_summary = df_2025_project.groupby(['标准项目名称', '精准材料归类'], dropna=False).agg({
            '含税金额': 'sum'
        }).reset_index()
        project_pivot = project_summary.pivot_table(
            values='含税金额',
            index='标准项目名称',
            columns='精准材料归类',
            aggfunc='sum',
            fill_value=0
        )
        
        return {
            '材料归类汇总': summary_by_category,
            '材料明细汇总': detail_by_item,
            '月度采购趋势': monthly_pivot,
            '物品明细汇总': top_items,
            '项目材料分布': project_pivot,
            '原始数据_2025': df_2025
        }
    
    def export_2025_report(self, summaries, output_path=None):
        """导出2025年专项报告"""
        if output_path is None:
            timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
            output_path = f'2025年材料采购汇总报告_{timestamp}.xlsx'
        
        print(f"\n正在导出2025年报告: {output_path}")
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 1. 材料归类汇总（主报告）
            summaries['材料归类汇总'].to_excel(writer, sheet_name='材料归类汇总', index=False)
            
            # 2. 材料明细汇总
            summaries['材料明细汇总'].to_excel(writer, sheet_name='材料明细汇总', index=False)
            
            # 3. 月度采购趋势
            summaries['月度采购趋势'].to_excel(writer, sheet_name='月度采购趋势')
            
            # 4. 物品明细（全量，含规格）
            summaries['物品明细汇总'].to_excel(writer, sheet_name='物品明细汇总', index=False)
            
            # 5. 项目材料分布
            summaries['项目材料分布'].to_excel(writer, sheet_name='项目材料分布')

            # 6. 相似归类映射
            if self.similarity_mapping:
                mapping_df = pd.DataFrame({
                    '原始物品名称': list(self.similarity_mapping.keys()),
                    '标准物品名称': list(self.similarity_mapping.values())
                }).sort_values('标准物品名称')
                mapping_df.to_excel(writer, sheet_name='物品名称相似归类', index=False)
            
            # 7. 原始数据（2025年，带精准分类）
            export_cols = ['标准项目名称', '年', '月',
                          '材料归类', '精准材料归类', '材料小类', '物品名称', '标准物品名称', '规格', 
                          '单位', '标准单位', '数量', '单价', '含税金额',
                          '供货商名称', '备注']
            df_raw = summaries['原始数据_2025']
            df_export = df_raw[[col for col in export_cols if col in df_raw.columns]]
            df_export.to_excel(writer, sheet_name='2025年原始数据', index=False)
        
        # 格式化Excel
        self._format_2025_report(output_path, summaries)
        
        print(f"✓ 报告导出完成: {output_path}")
        return output_path
    
    def _format_2025_report(self, file_path, summaries):
        """格式化2025年报告"""
        wb = openpyxl.load_workbook(file_path)
        
        # 定义样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        title_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        title_font = Font(bold=True, size=12)
        number_format = '#,##0.00'
        border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # 格式化每个工作表
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            
            # 设置标题行样式
            ws.row_dimensions[1].height = 40  # 设置表头行高
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
                cell.border = border
            
            # 自动调整列宽
            for column in ws.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if cell.value:
                            max_length = max(max_length, len(str(cell.value)))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 60)
                ws.column_dimensions[column_letter].width = adjusted_width
            
            # 数字列格式化
            if sheet_name in ['材料归类汇总', '材料明细汇总', 'Top100物品明细', '物品明细汇总']:
                for row in ws.iter_rows(min_row=2):
                    for cell in row:
                        if cell.column_letter in ['B', 'C', 'E', 'F', 'G', 'H']:  # 金额/数量列
                            cell.number_format = number_format
            
            # 添加自动筛选
            if ws.max_row > 1 and ws.max_column > 0:
                ws.auto_filter.ref = ws.dimensions
            
            # 冻结首行
            ws.freeze_panes = 'A2'
        
        # 在材料归类汇总表添加总计行
        ws_summary = wb['材料归类汇总']
        last_row = ws_summary.max_row
        ws_summary.cell(last_row + 1, 1).value = '总计'
        ws_summary.cell(last_row + 1, 1).font = Font(bold=True)
        ws_summary.cell(last_row + 1, 2).value = f'=SUM(B2:B{last_row})'
        ws_summary.cell(last_row + 1, 3).value = f'=SUM(C2:C{last_row})'
        ws_summary.cell(last_row + 1, 4).value = f'=SUM(D2:D{last_row})'
        
        for cell in ws_summary[last_row + 1]:
            cell.fill = title_fill
            cell.font = title_font
        
        wb.save(file_path)
    
    def print_2025_summary(self, summaries):
        """打印2025年汇总信息"""
        summary = summaries['材料归类汇总']
        
        print("\n" + "="*80)
        print("2025年材料采购汇总报告")
        print("="*80)
        
        total_amount = summary['总金额'].sum()
        total_records = summary['记录条数'].sum()
        
        print(f"\n总采购金额: {total_amount/1e8:.2f} 亿泰铢")
        print(f"总记录数: {total_records:,} 条")
        print(f"材料类别数: {len(summary)} 个")
        
        print("\n" + "-"*80)
        print(f"{'排名':<4} {'材料归类':<20} {'金额(万)':<15} {'占比%':<10} {'数量':<15}")
        print("-"*80)
        
        for idx, row in summary.head(20).iterrows():
            print(f"{idx+1:<4} {row['材料归类']:<20} {row['总金额']/10000:>13,.2f}  "
                  f"{row['占比%']:>8.1f}%  {row['总数量']:>13,.2f}")
        
        print("="*80)
    
    def run_full_process(self):
        """运行完整流程"""
        print("\n" + "="*80)
        print("精准材料归类 + 2025年材料采购汇总系统")
        print("="*80)
        
        # 1. 应用精准分类
        self._apply_similarity_grouping()
        self.apply_classification()
        
        # 2. 生成2025年汇总
        summaries_2025 = self.generate_2025_summary()
        
        # 3. 打印汇总信息
        self.print_2025_summary(summaries_2025)
        
        # 4. 导出报告
        output_file = self.export_2025_report(summaries_2025)
        
        # 5. 添加物品合并汇总工作表
        self.generate_merged_items_summary(output_file)
        
        # 6. 生成汇总分析工作簿
        self.generate_summary_analysis(output_file)
        
        # 7. 添加大小类按单位汇总工作表
        self.generate_category_subclass_unit_summary(output_file)
        
        # 7. 同时更新完整数据的分类
        self._export_full_data_with_classification()
        
        print("\n" + "="*80)
        print("处理完成！")
        print("="*80)
        print(f"\n已生成文件:")
        print(f"  1. {output_file} - 2025年材料采购专项报告")
        print(f"  2. 全数据_精准归类.xlsx - 所有年度数据（含精准归类）")
        
        return output_file
    
    def _export_full_data_with_classification(self):
        """导出所有数据（带精准分类）"""
        output_path = '全数据_精准归类.xlsx'
        print(f"\n正在导出完整数据: {output_path}")
        
        # 按年度汇总
        yearly_summary = self.df.groupby(['年', '精准材料归类']).agg({
            '含税金额': 'sum',
            '数量': 'sum',
            '序号': 'count'
        }).reset_index()
        yearly_summary.columns = ['年度', '材料归类', '总金额', '总数量', '记录数']
        
        yearly_pivot = yearly_summary.pivot_table(
            values='总金额',
            index='材料归类',
            columns='年度',
            aggfunc='sum',
            fill_value=0
        )
        yearly_pivot['总计'] = yearly_pivot.sum(axis=1)
        yearly_pivot = yearly_pivot.sort_values('总计', ascending=False)
        
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            # 年度汇总
            yearly_pivot.to_excel(writer, sheet_name='各年度材料归类汇总')
            
            # 完整数据
            export_cols = ['文件名称', '项目名称', '标准项目名称', '年', '月',
                          '材料归类', '精准材料归类', '物品名称', '标准物品名称', '规格',
                          '单位', '标准单位', '数量', '单价', '含税金额',
                          '供货商名称', '签订日期', '备注']
            df_export = self.df[[col for col in export_cols if col in self.df.columns]]
            df_export.to_excel(writer, sheet_name='完整数据', index=False)
        
        # 添加自动筛选和格式化
        wb = openpyxl.load_workbook(output_path)
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        for sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            # 设置表头行高和样式
            ws.row_dimensions[1].height = 40
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            if ws.max_row > 1 and ws.max_column > 0:
                ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = 'A2'
        wb.save(output_path)
        
        print(f"✓ 完整数据导出完成: {output_path}")
    
    def generate_merged_items_summary(self, output_file):
        """生成物品合并汇总工作表"""
        print(f"\n正在生成物品合并汇总...")
        
        # 读取2025年原始数据
        df_2025 = self.df[self.df['年'] == 2025].copy()
        
        # 确保标准物品名称存在
        if '标准物品名称' not in df_2025.columns:
            df_2025['标准物品名称'] = df_2025['物品名称']
        
        # 应用超级分类映射（45个大类）
        df_2025['超级物品分类'] = df_2025['标准物品名称'].apply(self._map_to_super_category)
        
        # 按超级分类+单位分组合并
        merged_data = []
        for (super_cat, unit), group in df_2025.groupby(['超级物品分类', '标准单位'], dropna=False):
            # 收集该组的所有数据
            specs = group['规格'].dropna().astype(str).unique()
            specs_str = '\n'.join(specs) if len(specs) > 0 else ''  # 列出所有规格，用换行符分隔
            
            # 收集标准物品名称（二级名称）
            std_names = group['标准物品名称'].dropna().unique()
            std_names_str = '、'.join(std_names) if len(std_names) > 0 else ''  # 列出所有标准物品名称
            
            # 收集原始物品名称（所有变体）
            original_names = group['物品名称'].dropna().unique()
            original_names_str = '、'.join(original_names) if len(original_names) > 0 else ''  # 列出所有原始名称
            
            suppliers = group['供货商名称'].dropna().unique()
            suppliers_str = '、'.join(suppliers[:5]) if len(suppliers) > 0 else ''  # 最多显示5个供应商
            
            merged_record = {
                '材料归类': group['精准材料归类'].mode()[0] if len(group['精准材料归类'].mode()) > 0 else group['精准材料归类'].iloc[0],
                '小类': group['材料小类'].mode()[0] if len(group['材料小类'].mode()) > 0 else group['材料小类'].iloc[0],
                '标准化物品名称': super_cat,
                '原始物品名称': original_names_str,
                '标准单位': unit,
                '总数量': group['数量'].sum(),
                '平均单价': group['单价'].mean(),
                '总金额': group['含税金额'].sum(),
                '规格型号': specs_str,
                '记录条数': len(group),
                '主要供货商': suppliers_str
            }
            merged_data.append(merged_record)
        
        merged_df = pd.DataFrame(merged_data)
        merged_df = merged_df.sort_values(['材料归类', '总金额'], ascending=[True, False])
        
        print(f"物品合并汇总: {len(merged_df)} 条汇总记录（按45个超级分类）")
        print(f"  总金额: {merged_df['总金额'].sum()/1e8:.2f} 亿泰铢")
        print(f"  覆盖超级分类数: {merged_df['标准化物品名称'].nunique()} 个")
        
        # 调整列顺序，优化显示
        column_order = ['材料归类', '小类', '标准化物品名称', '原始物品名称', 
                       '标准单位', '总数量', '平均单价', '总金额', '规格型号', 
                       '记录条数', '主要供货商']
        merged_df = merged_df[[col for col in column_order if col in merged_df.columns]]
        
        # 写入Excel
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            merged_df.to_excel(writer, sheet_name='物品合并汇总', index=False)
        
        # 格式化新工作表
        wb = openpyxl.load_workbook(output_file)
        ws = wb['物品合并汇总']
        
        # 设置表头样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        
        ws.row_dimensions[1].height = 40
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 设置列宽
        for column in ws.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if cell.value:
                        cell_text = str(cell.value)
                        max_length = max(max_length, max(len(line) for line in cell_text.split('\n')))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            ws.column_dimensions[column_letter].width = adjusted_width
        
        # 启用自动换行和自动筛选
        for row in ws.iter_rows(min_row=2):
            for cell in row:
                cell.alignment = Alignment(wrap_text=True, vertical='top')
                if cell.column_letter in ['G', 'H', 'I']:  # 数值列
                    cell.number_format = '#,##0.00'
        
        # 添加自动筛选
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = 'A2'
        
        wb.save(output_file)
        print(f"✓ 物品合并汇总工作表已添加")
    
    def generate_summary_analysis(self, output_file):
        """基于物品合并汇总数据生成汇总分析工作簿"""
        print(f"\n正在生成汇总分析工作簿...")
        
        # 读取物品合并汇总工作表
        merged_df = pd.read_excel(output_file, sheet_name='物品合并汇总')
        
        # 生成汇总分析数据（按材料归类+小类+标准化物品名称+标准单位聚合）
        summary_df = merged_df.groupby(['材料归类', '小类', '标准化物品名称', '标准单位'], dropna=False).agg({
            '总数量': 'sum',
            '总金额': 'sum'
        }).reset_index()
        
        # 排序
        summary_df = summary_df.sort_values(['材料归类', '小类', '总金额'], ascending=[True, True, False])
        
        print(f"汇总分析: {len(summary_df)} 条汇总记录")
        print(f"  总金额: {summary_df['总金额'].sum()/1e8:.2f} 亿泰铢")
        
        # 创建新的工作簿
        summary_output = output_file.replace('2025年材料采购汇总报告_', '品类汇总分析_')
        
        with pd.ExcelWriter(summary_output, engine='openpyxl') as writer:
            summary_df.to_excel(writer, sheet_name='品类汇总分析', index=False)
        
        # 格式化工作簿
        wb = openpyxl.load_workbook(summary_output)
        ws = wb['品类汇总分析']
        
        # 设置表头样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        ws.row_dimensions[1].height = 40
        
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 设置列宽
        ws.column_dimensions['A'].width = 15  # 材料归类
        ws.column_dimensions['B'].width = 15  # 小类
        ws.column_dimensions['C'].width = 20  # 标准化物品名称
        ws.column_dimensions['D'].width = 12  # 标准单位
        ws.column_dimensions['E'].width = 18  # 总数量
        ws.column_dimensions['F'].width = 18  # 总金额
        
        # 数据行样式
        number_format = '#,##0.00'
        for row in ws.iter_rows(min_row=2):
            for cell_idx, cell in enumerate(row, 1):
                if cell_idx == 5:  # 总数量
                    cell.number_format = number_format
                elif cell_idx == 6:  # 总金额
                    cell.number_format = number_format
                cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        # 添加自动筛选和冻结
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = 'A2'
        
        # 添加总计行
        last_row = ws.max_row
        ws.cell(last_row + 1, 1).value = '总计'
        ws.cell(last_row + 1, 1).font = Font(bold=True)
        ws.cell(last_row + 1, 5).value = f'=SUM(E2:E{last_row})'
        ws.cell(last_row + 1, 6).value = f'=SUM(F2:F{last_row})'
        
        title_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        title_font = Font(bold=True, size=12)
        for col in range(1, 7):
            cell = ws.cell(last_row + 1, col)
            cell.fill = title_fill
            cell.font = title_font
            if col in [5, 6]:
                cell.number_format = number_format
        
        wb.save(summary_output)
        print(f"✓ 汇总分析工作簿已生成: {summary_output}")
        return summary_output
    
    def generate_category_subclass_unit_summary(self, output_file):
        """生成大小类按单位汇总工作表"""
        print(f"\n正在生成大小类按单位汇总...")
        
        # 读取物品合并汇总工作表（已经按45个超级分类合并）
        merged_df = pd.read_excel(output_file, sheet_name='物品合并汇总')
        
        # 按材料归类+小类+标准化物品名称+标准单位分组汇总
        summary_data = merged_df.groupby(['材料归类', '小类', '标准化物品名称', '标准单位'], dropna=False).agg({
            '总金额': 'sum',
            '总数量': 'sum',
            '原始物品名称': 'first',  # 取第一个作为代表
            '记录条数': 'sum'
        }).reset_index()
        
        summary_data.columns = ['大类', '小类', '标准化物品名称', '单位', '总金额', '总数量', '主要物品', '记录条数']
        
        # 调整列顺序
        summary_data = summary_data[['大类', '小类', '标准化物品名称', '总金额', '总数量', '单位', '记录条数']]
        
        # 按大类、小类、标准化物品名称、金额排序
        summary_data = summary_data.sort_values(['大类', '小类', '标准化物品名称', '总金额'], ascending=[True, True, True, False])
        
        print(f"大小类按单位汇总: {len(summary_data)} 条汇总记录")
        print(f"  总金额: {summary_data['总金额'].sum()/1e8:.2f} 亿泰铢")
        
        # 写入Excel
        with pd.ExcelWriter(output_file, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
            summary_data.to_excel(writer, sheet_name='大小类按单位汇总', index=False)
        
        # 格式化新工作表
        wb = openpyxl.load_workbook(output_file)
        ws = wb['大小类按单位汇总']
        
        # 设置表头样式
        header_fill = PatternFill(start_color='4472C4', end_color='4472C4', fill_type='solid')
        header_font = Font(bold=True, color='FFFFFF', size=11)
        
        ws.row_dimensions[1].height = 40
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        
        # 设置列宽
        ws.column_dimensions['A'].width = 15  # 大类
        ws.column_dimensions['B'].width = 15  # 小类
        ws.column_dimensions['C'].width = 20  # 标准化物品名称
        ws.column_dimensions['D'].width = 18  # 总金额
        ws.column_dimensions['E'].width = 18  # 总数量
        ws.column_dimensions['F'].width = 12  # 单位
        ws.column_dimensions['G'].width = 12  # 记录条数
        
        # 数据行样式
        number_format = '#,##0.00'
        for row_idx, row in enumerate(ws.iter_rows(min_row=2), 2):
            for cell_idx, cell in enumerate(row, 1):
                # 数字列格式化
                if cell_idx == 4:  # 总金额
                    cell.number_format = number_format
                elif cell_idx == 5:  # 总数量
                    cell.number_format = number_format
                
                # 自动换行
                cell.alignment = Alignment(wrap_text=True, vertical='top')
        
        # 添加自动筛选
        if ws.max_row > 1:
            ws.auto_filter.ref = ws.dimensions
            ws.freeze_panes = 'A2'
        
        # 添加总计行
        last_row = ws.max_row
        ws.cell(last_row + 1, 1).value = '总计'
        ws.cell(last_row + 1, 1).font = Font(bold=True)
        ws.cell(last_row + 1, 4).value = f'=SUM(D2:D{last_row})'
        ws.cell(last_row + 1, 5).value = f'=SUM(E2:E{last_row})'
        ws.cell(last_row + 1, 7).value = f'=SUM(G2:G{last_row})'
        
        title_fill = PatternFill(start_color='D9E1F2', end_color='D9E1F2', fill_type='solid')
        title_font = Font(bold=True, size=12)
        for col in range(1, 8):
            cell = ws.cell(last_row + 1, col)
            cell.fill = title_fill
            cell.font = title_font
            if col in [4, 5]:
                cell.number_format = number_format
        
        wb.save(output_file)
        print(f"✓ 大小类按单位汇总工作表已添加")


def main():
    """主函数"""
    # 获取脚本所在目录的绝对路径
    script_dir = os.path.dirname(os.path.abspath(__file__))
    input_file = os.path.join(script_dir, 'merged_excel_日期已标准化.xlsx')
    
    # 创建分类器实例
    classifier = PreciseMaterialClassifier(input_file)
    
    # 运行完整流程
    output_file = classifier.run_full_process()
    
    print("\n" + "="*80)
    print("使用说明")
    print("="*80)
    print("""
    生成的报告包含以下工作表：
    
    【2025年材料采购汇总报告】
    1. 材料归类汇总 - 按材料类别的总体统计（重点查看）
    2. 材料明细汇总 - 每个材料类别下的具体物品明细
    3. 月度采购趋势 - 2025年各月的材料采购情况
    4. Top100物品明细 - 金额最高的100个物品
    5. 项目材料分布 - 各项目的材料采购分布
    6. 2025年原始数据 - 带精准归类的完整数据
    
    【全数据_精准归类.xlsx】
    - 各年度材料归类汇总 - 2023-2026年的对比
    - 完整数据 - 所有年度的数据（含精准归类）
    
    材料归类体系（共18类）：
    - 钢筋类、桩基类、混凝土及砂浆、砖砌体类
    - 骨料类、水泥制品、电气材料、钢材类
    - 模板及支撑、门窗类、防水保温、管材管件
    - 五金材料、井盖及附属、机械设备、装饰材料
    - 土方及回填、临建设施、其他材料
    """)


if __name__ == '__main__':
    main()
