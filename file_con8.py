# -*- coding: utf-8 -*-
import pandas as pd
import os
import glob
import re
import numpy as np
import sys
from datetime import datetime
from collections import defaultdict, Counter
import random
from typing import List, Tuple, Dict, Set, Optional
import warnings
warnings.filterwarnings('ignore')

# 导入用于写入 Excel 格式的库
try:
    import xlsxwriter
except ImportError:
    print("错误：需要 'xlsxwriter' 库来设置单元格颜色。")
    print("请使用 'pip install xlsxwriter' 命令安装后再试。")
    sys.exit("缺少必要的库: xlsxwriter")

# 导入用于聚类的库
try:
    from Levenshtein import distance as levenshtein_distance
except ImportError:
    print("警告：未安装 python-Levenshtein 库，将使用纯Python实现（速度较慢）")
    print("建议安装：pip install python-Levenshtein")
    def levenshtein_distance(s1, s2):
        """纯Python实现的编辑距离（备用）"""
        if len(s1) < len(s2):
            return levenshtein_distance(s2, s1)
        if len(s2) == 0:
            return len(s1)
        previous_row = range(len(s2) + 1)
        for i, c1 in enumerate(s1):
            current_row = [i + 1]
            for j, c2 in enumerate(s2):
                insertions = previous_row[j + 1] + 1
                deletions = current_row[j] + 1
                substitutions = previous_row[j] + (c1 != c2)
                current_row.append(min(insertions, deletions, substitutions))
            previous_row = current_row
        return previous_row[-1]


# ============================================================
# ===== 人工干预配置区域 =====
# ============================================================

# 1. 需要排除的AI回复示例（这些示例及其同类将被强制删除）
# 格式: {"使用场景": ["AI回复示例1", "AI回复示例2", ...]}
EXCLUDE_AI_REPLY_EXAMPLES = {
    # "单呼": ["请勿呼叫自己", "测试消息"],
    # "群呼": ["测试消息"],
}

# 2. 需要保留的AI回复示例（这些示例及其同类将被强制保留）
# 格式: {"使用场景": ["AI回复示例1", "AI回复示例2", ...]}
KEEP_AI_REPLY_EXAMPLES = {
    # "单呼": ["已接通张三，请按键说"],
    # "群呼": ["全体人员请注意"],
}

# 3. 完全跳过模板分析的场景（该场景所有数据全部保留）
# 格式: ["使用场景1", "使用场景2", ...]
SKIP_SCENE_ANALYSIS = [
    # "紧急呼叫",
    # "测试场景"
]

# 4. 需要强制保留的数据（不进入抽取流程，直接保留）
# 格式: {"使用场景": ["AI回复关键词1", "AI回复关键词2", ...]}
# 注意：AI回复使用包含匹配（contains），只要AI回复包含指定字符串即可
FORCE_KEEP_DATA = {
    # "单呼": ["紧急", "重要"],
    # "群呼": ["全体人员"],
}

# ===== 结束配置区域 ==========================================


# --- 帮助函数 ---
def parse_date_range(range_str):
    """解析用户输入的日期范围字符串"""
    parts = range_str.split('-')
    if len(parts) != 2:
        raise ValueError("日期范围格式错误，应用'-'分隔开始和结束日期。")
    
    start_str = parts[0].strip()
    end_str = parts[1].strip()
    
    try:
        start_date = pd.to_datetime(start_str, errors='raise')
        end_date = pd.to_datetime(end_str, errors='raise')
        end_date = end_date.replace(hour=23, minute=59, second=59, microsecond=999999)
        if start_date > end_date:
            raise ValueError("开始日期不能晚于结束日期。")
        return start_date, end_date
    except ValueError as e:
        raise ValueError(f"无法解析日期: '{range_str}'. 请使用 YYYY/MM/DD 或类似格式。错误详情: {e}")


def parse_and_validate_custom_filename_core(core_str):
    """解析并校验用户输入的自定义文件名核心部分"""
    valid_prefixes = ["WT", "TEL"]
    valid_middles = ["ALL", "VAC", "FILL"]
    
    if not core_str.endswith("_"):
        raise ValueError("自定义文件名核心部分必须以下划线 '_' 结尾。")
    
    parts = core_str.rstrip("_").split('_')
    if len(parts) != 3:
        raise ValueError(f"自定义文件名核心部分格式错误，应为 'PREFIX_MIDDLE_DATEPART' (例如 'WT_ALL_2025_' 或 'TEL_VAC_2023_')，当前部分数量: {len(parts)}。")
    
    prefix, middle, date_part_str = parts[0].upper(), parts[1].upper(), parts[2]
    
    if prefix not in valid_prefixes:
        raise ValueError(f"前缀部分 '{parts[0]}' 无效。有效选项: {', '.join(valid_prefixes)}。")
    if middle not in valid_middles:
        raise ValueError(f"中间部分 '{parts[1]}' 无效。有效选项: {', '.join(valid_middles)}。")
    
    cleaned_date_part_str = re.sub(r'[^0-9-]', '', date_part_str)
    if '-' in cleaned_date_part_str:
        date_range_parts = cleaned_date_part_str.split('-')
        if len(date_range_parts) != 2:
            raise ValueError(f"日期范围 '{date_part_str}' 格式错误，应为 YYYYMMDD-YYYYMMDD。")
        for dp_part in date_range_parts:
            if not (len(dp_part) == 8 and dp_part.isdigit()):
                raise ValueError(f"日期范围中的日期部分 '{dp_part}' 格式错误，应为 YYYYMMDD。")
            try:
                datetime.strptime(dp_part, "%Y%m%d")
            except ValueError:
                raise ValueError(f"日期范围中的日期 '{dp_part}' 无效。")
        start_dt = datetime.strptime(date_range_parts[0], "%Y%m%d")
        end_dt = datetime.strptime(date_range_parts[1], "%Y%m%d")
        if start_dt > end_dt:
            raise ValueError(f"日期范围中开始日期 '{date_range_parts[0]}' 不能晚于结束日期 '{date_range_parts[1]}'")
    elif len(cleaned_date_part_str) == 4 and cleaned_date_part_str.isdigit():
        try:
            datetime.strptime(cleaned_date_part_str, "%Y")
        except ValueError:
            raise ValueError(f"年份 '{date_part_str}' 格式错误，应为 YYYY。")
    elif len(cleaned_date_part_str) == 6 and cleaned_date_part_str.isdigit():
        try:
            datetime.strptime(cleaned_date_part_str, "%Y%m")
        except ValueError:
            raise ValueError(f"年月 '{date_part_str}' 格式错误，应为 YYYYMM。")
    elif len(cleaned_date_part_str) == 8 and cleaned_date_part_str.isdigit():
        try:
            datetime.strptime(cleaned_date_part_str, "%Y%m%d")
        except ValueError:
            raise ValueError(f"日期 '{date_part_str}' 格式错误，应为 YYYYMMDD。")
    else:
        raise ValueError(f"日期部分 '{date_part_str}' 格式无法识别。支持格式: YYYY, YYYYMM, YYYYMMDD, 或 YYYYMMDD-YYYYMMDD。")
    return core_str


# ============================================================
# ===== 聚类与模板提取模块 =====
# ============================================================

def fast_cluster_texts(texts: List[str], max_samples: int = 5000) -> Dict[str, List[str]]:
    """
    对文本列表进行快速聚类
    返回: {模板代表文本: [同类文本列表]}
    """
    if not texts:
        return {}
    
    if len(texts) > max_samples:
        sampled_texts = random.sample(texts, max_samples)
    else:
        sampled_texts = texts
    
    unique_texts = list(set(sampled_texts))
    if len(unique_texts) <= 1:
        return {unique_texts[0]: unique_texts if unique_texts else []}
    
    clusters = []
    distances = []
    sample_for_distance = unique_texts[:min(200, len(unique_texts))]
    for i in range(len(sample_for_distance)):
        for j in range(i + 1, len(sample_for_distance)):
            dist = levenshtein_distance(sample_for_distance[i], sample_for_distance[j])
            max_len = max(len(sample_for_distance[i]), len(sample_for_distance[j]))
            if max_len > 0:
                distances.append(dist / max_len)
    
    if distances:
        threshold_percentile = 25
        threshold = np.percentile(distances, threshold_percentile)
        threshold = max(0.15, min(0.5, threshold))
    else:
        threshold = 0.3
    
    for text in unique_texts:
        assigned = False
        for cluster_rep, cluster_members in clusters:
            rep_len = max(len(cluster_rep), 1)
            dist = levenshtein_distance(text, cluster_rep) / rep_len
            if dist <= threshold:
                cluster_members.append(text)
                assigned = True
                break
        if not assigned:
            clusters.append([text, [text]])
    
    result = {}
    for cluster_rep, members in clusters:
        result[cluster_rep] = members
    
    return result


def extract_core_templates(
    scene_texts: List[str],
    exclude_examples: List[str] = None,
    keep_examples: List[str] = None,
    skip_analysis: bool = False
) -> Tuple[Set[str], Set[str], Set[str]]:
    """
    提取核心模板
    返回: (核心模板集合, 需删除的文本集合, 强制保留的文本集合)
    """
    exclude_examples = exclude_examples or []
    keep_examples = keep_examples or []
    
    if skip_analysis or not scene_texts:
        return set(scene_texts), set(), set()
    
    clusters = fast_cluster_texts(scene_texts)
    
    template_freq = {}
    for rep, members in clusters.items():
        full_count = sum(1 for t in scene_texts if t in members or levenshtein_distance(t, rep) / max(len(t), len(rep), 1) <= 0.3)
        template_freq[rep] = full_count
    
    if template_freq:
        avg_freq = sum(template_freq.values()) / len(template_freq)
    else:
        avg_freq = 0
    
    core_templates = {rep for rep, freq in template_freq.items() if freq >= avg_freq}
    non_core_templates = set(template_freq.keys()) - core_templates
    force_keep_texts = set()
    
    for example in keep_examples:
        if example in clusters:
            core_templates.add(example)
            non_core_templates.discard(example)
        else:
            min_dist = float('inf')
            closest_rep = None
            for rep in clusters:
                dist = levenshtein_distance(example, rep) / max(len(example), len(rep), 1)
                if dist < min_dist:
                    min_dist = dist
                    closest_rep = rep
            if closest_rep and min_dist <= 0.4:
                core_templates.add(closest_rep)
                non_core_templates.discard(closest_rep)
                force_keep_texts.update(clusters.get(closest_rep, []))
    
    for example in exclude_examples:
        if example in clusters:
            non_core_templates.add(example)
            core_templates.discard(example)
        else:
            min_dist = float('inf')
            closest_rep = None
            for rep in clusters:
                dist = levenshtein_distance(example, rep) / max(len(example), len(rep), 1)
                if dist < min_dist:
                    min_dist = dist
                    closest_rep = rep
            if closest_rep and min_dist <= 0.4:
                non_core_templates.add(closest_rep)
                core_templates.discard(closest_rep)
    
    texts_to_delete = set()
    for rep in non_core_templates:
        if rep in clusters:
            texts_to_delete.update(clusters[rep])
    
    return core_templates, texts_to_delete, force_keep_texts


def balanced_sampling(
    non_empty_df: pd.DataFrame,
    empty_count: int,
    id_cols: List[str],
    scene_col: str,
    median_per_id: int,
    force_keep_indices: Set = None,
    original_content_col: str = '语音内容',      # ⭐ 新增参数
    corrected_content_col: str = '修正后的语音内容'  # ⭐ 新增参数
) -> pd.DataFrame:
    """
    分层随机抽取，使非空场景数据量与空场景数据量接近
    按轮次抽取：每轮每个ID在每个场景中抽取1条
    每次抽取前随机打乱ID顺序和场景顺序，保证随机性
    同ID同场景下，已抽取过的(语音内容, 修正后的语音内容)组合不再重复抽取
    """
    if non_empty_df.empty or empty_count == 0:
        return pd.DataFrame()
    
    force_keep_indices = force_keep_indices or set()
    available_df = non_empty_df[~non_empty_df.index.isin(force_keep_indices)].copy()
    if available_df.empty:
        return pd.DataFrame()
    
    grouped = available_df.groupby(id_cols)
    id_groups = {id_val: group for id_val, group in grouped}
    
    selected_indices = []
    selected_set = set()
    total_selected = 0
    id_selected_count = defaultdict(int)
    
    # 记录已抽取的组合
    used_combinations = set()
    
    id_list = list(id_groups.keys())
    
    round_num = 0
    while total_selected < empty_count:
        round_num += 1
        round_selected = False
        
        random.shuffle(id_list)
        
        for id_val in id_list:
            if id_selected_count[id_val] >= median_per_id:
                continue
            
            group = id_groups[id_val]
            
            available_scenes = group[scene_col].unique()
            scene_list = list(available_scenes)
            random.shuffle(scene_list)
            
            for scene in scene_list:
                id_scene_data = group[group[scene_col] == scene]
                available = id_scene_data[~id_scene_data.index.isin(selected_set)]
                if available.empty:
                    continue
                
                # ⭐ 过滤已抽取过的组合
                candidates = []
                for idx, row in available.iterrows():
                    original = str(row.get(original_content_col, '')).strip()
                    corrected = str(row.get(corrected_content_col, '')).strip()
                    combo = (id_val, scene, original, corrected)
                    if combo not in used_combinations:
                        candidates.append((idx, combo))
                
                if not candidates:
                    continue
                
                chosen_idx, chosen_combo = random.choice(candidates)
                selected_indices.append(chosen_idx)
                selected_set.add(chosen_idx)
                used_combinations.add(chosen_combo)
                total_selected += 1
                id_selected_count[id_val] += 1
                round_selected = True
                
                if total_selected >= empty_count:
                    break
                if id_selected_count[id_val] >= median_per_id:
                    break
            
            if total_selected >= empty_count:
                break
        
        if not round_selected:
            break
    
    if selected_indices:
        return available_df.loc[selected_indices]
    else:
        return pd.DataFrame()


def generate_final_format(df: pd.DataFrame, config: Dict) -> pd.DataFrame:
    """生成最终输出的6个字段格式"""
    if df.empty:
        return pd.DataFrame()
    
    corrected_content_col = config.get('corrected_content_col', '修正后的语音内容')
    voice_address_col = config.get('voice_address_col', '语音地址')
    user_defined_filename_core_part = config.get('user_defined_filename_core_part', 'WT_ALL_2025_')
    fixed_path_string = config.get('fixed_path_string', '/workspace/FunASR/data/list/audio/')
    file_extension = config.get('file_extension', '.wav')
    
    df = df.reset_index(drop=True)
    
    # 生成排序编号（从1开始）
    df['排序'] = np.arange(1, len(df) + 1)
    df['排序'] = df['排序'].astype(str)
    df['文件名核心'] = user_defined_filename_core_part + df['排序']
    
    # 确保 final_corrected_content 存在
    if 'final_corrected_content' not in df.columns:
        df['final_corrected_content'] = ''
    df['final_corrected_content'] = df['final_corrected_content'].fillna('').astype(str)
    
    # 逐行拼接，避免索引对齐问题
    id_content_list = []
    id_filename_list = []
    for idx, row in df.iterrows():
        id_content_list.append(str(row['排序']) + ' ' + str(row['final_corrected_content']))
        id_filename_list.append(str(row['排序']) + ' ' + fixed_path_string + str(row['文件名核心']) + file_extension)
    
    df['ID+内容'] = id_content_list
    df['ID+文件名'] = id_filename_list
    
    # 重命名
    df.rename(columns={'final_corrected_content': corrected_content_col}, inplace=True)
    
    # 选择最终列
    final_columns = ['排序', '文件名核心', 'ID+内容', 'ID+文件名', corrected_content_col, voice_address_col]
    for col in final_columns:
        if col not in df.columns:
            df[col] = ''
    
    return df[final_columns].copy()


def generate_filtered_sheets(
    source_df: pd.DataFrame,
    config: Dict
) -> Tuple[Optional[pd.DataFrame], Optional[pd.DataFrame], Optional[pd.DataFrame], str]:
    """
    生成三个表：
    1. 排序去除风险数据表（模板分析后、抽取前的数据）
    2. 风险数据表（被排除的非核心模板数据，原字段输出）
    3. 排序已筛选表（最终平衡后的数据）
    """
    risk_removed_df = None
    risk_data_df = None
    filtered_df = None
    status = ""

    corrected_content_col = config.get('corrected_content_col', '修正后的语音内容')
    original_content_col = config.get('original_content_col', '语音内容')
    voice_address_col = config.get('voice_address_col', '语音地址')
    error_analysis_col = config.get('error_analysis_col', '错误分析')
    correct_symbols_list = config.get('correct_symbols_list', ['✓', '√', '✔'])
    dirty_symbols_list = config.get('dirty_symbols_list', ['×', '✗', '✘', 'x', 'X', '❌'])
    error_analysis_exclusion_keywords = config.get('error_analysis_exclusion_keywords', [])
    
    scene_col = '使用场景'
    device_col = '设备序列号'
    user_col = '使用人'
    ai_reply_col = 'AI回复'
    
    df = source_df.copy()
    
    required_cols = [corrected_content_col, original_content_col, voice_address_col, 
                     scene_col, device_col, user_col, ai_reply_col]
    missing_cols = [col for col in required_cols if col not in df.columns]
    if missing_cols:
        return None, None, None, f"缺少必要列: {missing_cols}"
    
    df = df[df[voice_address_col].notna() & (df[voice_address_col].astype(str).str.strip() != '')].copy()
    
    df[corrected_content_col] = df[corrected_content_col].fillna('').astype(str).str.strip()
    df[original_content_col] = df[original_content_col].fillna('').astype(str).str.strip()
    df[voice_address_col] = df[voice_address_col].fillna('').astype(str).str.strip()
    df[scene_col] = df[scene_col].fillna('').astype(str).str.strip()
    df[device_col] = df[device_col].fillna('').astype(str).str.strip()
    df[user_col] = df[user_col].fillna('').astype(str).str.strip()
    df[ai_reply_col] = df[ai_reply_col].fillna('').astype(str).str.strip()
    
    is_dirty = df[corrected_content_col].isin(dirty_symbols_list)
    df = df[~is_dirty].copy()
    
    if error_analysis_col in df.columns and error_analysis_exclusion_keywords:
        df[error_analysis_col] = df[error_analysis_col].fillna('').astype(str).str.strip()
        condition_is_correct = df[corrected_content_col].isin(correct_symbols_list)
        escaped_keywords = [re.escape(keyword) for keyword in error_analysis_exclusion_keywords]
        pattern_for_exclusion = '|'.join(escaped_keywords)
        condition_has_exclusion_keyword = df[error_analysis_col].astype(str).str.contains(
            pattern_for_exclusion, case=False, regex=True, na=False
        )
        condition_to_exclude = condition_is_correct & condition_has_exclusion_keyword
        df = df[~condition_to_exclude].copy()
    
    condition_is_correct = df[corrected_content_col].isin(correct_symbols_list)
    df['final_corrected_content'] = np.where(
        condition_is_correct,
        df[original_content_col],
        df[corrected_content_col]
    )
    df['final_corrected_content'] = df['final_corrected_content'].fillna('').astype(str).str.strip()
    
    if df.empty:
        return None, None, None, "基础处理后无有效数据"
    
    df_valid_voice = df[df[voice_address_col].notna() & (df[voice_address_col].astype(str).str.strip() != '')].copy()
    df_valid_voice['scene_is_empty'] = df_valid_voice[scene_col].isna() | (df_valid_voice[scene_col].astype(str).str.strip() == '') | (df_valid_voice[scene_col].astype(str).str.strip() == 'nan')
    
    # ============================================================
    # ===== 修改点：三分类逻辑（空场景 / 伪空场景 / 真正非空场景） =====
    # ============================================================
    
    # 1. 分出原空场景
    empty_scene_df = df_valid_voice[df_valid_voice['scene_is_empty'] == True].copy()
    non_empty_scene_df = df_valid_voice[df_valid_voice['scene_is_empty'] == False].copy()
    
    # 2. 定义伪空场景判断函数：修正后的语音内容包含至少一个中文字符、英文字母或数字
    def is_pseudo_empty(content):
        """判断是否为伪空场景：包含至少一个中文字符、英文字母或数字"""
        if not content or not isinstance(content, str):
            return False
        return bool(re.search(r'[\u4e00-\u9fffa-zA-Z0-9]', content))
    
    # 3. 从非空场景中分离出"伪空场景"和"真正的非空场景"
    pseudo_empty_mask = non_empty_scene_df[corrected_content_col].astype(str).apply(is_pseudo_empty)
    pseudo_empty_df = non_empty_scene_df[pseudo_empty_mask].copy()
    true_non_empty_df = non_empty_scene_df[~pseudo_empty_mask].copy()
    
    # 4. 重新定义空场景总数 = 原空场景 + 伪空场景
    empty_count = len(empty_scene_df) + len(pseudo_empty_df)
    non_empty_count = len(true_non_empty_df)
    
    # 5. 直接保留的数据（原空场景 + 伪空场景），不参与任何后续处理
    direct_keep_df = pd.concat([empty_scene_df, pseudo_empty_df], ignore_index=True)
    
    print(f"  数据分类: 原空场景 {len(empty_scene_df)} 条, 伪空场景 {len(pseudo_empty_df)} 条, 真正的非空场景 {len(true_non_empty_df)} 条")
    print(f"  空场景总数(含伪空): {empty_count}, 真正非空: {non_empty_count}")
    
    # ============================================================
    # ===== 后续流程只针对 true_non_empty_df =====
    # ============================================================
    
    if empty_count >= non_empty_count:
        # 直接保留的数据已足够多，无需平衡，直接合并输出
        combined = pd.concat([direct_keep_df, true_non_empty_df], ignore_index=True)
        filtered_df = generate_final_format(combined, config)
        # 排序去除风险数据表：包含所有数据（无数据被删除）
        risk_removed_df = generate_final_format(combined, config) if not combined.empty else None
        status = f"空场景总数({empty_count}) >= 真正非空数据({non_empty_count})，无需平衡，总计 {len(filtered_df)} 条"
        print(f"  {status}")
        return risk_removed_df, None, filtered_df, status
    
    # 强制保留数据（仅针对真正的非空场景）
    force_keep_indices = set()
    force_keep_df = pd.DataFrame()
    
    for scene, keywords in FORCE_KEEP_DATA.items():
        if not keywords:
            continue
        for keyword in keywords:
            mask = (true_non_empty_df[scene_col] == scene) & (true_non_empty_df[ai_reply_col].str.contains(keyword, case=False, na=False))
            matched = true_non_empty_df[mask]
            if not matched.empty:
                force_keep_indices.update(matched.index.tolist())
                if force_keep_df.empty:
                    force_keep_df = matched.copy()
                else:
                    force_keep_df = pd.concat([force_keep_df, matched], ignore_index=True)
                print(f"    强制保留: 场景'{scene}' 关键词'{keyword}' 匹配 {len(matched)} 条")
    
    # 用于抽样的数据 = 真正的非空场景 - 强制保留的数据
    non_empty_for_sampling = true_non_empty_df[~true_non_empty_df.index.isin(force_keep_indices)].copy()
    
    # 场景模板分析（仅针对 true_non_empty_df）
    scene_groups = non_empty_for_sampling.groupby(scene_col)
    texts_to_delete = set()
    risk_data_indices = set()
    risk_data_full = []
    
    print(f"  开始场景模板分析，共 {len(scene_groups)} 个场景（仅针对真正的非空场景）...")
    
    for scene, group in scene_groups:
        print(f"    处理场景: {scene}, 数据量: {len(group)}")
        
        skip = scene in SKIP_SCENE_ANALYSIS
        exclude_examples = EXCLUDE_AI_REPLY_EXAMPLES.get(scene, [])
        keep_examples = KEEP_AI_REPLY_EXAMPLES.get(scene, [])
        
        ai_texts = group[ai_reply_col].dropna().astype(str).str.strip().tolist()
        ai_texts = [t for t in ai_texts if t]
        
        if not ai_texts:
            continue
        
        core_templates, delete_texts, force_keep_texts = extract_core_templates(
            ai_texts,
            exclude_examples=exclude_examples,
            keep_examples=keep_examples,
            skip_analysis=skip
        )
        
        if delete_texts and not skip:
            for text in delete_texts:
                matching_rows = group[group[ai_reply_col].astype(str).str.strip() == text].index.tolist()
                texts_to_delete.update(matching_rows)
                risk_data_indices.update(matching_rows)
        
        if risk_data_indices:
            risk_data_rows = df.loc[list(risk_data_indices)].copy()
            if not risk_data_rows.empty:
                risk_data_full.append(risk_data_rows)
        
        print(f"      核心模板数: {len(core_templates)}, 将删除 {len(delete_texts)} 条")
    
    if texts_to_delete:
        non_empty_after_clean = non_empty_for_sampling[~non_empty_for_sampling.index.isin(texts_to_delete)].copy()
        deleted_by_template = len(texts_to_delete)
    else:
        non_empty_after_clean = non_empty_for_sampling.copy()
        deleted_by_template = 0
    
    if not force_keep_df.empty:
        non_empty_after_clean = pd.concat([non_empty_after_clean, force_keep_df], ignore_index=True)
    
    # 生成"排序去除风险数据表" = direct_keep_df + 清理后的 true_non_empty_df
    # risk_removed_combined = pd.concat([direct_keep_df, non_empty_after_clean], ignore_index=True)
    # risk_removed_df = None
    # if not risk_removed_combined.empty:
    #     risk_removed_df = generate_final_format(risk_removed_combined, config)
    
    # 生成"风险数据表"
    risk_data_df = None
    if risk_data_full:
        risk_data_df = pd.concat(risk_data_full, ignore_index=True)
        risk_data_df = risk_data_df.drop_duplicates()
    
    non_empty_count_after = len(non_empty_after_clean)
    
    print(f"  模板清理完成: 删除 {deleted_by_template} 条，剩余真正非空数据: {non_empty_count_after}")
    
    if non_empty_count_after <= empty_count:
        # 真正非空数据已少于或等于空场景总数，合并所有数据生成最终表
        combined = pd.concat([direct_keep_df, non_empty_after_clean], ignore_index=True)
        filtered_df = generate_final_format(combined, config)
        status = f"模板清理后真正非空数据({non_empty_count_after}) <= 空场景总数({empty_count})，合并后总计 {len(filtered_df)} 条"
        print(f"  {status}")
        return risk_removed_df, risk_data_df, filtered_df, status
    else:
        print(f"  真正非空数据({non_empty_count_after})仍多于空场景总数({empty_count})，开始分层随机抽取...")
        
        non_empty_after_clean['id_key'] = non_empty_after_clean[device_col].astype(str) + '|' + non_empty_after_clean[user_col].astype(str)
        id_stats_after = non_empty_after_clean.groupby('id_key').size()
        if not id_stats_after.empty:
            median_count_after = int(id_stats_after.max())
        else:
            median_count_after = 1
        
        sampled_df = balanced_sampling(
            non_empty_after_clean,
            empty_count,
            [device_col, user_col],
            scene_col,
            median_count_after,
            force_keep_indices,
            original_content_col,   # ⭐ 新增
            corrected_content_col   # ⭐ 新增
        )
        
        if not sampled_df.empty:
            if not force_keep_df.empty:
                final_non_empty = pd.concat([force_keep_df, sampled_df], ignore_index=True)
            else:
                final_non_empty = sampled_df
        else:
            # 如果无法抽取，使用所有清理后的数据
            final_non_empty = non_empty_after_clean
        
        # 合并直接保留的数据 + 抽样后的真正非空数据
        combined = pd.concat([direct_keep_df, final_non_empty], ignore_index=True)
        filtered_df = generate_final_format(combined, config)
        
        if not sampled_df.empty:
            status = f"分层抽取完成，抽取 {len(sampled_df)} 条真正非空数据，空场景总数 {empty_count} 条，总计 {len(filtered_df)} 条"
        else:
            status = f"无法继续逐轮抽取，使用所有真正非空数据，空场景总数 {empty_count} 条，真正非空数据 {len(final_non_empty)} 条，总计 {len(filtered_df)} 条"
        print(f"  {status}")
        
        return risk_removed_df, risk_data_df, filtered_df, status

# ============================================================
# ===== 主脚本 =====
# ============================================================

while True:
    restart_script = False
    
    # --- 配置 ---
    date_filter_col = '开始时间'
    original_content_col = '语音内容'
    corrected_content_col = '修正后的语音内容'
    voice_address_col = '语音地址'
    error_analysis_col = '错误分析'
    
    correct_symbols_list = ['✓', '√', '✔']
    dirty_symbols_list = ['×', '✗', '✘', 'x', 'X', '❌']
    
    error_analysis_exclusion_keywords = [
        "对讲使用不正确",
        "口误",
        "ptt",
        "语义混淆",
        "吐字不清",
        "非对讲功能场景",
        "声音过轻"
    ]
    
    merged_sheet_name = 'Merged_Data'
    processed_sheet_name = '排序'
    output_sort_col_name = '排序'
    output_filename_core_col_name = '文件名核心'
    output_corrected_col_name = corrected_content_col
    output_id_content_col_name = 'ID+内容'
    output_id_filename_col_name = 'ID+文件名'
    output_voice_address_col_name = voice_address_col
    fixed_path_string = "/workspace/FunASR/data/list/audio/"
    file_extension = ".wav"
    start_sort_number = 1
    null_fill_value = "NULL"
    fill_color_hex = '#FFC7CE'
    user_defined_filename_core_part = ""
    
    print("Excel 文件按日期合并与空值填充脚本 (输入 'P' 可随时重新开始)")
    print("-" * 30)
    
    input_folder = None
    while True:
        folder_input_raw = input(f"请输入包含 Excel 文件的文件夹路径 (输入 'P' 重新开始): ").strip()
        if folder_input_raw.upper() == 'P':
            restart_script = True
            break
        if os.path.isdir(folder_input_raw):
            input_folder = folder_input_raw
            break
        else:
            print("错误：输入的路径不存在或不是一个有效的文件夹。")
    if restart_script:
        print("\n收到重新开始指令，脚本将重新启动...\n" + "="*30 + "\n")
        continue
    
    output_file = None
    while True:
        output_file_raw = input(f"请输入合并后的输出 Excel 文件完整路径 (例如 D:\\Data\\merged.xlsx) (输入 'P' 重新开始): ").strip()
        if output_file_raw.upper() == 'P':
            restart_script = True
            break
        output_dir = os.path.dirname(output_file_raw)
        temp_output_file = output_file_raw
        if not output_dir:
            output_dir = '.'
            temp_output_file = os.path.join(output_dir, output_file_raw)
            print(f"将在当前目录下创建文件: {os.path.abspath(temp_output_file)}")
        elif not os.path.isdir(output_dir):
            print(f"错误：输出文件所在的目录 '{output_dir}' 不存在。")
            continue
        base, ext = os.path.splitext(temp_output_file)
        if ext.lower() != '.xlsx':
            temp_output_file = base + '.xlsx'
            print(f"文件扩展名已强制修正为 .xlsx (用于格式设置): {os.path.basename(temp_output_file)}")
        output_file = temp_output_file
        break
    if restart_script:
        print("\n收到重新开始指令，脚本将重新启动...\n" + "="*30 + "\n")
        continue
    
    start_date = None
    end_date = None
    while True:
        date_range_raw = input(f"请输入要合并数据的日期范围 (格式: YYYY/MM/DD-YYYY/MM/DD) (输入 'P' 重新开始): ").strip()
        if date_range_raw.upper() == 'P':
            restart_script = True
            break
        try:
            start_date, end_date = parse_date_range(date_range_raw)
            print(f"将合并 '{date_filter_col}' 在 {start_date.strftime('%Y-%m-%d')} 到 {end_date.strftime('%Y-%m-%d')} 之间的数据。")
            break
        except ValueError as e:
            print(f"输入错误: {e} 请重新输入。")
    if restart_script:
        print("\n收到重新开始指令，脚本将重新启动...\n" + "="*30 + "\n")
        continue
    
    while True:
        start_num_raw = input(f"请输入 '{output_sort_col_name}' 列的起始数字 (默认为 {start_sort_number}, 输入 'P' 重新开始): ").strip()
        if start_num_raw.upper() == 'P':
            restart_script = True
            break
        if not start_num_raw:
            print(f"使用默认起始数字: {start_sort_number}")
            break
        try:
            temp_start_num = int(start_num_raw)
            if temp_start_num >= 1:
                start_sort_number = temp_start_num
                print(f"排序将从 {start_sort_number} 开始。")
                break
            else:
                print("错误：起始数字必须大于或等于 1。")
        except ValueError:
            print("错误：请输入一个有效的整数。")
    if restart_script:
        print("\n收到重新开始指令，脚本将重新启动...\n" + "="*30 + "\n")
        continue
    
    print("\n--- '文件名核心' 自定义部分配置 ---")
    print("这部分将与排序ID组合成文件名核心。格式: PREFIX_MIDDLE_DATEPART_")
    print("  PREFIX 可选: WT, TEL")
    print("  MIDDLE 可选: ALL, VAC, FILL")
    print("  DATEPART 可选: YYYY, YYYYMM, YYYYMMDD, YYYYMMDD-YYYYMMDD")
    print("  示例: WT_ALL_2025_ 或 TEL_VAC_202403_ 或 WT_FILL_20230101-20230115_")
    while True:
        filename_core_part_raw = input(f"请输入自定义文件名核心部分 (输入 'P' 重新开始): ").strip()
        if filename_core_part_raw.upper() == 'P':
            restart_script = True
            break
        if not filename_core_part_raw:
            print("错误：自定义文件名核心部分不能为空。")
            continue
        try:
            user_defined_filename_core_part = parse_and_validate_custom_filename_core(filename_core_part_raw)
            print(f"自定义文件名核心部分校验通过: '{user_defined_filename_core_part}'")
            break
        except ValueError as e:
            print(f"输入错误: {e} 请重新输入。")
    if restart_script:
        print("\n收到重新开始指令，脚本将重新启动...\n" + "="*30 + "\n")
        continue
    
    print("-" * 30)
    all_dataframes = []
    error_files = []
    print(f"开始处理文件夹: {input_folder}")
    print(f"合并结果将保存至: {output_file}")
    print(f"筛选条件: '{date_filter_col}' 在 {start_date.strftime('%Y-%m-%d')} 到 {end_date.strftime('%Y-%m-%d')} 之间")
    print(f"空日期处理: 若邻近日期相同则填充，否则填 '{null_fill_value}'，并标红单元格")
    print(f"用于'文件名核心'的自定义部分: '{user_defined_filename_core_part}'")
    if error_analysis_exclusion_keywords:
        print(f"对于“排序”工作表，若“{corrected_content_col}”为打勾行，且“{error_analysis_col}”列中包含以下任一关键词，则该行将被排除：\n  {', '.join(error_analysis_exclusion_keywords)}")
    else:
        print(f"信息：未配置“{error_analysis_col}”的排除关键词，不执行此特定筛选。")
    
    excel_files_raw = glob.glob(os.path.join(input_folder, '*.xlsx')) + glob.glob(os.path.join(input_folder, '*.xls'))
    if not excel_files_raw:
        print("错误：未找到任何 Excel 文件 (.xlsx, .xls)。")
        print("\n请检查输入文件夹路径后重试。\n" + "="*30 + "\n")
        continue
    print(f"找到 {len(excel_files_raw)} 个 Excel 文件，开始筛选和合并...")
    excel_files = []
    abs_output_file_path = os.path.abspath(output_file)
    for file_path in excel_files_raw:
        file_basename = os.path.basename(file_path)
        if file_basename.startswith('~$'):
            print(f"  跳过 Excel 临时文件: {file_basename}")
            continue
        try:
            abs_file_path = os.path.abspath(file_path)
            if abs_file_path == abs_output_file_path:
                print(f"  跳过输出文件自身: {file_basename}")
                continue
        except Exception as e_abs:
            print(f"  警告：获取文件绝对路径时出错 '{file_basename}': {e_abs}. 尝试继续...")
        excel_files.append(file_path)
    if not excel_files:
        print("错误：筛选后未找到有效的 Excel 文件进行处理（可能只有临时文件或输出文件）。")
        print("\n请检查输入文件夹内容后重试。\n" + "="*30 + "\n")
        continue
    print(f"筛选后剩余 {len(excel_files)} 个有效 Excel 文件进行处理。")
    processed_files_count = 0
    sheets_processed_count = 0
    sheets_skipped_no_date_col = 0
    rows_outside_range = 0
    rows_initially_invalid_date = 0
    for file_path in excel_files:
        file_basename = os.path.basename(file_path)
        file_processed_flag = False
        print(f"  处理文件: {file_basename}")
        try:
            engine = 'openpyxl' if file_path.lower().endswith('.xlsx') else None
            excel_file_obj = pd.ExcelFile(file_path, engine=engine)
            sheet_names_in_file = excel_file_obj.sheet_names
            print(f"    文件包含工作表: {', '.join(sheet_names_in_file)}")
            if not sheet_names_in_file:
                print(f"    警告：文件 '{file_basename}' 不包含任何工作表，跳过。")
                continue
            for sheet_name in sheet_names_in_file:
                print(f"      处理工作表: '{sheet_name}'")
                try:
                    df_sheet = pd.read_excel(excel_file_obj, sheet_name=sheet_name, engine=engine)
                    if df_sheet.empty:
                        print(f"        工作表 '{sheet_name}' 为空，跳过。")
                        continue
                    if date_filter_col not in df_sheet.columns:
                        print(f"        警告：工作表 '{sheet_name}' 缺少必需的日期列 '{date_filter_col}'，跳过此表。")
                        sheets_skipped_no_date_col += 1
                        continue
                    original_dates_series = df_sheet[date_filter_col]
                    df_sheet[date_filter_col] = pd.to_datetime(original_dates_series, errors='coerce')
                    invalid_date_count_in_sheet = df_sheet[date_filter_col].isna().sum()
                    if invalid_date_count_in_sheet > 0:
                        print(f"        工作表 '{sheet_name}' 的 '{date_filter_col}' 列中初始发现 {invalid_date_count_in_sheet} 个无效或空日期。")
                        rows_initially_invalid_date += invalid_date_count_in_sheet
                    date_condition_in_range = (df_sheet[date_filter_col] >= start_date) & (df_sheet[date_filter_col] <= end_date)
                    is_na_date = df_sheet[date_filter_col].isna()
                    df_filtered_by_date_or_na = df_sheet[date_condition_in_range | is_na_date].copy()
                    valid_dates_mask = original_dates_series.notna() & pd.to_datetime(original_dates_series, errors='coerce').notna()
                    rows_removed_by_date_range = (~date_condition_in_range & valid_dates_mask).sum()
                    rows_outside_range += rows_removed_by_date_range
                    if df_filtered_by_date_or_na.empty:
                        print(f"        工作表 '{sheet_name}' 中没有在指定日期范围内的数据，且没有无效日期行，跳过。")
                        continue
                    print(f"        保留 {len(df_filtered_by_date_or_na)} 行 (包含有效日期在范围内或日期无效/为空的行)，排除了 {rows_removed_by_date_range} 行有效日期但在范围外的数据。")
                    
                    current_required_cols = [original_content_col, corrected_content_col, voice_address_col, error_analysis_col]
                    for col in current_required_cols:
                        if col not in df_filtered_by_date_or_na.columns:
                            if col == error_analysis_col:
                                print(f"        信息：工作表 '{sheet_name}' (筛选后) 缺少可选列 '{error_analysis_col}'。将填充空值。")
                            else:
                                print(f"        警告：工作表 '{sheet_name}' (筛选后) 缺少核心列: {col}。尝试填充空值...")
                            df_filtered_by_date_or_na[col] = np.nan
                    
                    critical_missing_cols = []
                    for core_col in [original_content_col, corrected_content_col, voice_address_col]:
                        if core_col not in df_filtered_by_date_or_na.columns or df_filtered_by_date_or_na[core_col].isnull().all():
                            critical_missing_cols.append(core_col)
                    if critical_missing_cols:
                        print(f"        严重警告：工作表 '{sheet_name}' (筛选后) 仍然缺少或全为空的关键列: {critical_missing_cols}。此工作表可能无法提供有效数据用于“排序”表。")
                    
                    df_final_filtered = df_filtered_by_date_or_na
                    if corrected_content_col in df_final_filtered.columns:
                        corrected_col_str = df_final_filtered[corrected_content_col].fillna('').astype(str).str.strip()
                        condition_is_empty = (corrected_col_str == '')
                        df_final_filtered = df_final_filtered[~condition_is_empty].copy()
                        skipped_empty_rows = len(df_filtered_by_date_or_na) - len(df_final_filtered)
                        if skipped_empty_rows > 0:
                            print(f"        进一步跳过 {skipped_empty_rows} 行，因为 '{corrected_content_col}' 为空或仅含空格。")
                    else:
                        print(f"        警告：由于缺少 '{corrected_content_col}' 列，工作表 '{sheet_name}' 的所有剩余行将被视为空并跳过。")
                        df_final_filtered = pd.DataFrame(columns=df_filtered_by_date_or_na.columns)
                    
                    if not df_final_filtered.empty:
                        print(f"        添加 {len(df_final_filtered)} 行有效数据到合并列表。")
                        all_dataframes.append(df_final_filtered)
                        sheets_processed_count += 1
                        file_processed_flag = True
                    else:
                        print(f"        工作表 '{sheet_name}' 在所有筛选后没有剩余有效数据。")
                except Exception as e_read_sheet:
                    err_type = type(e_read_sheet).__name__
                    err_msg = f"读取或处理文件 '{file_basename}' 的工作表 '{sheet_name}' 时发生 {err_type} 错误: {e_read_sheet}"
                    print(f"      错误：{err_msg}")
                    error_files.append(f"{file_basename} - {sheet_name}: {err_msg}")
                    continue
            if file_processed_flag:
                processed_files_count += 1
        except FileNotFoundError:
            err_msg = f"文件未找到: {file_path}"
            print(f"  错误：{err_msg}")
            error_files.append(f"{file_basename}: {err_msg}")
            continue
        except ValueError as ve:
            err_msg = f"处理文件 '{file_basename}' 时发生值错误: {ve}"
            print(f"  错误：{err_msg}\n  将跳过此文件。")
            error_files.append(f"{file_basename}: {err_msg}")
            continue
        except Exception as e:
            err_type = type(e).__name__
            err_msg = f"处理文件 '{file_basename}' 时发生未知的 {err_type} 错误: {e}"
            print(f"  错误：{err_msg}\n  将跳过此文件。")
            error_files.append(f"{file_basename}: {err_msg}")
            continue
    
    print("-" * 30)
    print("处理汇总:")
    print(f"  总共检查文件数: {len(excel_files)}")
    print(f"  成功处理并提取到数据的文件数: {processed_files_count}")
    print(f"  成功处理并提取到数据的工作表数: {sheets_processed_count}")
    print(f"  因缺少 '{date_filter_col}' 列而跳过的工作表数: {sheets_skipped_no_date_col}")
    print(f"  初始读取时发现的无效/空日期总行数: {rows_initially_invalid_date}")
    print(f"  因日期在范围之外而被忽略的总行数 (仅计有效日期): {rows_outside_range}")
    if error_files:
        print("\n" + "="*30)
        print("处理过程中遇到以下非致命错误或警告：")
        for err in error_files:
            print(f"- {err}")
        print("="*30 + "\n")
    
    if not all_dataframes:
        print("错误：未能从任何文件中筛选到符合日期范围或包含待处理空日期的数据。脚本终止。")
        print("\n请检查输入文件内容、日期范围和错误信息后重试。\n" + "="*30 + "\n")
        continue
    
    print(f"\n正在合并来自 {sheets_processed_count} 个工作表的数据...")
    combined_df = pd.concat(all_dataframes, ignore_index=True, sort=False)
    print(f"合并完成，总共 {len(combined_df)} 行数据 (包含待处理的空日期)。")
    
    print(f"\n开始处理 '{merged_sheet_name}' 工作表中的 '{date_filter_col}' 空值...")
    cells_to_color = []
    nat_rows_indices = combined_df.index[combined_df[date_filter_col].isna()].tolist()
    filled_count = 0
    null_filled_count = 0
    if nat_rows_indices:
        print(f"  找到 {len(nat_rows_indices)} 行 '{date_filter_col}' 为空值，尝试填充...")
        try:
            date_col_idx = combined_df.columns.get_loc(date_filter_col)
        except KeyError:
            print(f"  严重错误：合并后的 DataFrame 中找不到列 '{date_filter_col}'。")
            sys.exit(f"合并后缺少关键列 '{date_filter_col}'")
        if combined_df[date_filter_col].dtype != 'object':
            combined_df[date_filter_col] = combined_df[date_filter_col].astype(object)
        for i in nat_rows_indices:
            prev_date = pd.NaT
            next_date = pd.NaT
            can_fill = False
            if i > 0:
                prev_val = combined_df.iloc[i - 1][date_filter_col]
                if pd.notna(prev_val) and prev_val != null_fill_value:
                    prev_date = prev_val
            if i < len(combined_df) - 1:
                next_val = combined_df.iloc[i + 1][date_filter_col]
                if pd.notna(next_val) and next_val != null_fill_value:
                    next_date = next_val
            if pd.notna(prev_date) and pd.notna(next_date) and prev_date == next_date:
                combined_df.loc[i, date_filter_col] = prev_date
                cells_to_color.append((i, date_col_idx))
                filled_count += 1
                can_fill = True
            if not can_fill:
                combined_df.loc[i, date_filter_col] = null_fill_value
                cells_to_color.append((i, date_col_idx))
                null_filled_count += 1
        print(f"  填充完成：{filled_count} 行填充了邻近日期，{null_filled_count} 行填充了 '{null_fill_value}'。")
    else:
        print(f"  '{date_filter_col}' 列中没有发现空值 (NaT)，无需填充。")
    
    print(f"\n检查合并后的数据 ({merged_sheet_name}) 是否包含后续处理所需的列...")
    required_processing_cols = [original_content_col, corrected_content_col, voice_address_col]
    
    if error_analysis_col not in combined_df.columns:
        print(f"  信息：合并后的数据中缺少 '{error_analysis_col}' 列。将添加此列并填充空值以便进行后续处理。")
        combined_df[error_analysis_col] = np.nan
    
    missing_processing_cols = [col for col in required_processing_cols if col not in combined_df.columns]
    if missing_processing_cols:
        error_msg = f"错误：合并后的数据中缺少创建“{processed_sheet_name}”工作表所需的关键列: {missing_processing_cols}"
        print(error_msg)
        print(f"将只保存包含所有合并和已处理空日期数据的“{merged_sheet_name}”工作表。")
        try:
            print(f"\n正在保存 Sheet '{merged_sheet_name}' ...")
            output_dir_final = os.path.dirname(output_file)
            os.makedirs(output_dir_final, exist_ok=True)
            with pd.ExcelWriter(output_file, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
                combined_df.to_excel(writer, sheet_name=merged_sheet_name, index=False, na_rep='')
                if cells_to_color:
                    workbook = writer.book
                    worksheet = writer.sheets[merged_sheet_name]
                    red_format = workbook.add_format({'bg_color': fill_color_hex})
                    print(f"  正在对 {len(cells_to_color)} 个单元格应用红色背景...")
                    for row_idx, col_idx in cells_to_color:
                        worksheet.write(row_idx + 1, col_idx, combined_df.iloc[row_idx, col_idx], red_format)
            print(f"文件保存成功！路径: {os.path.abspath(output_file)}")
        except PermissionError:
            print(f"错误：权限不足，无法写入文件 '{output_file}'。")
        except Exception as e:
            print(f"保存文件时出错: {e}")
        print("\n处理完毕（仅生成合并表）。")
        break
    
    print("所需列存在，继续创建“排序”工作表。")
    
    # --- 6. 创建并处理 Sheet 2 ('排序') ---
    print(f"\n开始创建和处理工作表: '{processed_sheet_name}' (基于已处理空日期的合并数据)...")
    sheet2_df = combined_df.copy()
    sheet2_df[corrected_content_col] = sheet2_df[corrected_content_col].fillna('').astype(str).str.strip()
    sheet2_df[original_content_col] = sheet2_df[original_content_col].fillna('').astype(str).str.strip()
    sheet2_df[voice_address_col] = sheet2_df[voice_address_col].fillna('').astype(str).str.strip()
    
    if error_analysis_col in sheet2_df.columns:
        sheet2_df[error_analysis_col] = sheet2_df[error_analysis_col].fillna('').astype(str).str.strip()
    else:
        print(f"  严重警告：在创建 '{processed_sheet_name}' 时，本应存在的 '{error_analysis_col}' 列意外缺失。将创建为空列，但相关筛选可能无效。")
        sheet2_df[error_analysis_col] = ""
    
    rows_before_delete = len(sheet2_df)
    is_dirty = sheet2_df[corrected_content_col].astype(str).isin(dirty_symbols_list)
    sheet2_df = sheet2_df[~is_dirty].copy()
    print(f"  已删除 {rows_before_delete - len(sheet2_df)} 行 ('错误'标记: {', '.join(dirty_symbols_list)})。")
    
    rows_before_error_analysis_filter = len(sheet2_df)
    if error_analysis_col in sheet2_df.columns and error_analysis_exclusion_keywords:
        print(f"  正在根据 '{corrected_content_col}' 的正确标记和 '{error_analysis_col}' 的内容进行额外筛选 (关键词匹配)...")
        condition_is_correct_symbol = sheet2_df[corrected_content_col].astype(str).isin(correct_symbols_list)
        escaped_keywords = [re.escape(keyword) for keyword in error_analysis_exclusion_keywords]
        pattern_for_exclusion = '|'.join(escaped_keywords)
        condition_has_exclusion_keyword = sheet2_df[error_analysis_col].astype(str).str.contains(
            pattern_for_exclusion,
            case=False,
            regex=True,
            na=False
        )
        condition_to_exclude_due_to_error_analysis = condition_is_correct_symbol & condition_has_exclusion_keyword
        sheet2_df = sheet2_df[~condition_to_exclude_due_to_error_analysis].copy()
        excluded_by_error_analysis_count = rows_before_error_analysis_filter - len(sheet2_df)
        if excluded_by_error_analysis_count > 0:
            print(f"  因 '{error_analysis_col}' 包含排除性关键词 (针对打勾行)，额外删除了 {excluded_by_error_analysis_count} 行。")
        else:
            print(f"  '{error_analysis_col}' 的内容未导致额外删除打勾行。")
    elif not error_analysis_exclusion_keywords:
        print(f"  信息：未定义 '{error_analysis_col}' 的排除关键词，跳过此特定筛选。")
    
    condition_is_correct = sheet2_df[corrected_content_col].astype(str).isin(correct_symbols_list)
    sheet2_df['final_corrected_content'] = np.where(
        condition_is_correct,
        sheet2_df[original_content_col],
        sheet2_df[corrected_content_col]
    )
    print(f"  已处理 '正确' 标记，生成 'final_corrected_content'。")
    sheet2_df['final_corrected_content'] = sheet2_df['final_corrected_content'].fillna('').astype(str).str.strip()
    
    sheet2_df.reset_index(drop=True, inplace=True)
    if not sheet2_df.empty:
        end_sort_number = start_sort_number + len(sheet2_df)
        sheet2_df[output_sort_col_name] = np.arange(start_sort_number, end_sort_number)
        print(f"  已生成 '{output_sort_col_name}' 列 ({start_sort_number} 到 {end_sort_number - 1})。")
    else:
        sheet2_df[output_sort_col_name] = pd.Series(dtype='object')
        print(f"  注意：处理后没有有效数据行，'{output_sort_col_name}' 列为空。")
    
    sheet2_df[output_sort_col_name] = sheet2_df[output_sort_col_name].astype(str)
    
    if not sheet2_df.empty:
        sheet2_df[output_filename_core_col_name] = user_defined_filename_core_part + sheet2_df[output_sort_col_name]
        print(f"  已创建 '{output_filename_core_col_name}' 列 (例如: {user_defined_filename_core_part}{start_sort_number})。")
    else:
        sheet2_df[output_filename_core_col_name] = pd.Series(dtype='object')
        print(f"  注意: '{output_filename_core_col_name}' 列为空因为没有数据。")
    
    if 'final_corrected_content' not in sheet2_df.columns:
        sheet2_df['final_corrected_content'] = pd.Series(dtype='object')
    
    sheet2_df[output_id_content_col_name] = sheet2_df[output_sort_col_name] + ' ' + sheet2_df['final_corrected_content']
    print(f"  已创建 '{output_id_content_col_name}' 列。")
    
    if not sheet2_df.empty and output_filename_core_col_name in sheet2_df.columns:
        sheet2_df[output_id_filename_col_name] = sheet2_df[output_sort_col_name] + ' ' + \
                                                 fixed_path_string + \
                                                 sheet2_df[output_filename_core_col_name] + \
                                                 file_extension
        print(f"  已创建 '{output_id_filename_col_name}' 列 (例如: {start_sort_number} {fixed_path_string}{user_defined_filename_core_part}{start_sort_number}{file_extension})。")
    else:
        sheet2_df[output_id_filename_col_name] = pd.Series(dtype='object')
        print(f"  注意: '{output_id_filename_col_name}' 列为空。")
    
    final_columns_in_order = [
        output_sort_col_name,
        output_filename_core_col_name,
        output_id_content_col_name,
        output_id_filename_col_name,
        'final_corrected_content',
        voice_address_col
    ]
    for col in final_columns_in_order:
        if col not in sheet2_df.columns:
            sheet2_df[col] = pd.Series(dtype='object')
    
    final_sheet2_df = sheet2_df[final_columns_in_order].copy()
    final_sheet2_df.rename(columns={'final_corrected_content': output_corrected_col_name}, inplace=True)
    print(f"  已提取并重命名最终列: '{output_sort_col_name}', '{output_filename_core_col_name}', '{output_id_content_col_name}', '{output_id_filename_col_name}', '{output_corrected_col_name}', '{output_voice_address_col_name}'。")
    print(f"  处理后的 '{processed_sheet_name}' 工作表包含 {len(final_sheet2_df)} 行数据。")
    
    # ============================================================
    # ===== 新增：生成三个筛选相关表 =====
    # ============================================================
    print(f"\n开始生成筛选相关表...")
    
    config = {
        'corrected_content_col': corrected_content_col,
        'original_content_col': original_content_col,
        'voice_address_col': voice_address_col,
        'error_analysis_col': error_analysis_col,
        'correct_symbols_list': correct_symbols_list,
        'dirty_symbols_list': dirty_symbols_list,
        'error_analysis_exclusion_keywords': error_analysis_exclusion_keywords,
        'user_defined_filename_core_part': user_defined_filename_core_part,
        'fixed_path_string': fixed_path_string,
        'file_extension': file_extension,
    }
    
    risk_removed_df, risk_data_df, filtered_df, status = generate_filtered_sheets(combined_df, config)
    
    print(f"  生成状态: {status}")
    if risk_removed_df is not None and not risk_removed_df.empty:
        print(f"  '排序去除风险数据表' 包含 {len(risk_removed_df)} 行数据。")
    else:
        print(f"  '排序去除风险数据表' 未生成或为空。")
    
    if risk_data_df is not None and not risk_data_df.empty:
        print(f"  '风险数据表' 包含 {len(risk_data_df)} 行数据。")
    else:
        print(f"  '风险数据表' 未生成或为空。")
    
    if filtered_df is not None and not filtered_df.empty:
        print(f"  '排序已筛选' 包含 {len(filtered_df)} 行数据。")
    else:
        print(f"  '排序已筛选' 未生成或为空。")
    
    # ============================================================
    # ===== 保存 =====
    # ============================================================
    print(f"\n正在将工作表保存到: {output_file}")
    try:
        output_dir_final = os.path.dirname(output_file)
        os.makedirs(output_dir_final, exist_ok=True)
        with pd.ExcelWriter(output_file, engine='xlsxwriter', engine_kwargs={'options': {'strings_to_urls': False}}) as writer:
            print(f"  写入工作表 '{merged_sheet_name}'...")
            combined_df.to_excel(writer, sheet_name=merged_sheet_name, index=False, na_rep='')
            if cells_to_color:
                workbook = writer.book
                worksheet = writer.sheets[merged_sheet_name]
                red_format = workbook.add_format({'bg_color': fill_color_hex})
                print(f"  正在对 {len(cells_to_color)} 个单元格应用红色背景...")
                for row_idx, col_idx in cells_to_color:
                    worksheet.write(row_idx + 1, col_idx, combined_df.iloc[row_idx, col_idx], red_format)
            else:
                print(f"  无需对 '{merged_sheet_name}' 中的单元格应用标红格式。")
            
            print(f"  写入工作表 '{processed_sheet_name}'...")
            final_sheet2_df.to_excel(writer, sheet_name=processed_sheet_name, index=False, na_rep='')
            
            if risk_removed_df is not None and not risk_removed_df.empty:
                print(f"  写入工作表 '排序去除风险数据表'...")
                risk_removed_df.to_excel(writer, sheet_name='排序去除风险数据表', index=False, na_rep='')
            
            if risk_data_df is not None and not risk_data_df.empty:
                print(f"  写入工作表 '风险数据表'...")
                risk_data_df.to_excel(writer, sheet_name='风险数据表', index=False, na_rep='')
            
            if filtered_df is not None and not filtered_df.empty:
                print(f"  写入工作表 '排序已筛选'...")
                filtered_df.to_excel(writer, sheet_name='排序已筛选', index=False, na_rep='')
        
        print(f"文件保存成功！路径: {os.path.abspath(output_file)}")
    except PermissionError:
        print(f"错误：权限不足，无法写入文件 '{output_file}'。")
        print("\n脚本执行因保存错误而中止。")
        break
    except Exception as e:
        err_type = type(e).__name__
        print(f"保存合并文件时发生 {err_type} 错误: {e}")
        print("\n脚本执行因保存错误而中止。")
        break
    
    print("\n脚本执行完毕。")
    break