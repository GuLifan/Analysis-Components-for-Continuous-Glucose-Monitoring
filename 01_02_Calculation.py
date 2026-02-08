# 01_02_Calculation.py
# 完整计算代码，具体需要在config里修改
# LAST UPDATE BY LIFANGU IN 202602082100

# --- 导入包（忽视 IDE 高亮，已是最简洁状态） ---
import os
import pandas as pd
import math
from datetime import datetime, time, timedelta
import openpyxl
import numpy as np
import argparse

# -- 导入配置 ---
import yaml

def load_config(config_path='config.yaml'):
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f)
    except Exception as e:
        print(f"Error loading config.yaml: {e}")
        exit(1)

config_data = load_config()

# Unpack config variables
mode = config_data.get('mode', 0)
interimday = config_data.get('interimday', 7)
duringday = config_data.get('duringday', 7)
nametag = config_data.get('nametag', '')
datetag = config_data.get('datetag', '')
patient_list_file = config_data.get('patient_list_file', '')
data_folder = config_data.get('data_folder', '')
output_folder = config_data.get('output_folder', '')
CALC_GROUPS = config_data.get('calc_groups', {})
MATCH_BY = config_data.get('match_by', 'sensor_id')
daily_output = config_data.get('daily_output', False)

def get_output_filename(current_mode):
    template = config_data.get('output_filename_template', "CGM_{nametag}_Mode{mode}_Results.xlsx")
    filename = template.format(
        datetag=datetag,
        nametag=nametag,
        mode=current_mode
    )
    return os.path.join(output_folder, filename)
# 每天最少需要点数，目前执行288/天（间隔5分钟）
minlennum_per_day = 12 * 24  # 每天所需的最小数据点数（每5分钟一个点）

# --- 功能函数 ---
""" (01) 计算基础指标组 (Basic Indicators Group)，包含：MEAN, SD, CV, GMI """
def calc_basic_stats(df):
    if df.empty: return {}
    # 获取计算队列中的血糖数据列
    glucose_series = df['glucose']
    # 计算基础指标
    mean_val = glucose_series.mean() # 1. MEAN (平均血糖): 血糖读数的算术平均值
    std_val = glucose_series.std() # 2. SD (标准差): 血糖读数的标准差，反映波动程度
    # 3. CV (变异系数): 标准差与平均值的比值 (SD/MEAN)，反映相对波动程度
    if mean_val != 0 and not np.isnan(mean_val):
        cv_val = round(std_val / mean_val, 4)
    else:
        cv_val = None
    # 4. GMI (血糖管理指标): 基于平均血糖估算的HbA1c水平，GMI = 3.31 + 0.02392 * Mean_Glucose(mg/dL)
    if not np.isnan(mean_val):
        gmi_val = round(3.31 + 0.02392 * 18 * mean_val, 4) # 注意: 输入单位为 mmol/L，需乘以 18 转换为 mg/dL
    else:
        gmi_val = None

    return {
        'MEAN': round(mean_val, 4) if not np.isnan(mean_val) else None,
        'SD': round(std_val, 4) if not np.isnan(std_val) else None,
        'CV': cv_val,
        'GMI': gmi_val
    }


""" (02) 计算风险指标组 (Risk Indicators Group)，包含：LBGI, HBGI, ADRR, MODD """
"""
    注意：
    1. 计算前需确保数据已按时间排序
    2. 风险指标基于 mg/dL 计算，需先将 mmol/L 转换为 mg/dL (x 18)
    3. 风险指标基于24小时分段计算，每个24小时计算一次指标, 默认输出每日指标的平均值 (或返回数组)
"""
def calc_lbgi_hbgi_adrr(df):
    if df.empty:
        return ([], [], []) if daily_output else (None, None, None)
    # 排序并计算相对天数索引
    df = df.sort_values('timestamp').copy()
    start_time = df['timestamp'].iloc[0]
    df['day_idx'] = ((df['timestamp'] - start_time).dt.total_seconds() // 86400).astype(int)
    # 过滤掉不足24h的最后一天，计算总时长包含的完整天数
    total_seconds = (df['timestamp'].iloc[-1] - start_time).total_seconds()
    num_full_days = int(total_seconds // 86400)
    # 仅保留完整天数的数据
    valid_df = df[df['day_idx'] < num_full_days].copy()
    
    if valid_df.empty:
        return ([], [], []) if daily_output else (None, None, None)

    # 过滤掉小于1.0的血糖值，防止 log 结果为负数导致幂运算错误 (log(1)=0)
    valid_df = valid_df[valid_df['glucose'] >= 1.0].copy()
    
    if valid_df.empty:
        return ([], [], []) if daily_output else (None, None, None)

    # 恢复用户习惯的“旧代码”公式，但使用新的（标准的）ADRR计算逻辑（Sum of Maxes）
    # 解释：
    # Old Logic (User's Code): ADRR = Mean(Risk)
    # New Logic (Standard): ADRR = Max(RL) + Max(RH) per day
    # User Request: "Follow New Logic, but use mmol/L unit"
    # Action: Use 1.794 formula (mmol/L specific) AND Max+Max aggregation.
    
    # Formula: 1.794 * (ln(bg)^1.026 - 1.861)  (Input: mmol/L)
    valid_df['fBG'] = 1.794 * (np.log(valid_df['glucose']) ** 1.026 - 1.861)
    valid_df['risk'] = 10 * (valid_df['fBG'] ** 2)
    
    valid_df['rl'] = np.where(valid_df['fBG'] < 0, valid_df['risk'], 0)
    valid_df['rh'] = np.where(valid_df['fBG'] > 0, valid_df['risk'], 0)
    
    if daily_output:
        lbgi_list = []
        hbgi_list = []
        adrr_list = []
        
        for day_id in sorted(valid_df['day_idx'].unique()):
            day_data = valid_df[valid_df['day_idx'] == day_id]
            
            # LBGI: Mean of RL (Risk Low)
            # Standard definition is Mean of RL (calculated over readings)
            low_risk = day_data['rl']
            lbgi = low_risk.mean() if not low_risk.empty else 0
            
            # HBGI: Mean of RH (Risk High)
            high_risk = day_data['rh']
            hbgi = high_risk.mean() if not high_risk.empty else 0
            
            # ADRR: Max(RL) + Max(RH)
            max_rl = day_data['rl'].max() if not day_data.empty else 0
            max_rh = day_data['rh'].max() if not day_data.empty else 0
            adrr = max_rl + max_rh
            
            lbgi_list.append(round(lbgi, 4))
            hbgi_list.append(round(hbgi, 4))
            adrr_list.append(round(adrr, 4))
            
        return lbgi_list, hbgi_list, adrr_list
    else:
        # Global Aggregation (Average of daily metrics)
        
        daily_stats = valid_df.groupby('day_idx').agg(
            mean_rl=('rl', 'mean'),
            mean_rh=('rh', 'mean'),
            max_rl=('rl', 'max'),
            max_rh=('rh', 'max')
        )
        
        # Calculate daily ADRR first
        daily_stats['daily_adrr'] = daily_stats['max_rl'] + daily_stats['max_rh']
        
        # Then average them
        final_lbgi = daily_stats['mean_rl'].mean()
        final_hbgi = daily_stats['mean_rh'].mean()
        final_adrr = daily_stats['daily_adrr'].mean()
        
        return round(final_lbgi, 4), round(final_hbgi, 4), round(final_adrr, 4)


"""
    计算日间血糖平均绝对差 (MODD)
    逻辑:
        1. 按24h分段 (Day 0, Day 1...)
        2. 比较 Day i 与 Day i-1 的同一时刻血糖差绝对值
        3. MODD 从第二天 (Day 1) 开始计算
        4. 若仅有1个完整24h，返回 N/A
"""
def calc_modd(df):
    if df.empty:
        return [] if daily_output else None
        
    df = df.sort_values('timestamp').copy()
    start_time = df['timestamp'].iloc[0]
    # 1. 计算完整天数
    total_seconds = (df['timestamp'].iloc[-1] - start_time).total_seconds()
    num_full_days = int(total_seconds // 86400)
    
    if num_full_days < 2:
        # 不足2天，无法计算MODD (需对比Day 1和Day 0)
        return [] if daily_output else None
        
    # 2. 标记天数索引
    df['day_idx'] = ((df['timestamp'] - start_time).dt.total_seconds() // 86400).astype(int)
    valid_df = df[df['day_idx'] < num_full_days].copy()
    
    # 3. 对齐时间 (使用 dt.time 可能在跨日时有歧义，但这里是相对24h，用time_offset更稳)
    # 实际上如果采样规律，dt.time在24h切片下是一致的。
    # 为了兼容不规则采样，使用 resample/pivot based on rounded time
    valid_df['time_key'] = valid_df['timestamp'].dt.strftime('%H:%M') # 简单按分钟对齐
    
    # Pivot: Index=time, Columns=day_idx
    pivot = valid_df.pivot_table(index='time_key', columns='day_idx', values='glucose')
    
    # 4. 计算差值 (Day i - Day i-1)
    diffs = pivot.diff(axis=1).abs()
    
    # 5. 计算每日 MODD (每一列的均值)
    # diffs 的列索引是 0, 1, 2...
    # column 0 将是 NaN (因为没有 -1)
    # column 1 是 Day 1 - Day 0
    daily_modd = diffs.mean()
    
    # 只要 Day 1 及以后的数据
    # Drop column 0 or NaN columns
    valid_modds = daily_modd.dropna()
    
    if daily_output:
        return valid_modds.round(4).tolist()
    else:
        val = valid_modds.mean()
        return round(val, 4) if not np.isnan(val) else None


""" (03) 计算血糖波动指标组 (Glucose Variability Indicators Group)，包含：MAGE, LAGE """
def calc_mage_daily(glucose_series):
    """
    计算单日的MAGE (Service 1970定义)
    :param glucose_series: 单日血糖数据 (pd.Series)
    :return: MAGE value or None
    """
    # 移除NaN
    data = glucose_series.dropna().values
    if len(data) < 3:
        return None
        
    # 1. 计算SD (24h内血糖值的SD)
    sd = np.std(data, ddof=1)
    if sd == 0:
        return 0.0
        
    # 2. 识别峰值和谷值 (Turning Points)
    peaks = [] # (index, value)
    nadirs = [] # (index, value)
    
    # 简单的局部极值检测
    for i in range(1, len(data) - 1):
        if data[i] > data[i-1] and data[i] > data[i+1]:
            peaks.append((i, data[i]))
        elif data[i] < data[i-1] and data[i] < data[i+1]:
            nadirs.append((i, data[i]))
            
    if not peaks or not nadirs:
        return None
        
    # 合并极值点并按索引排序
    turning_points = sorted(peaks + nadirs, key=lambda x: x[0])
    
    # 3. 寻找有效波动 (Effective AGEs)
    # Service 1970: 只有当血糖波动幅度 > 1 SD 时被认为是有效的
    # 以第1个有效的AGE方向为准 (上升支或下降支)
    
    first_valid_direction = None # 1: Up (Nadir->Peak), -1: Down (Peak->Nadir)
    mage_sum = 0
    mage_count = 0
    
    # 遍历相邻的转折点
    for i in range(1, len(turning_points)):
        current_val = turning_points[i][1]
        prev_val = turning_points[i-1][1]
        
        diff = current_val - prev_val
        amplitude = abs(diff)
        
        # 判断幅度是否大于SD
        if amplitude > sd:
            # 确定当前波动的方向
            direction = 1 if diff > 0 else -1
            
            if first_valid_direction is None:
                # 找到第一个有效波动，确定方向
                first_valid_direction = direction
                mage_sum += amplitude
                mage_count += 1
            elif direction == first_valid_direction:
                # 只统计与第一个有效波动同方向的波动
                mage_sum += amplitude
                mage_count += 1
                
    if mage_count == 0:
        return None
        
    # 4. 计算MAGE (有效AGE的均值)
    return mage_sum / mage_count


def calc_lage_mage(df):
    """计算LAGE和MAGE指标"""
    if df.empty:
         return ([], []) if daily_output else (None, None)

    # 确保按时间排序
    df = df.sort_values('timestamp').copy()
    start_time = df['timestamp'].iloc[0]
    
    # 计算完整天数 (按24h切分，忽略最后不足24h的部分)
    total_seconds = (df['timestamp'].iloc[-1] - start_time).total_seconds()
    num_full_days = int(total_seconds // 86400)
    
    # 标记天数索引
    df['day_idx'] = ((df['timestamp'] - start_time).dt.total_seconds() // 86400).astype(int)
    
    # 仅使用完整天数的数据
    valid_df = df[df['day_idx'] < num_full_days].copy()
    
    if valid_df.empty:
         return ([], []) if daily_output else (None, None)
    
    daily_lages = []
    daily_mages = []
    
    grouped = valid_df.groupby('day_idx')
    
    # 确保按天数顺序处理 (0, 1, 2...)
    # 如果某天没有数据，groupby会自动跳过，但我们需要保持索引对齐吗？
    # 现在的逻辑是返回存在的有效计算值。
    
    for day in sorted(grouped.groups.keys()):
        day_df = grouped.get_group(day)
        glucose_series = day_df['glucose']
        
        # 简单检查数据量，每天应有足够数据点 (例如 > 70% 覆盖，即 > 200点/288点)
        # 放宽一点，至少50% (144点)
        if len(glucose_series.dropna()) < 144: 
            # 数据不足，该天记为None还是跳过？
            # 为了保持daily_output数组长度对应天数，最好append None，但后续计算均值需注意
            # 这里简单跳过，或者视作该天无法计算
            continue
            
        # LAGE: Max - Min
        lage = glucose_series.max() - glucose_series.min()
        daily_lages.append(lage)
        
        # MAGE
        mage = calc_mage_daily(glucose_series)
        if mage is not None:
            daily_mages.append(mage)
            
    # 结果格式化
    if daily_output:
        return (
            [round(x, 4) for x in daily_lages], 
            [round(x, 4) for x in daily_mages]
        )
    else:
        mean_lage = np.mean(daily_lages) if daily_lages else None
        mean_mage = np.mean(daily_mages) if daily_mages else None
        return (
            round(mean_lage, 4) if mean_lage is not None else None,
            round(mean_mage, 4) if mean_mage is not None else None
        )


""" (04) 计算血糖范围指标组 (Range Group)，包含：TIR, TAR, TBR (all variants), TITR, 0-6/6-24 variants """
def calc_range_stats(df, prefix=''):
    """计算TIR/TAR/TBR等范围指标"""
    if df.empty:
        return {}
    
    total = len(df)
    g = df['glucose'].values
    
    res = {}
    # Helper for safe division
    def calc_ratio(count):
        return round(count / total, 4) if total > 0 else 0

    # Basic ranges
    res[f'TIR{prefix}'] = calc_ratio(np.sum((g >= 3.9) & (g <= 10.0)))
    res[f'TAR{prefix}'] = calc_ratio(np.sum(g > 10.0))
    res[f'TBR{prefix}'] = calc_ratio(np.sum(g < 3.9))
    
    # Detailed ranges
    # TAR Level 1: > 10.0 and <= 13.9
    res[f'TAR1{prefix}'] = calc_ratio(np.sum((g > 10.0) & (g <= 13.9)))
    # TAR Level 2: > 13.9
    res[f'TAR2{prefix}'] = calc_ratio(np.sum(g > 13.9))
    # TBR Level 1: < 3.9 and >= 3.0
    res[f'TBR1{prefix}'] = calc_ratio(np.sum((g >= 3.0) & (g < 3.9)))
    # TBR Level 2: < 3.0
    res[f'TBR2{prefix}'] = calc_ratio(np.sum(g < 3.0))
    # TITR: 3.9 - 7.8 (修正为标准定义，防止 TIR-TITR 出现负数)
    res[f'TITR{prefix}'] = calc_ratio(np.sum((g >= 3.9) & (g <= 7.8)))
    
    if f'TIR{prefix}' in res and f'TITR{prefix}' in res:
            res[f'TIR-TITR{prefix}'] = round(res[f'TIR{prefix}'] - res[f'TITR{prefix}'], 4)
             
    return res



""" (06) 事件统计指标组 (Event Stats Group) """
def find_simple_events(df, threshold, compare_func, min_duration_min=15):
    """
    识别简单事件 (连续N分钟满足条件)
    :param df: DataFrame with timestamp and glucose
    :param threshold: 阈值
    :param compare_func: 比较函数 (lambda x, th: x < th)
    :param min_duration_min: 最小持续时间(分)
    :return: list of (start_time, end_time)
    """
    if df.empty: return []
    
    # 标记满足条件的行
    is_event = df['glucose'].apply(lambda x: compare_func(x, threshold))
    
    # 分组连续满足的块
    # 逻辑: (is_event != is_event.shift()).cumsum() 创建组ID
    # 只取 is_event 为 True 的组
    
    events = []
    
    # 为处理时间间隔不均匀，我们迭代行可能更稳健，或者假设间隔均匀
    # 给定数据主要是5min间隔，使用时间差判断更准
    
    # 迭代法识别事件
    in_event = False
    start_time = None
    
    # 转换为列表加速
    times = df['timestamp'].tolist()
    values = df['glucose'].tolist()
    
    # 临时存储当前事件的所有时间点
    current_event_times = []
    
    for i in range(len(values)):
        val = values[i]
        t = times[i]
        
        if compare_func(val, threshold):
            if not in_event:
                in_event = True
                start_time = t
                current_event_times = [t]
            else:
                # 检查时间连续性 (允许一定误差，比如缺一个点? 这里暂严格按时间差累计)
                # 简单处理：只要满足条件就加入。计算duration时用 end-start
                # 但如果中间断了很久(比如几天)? 
                # 这里假设df是按时间排序的
                
                # 检查与上一个点的时间差，如果超过例如 15min，视为新事件?
                # 这里的定义是 "连续"，通常指中间没有不满足的点。
                # 至于时间断点，CGM通常会有。
                # 如果中间缺数据，视作不连续。
                
                prev_t = current_event_times[-1]
                if (t - prev_t).total_seconds() > 15 * 60: # 间隙 > 15min
                     # 结算上一个事件
                     duration = (current_event_times[-1] - start_time).total_seconds() / 60
                     # 单点duration视为0? 或者5min? 
                     # 通常 continuous 15min means >= 3 points (0, 5, 10 -> 10min interval covered, +5min duration?)
                     # 临床定义通常是 Duration = End - Start (+ SamplingInterval?)
                     # 这里为了保守，用 time difference。如果3个点: 0, 5, 10. Diff is 10min. 
                     # 这里的 min_duration_min=15.
                     # 如果我们要 "连续15min"，通常意味着覆盖了 15min 的跨度。
                     # 比如 0, 5, 10, 15. Diff=15. 4个点。
                     
                     if duration >= min_duration_min:
                         events.append((start_time, current_event_times[-1]))
                     
                     # 开启新事件
                     start_time = t
                     current_event_times = [t]
                else:
                    current_event_times.append(t)
        else:
            if in_event:
                # 事件结束
                duration = (current_event_times[-1] - start_time).total_seconds() / 60
                if duration >= min_duration_min:
                    events.append((start_time, current_event_times[-1]))
                
                in_event = False
                current_event_times = []
                
    # 循环结束后的处理
    if in_event:
        duration = (current_event_times[-1] - start_time).total_seconds() / 60
        if duration >= min_duration_min:
            events.append((start_time, current_event_times[-1]))
            
    return events


def find_complex_events(df, start_func, end_condition_func, min_event_duration=120):
    """
    识别复杂事件 (Extended Events)
    :param df: DataFrame
    :param start_func: lambda x: x < 3.9 (判断开始)
    :param end_condition_func: lambda val_list: 检查后续一段数据是否满足终止条件 (例如 >= 3.9 持续 15min)
    :param min_event_duration: 最小事件持续时间 (分)
    :return: list of (start_time, end_time)
    """
    if df.empty: return []
    
    events = []
    times = df['timestamp'].tolist()
    values = df['glucose'].tolist()
    n = len(values)
    
    i = 0
    while i < n:
        # 1. 寻找开始点
        if start_func(values[i]):
            start_idx = i
            start_time = times[i]
            
            # 2. 寻找结束点
            # 向后扫描，直到找到满足终止条件的位置
            end_idx = -1
            
            # 从 start_idx + 1 开始找
            j = start_idx + 1
            while j < n:
                # 检查从 j 开始的一段数据是否满足终止条件
                # 终止条件需要一段数据，我们假设 end_condition_func 接收 (values[j:], times[j:])
                # 但为了效率，我们在这里硬编码逻辑:
                # 终止: "以血糖...持续>=15min"
                # 意味着从 j 开始，未来15min内的数据都满足恢复条件
                
                # 获取从 j 开始覆盖 15min 的数据切片
                # 实际上是看 j, j+1, j+2... 是否都满足恢复阈值
                
                # 调用回调检查
                # 为了通用性，传递 index j 和 data context
                is_terminated, advance_steps = end_condition_func(values, times, j)
                
                if is_terminated:
                    # 事件在 j 点终止 (j 是恢复的开始点)
                    # 所以事件区间是 start_idx 到 j-1 (或者 j? 通常事件不包含恢复期)
                    # "以...作为事件终止" -> j 是终止点。事件持续到 j (不含 j 或含 j 的边界?)
                    # 持续时间通常计算为 times[j] - times[start_idx]
                    end_idx = j
                    break
                
                # 如果没有终止，继续下一个点
                # 如果数据断裂太长(如 > 30min)，是否强制终止?
                # 暂时假设数据相对连续。
                if j > start_idx and (times[j] - times[j-1]).total_seconds() > 30 * 60:
                    # 数据中断，事件被迫终止于 j-1
                    end_idx = j # 视为在 j 处结束(虽然 j 是很久以后)
                    # 或者 j-1?
                    # 简单起见，如果断了，就截止到 j-1
                    end_idx = j 
                    break
                    
                j += 1
            
            if end_idx == -1:
                # 到了文件末尾还没终止，视为持续到最后
                end_idx = n
            
            # 计算持续时间
            # 事件结束时间：如果是恢复终止，结束时间是 times[end_idx] (恢复开始的时间)
            # 如果是文件末尾，是 times[n-1]
            
            if end_idx < n:
                event_end_time = times[end_idx]
            else:
                event_end_time = times[n-1]
                
            duration = (event_end_time - start_time).total_seconds() / 60
            
            if duration >= min_event_duration:
                events.append((start_time, event_end_time))
            
            # 下一次搜索从 end_idx 开始 (或者是 end_idx + 1?)
            # 既然 end_idx 是恢复的开始，或者是断点，我们从 end_idx 开始找下一次 Start
            i = end_idx
        else:
            i += 1
            
    return events


def calc_event_stats(df):
    """计算Group 6 事件统计"""
    stats = {}
    if df.empty: return stats
    
    # 按时间排序
    df = df.sort_values('timestamp').copy()
    
    # 格式化输出函数
    def fmt_events(evt_list):
        if not evt_list: return 0, None
        count = len(evt_list)
        # Format: "start~end, start~end"
        time_strs = [f"{s.strftime('%Y-%m-%d %H:%M:%S')}~{e.strftime('%Y-%m-%d %H:%M:%S')}" for s, e in evt_list]
        return count, ",".join(time_strs)

    # 1. Hypoglycemic events (<3.9, >=15min)
    hypo_events = find_simple_events(
        df, 
        threshold=3.9, 
        compare_func=lambda x, th: x < th, 
        min_duration_min=15
    )
    c, t = fmt_events(hypo_events)
    stats['HYPO'] = c
    stats['Time-HYPO'] = t
    
    # 2. HYPO 0TO6AM (Start time in 0-6)
    hypo_0to6 = [e for e in hypo_events if 0 <= e[0].hour < 6]
    c, t = fmt_events(hypo_0to6)
    stats['HYPO 0TO6AM'] = c
    stats['Time-HYPO 0TO6AM'] = t
    
    # 3. Extended Hypo (<3.9, >120min, End >=3.9 for 15min)
    def check_hypo_recovery(vals, ts, idx):
        # 检查从 idx 开始是否满足 >= 3.9 持续 15min
        # 需覆盖时间跨度 >= 15min 的点都 >= 3.9
        if idx >= len(vals): return False, 0
        
        start_t = ts[idx]
        curr_idx = idx
        
        while curr_idx < len(vals):
            # 如果值 < 3.9，恢复失败
            if vals[curr_idx] < 3.9:
                return False, 0
            
            # 如果值 >= 3.9，检查时间跨度
            span = (ts[curr_idx] - start_t).total_seconds() / 60
            if span >= 15:
                return True, 0 # 成功恢复
            
            # 如果相邻点断裂 > 30min，视为无法判断连续性，返回False? 
            # 或者认为恢复中断。这里简单处理：只要遇到不满足就Fail。
            curr_idx += 1
            
        return False, 0 # 到了末尾也没满足15min
        
    ex_hypo_events = find_complex_events(
        df,
        start_func=lambda x: x < 3.9,
        end_condition_func=check_hypo_recovery,
        min_event_duration=120
    )
    c, t = fmt_events(ex_hypo_events)
    stats['EX HYPO'] = c
    stats['Time-EX HYPO'] = t
    
    # 4. EX HYPO 0TO6AM
    ex_hypo_0to6 = [e for e in ex_hypo_events if 0 <= e[0].hour < 6]
    c, t = fmt_events(ex_hypo_0to6)
    stats['EX HYPO 0TO6AM'] = c
    stats['Time-EX HYPO 0TO6AM'] = t
    
    # 5. Extended Hyper (>13.9, >120min, End <=10.0 for 15min)
    def check_hyper_recovery(vals, ts, idx):
        if idx >= len(vals): return False, 0
        start_t = ts[idx]
        curr_idx = idx
        while curr_idx < len(vals):
            if vals[curr_idx] > 10.0: # 必须 <= 10.0
                return False, 0
            span = (ts[curr_idx] - start_t).total_seconds() / 60
            if span >= 15:
                return True, 0
            curr_idx += 1
        return False, 0

    ex_hyper_events = find_complex_events(
        df,
        start_func=lambda x: x > 13.9,
        end_condition_func=check_hyper_recovery,
        min_event_duration=120
    )
    c, t = fmt_events(ex_hyper_events)
    stats['EX HYPER'] = c
    stats['Time-EX HYPER'] = t
    
    return stats


def time_period_stats(df, start_hour, end_hour):
    """计算指定时间范围内的统计指标"""
    start_time = datetime.strptime(f"{start_hour:02d}:00:00", "%H:%M:%S").time()
    end_time = datetime.strptime(f"{end_hour:02d}:59:59", "%H:%M:%S").time()
    # 筛选时间段内的数据
    mask = (df['timestamp'].dt.time >= start_time) & (df['timestamp'].dt.time <= end_time)
    period_df = df[mask].copy()
    if period_df.empty:
        return {
            'mean': None,
            'std': None,
            'cv': None,
            'median': None,
            'vv_list': None,
            'vv_time_list': None
        }
    # 计算基本统计量
    glucose_series = period_df['glucose'].dropna()
    if glucose_series.empty:
        return {
            'mean': None,
            'std': None,
            'cv': None,
            'median': None,
            'vv_list': None,
            'vv_time_list': None
        }

    mean = glucose_series.mean()
    std = glucose_series.std()
    median = glucose_series.median()
    cv = std / mean if (mean != 0 and not np.isnan(mean)) else None

    # 四舍五入处理
    mean = round(mean, 4) if not np.isnan(mean) else None
    std = round(std, 4) if not np.isnan(std) else None
    cv = round(cv, 4) if cv is not None and not np.isnan(cv) else None
    median = round(median, 4) if not np.isnan(median) else None

    # 按天处理每日最小值及时间
    period_df['date'] = period_df['timestamp'].dt.date
    daily_groups = period_df.groupby('date')

    daily_mins = []
    daily_min_times = []

    for date, group in daily_groups:
        valid_group = group.dropna(subset=['glucose'])
        if valid_group.empty:
            continue
        min_value = valid_group['glucose'].min()
        min_rows = valid_group[valid_group['glucose'] == min_value]
        min_time = min_rows['timestamp'].iloc[0].strftime('%H:%M:%S')
        daily_mins.append(round(min_value, 4))
        daily_min_times.append(min_time)

    daily_min_str = ','.join(map(str, daily_mins)) if daily_mins else None
    daily_min_time_str = ','.join(daily_min_times) if daily_min_times else None

    return {
        'mean': mean,
        'std': std,
        'cv': cv,
        'median': median,
        'vv_list': daily_min_str,
        'vv_time_list': daily_min_time_str
    }


def daily_closest_time_stats(df, target_times):
    """计算每天最接近指定时间的血糖统计量"""
    stats = {}
    for target_time_str in target_times:
        target_time = datetime.strptime(target_time_str, '%H:%M:%S').time()
        glucose_values = []

        for date, date_group in df.groupby(df['timestamp'].dt.date):
            target_datetime = datetime.combine(date, target_time)
            date_group = date_group.copy()
            date_group['time_diff'] = (date_group['timestamp'] - target_datetime).abs()

            if not date_group.empty:
                min_row = date_group.loc[date_group['time_diff'].idxmin()]
                glucose_values.append(min_row['glucose'])

        if glucose_values:
            mean = np.mean(glucose_values)
            std = np.std(glucose_values, ddof=1)
            cv = std / mean if mean != 0 else np.nan
            median = np.median(glucose_values)
        else:
            mean = std = cv = median = np.nan

        prefix = target_time_str.replace(':', '')
        stats[f'{prefix}_mean'] = round(mean, 4) if not np.isnan(mean) else None
        stats[f'{prefix}_std'] = round(std, 4) if not np.isnan(std) else None
        stats[f'{prefix}_cv'] = round(cv, 4) if not np.isnan(cv) else None
        stats[f'{prefix}_median'] = round(median, 4) if not np.isnan(median) else None

    return stats


def analyze_hypo_events(df, threshold, min_duration=15, min_val_threshold=None, time_range=None):
    """
    通用低血糖事件分析
    :param df: DataFrame
    :param threshold: 低血糖阈值 (例如 3.0)
    :param min_duration: 持续时间 (例如 15分钟)
    :param min_val_threshold: 事件期间最低值需低于此值 (可选)
    :param time_range: 时间范围 (start_time, end_time)
    :return: (count, time_list_str)
    """
    if df.empty:
        return 0, None

    # Filter by time range
    if time_range:
        start_time, end_time = time_range
        if start_time <= end_time:
            df = df[df['timestamp'].dt.time.between(start_time, end_time)].copy()
        else:
            # Cross midnight
            df = df[(df['timestamp'].dt.time >= start_time) | (df['timestamp'].dt.time <= end_time)].copy()
    else:
        df = df.copy()

    if df.empty:
        return 0, None

    # Identify rows below threshold
    df['is_low'] = df['glucose'] < threshold
    
    # Identify events (consecutive True)
    df['group_id'] = (df['is_low'] != df['is_low'].shift()).cumsum()
    
    events_count = 0
    events_times = []
    
    # Only iterate over groups where is_low is True
    low_groups = df[df['is_low']].groupby('group_id')
    
    for _, group in low_groups:
        duration = len(group) * 5 # Assuming 5 min intervals
        
        if duration >= min_duration:
            # Check min value condition if specified
            if min_val_threshold is not None:
                min_val = group['glucose'].min()
                if not (min_val < min_val_threshold):
                    continue
            
            events_count += 1
            events_times.append(group['timestamp'].iloc[0].strftime('%Y-%m-%d %H:%M:%S'))
            
    return events_count, ','.join(events_times) if events_times else None


def process_patient_file(file_path, admission_time=None, discharge_time=None, pump_start_time=None, pump_end_time=None, current_mode=None):
    """
    处理患者CGM数据文件
    """
    # 使用传入的模式或全局模式
    selected_mode = current_mode if current_mode is not None else mode
    
    try:
        df = pd.read_excel(file_path, header=None)
    except Exception as e:
        print(f"无法读取文件 {file_path}: {e}")
        return None

    df = df.drop(0).reset_index(drop=True)
    df.columns = ['timestamp', 'glucose']

    try:
        df['timestamp'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%d %H:%M:%S')
    except Exception as e:
        print(f"文件 {file_path} 时间格式转换失败: {e}")
        return None

    df['glucose'] = pd.to_numeric(df['glucose'], errors='coerce')
    df['glucose'] = df['glucose'].clip(1.8, 33.3)

    # 获取数据的开始和结束时间
    data_start_time = df['timestamp'].min()
    data_end_time = df['timestamp'].max()
    
    # 根据不同模式设置计算的开始和结束时间
    if selected_mode == 0:  # 开始时间~开始时间+duringday天
        start_time = data_start_time
        end_time = start_time + timedelta(days=duringday)
    elif selected_mode == 1:  # 开始时间+interimday天~开始时间+duringday天
        start_time = data_start_time + timedelta(days=interimday)
        end_time = data_start_time + timedelta(days=duringday)
        # 确保结束时间大于开始时间
        if end_time <= start_time:
            print(f"警告：模式1中结束时间不能小于等于开始时间，请确保duringday > interimday")
            # 强制设置结束时间为开始时间+1天
            end_time = start_time + timedelta(days=1)
    elif selected_mode == 2:  # 开始时间~结束时间
        start_time = data_start_time
        end_time = data_end_time
    elif selected_mode == 3:  # 出院时间-duringday天~出院时间
        # 检查出院时间是否有效
        if pd.isna(discharge_time):
            print(f"文件 {file_path} 出院时间无效，无法计算mode3")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
            
        start_time = discharge_time - timedelta(days=duringday)
        end_time = discharge_time
        
        # 检查出院时间是否在数据范围内
        if end_time < data_start_time:
            print(f"警告：出院时间 {end_time} 早于数据开始时间 {data_start_time}，无法计算mode3")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': start_time,
                'end_time': end_time,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
    elif selected_mode == 4:  # 佩戴血糖仪开始时间~出院时间
        # 检查佩戴开始时间和出院时间是否有效
        if pd.isna(data_start_time) or pd.isna(discharge_time):
            print(f"文件 {file_path} 佩戴开始时间或出院时间无效，无法计算mode4")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
        
        # 检查佩戴开始时间和出院时间是否合理
        if data_start_time >= discharge_time:
            print(f"警告：佩戴开始时间 {data_start_time} 不早于出院时间 {discharge_time}，无法计算mode4")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
            
        # 检查佩戴开始时间和出院时间是否在数据范围内
        if discharge_time < data_start_time or data_start_time > data_end_time:
            print(f"警告：佩戴开始时间 {data_start_time} 到出院时间 {discharge_time} 与数据时间范围 {data_start_time} 到 {data_end_time} 没有重叠，无法计算mode4")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': data_start_time,
                'end_time': discharge_time,
                'device_start_time': data_start_time,
                'discharge_time': discharge_time
            }
            
        start_time = data_start_time  # 使用佩戴开始时间作为计算起始点
        end_time = discharge_time
    elif selected_mode == 5:  # 出院时间~出院时间+duringday天
        # 检查出院时间是否有效
        if pd.isna(discharge_time):
            print(f"文件 {file_path} 出院时间无效，无法计算mode5")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
        
        start_time = discharge_time
        end_time = discharge_time + timedelta(days=duringday)
        
        # 检查出院时间是否在数据范围内
        if start_time > data_end_time:
            print(f"警告：出院时间 {start_time} 晚于数据结束时间 {data_end_time}，无法计算mode5")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': start_time,
                'end_time': end_time,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
    elif selected_mode == 6:  # 出院时间~结束时间
        # 检查出院时间是否有效
        if pd.isna(discharge_time):
            print(f"文件 {file_path} 出院时间无效，无法计算mode6")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
        
        start_time = discharge_time
        end_time = data_end_time
        
        # 检查出院时间是否在数据范围内
        if start_time > data_end_time:
            print(f"警告：出院时间 {start_time} 晚于数据结束时间 {data_end_time}，无法计算mode6")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': start_time,
                'end_time': end_time,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            }
    elif selected_mode == 7: # 胰岛素泵使用时间
        if pd.isna(pump_start_time) or pd.isna(pump_end_time):
            print(f"文件 {file_path} 胰岛素泵时间无效，无法计算mode7")
            return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': None,
                'end_time': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time,
                'note': 'Invalid Pump Time'
            }
            
        # 1. 在血糖数据中找到最接近泵开始/结束时间的两个时刻
        # 计算所有时间戳与 pump_start/end 的差值
        time_diff_start = (df['timestamp'] - pump_start_time).abs()
        time_diff_end = (df['timestamp'] - pump_end_time).abs()
        
        # 找到最小差值的索引
        closest_start_idx = time_diff_start.idxmin()
        closest_end_idx = time_diff_end.idxmin()
        
        # 获取对应的实际血糖时间
        # 优化逻辑：确保这两个时刻之间包含住泵使用时间（Start <= PumpStart, End >= PumpEnd 吗？)
        # 用户要求：“且一般来说这两个时刻之间应当包含住胰岛素泵使用时间”
        # 实际上，通常我们希望找到数据中覆盖该段的时间点。
        # 如果直接找最接近点，可能是在PumpStart之后一点点。
        # 为了“包含住”，我们可以尝试找 <= PumpStart 的最后一个点，和 >= PumpEnd 的第一个点。
        # 但如果数据本身不对齐，可能找不到。
        # 这里采用最接近点作为计算区间的边界，通常是最合理的近似。
        # 实际上，计算区间应当是 [closest_start, closest_end]。
        
        start_time = df.loc[closest_start_idx, 'timestamp']
        end_time = df.loc[closest_end_idx, 'timestamp']
        
        # 如果最接近的开始时间比结束时间还晚（数据异常或泵时间极短且错位），做个防御
        if start_time > end_time:
            # 尝试交换或者报错
            # 这里简单做个修正，确保 start <= end
             start_time, end_time = end_time, start_time

        # 2. 检查胰岛素泵使用时间是否小于24小时
        pump_duration_hours = (pump_end_time - pump_start_time).total_seconds() / 3600
        duration_note = None
        if pump_duration_hours < 24:
            duration_note = f'Pump Duration < 24h ({pump_duration_hours:.1f}h)'
            print(f"警告：胰岛素泵使用时间过短: {pump_duration_hours:.1f}小时")

        if start_time >= end_time: # 仍然无法构成有效区间
             return {
                'patient_id': os.path.splitext(os.path.basename(file_path))[0],
                'start_time': start_time,
                'end_time': end_time,
                'admission_time': admission_time,
                'discharge_time': discharge_time,
                'note': duration_note if duration_note else 'Invalid Time Range'
            }
            
    else:
        print(f"未知模式: {selected_mode}")
        return None

    # 筛选指定时间范围内的数据
    df = df[(df['timestamp'] >= start_time) & (df['timestamp'] <= end_time)]

    # 计算该时间段所需的最小数据点数
    # 首先检查时间范围是否有效
    if pd.isna(start_time) or pd.isna(end_time):
        print(f"警告：文件 {file_path} 的时间范围无效（开始时间={start_time}, 结束时间={end_time}）")
        return {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        }
    
    # 计算时间差（天数）
    time_diff = end_time - start_time
    time_diff_days = time_diff.total_seconds() / (24 * 3600)
    
    # 检查时间差是否有效
    if math.isnan(time_diff_days) or time_diff_days <= 0:
        print(f"警告：文件 {file_path} 的时间差无效（{time_diff_days}天）")
        return {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        }
    
    # 安全计算所需点数
    try:
        required_min_points = max(0, int(time_diff_days * minlennum_per_day) - 36)
    except Exception as e:
        print(f"计算所需点数时出错: {e}")
        required_min_points = 0  # 设置默认值

    if df.empty:
        print(f"文件 {file_path} 在指定时间范围内无数据")
        return {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        }
    
    if len(df) < required_min_points:
        print(f"文件 {file_path} 数据不足 {time_diff_days:.1f} 天所需的数据量（需要至少 {required_min_points} 个数据点）")
        return {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        }
    


    # 计算统计指标
    total_rows = len(df)
    if total_rows == 0:
        # 返回基本信息但不包含计算结果
        return {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        }

    # 初始化结果字典
    stats = {
        'patient_id': os.path.splitext(os.path.basename(file_path))[0],
        'start_time': data_start_time,
        'end_time': data_end_time,
        'admission_time': admission_time,
        'discharge_time': discharge_time,
    }
    
    # Mode 7 specific note
    if selected_mode == 7 and 'duration_note' in locals() and duration_note:
         stats['Note'] = duration_note

    # Group 1: 基础指标组 (Basic Indicators Group) - 强制开启
    # 无论配置如何，始终计算基础指标
    stats.update(calc_basic_stats(df))

    # Group 2: 风险指标组 (Risk Indicators Group)
    if CALC_GROUPS.get(2, True):
        lbgi, hbgi, adrr = calc_lbgi_hbgi_adrr(df)
        modd = calc_modd(df)
        stats.update({
            'LBGI': round(lbgi, 4) if lbgi is not None else None,
            'HBGI': round(hbgi, 4) if hbgi is not None else None,
            'ADRR': round(adrr, 4) if adrr is not None else None,
            'MODD': modd
        })

    # Group 3: 血糖波动指标组 (MAGE, LAGE)
    if CALC_GROUPS.get(3, True):
        lage, mage = calc_lage_mage(df)
        stats.update({
            'LAGE': lage,
            'MAGE': mage
        })

    # Group 4: 血糖范围指标组 (Range Group)
    if CALC_GROUPS.get(4, True):
        # 1. 全天数据 (All Time)
        stats.update(calc_range_stats(df, prefix=''))
        
        # 2. 00:00 - 06:00 (Night/Early Morning)
        # 筛选 0点 <= 时间 < 6点 的数据
        df_0to6 = df[df['timestamp'].dt.hour < 6]
        stats.update(calc_range_stats(df_0to6, prefix='-0TO6AM'))
        
        # 3. 06:00 - 24:00 (Day/Evening)
        # 筛选 6点 <= 时间 <= 23点 的数据
        df_6to0 = df[df['timestamp'].dt.hour >= 6]
        stats.update(calc_range_stats(df_6to0, prefix='-6AMTO0'))

    # Group 5: Hourly Stats (0-6, 6-24)
    if CALC_GROUPS.get(5, True):
        period_00_06 = time_period_stats(df, 0, 5)
        period_06_24 = time_period_stats(df, 6, 23)
        
        # Mapping for 0-6AM
        stats['MEAN-0TO6AM'] = period_00_06['mean']
        stats['SD-0TO6AM'] = period_00_06['std']
        stats['CV-0TO6AM'] = period_00_06['cv']
        stats['VV-0TO6AM'] = period_00_06['vv_list']
        stats['VVtime-0TO6AM'] = period_00_06['vv_time_list']
        
        # Mapping for 6AM-0 (using -6AMTO0 suffix as requested)
        stats['MEAN-6AMTO0'] = period_06_24['mean']
        stats['SD-6AMTO0'] = period_06_24['std']
        stats['CV-6AMTO0'] = period_06_24['cv']
        # Note: VV/VVtime for 6AM-0 were not explicitly requested in the list but we can add them or skip
        # User list: SD-0TO6AM、SD-6AMTO0、MEAN-0TO6AM、MEAN-6AMTO、CV-0TO6AM、CV-6AMTO0、VV-0TO6AM、VVtime-0TO6AM、FBS-CV
        
        # FBS_CV (at 06:30)
        target_times = ['06:30:00']
        closest_time_stats = daily_closest_time_stats(df, target_times)
        stats['FBS-CV'] = closest_time_stats.get('063000_cv')

    # Group 6: Event Stats (Hypo/Hyper Events)
    if CALC_GROUPS.get(6, True):
        stats.update(calc_event_stats(df))

    # Group 7: Level 2 Hypo Events (<3.0, >15min)
    if CALC_GROUPS.get(7, True):
        # 使用 find_simple_events 计算
        # 1. 全天 2级低血糖
        l2_hypo_events = find_simple_events(
            df, 
            threshold=3.0, 
            compare_func=lambda x, th: x < th, 
            min_duration_min=15
        )
        
        # 格式化输出
        def fmt_events(evt_list):
            if not evt_list: return 0, None
            count = len(evt_list)
            time_strs = [f"{s.strftime('%Y-%m-%d %H:%M:%S')}~{e.strftime('%Y-%m-%d %H:%M:%S')}" for s, e in evt_list]
            return count, ",".join(time_strs)

        c, t = fmt_events(l2_hypo_events)
        stats['LV2 HYPO'] = c
        stats['Time-LV2 HYPO'] = t
        
        # 2. 夜间 2级低血糖 (Start in 0-6AM)
        l2_hypo_0to6 = [e for e in l2_hypo_events if 0 <= e[0].hour < 6]
        c_night, t_night = fmt_events(l2_hypo_0to6)
        stats['LV2 HYPO 0TO6AM'] = c_night
        stats['Time-LV2 HYPO 0TO6AM'] = t_night

    # Group 8: Conditional Hypo Events (Based on Global Min)
    if CALC_GROUPS.get(8, True):
        # 1. 计算全局最小值 (Global Minimum)
        if df.empty:
            global_min = None
        else:
            global_min = df['glucose'].min()

        # 2. 基础低血糖事件: 血糖 < 3.0, 持续 >= 15min
        # 如果 Group 7 已计算，直接复用，否则重算
        if 'l2_hypo_events' not in locals():
            l2_hypo_events = find_simple_events(
                df, 
                threshold=3.0, 
                compare_func=lambda x, th: x < th, 
                min_duration_min=15
            )

        # 辅助函数: 格式化输出 (Count, TimeString)
        def fmt_events(evt_list):
            if not evt_list: return 0, None
            count = len(evt_list)
            time_strs = [f"{s.strftime('%Y-%m-%d %H:%M:%S')}~{e.strftime('%Y-%m-%d %H:%M:%S')}" for s, e in evt_list]
            return count, ",".join(time_strs)

        # 辅助函数: 根据全局最小值判断输出
        # 如果 global_min < min_condition，则输出事件统计；否则输出 "#N/A"
        def get_conditional_stats(events, min_condition, is_night_only=False):
            if global_min is None or global_min >= min_condition:
                return "#N/A", "#N/A"
            
            # 筛选事件
            if is_night_only:
                filtered_events = [e for e in events if 0 <= e[0].hour < 6]
            else:
                filtered_events = events
                
            return fmt_events(filtered_events)

        # --- 指标 1-4: 前提是 全局最小值 < 3.0 ---
        # 1. 次数：低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3
        # 2. 出现时间：低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3
        c, t = get_conditional_stats(l2_hypo_events, min_condition=3.0, is_night_only=False)
        stats['HYPO_COND_3.0'] = c
        stats['Time-HYPO_COND_3.0'] = t

        # 3. 次数：夜间低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3
        # 4. 出现时间：夜间低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3
        c, t = get_conditional_stats(l2_hypo_events, min_condition=3.0, is_night_only=True)
        stats['HYPO_COND_3.0 0TO6AM'] = c
        stats['Time-HYPO_COND_3.0 0TO6AM'] = t

        # --- 指标 5-8: 前提是 全局最小值 < 3.5 ---
        # 5. 次数：低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3.5
        # 6. 出现时间：低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3.5
        c, t = get_conditional_stats(l2_hypo_events, min_condition=3.5, is_night_only=False)
        stats['HYPO_COND_3.5'] = c
        stats['Time-HYPO_COND_3.5'] = t

        # 7. 次数：夜间低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3.5
        # 8. 出现时间：低血糖事件（血糖<3持续超过15分钟）+血糖最低值<3.5 (Implies Night Time)
        c, t = get_conditional_stats(l2_hypo_events, min_condition=3.5, is_night_only=True)
        stats['HYPO_COND_3.5 0TO6AM'] = c
        stats['Time-HYPO_COND_3.5 0TO6AM'] = t

    # (10) Count A: Event(<3.0) + min<3.0 -> Merged to 8
    # (11) Night Count A -> Merged to 8
    # (12) Count B: Event(<3.0) + min<3.5 -> Merged to 8
    # (13) Night Count B -> Merged to 8

    return stats


def find_patient_file(device_id, folder_path):
    """
    在指定文件夹中查找包含设备ID的文件
    """
    # 检查是否为 NaN 值
    if pd.isna(device_id):
        return None  # 如果是 NaN，直接返回 None
    
    # 将设备ID转换为字符串
    if isinstance(device_id, float):
        # 浮点数先转为整数再转字符串（去除小数部分）
        device_id_str = str(int(device_id))
    else:
        # 其他类型直接转为字符串
        device_id_str = str(device_id)
    
    for filename in os.listdir(folder_path):
        # 使用转换后的字符串进行比较
        if device_id_str in filename and filename.endswith(('.xls', '.xlsx')):
            return os.path.join(folder_path, filename)
    return None


def main(selected_mode=None):
    # 使用命令行参数指定的模式，如果有的话
    current_mode = selected_mode if selected_mode is not None else mode
    
    # 生成简化的输出文件名
    output_file = get_output_filename(current_mode)

    
    # 读取患者列表
    try:
        patient_df = pd.read_excel(patient_list_file)
        # 确保列名正确，如果没有列名，则设置列名
        if len(patient_df.columns) >= 7:
            # Columns: Hospital ID, Pump Start, Pump End, Discharge, Admission, Sensor ID, Phone Number
            patient_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number'] + list(patient_df.columns[7:])
        elif len(patient_df.columns) >= 6:
             # Just in case phone is missing but we have sensor_id
            patient_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id'] + list(patient_df.columns[6:])
        elif len(patient_df.columns) >= 5:
            # Columns: Hospital ID, Pump Start, Pump End, Discharge, Admission
            patient_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time'] + list(patient_df.columns[5:])
        else:
            print("患者列表文件格式不正确，至少需要5列数据")
            return
    except Exception as e:
        print(f"无法读取患者列表文件 {patient_list_file}: {e}")
        return
    
    # 转换时间格式
    patient_df['admission_time'] = pd.to_datetime(patient_df['admission_time'], errors='coerce')
    patient_df['discharge_time'] = pd.to_datetime(patient_df['discharge_time'], errors='coerce')
    patient_df['pump_start_time'] = pd.to_datetime(patient_df['pump_start_time'], errors='coerce')
    patient_df['pump_end_time'] = pd.to_datetime(patient_df['pump_end_time'], errors='coerce')
    
    # 定义一个空列表，用于存储处理结果
    results = []
    
    # 遍历患者列表
    for index, row in patient_df.iterrows():
        hospital_id = row['hospital_id']
        admission_time = row['admission_time']
        discharge_time = row['discharge_time']
        pump_start_time = row['pump_start_time']
        pump_end_time = row['pump_end_time']
        
        # Determine matching key based on configuration
        match_key = None
        if MATCH_BY == 'sensor_id' and 'sensor_id' in row:
             match_key = row['sensor_id']
        elif MATCH_BY == 'phone_number' and 'phone_number' in row:
             match_key = row['phone_number']
        else:
             match_key = hospital_id
             
        if pd.isna(match_key):
             print(f"警告: 患者 {hospital_id} 的匹配键 {MATCH_BY} 为空，跳过匹配")
             continue

        # 查找患者数据文件
        # Use configured match_key
        patient_file = find_patient_file(match_key, data_folder)
        
        if patient_file:
            print(f"正在处理患者 {hospital_id} 的文件: {patient_file} (匹配键: {match_key})")
            # 处理文件，获取患者统计数据
            patient_stats = process_patient_file(
                patient_file, 
                admission_time, 
                discharge_time, 
                pump_start_time,
                pump_end_time,
                current_mode
            )
            
            # 如果患者统计数据不为空
            if patient_stats:
                # 将患者ID添加到统计数据中
                patient_stats['hospital_id'] = hospital_id
                # 将患者统计数据添加到结果列表中
                results.append(patient_stats)
            else:
                # 如果处理失败，添加基本信息
                results.append({
                    'hospital_id': hospital_id,
                    'patient_id': os.path.splitext(os.path.basename(patient_file))[0],
                    'admission_time': admission_time,
                    'discharge_time': discharge_time
                })
        else:
            print(f"未找到患者 {hospital_id} 的数据文件 (匹配键: {match_key})")
            # 添加基本信息到结果中
            results.append({
                'hospital_id': hospital_id,
                'patient_id': None,
                'admission_time': admission_time,
                'discharge_time': discharge_time
            })
    
    # 如果结果列表为空
    if not results:
        print("没有处理任何文件，结果为空")
        return
    
    # 将结果列表转换为DataFrame
    result_df = pd.DataFrame(results)
    
    # 调整列顺序，将 hospital_id 放在第一列
    cols = list(result_df.columns)
    if 'hospital_id' in cols:
        cols.insert(0, cols.pop(cols.index('hospital_id')))
        result_df = result_df[cols]
    
    # 定义时间列
    time_cols = ['start_time', 'end_time', 'admission_time', 'discharge_time']
    
    # 遍历时间列
    for col in time_cols:
        if col in result_df.columns:
            # 将时间列转换为datetime类型，无法转换的将变为NaT
            result_df[col] = pd.to_datetime(result_df[col], errors='coerce')
            # 对非NaT值进行格式化，NaT值转换为空字符串
            result_df[col] = result_df[col].apply(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if pd.notnull(x) else '')
    
    # 将结果保存为Excel文件
    result_df.to_excel(output_file, index=False)
    print(f"处理完成，结果已保存至：{output_file}")


if __name__ == "__main__":
    # 创建命令行参数解析器
    parser = argparse.ArgumentParser(description='CGM数据分析工具')
    parser.add_argument('-m', '--mode', type=int, choices=[0, 1, 2, 3, 4, 5, 6, 7], 
                        help='分析模式: 0(开始时间~开始时间+duringday天), 1(开始时间+interimday天~开始时间+duringday天), '
                             '2(开始时间~结束时间), 3(出院时间-duringday天~出院时间), 4(入院时间~出院时间), '
                             '5(出院时间~出院时间+duringday天), 6(出院时间~结束时间), 7(胰岛素泵使用时间)')
    parser.add_argument('-i', '--interim', type=int, help='中间天数(用于mode1)')
    parser.add_argument('-d', '--during', type=int, help='计算的时间段长度(天)')
    parser.add_argument('--nametag', type=str, help='Nametag for output filename')
    parser.add_argument('--datetag', type=str, help='Datetag for output filename')
    parser.add_argument('--patient_list', type=str, help='Patient list file path')
    parser.add_argument('--data_folder', type=str, help='Data folder path')
    parser.add_argument('--output_folder', type=str, help='Output folder path')
    parser.add_argument('--match_by', type=str, help='Match by: sensor_id, phone_number, or hospital_id')
    parser.add_argument('--daily_output', action='store_true', help='Enable daily output')
    parser.add_argument('--no_daily_output', action='store_false', dest='daily_output', help='Disable daily output')
    parser.add_argument('--calc_groups', type=str, help='JSON string for calc_groups')
    
    args = parser.parse_args()
    
    # 如果命令行指定了参数，则覆盖配置文件中的值
    selected_mode = args.mode
    if args.interim is not None:
# 此处不需要global声明，因为interimday已经是全局变量
        interimday = args.interim
    if args.during is not None:
# 此处不需要global声明，因为duringday已经是全局变量
        duringday = args.during
    if args.nametag is not None:
        nametag = args.nametag
    if args.datetag is not None:
        datetag = args.datetag
    if args.patient_list is not None:
        patient_list_file = args.patient_list
    if args.data_folder is not None:
        data_folder = args.data_folder
    if args.output_folder is not None:
        output_folder = args.output_folder
    if args.match_by is not None:
        MATCH_BY = args.match_by
    if args.daily_output is not None:
        daily_output = args.daily_output
    if args.calc_groups is not None:
        import json
        try:
            # Parse JSON string: e.g. '{"1": true, "2": false}'
            # Keys in JSON are always strings, need to convert to int if CALC_GROUPS uses int keys
            groups_dict = json.loads(args.calc_groups)
            # Convert keys to int
            CALC_GROUPS = {int(k): v for k, v in groups_dict.items()}
        except Exception as e:
            print(f"Error parsing --calc_groups: {e}")
    
    # 打印当前模式和参数
    mode_desc = {
        0: f"模式0: 开始时间 ~ 开始时间+{duringday}天",
        1: f"模式1: 开始时间+{interimday}天 ~ 开始时间+{duringday}天",
        2: "模式2: 开始时间 ~ 结束时间",
        3: f"模式3: 出院时间-{duringday}天 ~ 出院时间",
        4: "模式4: 入院时间 ~ 出院时间",
        5: f"模式5: 出院时间 ~ 出院时间+{duringday}天",
        6: "模式6: 出院时间 ~ 结束时间",
        7: "模式7: 胰岛素泵使用时间"
    }
    
    current_mode = selected_mode if selected_mode is not None else mode
    print(f"当前使用{mode_desc[current_mode]}")
    
    main(current_mode)
    