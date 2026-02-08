# 01_03_DailyCalculation.py
# 基于 01_02_Calculation.py 修改
# 功能：Mode 0 模式下，指定天数（duringday），输出每日详细指标（列展开）
# LAST UPDATE BY LIFANGU IN 202602090032

import os
import pandas as pd
import math
from datetime import datetime, time, timedelta
import numpy as np
import argparse
import json

# --- 1. 内嵌配置 (Hardcoded Configuration) ---

# 模式设置 (强制 Mode 0)
MODE = 0

# 日期参数
DURING_DAY = 14  # 计算范围（天数），例如 14 天，科内14天，科外3天
INTERIM_DAY = 0 # Mode 0 不使用

# 路径设置
PATIENT_LIST_FILE = r"C:\Users\lifan\Desktop\03_Requests\DataRequest2505IN.xlsx"
DATA_FOLDER = r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2505IN"
OUTPUT_FOLDER = r"C:\Users\lifan\Desktop\00_Outputs"

# 标签设置
DATETAG = "DRQ260122"
NAMETAG = "2505IN_Daily_Day1-14"

# 匹配设置
MATCH_BY = 'sensor_id' # 'sensor_id', 'phone_number', 'hospital_id'

# 计算开关 (Feature Toggles)
CALC_GROUPS = {
    1: True,   # Basic Stats
    2: True,   # Risk Stats (LBGI/HBGI/ADRR)
    3: True,   # Variability (MAGE/LAGE)
    4: True,   # Ranges (TIR/TAR/TBR)
    5: True,   # Hourly Stats
    6: True,   # Events (Hypo/Hyper)
    7: True,   # Level 2 Hypo
    8: True    # Conditional Hypo
}

# 每天最少需要点数
MIN_LEN_NUM_PER_DAY = 12 * 24

# --- 2. 核心计算函数 (复用自 01_02_Calculation.py) ---

def calc_basic_stats(df):
    if df.empty: return {}
    glucose_series = df['glucose']
    mean_val = glucose_series.mean()
    std_val = glucose_series.std()
    if mean_val != 0 and not np.isnan(mean_val):
        cv_val = round(std_val / mean_val, 4)
    else:
        cv_val = None
    if not np.isnan(mean_val):
        gmi_val = round(3.31 + 0.02392 * 18 * mean_val, 4)
    else:
        gmi_val = None

    return {
        'MEAN': round(mean_val, 4) if not np.isnan(mean_val) else None,
        'SD': round(std_val, 4) if not np.isnan(std_val) else None,
        'CV': cv_val,
        'GMI': gmi_val
    }

def calc_lbgi_hbgi_adrr(df, daily_output=False):
    # Modified to accept daily_output arg, though we usually call it with False for single day
    if df.empty:
        return ([], [], []) if daily_output else (None, None, None)
    
    df = df.sort_values('timestamp').copy()
    start_time = df['timestamp'].iloc[0]
    df['day_idx'] = ((df['timestamp'] - start_time).dt.total_seconds() // 86400).astype(int)
    
    # 注意：如果是单日计算，day_idx 都是 0。
    # 为了防止 "过滤掉不足24h的最后一天" 逻辑误杀单日数据，
    # 我们检查数据跨度。如果只有1天数据，我们假设它是有效的（由外部逻辑保证长度）。

    valid_df = df.copy()

    valid_df = valid_df[valid_df['glucose'] >= 1.0].copy()
    
    if valid_df.empty:
        return ([], [], []) if daily_output else (None, None, None)

    # Formula: 1.794 * (ln(bg)^1.026 - 1.861)  (Input: mmol/L)
    valid_df['fBG'] = 1.794 * (np.log(valid_df['glucose']) ** 1.026 - 1.861)
    valid_df['risk'] = 10 * (valid_df['fBG'] ** 2)
    
    valid_df['rl'] = np.where(valid_df['fBG'] < 0, valid_df['risk'], 0)
    valid_df['rh'] = np.where(valid_df['fBG'] > 0, valid_df['risk'], 0)
    
    # Aggregation
    daily_stats = valid_df.groupby('day_idx').agg(
        mean_rl=('rl', 'mean'),
        mean_rh=('rh', 'mean'),
        max_rl=('rl', 'max'),
        max_rh=('rh', 'max')
    )
    daily_stats['daily_adrr'] = daily_stats['max_rl'] + daily_stats['max_rh']
    
    final_lbgi = daily_stats['mean_rl'].mean()
    final_hbgi = daily_stats['mean_rh'].mean()
    final_adrr = daily_stats['daily_adrr'].mean()
    
    return round(final_lbgi, 4), round(final_hbgi, 4), round(final_adrr, 4)

def calc_daily_modd(current_day_df, prev_day_df):
    """计算单日 MODD (需提供前一日数据)"""
    if prev_day_df is None or prev_day_df.empty or current_day_df.empty:
        return None
        
    # 获取当天和前一天的唯一时间点
    curr_times = pd.unique(current_day_df['timestamp'].dt.time)
    prev_times = pd.unique(prev_day_df['timestamp'].dt.time)

    # 找到共同的时间点
    common_times = set(curr_times) & set(prev_times)
    
    daily_differences = []
    for time_point in common_times:
        curr_value = current_day_df[current_day_df['timestamp'].dt.time == time_point]['glucose'].values
        prev_value = prev_day_df[prev_day_df['timestamp'].dt.time == time_point]['glucose'].values

        if len(curr_value) > 0 and len(prev_value) > 0:
            daily_differences.append(abs(curr_value[0] - prev_value[0]))

    if daily_differences:
        return round(pd.Series(daily_differences).mean(), 4)
    else:
        return None

def calc_mage_daily(glucose_series):
    data = glucose_series.dropna().values
    if len(data) < 3: return None
    sd = np.std(data, ddof=1)
    if sd == 0: return 0.0
    
    peaks = []
    nadirs = []
    for i in range(1, len(data) - 1):
        if data[i] > data[i-1] and data[i] > data[i+1]:
            peaks.append((i, data[i]))
        elif data[i] < data[i-1] and data[i] < data[i+1]:
            nadirs.append((i, data[i]))
            
    if not peaks or not nadirs: return None
    turning_points = sorted(peaks + nadirs, key=lambda x: x[0])
    
    first_valid_direction = None
    mage_sum = 0
    mage_count = 0
    
    for i in range(1, len(turning_points)):
        current_val = turning_points[i][1]
        prev_val = turning_points[i-1][1]
        diff = current_val - prev_val
        amplitude = abs(diff)
        
        if amplitude > sd:
            direction = 1 if diff > 0 else -1
            if first_valid_direction is None:
                first_valid_direction = direction
                mage_sum += amplitude
                mage_count += 1
            elif direction == first_valid_direction:
                mage_sum += amplitude
                mage_count += 1
                
    if mage_count == 0: return None
    return mage_sum / mage_count

def calc_lage_mage(df, daily_output=False):
    if df.empty: return (None, None)
    # 单日计算
    glucose_series = df['glucose']
    if len(glucose_series.dropna()) < 144:
        return None, None
        
    lage = glucose_series.max() - glucose_series.min()
    mage = calc_mage_daily(glucose_series)
    
    return round(lage, 4) if lage is not None else None, round(mage, 4) if mage is not None else None

def calc_range_stats(df, prefix=''):
    if df.empty: return {}
    total = len(df)
    g = df['glucose'].values
    res = {}
    def calc_ratio(count):
        return round(count / total, 4) if total > 0 else 0

    res[f'TIR{prefix}'] = calc_ratio(np.sum((g >= 3.9) & (g <= 10.0)))
    res[f'TAR{prefix}'] = calc_ratio(np.sum(g > 10.0))
    res[f'TBR{prefix}'] = calc_ratio(np.sum(g < 3.9))
    res[f'TAR1{prefix}'] = calc_ratio(np.sum((g > 10.0) & (g <= 13.9)))
    res[f'TAR2{prefix}'] = calc_ratio(np.sum(g > 13.9))
    res[f'TBR1{prefix}'] = calc_ratio(np.sum((g >= 3.0) & (g < 3.9)))
    res[f'TBR2{prefix}'] = calc_ratio(np.sum(g < 3.0))
    res[f'TITR{prefix}'] = calc_ratio(np.sum((g >= 3.9) & (g <= 7.8)))
    
    if f'TIR{prefix}' in res and f'TITR{prefix}' in res:
            res[f'TIR-TITR{prefix}'] = round(res[f'TIR{prefix}'] - res[f'TITR{prefix}'], 4)
    return res

def find_simple_events(df, threshold, compare_func, min_duration_min=15):
    if df.empty: return []
    times = df['timestamp'].tolist()
    values = df['glucose'].tolist()
    events = []
    in_event = False
    start_time = None
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
                prev_t = current_event_times[-1]
                if (t - prev_t).total_seconds() > 15 * 60:
                     duration = (current_event_times[-1] - start_time).total_seconds() / 60
                     if duration >= min_duration_min:
                         events.append((start_time, current_event_times[-1]))
                     start_time = t
                     current_event_times = [t]
                else:
                    current_event_times.append(t)
        else:
            if in_event:
                duration = (current_event_times[-1] - start_time).total_seconds() / 60
                if duration >= min_duration_min:
                    events.append((start_time, current_event_times[-1]))
                in_event = False
                current_event_times = []
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
            end_idx = -1
            j = start_idx + 1
            while j < n:
                # 调用回调检查
                is_terminated, advance_steps = end_condition_func(values, times, j)
                
                if is_terminated:
                    end_idx = j
                    break
                
                # 如果数据断裂太长(如 > 30min)，强制终止
                if j > start_idx and (times[j] - times[j-1]).total_seconds() > 30 * 60:
                    end_idx = j 
                    break
                    
                j += 1
            
            if end_idx == -1:
                end_idx = n
            
            if end_idx < n:
                event_end_time = times[end_idx]
            else:
                event_end_time = times[n-1]
                
            duration = (event_end_time - start_time).total_seconds() / 60
            
            if duration >= min_event_duration:
                events.append((start_time, event_end_time))
            
            i = end_idx
        else:
            i += 1
            
    return events

def calc_event_stats(df):
    stats = {}
    if df.empty: return stats
    df = df.sort_values('timestamp').copy()
    
    def fmt_events(evt_list):
        if not evt_list: return 0, None
        count = len(evt_list)
        time_strs = [f"{s.strftime('%Y-%m-%d %H:%M:%S')}~{e.strftime('%Y-%m-%d %H:%M:%S')}" for s, e in evt_list]
        return count, ",".join(time_strs)

    # 1. Hypoglycemic events (<3.9, >=15min)
    hypo_events = find_simple_events(df, 3.9, lambda x, th: x < th, 15)
    c, t = fmt_events(hypo_events)
    stats['HYPO'] = c
    stats['Time-HYPO'] = t
    
    # 2. HYPO 0TO6AM
    hypo_0to6 = [e for e in hypo_events if 0 <= e[0].hour < 6]
    c, t = fmt_events(hypo_0to6)
    stats['HYPO 0TO6AM'] = c
    stats['Time-HYPO 0TO6AM'] = t
    
    # 3. Extended Hypo (<3.9, >120min, End >=3.9 for 15min)
    def check_hypo_recovery(vals, ts, idx):
        if idx >= len(vals): return False, 0
        start_t = ts[idx]
        curr_idx = idx
        while curr_idx < len(vals):
            if vals[curr_idx] < 3.9: return False, 0
            span = (ts[curr_idx] - start_t).total_seconds() / 60
            if span >= 15: return True, 0
            curr_idx += 1
        return False, 0
        
    ex_hypo_events = find_complex_events(df, lambda x: x < 3.9, check_hypo_recovery, 120)
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
            if vals[curr_idx] > 10.0: return False, 0
            span = (ts[curr_idx] - start_t).total_seconds() / 60
            if span >= 15: return True, 0
            curr_idx += 1
        return False, 0

    ex_hyper_events = find_complex_events(df, lambda x: x > 13.9, check_hyper_recovery, 120)
    c, t = fmt_events(ex_hyper_events)
    stats['EX HYPER'] = c
    stats['Time-EX HYPER'] = t
    
    return stats

def time_period_stats(df, start_hour, end_hour):
    start_time = datetime.strptime(f"{start_hour:02d}:00:00", "%H:%M:%S").time()
    end_time = datetime.strptime(f"{end_hour:02d}:59:59", "%H:%M:%S").time()
    mask = (df['timestamp'].dt.time >= start_time) & (df['timestamp'].dt.time <= end_time)
    period_df = df[mask].copy()
    
    if period_df.empty:
        return {'mean': None, 'std': None, 'cv': None, 'vv_list': None, 'vv_time_list': None}
        
    glucose_series = period_df['glucose'].dropna()
    if glucose_series.empty:
        return {'mean': None, 'std': None, 'cv': None, 'vv_list': None, 'vv_time_list': None}

    mean = glucose_series.mean()
    std = glucose_series.std()
    cv = std / mean if (mean != 0 and not np.isnan(mean)) else None

    # Daily mins
    period_df['date'] = period_df['timestamp'].dt.date
    daily_groups = period_df.groupby('date')
    daily_mins = []
    daily_min_times = []
    for date, group in daily_groups:
        valid_group = group.dropna(subset=['glucose'])
        if valid_group.empty: continue
        min_value = valid_group['glucose'].min()
        min_rows = valid_group[valid_group['glucose'] == min_value]
        min_time = min_rows['timestamp'].iloc[0].strftime('%H:%M:%S')
        daily_mins.append(round(min_value, 4))
        daily_min_times.append(min_time)
        
    daily_min_str = ','.join(map(str, daily_mins)) if daily_mins else None
    daily_min_time_str = ','.join(daily_min_times) if daily_min_times else None

    return {
        'mean': round(mean, 4),
        'std': round(std, 4),
        'cv': round(cv, 4) if cv else None,
        'vv_list': daily_min_str,
        'vv_time_list': daily_min_time_str
    }

def find_patient_file(device_id, folder_path):
    if pd.isna(device_id): return None
    if isinstance(device_id, float): device_id_str = str(int(device_id))
    else: device_id_str = str(device_id)
    for filename in os.listdir(folder_path):
        if device_id_str in filename and filename.endswith(('.xls', '.xlsx')):
            return os.path.join(folder_path, filename)
    return None

# --- 3. 核心流程 ---

def process_patient_daily_mode0(file_path, admission_time=None, discharge_time=None):
    try:
        df = pd.read_excel(file_path, header=None)
    except:
        return None
        
    df = df.drop(0).reset_index(drop=True)
    df.columns = ['timestamp', 'glucose']
    df['timestamp'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%d %H:%M:%S', errors='coerce')
    df['glucose'] = pd.to_numeric(df['glucose'], errors='coerce')
    df['glucose'] = df['glucose'].clip(1.8, 33.3)
    
    data_start = df['timestamp'].min()
    data_end = df['timestamp'].max()
    
    # Mode 0: Start from first day's midnight (if possible) or just start_time?
    # User request: "日期切分从第1个自然日最接近零点开始切分"
    # Logic: Start time should be the first day's 00:00:00 (if data covers it)
    # or align to the first available data point's date at 00:00:00?
    # Typically, Mode 0 starts at "Start Time" (admission/device start).
    # If we want natural days, we should align start_time to data_start's Date 00:00.
    
    # Align to midnight of the first data day
    # But wait, if admission is 14:00, do we include 00:00-14:00 (missing)?
    # Or start from next day?
    # "最接近零点" -> Likely means align to the 00:00 of the starting date.
    
    start_date = data_start.normalize()
    start_time = start_date # 00:00:00
    
    end_time = start_time + timedelta(days=DURING_DAY)
    
    # Filter global range
    # Note: data might start at 14:00, so Day 1 (00:00-24:00) will be partial.
    df = df[(df['timestamp'] >= start_time) & (df['timestamp'] < end_time)]
    if df.empty: return None
    
    daily_results = []
    prev_day_df = None
    
    # Calculate Global Min for Group 8
    global_min = df['glucose'].min() if not df.empty else None
    
    for day in range(DURING_DAY):
        day_start = start_time + timedelta(days=day)
        day_end = day_start + timedelta(days=1)
        
        day_df = df[(df['timestamp'] >= day_start) & (df['timestamp'] < day_end)].copy()
        
        # Check data sufficiency (half day)
        # For Day 1, it might be partial if data starts late.
        # User implies we want natural day alignment.
        if day_df.empty or len(day_df) < MIN_LEN_NUM_PER_DAY // 2:
            daily_results.append(None)
            prev_day_df = day_df # Update prev even if empty
            continue
            
        stats = {}
        # Group 1
        stats.update(calc_basic_stats(day_df))
        
        # Group 2 (Risk)
        if CALC_GROUPS.get(2, True):
            lbgi, hbgi, adrr = calc_lbgi_hbgi_adrr(day_df)
            modd = calc_daily_modd(day_df, prev_day_df)
            stats.update({'LBGI': lbgi, 'HBGI': hbgi, 'ADRR': adrr, 'MODD': modd})
            
        # Group 3 (Var)
        if CALC_GROUPS.get(3, True):
            lage, mage = calc_lage_mage(day_df)
            stats.update({'LAGE': lage, 'MAGE': mage})
            
        # Group 4 (Range)
        if CALC_GROUPS.get(4, True):
            stats.update(calc_range_stats(day_df, prefix=''))
            df_0to6 = day_df[day_df['timestamp'].dt.hour < 6]
            stats.update(calc_range_stats(df_0to6, prefix='-0TO6AM'))
            df_6to0 = day_df[day_df['timestamp'].dt.hour >= 6]
            stats.update(calc_range_stats(df_6to0, prefix='-6AMTO0'))
            
        # Group 5 (Hourly)
        if CALC_GROUPS.get(5, True):
            p_0to6 = time_period_stats(day_df, 0, 5)
            p_6to24 = time_period_stats(day_df, 6, 23)
            stats['MEAN-0TO6AM'] = p_0to6['mean']
            stats['SD-0TO6AM'] = p_0to6['std']
            stats['CV-0TO6AM'] = p_0to6['cv']
            stats['VV-0TO6AM'] = p_0to6['vv_list']
            stats['VVtime-0TO6AM'] = p_0to6['vv_time_list']
            stats['MEAN-6AMTO0'] = p_6to24['mean']
            stats['SD-6AMTO0'] = p_6to24['std']
            stats['CV-6AMTO0'] = p_6to24['cv']

            
        # Group 6 (Event)
        if CALC_GROUPS.get(6, True):
            evt = calc_event_stats(day_df)
            stats.update(evt)
            
        # Group 7 (Level 2 Hypo)
        if CALC_GROUPS.get(7, True):
            l2_events = find_simple_events(day_df, 3.0, lambda x, th: x < th, 15)
            
            def fmt_events(evt_list):
                if not evt_list: return 0, None
                count = len(evt_list)
                time_strs = [f"{s.strftime('%Y-%m-%d %H:%M:%S')}~{e.strftime('%Y-%m-%d %H:%M:%S')}" for s, e in evt_list]
                return count, ",".join(time_strs)
                
            c, t = fmt_events(l2_events)
            stats['LV2 HYPO'] = c
            stats['Time-LV2 HYPO'] = t
            
            l2_0to6 = [e for e in l2_events if 0 <= e[0].hour < 6]
            c_night, t_night = fmt_events(l2_0to6)
            stats['LV2 HYPO 0TO6AM'] = c_night
            stats['Time-LV2 HYPO 0TO6AM'] = t_night
            
        # Group 8 (Conditional Hypo)
        if CALC_GROUPS.get(8, True):
            # Using day_df specific events but Global Min?
            # User requirement: "Group 8 metrics... require Global Min check first"
            # And "Based on Global Min"
            # But calculated daily?
            # Usually Group 8 is a global metric. But if we split by day:
            # We should probably use the DAILY min or GLOBAL min?
            # Re-reading prompt: "你的日期切分...也方便你补全缺失的功能组"
            # "Group 8 metrics... output '#N/A' if global min >= threshold"
            # So the condition is Global Min. The count is Daily Events.
            
            if 'l2_events' not in locals():
                l2_events = find_simple_events(day_df, 3.0, lambda x, th: x < th, 15)
                
            def get_cond_stats(events, min_cond, night_only=False):
                # Use GLOBAL min for the condition
                if global_min is None or global_min >= min_cond:
                    return "#N/A", "#N/A"
                
                evts = [e for e in events if 0 <= e[0].hour < 6] if night_only else events
                return fmt_events(evts)
                
            c, t = get_cond_stats(l2_events, 3.0, False)
            stats['HYPO_COND_3.0'] = c
            stats['Time-HYPO_COND_3.0'] = t
            
            c, t = get_cond_stats(l2_events, 3.0, True)
            stats['HYPO_COND_3.0 0TO6AM'] = c
            stats['Time-HYPO_COND_3.0 0TO6AM'] = t
            
            c, t = get_cond_stats(l2_events, 3.5, False)
            stats['HYPO_COND_3.5'] = c
            stats['Time-HYPO_COND_3.5'] = t
            
            c, t = get_cond_stats(l2_events, 3.5, True)
            stats['HYPO_COND_3.5 0TO6AM'] = c
            stats['Time-HYPO_COND_3.5 0TO6AM'] = t
            
        daily_results.append(stats)
        prev_day_df = day_df
        
    # Return structured result: { 'info': {...}, 'days': [day1_stats, day2_stats, ...] }
    result_structure = {
        'info': {
            'patient_id': os.path.splitext(os.path.basename(file_path))[0],
            'start_time': start_time,
            'end_time': end_time,
            'admission_time': admission_time,
            'discharge_time': discharge_time
        },
        'days': daily_results
    }
    return result_structure

def main():
    print(f"开始计算... Mode={MODE}, During={DURING_DAY}")
    print(f"输出文件模板: CGM_{DATETAG}_{NAMETAG}_Mode{MODE}_Results.xlsx")
    
    try:
        patient_df = pd.read_excel(PATIENT_LIST_FILE)
        # 尝试匹配列名结构
        if len(patient_df.columns) >= 7:
            # 标准结构: Hospital ID, Pump Start, Pump End, Discharge, Admission, Sensor ID, Phone Number
            patient_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number'] + list(patient_df.columns[7:])
        elif len(patient_df.columns) >= 5:
            # 简化结构: Hospital ID, Pump Start, Pump End, Discharge, Admission
             patient_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time'] + list(patient_df.columns[5:])
    except Exception as e:
        print(f"读取患者列表失败: {e}")
        return

    # Store all patient data
    all_patients_data = []
    
    for index, row in patient_df.iterrows():
        hospital_id = row['hospital_id']
        admission_time = row['admission_time'] if 'admission_time' in row else None
        discharge_time = row['discharge_time'] if 'discharge_time' in row else None
        
        # Determine match_val
        match_val = None
        if MATCH_BY in row and pd.notna(row[MATCH_BY]): 
            match_val = row[MATCH_BY]
        elif 'sensor_id' in row and pd.notna(row['sensor_id']): 
            match_val = row['sensor_id']
        elif 'phone_number' in row and pd.notna(row['phone_number']): 
            match_val = row['phone_number']
        else: 
            match_val = hospital_id
        
        patient_file = find_patient_file(match_val, DATA_FOLDER)
        
        if patient_file:
            print(f"Processing {hospital_id}...")
            res = process_patient_daily_mode0(patient_file, admission_time, discharge_time)
            if res:
                res['info']['hospital_id'] = hospital_id
                all_patients_data.append(res)
        else:
            print(f"File not found for {hospital_id} (Match: {match_val})")

    if not all_patients_data:
        print("No results generated.")
        return

    # Create Output Excel
    out_path = os.path.join(OUTPUT_FOLDER, f"CGM_{DATETAG}_{NAMETAG}_Mode{MODE}_Results.xlsx")
    
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        # Create Daily Sheets
        for day_idx in range(DURING_DAY):
            sheet_name = f"Day {day_idx + 1}"
            day_data_list = []
            
            for p in all_patients_data:
                # Basic Info + Day Stats
                row_data = p['info'].copy()
                
                # Format Dates in Info
                for key in ['start_time', 'end_time', 'admission_time', 'discharge_time']:
                    if key in row_data and isinstance(row_data[key], (datetime, pd.Timestamp)):
                        row_data[key] = row_data[key].strftime('%Y-%m-%d %H:%M:%S')
                
                # Get stats for this day
                if day_idx < len(p['days']):
                    day_stats = p['days'][day_idx]
                    if day_stats:
                        row_data.update(day_stats)
                    else:
                        # Day missing or insufficient data
                        pass 
                
                day_data_list.append(row_data)
            
            if day_data_list:
                df_day = pd.DataFrame(day_data_list)
                
                # Reorder columns: identifiers first
                cols = list(df_day.columns)
                
                # Desired order for first few columns (matching 01_02 usually)
                # hospital_id, patient_id, admission_time, discharge_time, start_time, end_time
                priority_cols = ['hospital_id', 'patient_id', 'admission_time', 'discharge_time', 'start_time', 'end_time']
                
                first_cols = []
                for c in priority_cols:
                    if c in cols:
                        first_cols.append(c)
                        cols.remove(c)
                
                # Reconstruct column list
                final_cols = first_cols + cols
                
                df_day = df_day[final_cols]
                
                df_day.to_excel(writer, sheet_name=sheet_name, index=False)
                
    print(f"Done! Saved to {out_path}")

if __name__ == "__main__":
    main()
