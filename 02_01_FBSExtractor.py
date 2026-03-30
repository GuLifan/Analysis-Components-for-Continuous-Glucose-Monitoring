# 01_03_FBS_Extractor.py
# 提取每天最接近06:30的血糖值
# LAST UPDATE BY LIFANGU IN 202602061920

import os
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import yaml

# -- 硬编码配置 ---
# 请在此处直接修改路径


nametag = "2505IN" 
patient_list_file = fr"C:\Users\lifan\Desktop\03_Requests\{nametag}.xlsx"  # 患者列表文件
data_folder = fr"C:\Users\lifan\Desktop\04_SelectedData\DataSelectedFor{nametag}" # 数据文件夹
output_folder = fr"C:\Users\lifan\Desktop"  # 输出文件夹路径

MATCH_BY = 'sensor_id' # 匹配键: 'sensor_id', 'phone_number', 'hospital_id'
datetag = "DRQ260122"
REQUEST_COLUMNS = {}


def load_config(config_path):
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}


_config_dir = os.path.dirname(__file__)
_config_candidates = [
    os.path.join(_config_dir, 'config.yaml'),
    os.path.abspath(os.path.join(_config_dir, '..', 'config.yaml'))
]

config_data = {}
for _p in _config_candidates:
    config_data = load_config(_p)
    if config_data:
        break
REQUEST_COLUMNS = config_data.get('request_columns', {})


def _extract_request_series(df, spec):
    if spec is None:
        return None

    if isinstance(spec, str):
        spec_str = spec.strip()
        if spec_str in df.columns:
            return df[spec_str]
        if spec_str.isdigit():
            spec = int(spec_str)
        else:
            return None

    if isinstance(spec, (int, np.integer)):
        idx = int(spec)
        if idx >= 1:
            idx = idx - 1
        if 0 <= idx < df.shape[1]:
            return df.iloc[:, idx]
        return None

    return None


def normalize_request_df(patient_df, request_columns):
    standard_fields = [
        'hospital_id',
        'discharge_time',
        'admission_time',
        'sensor_id',
        'phone_number',
        'pump_start_time',
        'pump_end_time'
    ]

    if isinstance(request_columns, dict) and request_columns:
        out_df = patient_df.copy()
        for field in standard_fields:
            series = _extract_request_series(patient_df, request_columns.get(field))
            if series is not None:
                out_df[field] = series
            elif field not in out_df.columns:
                out_df[field] = np.nan
        return out_df

    if set(standard_fields).issubset(set(patient_df.columns)):
        return patient_df

    out_df = patient_df.copy()
    if len(out_df.columns) >= 7:
        out_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number'] + list(out_df.columns[7:])
    elif len(out_df.columns) >= 6:
        out_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id'] + list(out_df.columns[6:])
        if 'phone_number' not in out_df.columns:
            out_df['phone_number'] = np.nan
    elif len(out_df.columns) >= 5:
        out_df.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time'] + list(out_df.columns[5:])
        if 'sensor_id' not in out_df.columns:
            out_df['sensor_id'] = np.nan
        if 'phone_number' not in out_df.columns:
            out_df['phone_number'] = np.nan
    else:
        out_df['hospital_id'] = out_df.iloc[:, 0] if len(out_df.columns) >= 1 else np.nan
        out_df['admission_time'] = np.nan
        out_df['discharge_time'] = np.nan
        out_df['pump_start_time'] = np.nan
        out_df['pump_end_time'] = np.nan
        out_df['sensor_id'] = np.nan
        out_df['phone_number'] = np.nan

    return out_df

def get_output_filename():
    # 自定义输出文件名
    filename = f"CGM_{datetag}_{nametag}_FBS_Extraction.xlsx"
    return os.path.join(output_folder, filename)

def find_patient_file(device_id, folder_path):
    """
    在指定文件夹中查找包含设备ID的文件
    """
    if pd.isna(device_id):
        return None
    
    if isinstance(device_id, float):
        device_id_str = str(int(device_id))
    else:
        device_id_str = str(device_id)
    
    for filename in os.listdir(folder_path):
        if device_id_str in filename and filename.endswith(('.xls', '.xlsx')):
            return os.path.join(folder_path, filename)
    return None

def extract_daily_fbs(file_path):
    """提取每天最接近 06:30 的血糖值"""
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
    
    # 目标时间 06:30:00
    target_time = time(6, 30, 00)
    
    daily_fbs = []
    
    # 按天分组
    # 使用 timestamp.date 进行分组
    df['date'] = df['timestamp'].dt.date
    grouped = df.groupby('date')
    
    # 排序日期以确保顺序
    sorted_dates = sorted(grouped.groups.keys())
    
    for date in sorted_dates:
        day_df = grouped.get_group(date).copy()
        
        # 计算与目标时间的时间差
        # 构造当天的目标时间 datetime 对象
        target_datetime = datetime.combine(date, target_time)
        
        # 计算绝对时间差
        day_df['time_diff'] = (day_df['timestamp'] - target_datetime).abs()
        
        # 找到最小时间差的行
        if not day_df.empty:
            min_diff_idx = day_df['time_diff'].idxmin()
            closest_row = day_df.loc[min_diff_idx]            
            daily_fbs.append(closest_row['glucose'])
        else:
            daily_fbs.append(None) # 应该不会发生，因为是按存在的日期group的
            
    return daily_fbs

def main():
    output_file = get_output_filename()
    
    # 读取患者列表
    try:
        patient_df_raw = pd.read_excel(patient_list_file)
        patient_df = normalize_request_df(patient_df_raw, REQUEST_COLUMNS)
    except Exception as e:
        print(f"无法读取患者列表文件 {patient_list_file}: {e}")
        return

    # 准备结果列表
    results = []
    
    base_fields = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number']
    print(f"开始处理，匹配依据: {MATCH_BY}")
    
    max_days = 0
    
    for index, row in patient_df.iterrows():
        base_info = [row.get(f, None) for f in base_fields]
        match_key = row.get(MATCH_BY, None) if MATCH_BY in patient_df.columns else None
        if pd.isna(match_key) or match_key is None:
            match_key = row.get('hospital_id', None)
            
        if pd.isna(match_key):
            print(f"警告: 第 {index+1} 行患者的匹配键 {MATCH_BY} 为空，跳过匹配")
            # 依然添加基础信息，后面补空
            results.append(base_info + [])
            continue

        # 查找文件
        patient_file = find_patient_file(match_key, data_folder)
        
        fbs_values = []
        if patient_file:
            print(f"正在处理: {match_key} -> {os.path.basename(patient_file)}")
            fbs_values = extract_daily_fbs(patient_file)
            if fbs_values:
                max_days = max(max_days, len(fbs_values))
        else:
            print(f"未找到文件: {match_key}")
            
        # 将基础信息和血糖值合并
        # 注意：这里 fbs_values 是一个列表，可能长度不一
        # 我们先存列表，最后统一对齐
        results.append(base_info + (fbs_values if fbs_values else []))

    # 构建最终 DataFrame
    # 1. 确定列名
    header_cols = base_fields
    # 后续列名为 Day 1, Day 2, ...
    fbs_cols = [f"Day {i+1}" for i in range(max_days)]
    final_cols = header_cols + fbs_cols
    
    # 2. 填充数据
    # results 是 list of lists。需要补齐长度
    padded_results = []
    for row_data in results:
        # 基础部分长度
        base_len = 7
        # 当前数据总长
        curr_len = len(row_data)
        # 需要的血糖数据长度 = max_days
        # 当前血糖数据长度 = curr_len - base_len
        fbs_part = row_data[base_len:]
        # 补齐 None
        padded_fbs = fbs_part + [None] * (max_days - len(fbs_part))
        padded_results.append(row_data[:base_len] + padded_fbs)
    final_df = pd.DataFrame(padded_results, columns=final_cols)
    
    # 保存
    final_df.to_excel(output_file, index=False)
    print(f"处理完成，结果已保存至：{output_file}")

if __name__ == "__main__":
    main()
