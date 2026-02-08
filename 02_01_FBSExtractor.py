# 01_03_FBS_Extractor.py
# 提取每天最接近06:30的血糖值
# LAST UPDATE BY LIFANGU IN 202602061920

import os
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta

# -- 硬编码配置 ---
# 请在此处直接修改路径


nametag = "2407IN" 
patient_list_file = fr"C:\Users\lifan\Desktop\03_Requests\DataRequest{nametag}.xlsx"  # 患者列表文件
data_folder = fr"C:\Users\lifan\Desktop\04_SelectedData\DataSelected{nametag}" # 数据文件夹
output_folder = fr"C:\Users\lifan\Desktop"  # 输出文件夹路径
MATCH_BY = 'phone_number' # 匹配键: 'sensor_id', 'phone_number', 'hospital_id'
datetag = "DRQ260122"

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
        patient_df = pd.read_excel(patient_list_file)
        
        # 规范化列名 (与 01_02_Calculation.py 保持一致)
        # 假设前7列是固定的元数据
        if len(patient_df.columns) >= 7:
            base_cols = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number']
            # 保留原始列名还是重命名？为了处理方便，我们重命名，但在输出时可以考虑还原，或者就用规范名。
            # 题目要求：文档的前七列...一致。
            # 所以我们应该保留原始 DataFrame 的前7列。
            pass 
        else:
            print("患者列表文件格式不正确，至少需要7列数据")
            # 尝试继续，只要能找到匹配键
    except Exception as e:
        print(f"无法读取患者列表文件 {patient_list_file}: {e}")
        return

    # 准备结果列表
    results = []
    
    # 根据 01_02_Calculation.py 的逻辑：
    # Hospital ID (0), Pump Start (1), Pump End (2), Discharge (3), Admission (4), Sensor ID (5), Phone Number (6)
    # 使用 iloc 来获取这些列，避免列名依赖
    # 但为了匹配，我们需要知道哪一列是 sensor_id 等。
    # 假设列顺序固定：
    # 0: hospital_id
    # 5: sensor_id
    # 6: phone_number
    col_map = {
        'hospital_id': 0,
        'sensor_id': 5,
        'phone_number': 6
    }
    
    match_col_idx = col_map.get(MATCH_BY, 0)
    print(f"开始处理，匹配依据: {MATCH_BY} (第 {match_col_idx+1} 列)")
    
    max_days = 0
    
    for index, row in patient_df.iterrows():
        # 获取前7列数据作为基础信息
        base_info = row.iloc[:7].tolist()
        
        # 获取匹配键的值
        if match_col_idx < len(row):
            match_key = row.iloc[match_col_idx]
        else:
            match_key = None
            
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
    # 前7列使用原文件的列名
    header_cols = patient_df.columns[:7].tolist()
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
