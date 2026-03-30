# 02_02_CGMDaysExtractor.py
# 作用：统计每个 CGM 血糖文件包含的总时长（小时）以及按 24h 为一自然日的天数（从数据开始时间起算）
# LAST UPDATE BY LIFANGU IN 20260325

import os
import pandas as pd
import numpy as np
from datetime import datetime, time, timedelta
import yaml

# --- 硬编码配置（可根据需要修改，或迁移到独立配置文件） ---
# nametag 用于拼接请求文件/数据目录名
# patient_list_file 为 request 列表文件（Excel），包含住院号、匹配键等基础信息
# data_folder 为已筛选后的 CGM 数据目录，文件名中应包含匹配键（探头号/手机号/住院号）
# output_folder 为输出目录
# datetag 用于输出文件名标识
nametag = "2412IN"
patient_list_file = fr"C:\\Users\\lifan\\Desktop\\03_Requests\\{nametag}.xlsx"
data_folder = fr"C:\\Users\\lifan\\Desktop\\04_SelectedData\\DataSelectedFor{nametag}"
output_folder = fr"C:\\Users\\lifan\\Desktop"
datetag = "DRQ260122Nova"

# --- 配置读取：优先使用当前目录下的 config.yaml；如无则尝试项目根目录的 config.yaml ---
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
# 从配置中读取 request_columns（列映射）和 match_by（匹配键类型）
REQUEST_COLUMNS = config_data.get('request_columns', {})
MATCH_BY = config_data.get('match_by', 'sensor_id')

# --- 工具函数：从 DataFrame 中按列名或列号（1-based）提取 Series ---
# 允许使用列名字符串或“1 开始计数”的整数来指定列，以兼容带表头/不带表头的 Excel
def _extract_request_series(df, spec):
    if spec is None:
        return None
    if isinstance(spec, str):
        s = spec.strip()
        if s in df.columns:
            return df[s]
        if s.isdigit():
            spec = int(s)
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

# --- 归一化 request 表：生成标准字段列 ---
# 标准字段与 01_02_Calculation.py 保持一致：
# hospital_id, pump_start_time, pump_end_time, discharge_time, admission_time, sensor_id, phone_number
# 优先使用 config.yaml 的 request_columns；否则尝试按列数量推断；若仍不足则补 NaN
def normalize_request_df(patient_df, request_columns):
    fields = [
        'hospital_id',
        'pump_start_time',
        'pump_end_time',
        'discharge_time',
        'admission_time',
        'sensor_id',
        'phone_number'
    ]
    if isinstance(request_columns, dict) and request_columns:
        out = patient_df.copy()
        for f in fields:
            s = _extract_request_series(patient_df, request_columns.get(f))
            if s is not None:
                out[f] = s
            elif f not in out.columns:
                out[f] = np.nan
        return out
    if set(fields).issubset(set(patient_df.columns)):
        return patient_df
    out = patient_df.copy()
    if len(out.columns) >= 7:
        out.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number'] + list(out.columns[7:])
    elif len(out.columns) >= 6:
        out.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id'] + list(out.columns[6:])
        if 'phone_number' not in out.columns:
            out['phone_number'] = np.nan
    elif len(out.columns) >= 5:
        out.columns = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time'] + list(out.columns[5:])
        if 'sensor_id' not in out.columns:
            out['sensor_id'] = np.nan
        if 'phone_number' not in out.columns:
            out['phone_number'] = np.nan
    else:
        out['hospital_id'] = out.iloc[:, 0] if len(out.columns) >= 1 else np.nan
        out['admission_time'] = np.nan
        out['discharge_time'] = np.nan
        out['pump_start_time'] = np.nan
        out['pump_end_time'] = np.nan
        out['sensor_id'] = np.nan
        out['phone_number'] = np.nan
    return out

# --- 输出文件名构造 ---
def get_output_filename():
    filename = f"CGM_{datetag}_{nametag}_Days_Extraction.xlsx"
    return os.path.join(output_folder, filename)

# --- 在数据目录中查找包含匹配键的 CGM 文件 ---
# 要求：文件名包含匹配键字符串，且扩展名为 .xls 或 .xlsx
def find_patient_file(device_id, folder_path):
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

# --- “四舍五入到最近整数”的辅助函数 ---
# numpy.round 在 .5 边界的行为可能与期望不同，这里用 floor(x + 0.5) 实现传统四舍五入
def round_half_up(x):
    return int(np.floor(x + 0.5))

# --- 从单个 CGM 文件中提取总小时数与天数 ---
# 规则：
# - 使用第一行之后的数据作为实际记录（与 02_01 保持一致，首行通常是表头）
# - 将 timestamp 转换为 datetime，取最小/最大时间
# - 总小时数 = (max - min) 的秒数 / 3600，并四舍五入为整数
# - 天数 = 总小时数 / 24，并四舍五入为整数
def extract_days_info(file_path):
    try:
        df = pd.read_excel(file_path, header=None)
    except Exception:
        return None, None
    if df.shape[0] <= 1:
        return None, None
    df = df.drop(0).reset_index(drop=True)
    df.columns = ['timestamp', 'glucose']
    try:
        df['timestamp'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%d %H:%M:%S')
    except Exception:
        return None, None
    start_time = df['timestamp'].min()
    end_time = df['timestamp'].max()
    if pd.isna(start_time) or pd.isna(end_time):
        return None, None
    hours = (end_time - start_time).total_seconds() / 3600.0
    total_hours = round_half_up(hours)
    total_days = round_half_up(hours / 24.0)
    return total_hours, total_days

# --- 主流程 ---
# 1) 读取 request 列表并归一化前 7 列基础信息
# 2) 逐行根据 match_by（来自 config.yaml）找到对应 CGM 文件
# 3) 对每个文件计算 TotalHours 与 TotalDays
# 4) 输出为 Excel：前 7 列基础信息 + 两个统计列
def main():
    output_file = get_output_filename()
    try:
        patient_df_raw = pd.read_excel(patient_list_file)
        patient_df = normalize_request_df(patient_df_raw, REQUEST_COLUMNS)
    except Exception as e:
        print(f"读取患者列表失败: {e}")
        return
    results = []
    base_fields = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number']
    print(f"开始处理，匹配依据: {MATCH_BY}")
    for index, row in patient_df.iterrows():
        # 采集基础信息（与主计算保持一致的前 7 列）
        base_info = [row.get(f, None) for f in base_fields]
        # 优先使用配置中指定的 MATCH_BY 列；若为空则回退到 hospital_id
        match_key = row.get(MATCH_BY, None) if MATCH_BY in patient_df.columns else None
        if pd.isna(match_key) or match_key is None:
            match_key = row.get('hospital_id', None)
        if pd.isna(match_key):
            # 匹配键缺失时，保留基础信息并填空的统计列
            results.append(base_info + [None, None])
            continue
        patient_file = find_patient_file(match_key, data_folder)
        total_hours, total_days = (None, None)
        if patient_file:
            print(f"正在处理: {match_key} -> {os.path.basename(patient_file)}")
            total_hours, total_days = extract_days_info(patient_file)
        else:
            print(f"未找到文件: {match_key}")
        results.append(base_info + [total_hours, total_days])
    # 构造输出 DataFrame（前 7 列 + 统计列）
    header_cols = base_fields + ['TotalHours', 'TotalDays']
    final_df = pd.DataFrame(results, columns=header_cols)
    final_df.to_excel(output_file, index=False)
    print(f"处理完成，结果已保存至：{output_file}")

if __name__ == "__main__":
    main()

