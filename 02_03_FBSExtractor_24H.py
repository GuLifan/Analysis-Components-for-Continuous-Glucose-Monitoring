#
# 02_02_FBSExtractor_24H.py
# 作用：
#   - 对每个 CGM 血糖文件，按“从数据开始时间起算”的连续 24h 窗口切分（而不是按自然日 date）
#   - 每个 24h 窗口内，寻找该窗口内“最接近 FBS 时刻（默认 06:30）”的血糖值
#   - 固定输出上限 14 天（Day 1 ~ Day 14），不会出现 Day 15
# LAST UPDATE BY LIFANGU IN 20260330
#
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
nametag = "2505IN"
patient_list_file = fr"C:\Users\lifan\Desktop\03_Requests\{nametag}.xlsx"
data_folder = fr"C:\Users\lifan\Desktop\04_SelectedData\DataSelectedFor{nametag}"
output_folder = fr"C:\Users\lifan\Desktop"

# 匹配键列名：与 request 表中的列一致（例如 sensor_id / phone_number / hospital_id）
# 注：该脚本与 02_02_CGMDaysExtractor.py 略有不同；这里 MATCH_BY 为硬编码，
# 如需统一由 config.yaml 控制，可参照 02_02_CGMDaysExtractor.py 的写法改造。
MATCH_BY = 'sensor_id'
datetag = "DRQ260122"
REQUEST_COLUMNS = {}

# 输出上限：最多输出 14 个 24h 窗口（Day 1 ~ Day 14）
MAX_DAYS = 14
# FBS 目标时刻（“空腹血糖”通常取清晨固定时刻；这里默认 06:30）
FBS_CLOCK_TIME = time(6, 30, 0)


# --- 配置读取：优先使用当前目录下的 config.yaml；如无则尝试项目根目录的 config.yaml ---
# 用途：读取 request_columns（列映射），用于兼容不同 request Excel 的列名/列序
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


# --- 工具函数：从 DataFrame 中按列名或列号（1-based）提取 Series ---
# 允许使用列名字符串或“1 开始计数”的整数来指定列，以兼容带表头/不带表头的 Excel
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


# --- 归一化 request 表：生成标准字段列 ---
# 标准字段与主计算脚本保持一致：
# hospital_id, pump_start_time, pump_end_time, discharge_time, admission_time, sensor_id, phone_number
# 优先使用 config.yaml 的 request_columns；否则尝试按列数量推断；若仍不足则补 NaN
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


# --- 输出文件名构造 ---
def get_output_filename():
    filename = f"CGM_{datetag}_{nametag}_FBS_Extraction_24H_Max14.xlsx"
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


# --- 将“窗口起点”映射到“窗口内应选取的 FBS 目标时刻” ---
# 规则：
# - 目标时刻为当天的 clock_t（默认 06:30）
# - 如果窗口起点 start_dt 已经过了当天的目标时刻，则目标时刻顺延到下一天的 clock_t
# 举例：
# - start_dt = 2026-03-01 02:00，则目标时刻 = 2026-03-01 06:30
# - start_dt = 2026-03-01 10:00，则目标时刻 = 2026-03-02 06:30
def _next_clock_time_on_or_after(start_dt, clock_t):
    candidate = datetime.combine(start_dt.date(), clock_t)
    if candidate < start_dt:
        candidate = candidate + timedelta(days=1)
    return candidate


# --- 从单个 CGM 文件中提取“按连续 24h 分段”的 FBS 值列表（最多 max_days 个）---
# 与旧版“按自然日 date 分组”的核心区别：
# - 旧版：timestamp.dt.date 分组，跨日边界会导致“14x24h 覆盖 15 个日期”的情况出现
# - 新版：以 t0 为起点每 24h 一段，输出段数与持续时长一致；并且强制最多 14 段
#
# 处理细节：
# - 读取 Excel：默认无表头（header=None），并丢弃首行（与本项目其它脚本保持一致）
# - 时间解析：要求格式为 '%Y-%m-%d %H:%M:%S'
# - 每段窗口：[t0 + i*24h, t0 + (i+1)*24h)
# - 每段目标时刻：取“窗口起点之后的下一次 06:30”
# - 选值方式：在窗口内找与目标时刻绝对时间差最小的记录，取其 glucose
# - 若某段窗口内没有任何记录，则该段返回 None（保证列对齐）
def extract_fbs_by_24h_windows(file_path, max_days=MAX_DAYS, fbs_clock_time=FBS_CLOCK_TIME):
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

    # glucose 强制转为数值；无法转换的置为 NaN
    df['glucose'] = pd.to_numeric(df['glucose'], errors='coerce')
    # 时间为空的行丢弃；按时间排序，确保窗口切分与 idxmin 选择稳定
    df = df.dropna(subset=['timestamp']).sort_values('timestamp').reset_index(drop=True)

    if df.empty:
        return []

    # 以第一条记录为窗口起点 t0；用最后一条时间戳作为“是否还需要继续切窗”的停止条件
    t0 = df['timestamp'].iloc[0].to_pydatetime()
    t_last = df['timestamp'].iloc[-1].to_pydatetime()

    fbs_values = []
    for i in range(int(max_days)):
        # 第 i 段窗口起止（闭开区间）
        window_start = t0 + timedelta(days=i)
        if window_start > t_last:
            break
        window_end = window_start + timedelta(days=1)

        # 取该窗口内的所有记录
        window_df = df[(df['timestamp'] >= window_start) & (df['timestamp'] < window_end)].copy()
        if window_df.empty:
            fbs_values.append(None)
            continue

        # 计算窗口内应选择的目标时刻（窗口起点之后的下一次 FBS_CLOCK_TIME）
        target_dt = _next_clock_time_on_or_after(window_start, fbs_clock_time)
        # 找到窗口内最接近目标时刻的记录
        window_df['time_diff'] = (window_df['timestamp'] - target_dt).abs()
        closest_row = window_df.loc[window_df['time_diff'].idxmin()]
        fbs_values.append(closest_row['glucose'])

    return fbs_values


# --- 主流程 ---
# 1) 读取 request 列表并归一化前 7 列基础信息
# 2) 逐行根据 MATCH_BY 找到对应 CGM 文件（文件名包含匹配键）
# 3) 对每个文件提取最多 14 段的 FBS 值，并用 None 补齐到固定 14 列
# 4) 输出为 Excel：前 7 列基础信息 + Day 1 ~ Day 14
def main():
    output_file = get_output_filename()

    try:
        patient_df_raw = pd.read_excel(patient_list_file)
        patient_df = normalize_request_df(patient_df_raw, REQUEST_COLUMNS)
    except Exception as e:
        print(f"无法读取患者列表文件 {patient_list_file}: {e}")
        return

    results = []
    base_fields = ['hospital_id', 'pump_start_time', 'pump_end_time', 'discharge_time', 'admission_time', 'sensor_id', 'phone_number']
    print(f"开始处理，匹配依据: {MATCH_BY}")

    for index, row in patient_df.iterrows():
        base_info = [row.get(f, None) for f in base_fields]

        # 优先使用 MATCH_BY 指定列作为匹配键；若为空则回退到 hospital_id
        match_key = row.get(MATCH_BY, None) if MATCH_BY in patient_df.columns else None
        if pd.isna(match_key) or match_key is None:
            match_key = row.get('hospital_id', None)

        if pd.isna(match_key):
            print(f"警告: 第 {index+1} 行患者的匹配键 {MATCH_BY} 为空，跳过匹配")
            # 匹配键缺失时，仍保留基础信息，并填空的 14 个窗口值
            results.append(base_info + [None] * MAX_DAYS)
            continue

        # 在数据目录中按“文件名包含匹配键字符串”定位 CGM 文件
        patient_file = find_patient_file(match_key, data_folder)

        if patient_file:
            print(f"正在处理: {match_key} -> {os.path.basename(patient_file)}")
            fbs_values = extract_fbs_by_24h_windows(patient_file, max_days=MAX_DAYS, fbs_clock_time=FBS_CLOCK_TIME) or []
        else:
            print(f"未找到文件: {match_key}")
            fbs_values = []

        # 强制输出固定 14 列：先截断，再用 None 补齐
        fbs_values = list(fbs_values)[:MAX_DAYS]
        fbs_values = fbs_values + [None] * (MAX_DAYS - len(fbs_values))
        results.append(base_info + fbs_values)

    # 固定列名 Day 1 ~ Day 14（严格上限，保证不会生成 Day 15）
    fbs_cols = [f"Day {i+1}" for i in range(MAX_DAYS)]
    final_cols = base_fields + fbs_cols
    final_df = pd.DataFrame(results, columns=final_cols)
    final_df.to_excel(output_file, index=False)
    print(f"处理完成，结果已保存至：{output_file}")


if __name__ == "__main__":
    main()
