# 04_06_CGMDaysExtractor.py
# 统计每个 CGM 文件的总时长（小时）及按 24h 折算的天数
# 基于 02_02_CGMDaysExtractor.py 重写：CSV + config.yaml + 统一 5 分钟采样
# LAST UPDATE BY LIFANGU IN 20260618

import os, yaml
import pandas as pd, numpy as np

ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
config = yaml.safe_load(open(os.path.join(ROOT_DIR, 'config.yaml'), 'r', encoding='utf-8'))
MATCH_BY      = config.get('match_by', 'sensor_id')
DATA_FOLDER   = config.get('data_folder', '')
OUTPUT_FOLDER = config.get('output_folder', '')
PATIENT_LIST  = config.get('patient_list_file', '')
DATETAG       = config.get('datetag', '')
NAMETAG       = config.get('nametag', '')

CSV_COLUMN_MAP = {
    '住院号': 'hospital_id', '一次性探头编号': 'sensor_id',
    '入院日期': 'admission_time', '出院日期': 'discharge_time',
    '胰岛素泵开始时间': 'pump_start_time', '胰岛素泵停止时间': 'pump_end_time',
}

def read_patient_csv(csv_path):
    df = pd.read_csv(csv_path, dtype=object, encoding='utf-8-sig')
    rn = {k: v for k, v in CSV_COLUMN_MAP.items() if k in df.columns}
    df = df.rename(columns=rn)
    for c in ['hospital_id', 'sensor_id', 'admission_time', 'discharge_time',
              'pump_start_time', 'pump_end_time']:
        if c not in df.columns: df[c] = np.nan
    df = df[df['hospital_id'].notna() & (df['hospital_id'] != '')].copy()
    if MATCH_BY == 'sensor_id':
        df = df[df['sensor_id'].notna() & (df['sensor_id'] != '')].copy()
    return df

def find_patient_file(device_id, folder_path):
    if pd.isna(device_id): return None
    s = str(int(device_id)) if isinstance(device_id, float) else str(device_id)
    for fn in os.listdir(folder_path):
        if s in fn and fn.endswith(('.xls', '.xlsx')):
            return os.path.join(folder_path, fn)
    return None

def extract_days(file_path):
    """计算总小时数（四舍五入）和天数"""
    try:
        df = pd.read_excel(file_path, header=None)
    except Exception:
        return None, None
    if df.shape[0] <= 1: return None, None
    df = df.drop(0).reset_index(drop=True)
    df.columns = ['timestamp', 'glucose']
    df['timestamp'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%d %H:%M:%S')
    start = df['timestamp'].min(); end = df['timestamp'].max()
    if pd.isna(start) or pd.isna(end): return None, None
    hours = (end - start).total_seconds() / 3600.0
    return int(np.floor(hours + 0.5)), int(np.floor(hours / 24.0 + 0.5))

def main():
    print("04_06 CGMDaysExtractor", flush=True)
    pdf = read_patient_csv(PATIENT_LIST)
    base = ['hospital_id', 'admission_time', 'discharge_time', 'sensor_id',
            'pump_start_time', 'pump_end_time']
    results = []
    for _, row in pdf.iterrows():
        bi = [row.get(c, None) for c in base]
        mk = row.get(MATCH_BY) if MATCH_BY in pdf.columns else row.get('hospital_id')
        if pd.isna(mk): results.append(bi + [None, None]); continue
        pf = find_patient_file(mk, DATA_FOLDER)
        h, d = extract_days(pf) if pf else (None, None)
        if pf and h: print(f"  {os.path.basename(pf)}: {h}h, {d}d", flush=True)
        results.append(bi + [h, d])
    out = os.path.join(OUTPUT_FOLDER, f'CGM_{DATETAG}_{NAMETAG}_Days_Extraction.xlsx')
    pd.DataFrame(results, columns=base + ['TotalHours', 'TotalDays']).to_excel(out, index=False)
    print(f"Done: {out}")

if __name__ == '__main__':
    main()
