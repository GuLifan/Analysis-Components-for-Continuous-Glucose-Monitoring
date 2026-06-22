# 04_05_FBSExtractor.py
# 提取每人每天最接近 06:30 的血糖值（按自然日 date 分组）
# 基于 02_01_FBSExtractor.py 重写：CSV + config.yaml + 统一 5 分钟采样
# LAST UPDATE BY LIFANGU IN 20260618

import os, yaml
import pandas as pd, numpy as np
from datetime import datetime, time

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

def load_and_sample(file_path):
    df = pd.read_excel(file_path, header=None)
    df = df.drop(0).reset_index(drop=True)
    df.columns = ['timestamp', 'glucose']
    df['timestamp'] = pd.to_datetime(df['timestamp'], format='%Y-%m-%d %H:%M:%S')
    df['glucose'] = pd.to_numeric(df['glucose'], errors='coerce')
    nr = (df['glucose'].dropna() <= 0.5).sum() / max(len(df['glucose'].dropna()), 1)
    if nr > 0.7:
        df = df[df['glucose'] > 0.5].copy()
    else:
        df = df.iloc[::5].copy()
    return df

def extract_daily_fbs(file_path):
    """每人每天最接近 06:30 的血糖值"""
    try:
        df = load_and_sample(file_path)
    except Exception:
        return []
    TARGET = time(6, 30, 0)
    daily = []
    for date, g in df.groupby(df['timestamp'].dt.date):
        td = datetime.combine(date, TARGET)
        g = g.copy()
        g['td'] = (g['timestamp'] - td).abs()
        daily.append(g.loc[g['td'].idxmin(), 'glucose'])
    return daily

def main():
    print("04_05 FBSExtractor", flush=True)
    pdf = read_patient_csv(PATIENT_LIST)
    base = ['hospital_id', 'admission_time', 'discharge_time', 'sensor_id',
            'pump_start_time', 'pump_end_time']
    results, max_days = [], 0
    for _, row in pdf.iterrows():
        bi = [row.get(c, None) for c in base]
        mk = row.get(MATCH_BY) if MATCH_BY in pdf.columns else row.get('hospital_id')
        if pd.isna(mk): results.append(bi + []); continue
        pf = find_patient_file(mk, DATA_FOLDER)
        fbs = extract_daily_fbs(pf) if pf else []
        if fbs: max_days = max(max_days, len(fbs))
        results.append(bi + (fbs if fbs else []))
    fcols = [f'Day {i+1}' for i in range(max_days)]
    padded = []
    for rd in results:
        fb = rd[len(base):]; padded.append(rd[:len(base)] + fb + [None]*(max_days - len(fb)))
    out = os.path.join(OUTPUT_FOLDER, f'CGM_{DATETAG}_{NAMETAG}_FBS_Extraction.xlsx')
    pd.DataFrame(padded, columns=base + fcols).to_excel(out, index=False)
    print(f"Done: {out}")

if __name__ == '__main__':
    main()
