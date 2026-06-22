# 04_01_RequestBuilder.py
# 功能：以住院号为唯一匹配键，将患者信息.xlsx中MatchingRelationship未包含的列合并进来
#       输出 requestall.csv，第一列为住院号，日期格式统一为 YYYY-MM-DD HH:MM:SS
# 匹配逻辑：
#   - MatchingRelationship 第6列（住院号）⇔ 患者信息 第1列（住院号）
#   - 以 MatchingRelationship 为基准左连接，匹配不到的填空白
#   - MatchingRelationship CSV 文件的领取日期已为 YYYY-MM-DD 标准格式
# LAST UPDATE BY LIFANGU IN 20260618

# --- 导入包 ---
import os
import sys
import re
import pandas as pd
import numpy as np
from datetime import datetime

# --- 项目根目录 ---
ROOT_DIR = os.path.dirname(os.path.abspath(__file__))
# --- 文件路径 ---
MATCH_FILE = os.path.join(ROOT_DIR, "02_OriginalData", "MatchingRelationship250601_260331.csv")
INFO_FILE = os.path.join(ROOT_DIR, "患者信息.xlsx")
OUTPUT_FILE = os.path.join(ROOT_DIR, "02_OriginalData", "RequestList250601_260331.csv")

# --- MatchingRelationship CSV 文件（标准 csv，首行为列名）---

# --- 患者信息 文件结构（0-based 行/列索引）---
INFO_HEADER_ROW = 0         # 列名行
INFO_DATA_START = 1         # 数据起始行

# --- 输出列顺序（依次从左到右）---
OUTPUT_COLUMNS = [
    '序号',
    '住院号',   
    '领取日期',
    '一次性探头编号',
    '电话号码',
    '患者科室',
    '动态血糖开立时间',    # 以下来自患者信息
    '入院日期',
    '出院日期',
    '胰岛素泵',
    '胰岛素泵开始时间',
    '胰岛素泵停止时间',
    '德谷开始时间',
    '甘精开立时间',
]

# =============================================================================

# 通用值标准化：NaN → ""，去首尾空白和换行符
def normalize_str(val):
    if val is None: return ""
    if isinstance(val, float) and np.isnan(val): return ""
    s = str(val).strip().replace('\n', '').replace('\r', '')
    return "" if s.lower() == 'nan' else s

# 电话号码标准化：int/float → 去小数点的字符串
def normalize_phone(val):
    if val is None:return ""
    if isinstance(val, float) and np.isnan(val):return ""
    if isinstance(val, (int, float)): s = str(int(val))
    else: s = str(val).strip()
    return s

# 确保输出为 'YYYY-MM-DD HH:MM:SS' 格式
def normalize_datetime(val):
    if val is None: return ""
    if isinstance(val, float) and np.isnan(val): return ""
    if isinstance(val, (pd.Timestamp, datetime)): return val.strftime('%Y-%m-%d %H:%M:%S')
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'nat'): return ""
    if re.match(r'^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$', s): return s
    if re.match(r'^\d{4}-\d{2}-\d{2}$', s): return s + ' 00:00:00'
    try:
        dt = pd.to_datetime(s, errors='coerce')
        if not pd.isna(dt):
            return dt.strftime('%Y-%m-%d %H:%M:%S')
    except Exception:
        pass
    return s

# 领取日期标准化：CSV 中已为 YYYY-MM-DD，保持原样
def fix_get_date(val):
    return normalize_str(val)

# =============================================================================

def main():
    print("=" * 60)
    print("04_01_RequestBuilder")
    print("=" * 60)

    # ===== 步骤1: 读取 MatchingRelationship CSV =====
    print(f"\n[步骤1] 读取 MatchingRelationship: {os.path.basename(MATCH_FILE)}")
    df_match = pd.read_csv(MATCH_FILE, dtype=object, encoding='utf-8-sig')
    print(f"  列名: {list(df_match.columns)}")
    print(f"  数据行数: {len(df_match)}")

    # ===== 步骤2: 读取患者信息 =====
    print(f"\n[步骤2] 读取患者信息: {os.path.basename(INFO_FILE)}")
    df_info_raw = pd.read_excel(INFO_FILE, header=None, dtype=object)
    info_headers = df_info_raw.iloc[INFO_HEADER_ROW].tolist()
    print(f"  原始列名: {info_headers}")
    df_info = df_info_raw.iloc[INFO_DATA_START:].reset_index(drop=True)
    df_info.columns = info_headers
    print(f"  数据行数: {len(df_info)}")

    # ===== 步骤3: 构建患者信息查找字典 =====
    # key=住院号, value=dict of {列名: 原始值}
    print(f"\n[步骤3] 构建患者信息索引...")
    info_dict = {}
    for _, row in df_info.iterrows():
        hid = normalize_str(row['住院号'])
        if not hid:
            continue
        if hid not in info_dict:
            info_dict[hid] = {c: row[c] for c in info_headers if c != '住院号'}
    print(f"  索引 {len(info_dict)} 条唯一住院号记录")

    # ===== 步骤4: 逐行匹配合并 =====
    print(f"\n[步骤4] 匹配合并...")
    output_rows = []
    matched_count = 0
    unmatched_hids = []

    for _, row in df_match.iterrows():
        hid = normalize_str(row['住院号'])
        if not hid:
            continue

        # 基础列（来自 MatchingRelationship）
        row_data = {

            '序号': normalize_str(row['序号']),
            '住院号': hid,
            '领取日期': fix_get_date(row['领取日期']),
            '一次性探头编号': normalize_str(row['一次性探头编号']),
            '电话号码': normalize_phone(row['电话号码']),
            '患者科室': normalize_str(row['患者科室']),
        }

        # 新增列（来自患者信息查找）
        if hid in info_dict:
            matched_count += 1
            for col_name, val in info_dict[hid].items():
                row_data[col_name] = normalize_datetime(val) if '时间' in col_name or '日期' in col_name else (normalize_str(val) if val is not None and not (isinstance(val, float) and np.isnan(val)) else "")
        else:
            unmatched_hids.append(hid)
            for c in info_headers:
                if c != '住院号':
                    row_data[c] = ""

        output_rows.append(row_data)

    # ===== 步骤5: 构建输出并巡检日期格式 =====
    print(f"\n[步骤5] 构建输出、检查日期格式...")
    df_out = pd.DataFrame(output_rows, columns=OUTPUT_COLUMNS)

    date_cols_to_check = [
        '动态血糖开立时间', '入院日期', '出院日期',
        '胰岛素泵开始时间', '胰岛素泵停止时间', '德谷开始时间', '甘精开立时间'
    ]
    fixes_log = []
    for col in date_cols_to_check:
        if col not in df_out.columns:
            continue
        for idx in df_out.index:
            original = df_out.at[idx, col]
            fixed = normalize_datetime(original)
            if str(original) != str(fixed):
                fixes_log.append(f"    行{idx+1} {col}: '{original}' -> '{fixed}'")
                df_out.at[idx, col] = fixed

    if fixes_log:
        print(f"  格式修复 {len(fixes_log)} 处:")
        for log in fixes_log[:10]:
            print(log)
        if len(fixes_log) > 10:
            print(f"  ... 及另外 {len(fixes_log) - 10} 处")
    else:
        print(f"  所有日期格式均已正确，无需修复")

    # ===== 步骤6: 保存输出 =====
    print(f"\n[步骤6] 保存: {OUTPUT_FILE}")
    # 先写到临时目录再复制，避免文件锁定问题
    tmp_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "requestall_tmp.csv")
    df_out.to_csv(tmp_path, index=False, encoding='utf-8-sig')
    os.replace(tmp_path, OUTPUT_FILE)

    # ===== 统计报告 =====
    print(f"\n{'=' * 60}")
    print(f"    完成!")
    print(f"    MatchingRelationship 行数: {len(df_match)}")
    print(f"    输出行数:                 {len(df_out)}")
    print(f"    匹配到患者信息:           {matched_count}")
    print(f"    未匹配到:                 {len(unmatched_hids)}")
    if unmatched_hids:
        print(f"    未匹配住院号列表 ({len(unmatched_hids)}):")
        for h in unmatched_hids:
            print(f"      - {h}")
    print(f"    输出文件:                 {OUTPUT_FILE}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
