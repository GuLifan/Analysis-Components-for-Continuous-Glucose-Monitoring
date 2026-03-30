# 00_05_ExcelColumnTransfer.py
# 将第二个Excel的指定列，按前三列匹配后写入第一个Excel的指定列
# LAST UPDATE BY LIFANGU IN 20260324
#
# 规则：
# 1) 当 file1 每行前3列 与 file2 前3列完全匹配时
# 2) 复制 file2 的第4、5列 -> 写入 file1 的第7、8列

# --- 导入包（忽视 IDE 高亮，已是最简洁状态） ---
import os
import sys
from datetime import datetime

import numpy as np
import pandas as pd

# --- 配置路径（硬编码） ---
file1_path = r"C:\\Users\\lifan\\Desktop\\IN.xlsx"
file2_path = r"C:\\Users\\lifan\\Desktop\\IN_all.xlsx"
output_path = r"C:\\Users\\lifan\\Desktop\\File1_Merged.xlsx"

# --- 列参数（按 Excel 的 1-based 列号填写）---
match_cols_1based = [1, 2, 3]  # 前三列作为匹配键
src_cols_1based = [4, 5]       # 从第二个文件复制的列
dst_cols_1based = [7, 8]       # 写入第一个文件的列


def normalize_cell(val):
    if pd.isna(val):
        return None
    if isinstance(val, (pd.Timestamp, datetime)):
        return val.strftime('%Y-%m-%d %H:%M:%S')
    text = str(val).strip()
    return text if text != '' else None


def main():
    try:
        print(f"正在读取文件1: {file1_path}")
        df1 = pd.read_excel(file1_path, engine='openpyxl')
        print(f"正在读取文件2: {file2_path}")
        df2 = pd.read_excel(file2_path, engine='openpyxl')

        if df1.empty:
            print("错误：文件1为空，无法处理")
            return
        if df2.empty:
            print("错误：文件2为空，无法处理")
            return

        match_cols = [c - 1 for c in match_cols_1based]
        src_cols = [c - 1 for c in src_cols_1based]
        dst_cols = [c - 1 for c in dst_cols_1based]

        if df1.shape[1] < max(match_cols) + 1:
            print(f"错误：文件1列数不足，至少需要 {max(match_cols) + 1} 列用于匹配")
            return
        if df2.shape[1] < max(match_cols + src_cols) + 1:
            print(f"错误：文件2列数不足，至少需要 {max(match_cols + src_cols) + 1} 列用于匹配和取值")
            return

        while df1.shape[1] < max(dst_cols) + 1:
            df1[f'Extra_{df1.shape[1] + 1}'] = np.nan

        key_names = ['__k1', '__k2', '__k3']
        for idx, key_name in zip(match_cols, key_names):
            df1[key_name] = df1.iloc[:, idx].apply(normalize_cell)
            df2[key_name] = df2.iloc[:, idx].apply(normalize_cell)

        df2_map = df2[key_names + [df2.columns[src_cols[0]], df2.columns[src_cols[1]]]].copy()
        df2_map.columns = key_names + ['__v4', '__v5']

        duplicate_count = int(df2_map.duplicated(subset=key_names, keep='first').sum())
        if duplicate_count > 0:
            print(f"警告：文件2中存在 {duplicate_count} 组重复匹配键，将只使用首条记录")
            df2_map = df2_map.drop_duplicates(subset=key_names, keep='first')

        merged = df1.merge(df2_map, on=key_names, how='left')

        matched_rows = int(((merged['__v4'].notna()) | (merged['__v5'].notna())).sum())
        print(f"匹配到的行数（至少有一列可写入）: {matched_rows}")

        dst_col_name_1 = df1.columns[dst_cols[0]]
        dst_col_name_2 = df1.columns[dst_cols[1]]
        df1[dst_col_name_1] = df1[dst_col_name_1].astype('object')
        df1[dst_col_name_2] = df1[dst_col_name_2].astype('object')

        df1[dst_col_name_1] = merged['__v4'].astype('object').to_numpy()
        df1[dst_col_name_2] = merged['__v5'].astype('object').to_numpy()

        df1 = df1.drop(columns=key_names, errors='ignore')

        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        df1.to_excel(output_path, index=False, engine='openpyxl')
        print(f"完成！输出文件已保存: {output_path}")

    except FileNotFoundError as e:
        print(f"错误: 文件不存在 - {str(e)}")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        print("错误: Excel文件为空或格式不正确")
        sys.exit(1)
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()

