# 00_01_DataCleaning.py
# 删除信息不完善无法纳入计算的患者
# LAST UPDATE BY LIFANGU IN 202601302012

# --- 导入包（忽视 IDE 高亮，已是最简洁状态） ---
import os
import pandas as pd
import numpy as np
import openpyxl

# --- 配置区域 ---
# 在此处修改文件路径，注意路径中的斜杠为反斜杠
input_path = r"C:\Users\lifan\Desktop\2407-2505科外.xlsx"
name_tag = "2505EX"
output_path = f"C:\Users\lifan\Desktop\DataRequest{name_tag}.xlsx"
# 手动指定有效信息的列数
total_lens = 7 
# 设定文件名所在列，注意无需将第一列算作第 0 列，若此功能无用，则设为 1 
device_name_len = 2

# --- 执行区域 ---
print(f"ATTENTION! 已设定需要 {total_lens} 列数据，文件名所在列为第 {device_name_len} 列")

# 读取Excel文件（不自动识别列名）
df = pd.read_excel(input_path, header=None)
device_name_len = device_name_len - 1  # 列索引从0开始，所以需要减1

# 删除首行（索引0），此功能暂时不需要，因为首行是列名，不是数据行。已注释该功能
# df = df.drop(0)

# 重置索引（避免出现跳号）
df.reset_index(drop=True, inplace=True)

# 筛选条件：检查每行前 total_lens 列是否包含空值或"#N/A"
def is_valid_row(row):
    for i in range(total_lens):  # 检查前n列, 如果新的住院号等信息的xlsx文件里面的有效信息不止前四列,就修改这个total_lens
        cell_value = row[i]
        # 如果单元格值为空或者为"#N/A"，则返回False
        if pd.isna(cell_value) or (isinstance(cell_value, str) and cell_value.strip() == "#N/A"):
            return False
    return True

# 应用筛选条件
valid_mask = df.apply(is_valid_row, axis=1)
df = df[valid_mask]

# 处理文件名称列，将'-'替换为'_'，并移除'.xlsx'后缀
if len(df) > 0:  # 确保有有效数据
    # 转换为字符串类型以便处理
    df[device_name_len] = df[device_name_len].astype(str)
    # 替换操作：'-' 替换为 '_'，移除'.xlsx'
    df[device_name_len] = (
        df[device_name_len]
        .str.replace('-', '_', regex=False)  # 替换横线
        .str.replace('.xlsx', '', regex=False)  # 移除后缀
    )

# 保存处理后的数据到新文件（不包含索引和列名）
df.to_excel(output_path, index=False, header=False)
print(f"文件处理完成，已保存至: {output_path}")