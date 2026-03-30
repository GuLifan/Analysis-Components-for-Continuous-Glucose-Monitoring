# 00_02_FileFetching.py
# 抓取需要计算的患者的CGM记录文件
# LAST UPDATE BY LIFANGU IN 20260324

# --- 导入包（忽视 IDE 高亮，已是最简洁状态） ---
import os
import shutil
import pandas as pd
import openpyxl
import numpy as np
import sys
from datetime import datetime

# -- 配置路径 ---
USE_PUMP = True
name_tag = "2505IN"
if USE_PUMP:
    name_tag = name_tag + "_Pump"

excel_path = fr"C:\\Users\\lifan\\Desktop\\{name_tag}.xlsx"
input_folder = r"D:\\LifanDataTrae\\CGMCT20260122neo\\02_OriginalData\\CGMOriginalDataAll_250531"
output_folder = fr"C:\\Users\\lifan\\Desktop\\DataSelectedFor{name_tag}"

# 以下为匹配标识符，2024年12月之前的文件应该填写 5 （手机号码），2024年12月之后的文件应该填写 4 （探头号）
if name_tag == "2412IN" or name_tag == "2412EX" or name_tag == "2412IN_Pump" or name_tag == "2412EX_Pump":
    identifier_col = 5  # 匹配标识符所在列，注意文件和时间关系，不需要将第一列当作第 0 列
else:
    identifier_col = 4  # 匹配标识符所在列，注意文件和时间关系，不需要将第一列当作第 0 列

adm_time_col = 3  # 入院时间所在列
dis_time_col = 2  # 出院时间所在列

# 最小文件大小判定：最小有效行数（<24h则跳过）
min_valid_hours = 12  # 规则1：最小有效行数（<24h则跳过）

# 胰岛素泵时间判定：最接近胰岛素泵时间匹配
enable_pump_match = USE_PUMP
pump_start_col = 6  # 胰岛素泵开始时间所在列
pump_end_col = 7    # 胰岛素泵停止时间所在列

# --- 执行区域 ---
identifier_col = identifier_col - 1  # excel中首列为0
adm_time_col = adm_time_col - 1  # excel中首列为0
dis_time_col = dis_time_col - 1  # excel中首列为0
pump_start_col = pump_start_col - 1
pump_end_col = pump_end_col - 1
min_valid_rows = int( min_valid_hours * 12 )  # 转换为行数

# 检查时间重叠函数，有重叠则选择的文件合法
def is_time_overlap(adm_time, dis_time, file_start, file_end):
    try:
        # 转换所有时间为datetime对象
        adm_time = pd.to_datetime(adm_time, errors='coerce')
        dis_time = pd.to_datetime(dis_time, errors='coerce')
        file_start = pd.to_datetime(file_start, errors='coerce')
        file_end = pd.to_datetime(file_end, errors='coerce')

        # 如果有任何时间为NaT，则返回False
        if pd.isna(adm_time) or pd.isna(dis_time) or pd.isna(file_start) or pd.isna(file_end):
            return False

        # 检查时间重叠: (A_start <= B_end) and (B_start <= A_end)
        return (adm_time <= file_end) and (file_start <= dis_time)
    except Exception as e:
        print(f"时间比较错误: {str(e)}")
        return False


def calculate_match_score(pump_start, pump_end, file_start, file_end):
    """
    计算文件时间段与胰岛素泵使用时间段的匹配分数。
    分数越高表示越接近。
    逻辑：
    1. 计算重叠时长（Overlap Duration）。
    2. 计算距离（Gap Duration）。
    3. 如果有重叠，分数为重叠时长的正值（秒）。
    4. 如果无重叠，分数为距离的负值（秒）。
    这样 max(score) 就能选出重叠最多或者距离最近的文件。
    """
    try:
        pump_start = pd.to_datetime(pump_start)
        pump_end = pd.to_datetime(pump_end)
        file_start = pd.to_datetime(file_start)
        file_end = pd.to_datetime(file_end)

        if pd.isna(pump_start) or pd.isna(pump_end):
            # 如果没有泵的时间数据，无法比较，返回一个中性分数或者最低分
            # 这里假设没有泵数据时，不具备比较优势
            return -float('inf')

        # 计算交集
        latest_start = max(pump_start, file_start)
        earliest_end = min(pump_end, file_end)
        
        delta = (earliest_end - latest_start).total_seconds()

        if delta > 0:
            # 有重叠，返回重叠时长（正数）
            return delta
        else:
            # 无重叠，返回距离的负数（距离越小，分数越大）
            # delta 为负数，其绝对值就是两个区间最近点的距离
            return delta
    except Exception:
        return -float('inf')


def main():
    try:
        # 1. 从Excel文件读取数据
        print(f"正在从Excel文件读取数据: {excel_path}")
        df_excel = pd.read_excel(excel_path, engine='openpyxl')

        # 创建字典以便快速查找：键为设备文件名，值包含所有时间信息
        excel_data = {}
        for _, row in df_excel.iterrows():
            identifier = str(row.iloc[identifier_col])
            if pd.notna(identifier) and identifier != 'nan':
                adm_time = row.iloc[adm_time_col]
                dis_time = row.iloc[dis_time_col]
                
                pump_start = row.iloc[pump_start_col] if enable_pump_match else None
                pump_end = row.iloc[pump_end_col] if enable_pump_match else None

                excel_data[identifier] = {
                    'adm_time': adm_time,
                    'dis_time': dis_time,
                    'pump_start': pump_start,
                    'pump_end': pump_end
                }

        print(f"找到 {len(excel_data)} 个有效设备ID及其时间数据")
        if enable_pump_match:
            print("规则2已启用：将保留与胰岛素泵时间最接近的文件。")
        else:
            print("规则2未启用：将保留所有符合条件的重叠文件。")

        if not excel_data:
            print("警告: 未在Excel文件中找到有效的设备ID和时间数据!")
            return

        # 2. 创建目标文件夹
        os.makedirs(output_folder, exist_ok=True)
        print(f"目标文件夹已创建: {output_folder}")

        # 3. 遍历源文件夹查找匹配文件并检查时间重叠
        # 使用字典存储候选文件： key=identifier, value=list of (src_path, filename, score)
        candidates = {} 
        total_files = 0
        
        # 遍历源文件夹中的所有xlsx文件
        for root, _, files in os.walk(input_folder):
            for filename in files:
                if not filename.lower().endswith('.xlsx'):
                    continue

                total_files += 1
                file_base = os.path.splitext(filename)[0]

                # 检查文件名是否包含Excel中的任意设备ID
                matched_device = None
                for device_id in excel_data.keys():
                    if device_id in file_base:
                        matched_device = device_id
                        break

                if matched_device:
                    src_path = os.path.join(root, filename)

                    try:
                        # 读取文件
                        df_file = pd.read_excel(src_path, engine='openpyxl')

                        # 检查文件是否为空或行数不足
                        if df_file.empty or len(df_file) == 0:
                            print(f"跳过空文件: {filename}")
                            continue

                        # 获取文件时间范围
                        file_start = df_file.iloc[0, 0]
                        file_end = df_file.iloc[-1, 0]
                        row_count = len(df_file) # 记录行数供后续筛选

                        # 获取Excel数据
                        data = excel_data[matched_device]
                        adm_time = data['adm_time']
                        dis_time = data['dis_time']

                        # 检查时间是否有重叠
                        if is_time_overlap(adm_time, dis_time, file_start, file_end):
                            # 计算匹配分数（仅在启用规则2时有用）
                            score = 0
                            if enable_pump_match:
                                score = calculate_match_score(data['pump_start'], data['pump_end'], file_start, file_end)
                            
                            if matched_device not in candidates:
                                candidates[matched_device] = []
                            
                            candidates[matched_device].append({
                                'src_path': src_path,
                                'filename': filename,
                                'score': score,
                                'row_count': row_count
                            })
                        else:
                            # 时间不匹配
                            pass
                    except Exception as e:
                        print(f"处理文件 {filename} 时出错: {str(e)}")
                        continue

        # 4. 处理候选文件并复制
        matched_files = []
        matched_count = 0

        print(f"\n开始处理候选文件...")
        
        for device_id, file_list in candidates.items():
            files_to_copy = []
            
            # 只有当匹配文件数量超过1时，才应用筛选规则
            if len(file_list) > 1:
                # 规则1：过滤掉行数不足的文件
                valid_files = [f for f in file_list if f['row_count'] >= min_valid_rows]
                
                if not valid_files:
                     # 过滤后没有文件了，全部跳过
                     print(f"警告：设备 {device_id} 的所有 {len(file_list)} 个候选文件行数均不足 {min_valid_rows}，全部跳过。")
                     files_to_copy = []
                
                # 如果过滤后只剩一个，直接使用
                elif len(valid_files) == 1:
                    files_to_copy = valid_files
                    print(f"设备 {device_id}：经规则1筛选后保留唯一文件 {valid_files[0]['filename']}")
                
                # 如果过滤后仍有多个，且启用了规则2
                elif enable_pump_match:
                    # 规则2：只保留分数最高的一个
                    # 按分数降序排序
                    valid_files.sort(key=lambda x: x['score'], reverse=True)
                    best_match = valid_files[0]
                    files_to_copy = [best_match]
                    print(f"设备 {device_id}：经规则1筛选后仍有 {len(valid_files)} 个文件，根据规则2选择了最接近的一个: {best_match['filename']}")
                
                else:
                    # 规则2未启用，保留所有通过规则1的文件
                    files_to_copy = valid_files
                    print(f"设备 {device_id}：经规则1筛选后保留 {len(valid_files)} 个文件（规则2未启用）")
            
            else:
                # 只有一个文件，但也需要应用规则1检查行数
                valid_files = [f for f in file_list if f['row_count'] >= min_valid_rows]
                if valid_files:
                    files_to_copy = valid_files
                    print(f"设备 {device_id}：找到1个匹配文件且符合行数要求，保留")
                else:
                    files_to_copy = []
                    print(f"警告：设备 {device_id} 的唯一匹配文件行数不足 {min_valid_rows}，跳过")
            
            # 执行复制
            for item in files_to_copy:
                src_path = item['src_path']
                filename = item['filename']
                
                dst_path = os.path.join(output_folder, filename)
                
                # 处理重复文件名
                counter = 1
                base, ext = os.path.splitext(filename)
                while os.path.exists(dst_path):
                    new_filename = f"{base}_{counter}{ext}"
                    dst_path = os.path.join(output_folder, new_filename)
                    counter += 1
                
                shutil.copy2(src_path, dst_path)
                matched_files.append((src_path, dst_path, device_id))
                matched_count += 1
                print(f"复制文件: {filename} (设备ID: {device_id})")

        # 输出匹配文件列表
        if matched_files:
            print("\n已处理文件列表:")
            for src, dst, device in matched_files:
                print(f"- {os.path.basename(src)} (设备ID: {device}) -> {os.path.basename(dst)}")
        else:
            print("警告: 未找到任何匹配文件!")

        # 5. 输出结果
        print("\n操作完成!")
        print(f"扫描文件总数: {total_files}")
        print(f"匹配文件数量: {matched_count}")
        print(f"已复制到: {output_folder}")

    except FileNotFoundError as e:
        print(f"错误: 文件或目录不存在 - {str(e)}")
        sys.exit(1)
    except pd.errors.EmptyDataError:
        print("错误: Excel文件为空或格式不正确")
        sys.exit(1)
    except Exception as e:
        print(f"处理过程中发生错误: {str(e)}")
        sys.exit(1)


if __name__ == "__main__":
    main()
