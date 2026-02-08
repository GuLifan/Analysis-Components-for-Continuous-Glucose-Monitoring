# 00_04_FixDateFormat.py
# 用于修复 DataRequest 文件中胰岛素泵时间格式错误 (2025_05_28 -> 2025-05-28)
# LAST UPDATE BY LIFANGU IN 202602082000

# --- 导入包 ---
import pandas as pd
import os

# --- 配置区域 ---
# 需要处理的文件列表
target_files = [
    r"C:\Users\lifan\Desktop\03_Requests\DataRequest2505EX.xlsx",
    r"C:\Users\lifan\Desktop\03_Requests\DataRequest2412EX.xlsx",
    # 如果还有其他文件需要处理，请在此处添加
    # r"C:\Users\lifan\Desktop\03_Requests\DataRequest2412EX.xlsx",
]

# 胰岛素泵时间所在的列索引 (从0开始)
# 根据 01_02_Calculation.py 的推断：
# Column 0: Hospital ID
# Column 1: Pump Start Time
# Column 2: Pump End Time
pump_start_col = 1
pump_end_col = 2

# --- 执行区域 ---
print(f"开始处理时间格式修复任务...")

for file_path in target_files:
    if not os.path.exists(file_path):
        print(f"文件不存在，跳过: {file_path}")
        continue
        
    print(f"正在处理: {file_path}")
    
    try:
        # 读取Excel文件（不自动识别列名，保持 header=None 以便按索引操作）
        df = pd.read_excel(file_path, header=None)
        
        # 定义修复函数
        def fix_date_format(val):
            if pd.isna(val):
                return val
            
            # 如果是字符串，尝试替换
            if isinstance(val, str):
                # 替换 '2025_05_28' -> '2025-05-28'
                # 同时处理可能存在的其他分隔符问题
                new_val = val.replace('_', '-')
                return new_val
            
            # 如果已经是时间对象，直接返回
            return val

        # 修复泵开始时间
        if len(df.columns) > pump_start_col:
            # 跳过第一行（如果是表头）
            # 通常 DataRequest 文件第一行是表头，我们需要避开吗？
            # 00_01_DataCleaning.py 中没有避开，因为 header=None 读取，第一行也是数据。
            # 如果第一行是 "Pump Start"，replace '_' 也没影响。
            
            original_start = df[pump_start_col].copy()
            df[pump_start_col] = df[pump_start_col].apply(fix_date_format)
            
            # 检查是否有变化
            changes = (df[pump_start_col] != original_start).sum()
            if changes > 0:
                print(f"  - 修复了 {changes} 个开始时间格式")
        
        # 修复泵结束时间
        if len(df.columns) > pump_end_col:
            original_end = df[pump_end_col].copy()
            df[pump_end_col] = df[pump_end_col].apply(fix_date_format)
            
            changes = (df[pump_end_col] != original_end).sum()
            if changes > 0:
                print(f"  - 修复了 {changes} 个结束时间格式")

        # 保存回原文件
        # 注意：不带 header 和 index，保持原格式
        df.to_excel(file_path, index=False, header=False)
        print(f"  ✅ 保存成功")
        
    except Exception as e:
        print(f"  ❌ 处理失败: {e}")

print("所有任务完成。")
