# 00_03_DataMerger.py
# 用于将多个 Excel 文件按列名合并，并按住院号排序去重
# LAST UPDATE BY LIFANGU IN 202602090012

import pandas as pd
import os

# --- 配置区域 ---
ifresult = False # 合并结果文件

# 待合并的文件列表 (绝对路径)

if ifresult == True:

    locationtag = "EX" # EX 科外 IN 科内
    rangetag = "D7afterDischarge"
    mode = 5 # 0~7
    medfolder = "Mode5_D7afterDischarge"
    file_list = [
        fr"C:\Users\lifan\Desktop\00_Outputs\{medfolder}\CGM_DRQ260122_2407{locationtag}_{rangetag}_Mode{mode}_Results.xlsx",
        fr"C:\Users\lifan\Desktop\00_Outputs\{medfolder}\CGM_DRQ260122_2412{locationtag}_{rangetag}_Mode{mode}_Results.xlsx",
        fr"C:\Users\lifan\Desktop\00_Outputs\{medfolder}\CGM_DRQ260122_2505{locationtag}_{rangetag}_Mode{mode}_Results.xlsx"
    ]
    output_file = fr"C:\Users\lifan\Desktop\CGM_DRQ260122_{locationtag}_{medfolder}.xlsx"

else:
    file_list = [
        fr"C:\Users\lifan\Desktop\00_Outputs\CGM_DRQ260122_2505EX_Daily_Day1-3_Mode0_Results.xlsx",
        fr"C:\Users\lifan\Desktop\00_Outputs\CGM_DRQ260122_2412EX_Daily_Day1-3_Mode0_Results.xlsx",
        fr"C:\Users\lifan\Desktop\00_Outputs\CGM_DRQ260122_2407EX_Daily_Day1-3_Mode0_Results.xlsx"
    ]

    output_file = fr"C:\Users\lifan\Desktop\CGM_DRQ260122_EX_Mode0_Daily_Day1-3.xlsx"
# 默认住院号列名 (如果是第一列，代码逻辑会自动处理，这里用于显式指定)
# 假设所有文件第一列都是住院号，且列名一致
# 如果列名不一致，代码会以第一个文件的列名为准
hospital_id_col_index = 0 

# --- 执行区域 ---

def merge_excel_files():
    if not file_list:
        print("错误：文件列表为空")
        return

    print(f"准备合并 {len(file_list)} 个文件...")
    
    # 使用 ExcelFile 读取所有文件的所有 sheets
    # 结构: { sheet_name: [df1, df2, ...] }
    sheets_data = {}
    
    for i, file_path in enumerate(file_list):
        if not os.path.exists(file_path):
            print(f"警告：文件不存在，跳过: {file_path}")
            continue
            
        print(f"读取文件 [{i+1}/{len(file_list)}]: {file_path}")
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet_name)
                if sheet_name not in sheets_data:
                    sheets_data[sheet_name] = []
                sheets_data[sheet_name].append(df)
                
        except Exception as e:
            print(f"读取失败: {e}")
            return

    if not sheets_data:
        print("没有成功读取任何数据")
        return

    print(f"检测到 {len(sheets_data)} 个工作表: {list(sheets_data.keys())}")
    print("正在合并数据...")
    
    # 创建 ExcelWriter 对象用于写入多个 sheets
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        for sheet_name, dfs in sheets_data.items():
            print(f"处理工作表: {sheet_name}")
            
            # 纵向合并该 sheet 下的所有 dfs
            merged_df = pd.concat(dfs, ignore_index=True)
            
            # 排序和去重
            if not merged_df.empty:
                key_col = merged_df.columns[0] # 假设第一列是 Key
                
                # 排序
                merged_df = merged_df.sort_values(by=key_col, ascending=True)
                
                # 去重 (按 Key 列去重，保留第一个)
                initial_rows = len(merged_df)
                merged_df = merged_df.drop_duplicates(subset=[key_col], keep='first')
                removed_rows = initial_rows - len(merged_df)
                
                if removed_rows > 0:
                    print(f"  - 移除了 {removed_rows} 行重复数据")
            
            # 写入到输出文件的对应 sheet
            merged_df.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"保存结果至: {output_file}")
    print("完成!")

if __name__ == "__main__":
    merge_excel_files()
