# 01_01_BatchRunner.py
# CGM 批量计算工具 (Enhanced Version)
# 功能：按任务列表批量运行计算，支持全参数覆盖和进度显示
# 使用方法：配置 GLOBAL_CONFIG 和 TASKS，然后直接运行
# LAST UPDATE BY LIFANGU IN 202602081735

import subprocess
import sys
import time
import json
import os
from datetime import timedelta

# --- 1. 全局配置 (Global Configuration) ---
# 这些配置将应用于所有任务，除非在具体任务中被覆盖
# 路径相关
OUTPUT_FOLDER = r"C:\Users\lifan\Desktop\00_Outputs"
DATETAG = "DRQ260122"

# 计算开关 (指定一次即可)
GLOBAL_CALC_GROUPS = {
    1: True,   # Basic Stats
    2: True,   # Risk Stats (LBGI/HBGI/ADRR)
    3: True,   # Variability (MAGE/LAGE)
    4: True,   # Ranges (TIR/TAR/TBR)
    5: True,   # Hourly Stats
    6: True,   # Events (Hypo/Hyper)
    7: True,   # Level 2 Hypo
    8: True    # Conditional Hypo
}

GLOBAL_DAILY_OUTPUT = False  # 是否输出每日详细数据

# --- 计算参数 (全局设置) ---
GLOBAL_MODE = 5
GLOBAL_DURING = 7
GLOBAL_INTERIM = 6
COMMON_TAG = "D7afterDischarge"
# mode0: 开始时间~开始时间+duringday天，对应 D1-3、D1-5、D1-7、D1-14
# mode1: 开始时间+interimday天~开始时间+duringday天，对应D9-14、D7-14、D6-11
# mode2: 开始时间~结束时间，对应D all
# mode3: 出院时间-duringday天~出院时间，对应出院前3天
# mode4: 入院时间(实际上是血糖监测的开始时间)~出院时间，对应住院期间
# mode5: 出院时间~出院时间+duringday天，对应出院后1周
# mode6: 出院时间~结束时间，对应出院后全部时间
# mode7: 胰岛素泵使用时间，对应胰岛素泵期间


# --- 2. 任务列表 (Task List) ---
# 每个任务是一个字典，指定该批次特有的参数
# 必须包含: description, nametag, patient_list, data_folder
# 可选包含: match_by (默认 sensor_id，phone_number)
# 注意: mode, during, interim 将使用上述全局设置
TASKS = [
    # 示例任务 1: 处理 DataRequest2505EX
    {
        "description": f"Source 1 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2505EX.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2505EX",
        "match_by": "sensor_id",
        "nametag": f"2505EX_{COMMON_TAG}",
    },
    {
        "description": f"Source 2 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2505IN.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2505IN",
        "match_by": "sensor_id",
        "nametag": f"2505IN_{COMMON_TAG}",
    },
    {
        "description": f"Source 3 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2412EX.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2412EX",
        "match_by": "phone_number",
        "nametag": f"2412EX_{COMMON_TAG}",
    },
    {
        "description": f"Source 4 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2412IN.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2412IN",
        "match_by": "phone_number",
        "nametag": f"2412IN_{COMMON_TAG}",
    },
    {
        "description": f"Source 5 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2407EX.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2407EX",
        "match_by": "phone_number",
        "nametag": f"2407EX_{COMMON_TAG}",
    },
    {
        "description": f"Source 6 - MODE{GLOBAL_MODE} {COMMON_TAG}",
        "patient_list": r"C:\Users\lifan\Desktop\03_Requests\DataRequest2407IN.xlsx",
        "data_folder": r"C:\Users\lifan\Desktop\04_SelectedData\DataSelected2407IN",
        "match_by": "phone_number",
        "nametag": f"2407IN_{COMMON_TAG}",
    }
    # 示例任务 2: 处理 Source 2 (请修改路径)
    # {
    #     "description": "Source 2 - Standard Config",
    #     "patient_list": r"C:\Path\To\Source2.xlsx",
    #     "data_folder": r"C:\Path\To\Data2",
    #     "match_by": "phone_number",
    #     "nametag": "Source2_Standard",
    # },
]

# --- 3. 执行引擎 (Execution Engine) ---
SCRIPT_PATH = "01_02_Calculation.py"
PYTHON_EXE = sys.executable

def format_time(seconds):
    return str(timedelta(seconds=int(seconds)))

def run():
    total_tasks = len(TASKS)
    print(f"\n{'='*60}")
    print(f"CGM 批量计算任务开始")
    print(f"总任务数: {total_tasks}")
    print(f"全局参数: Mode={GLOBAL_MODE}, During={GLOBAL_DURING}, Interim={GLOBAL_INTERIM}")
    print(f"全局设置: Daily Output={GLOBAL_DAILY_OUTPUT}, Groups={list(GLOBAL_CALC_GROUPS.keys())}")
    print(f"{'='*60}\n")
    
    start_time_all = time.time()
    
    for i, task in enumerate(TASKS):
        task_num = i + 1
        desc = task.get("description", f"Task {task_num}")
        
        # 1. 进度显示
        elapsed = time.time() - start_time_all
        if i > 0:
            avg_time = elapsed / i
            remaining = avg_time * (total_tasks - i)
            eta_str = format_time(remaining)
        else:
            eta_str = "计算中..."
            
        print(f"[{task_num}/{total_tasks}] 正在执行: {desc}")
        print(f"      已用时间: {format_time(elapsed)} | 预计剩余: {eta_str}")
        print(f"      --------------------------------------------------")
        
        # 2. 构建命令
        cmd = [PYTHON_EXE, SCRIPT_PATH]
        
        # 添加固定参数
        cmd.extend(["--output_folder", OUTPUT_FOLDER])
        cmd.extend(["--datetag", DATETAG])
        
        # 添加全局计算参数
        cmd.extend(["--mode", str(GLOBAL_MODE)])
        cmd.extend(["--during", str(GLOBAL_DURING)])
        cmd.extend(["--interim", str(GLOBAL_INTERIM)])
        
        # 添加全局开关参数
        if GLOBAL_DAILY_OUTPUT:
            cmd.append("--daily_output")
        else:
            cmd.append("--no_daily_output")
            
        # 序列化 calc_groups
        cmd.extend(["--calc_groups", json.dumps(GLOBAL_CALC_GROUPS)])
        
        # 添加任务特定参数
        for key, value in task.items():
            if key == "description":
                continue
            # 只有当任务明确指定了 mode/during/interim 时才覆盖全局设置（虽然本需求是全局设置，但保留灵活性无妨）
            # 实际上，上面的代码已经添加了全局参数。如果这里再次添加同名参数，argparse通常会取最后一个。
            # 所以如果 task 中有 mode，它会覆盖上面的 GLOBAL_MODE。这符合“任务特定设置优先”的通用逻辑，
            # 但既然用户要求“只用指定一次”，在 TASKS 里不写这些字段即可。
            cmd.extend([f"--{key}", str(value)])
            
        # 3. 执行命令
        # print(f"      Cmd: {' '.join(cmd)}") # 调试用，可取消注释
        
        task_start = time.time()
        try:
            # 捕获输出以免刷屏，只显示结果
            result = subprocess.run(cmd, capture_output=True, text=True)
            task_dur = time.time() - task_start
            
            if result.returncode == 0:
                print(f"      ✅ 完成 (耗时 {task_dur:.1f}s)")
            else:
                print(f"      ❌ 失败 (耗时 {task_dur:.1f}s)")
                print(f"      错误信息:\n{result.stderr}")
                
        except Exception as e:
            print(f"      ❌ 执行异常: {e}")
            
        print("")
        
    total_time = time.time() - start_time_all
    print(f"{'='*60}")
    print(f"所有任务完成! 总耗时: {format_time(total_time)}")
    print(f"输出目录: {OUTPUT_FOLDER}")
    print(f"{'='*60}")

if __name__ == "__main__":
    # 确保输出目录存在
    if not os.path.exists(OUTPUT_FOLDER):
        try:
            os.makedirs(OUTPUT_FOLDER)
        except Exception as e:
            print(f"警告: 无法创建输出目录 {OUTPUT_FOLDER}: {e}")
            
    run()
