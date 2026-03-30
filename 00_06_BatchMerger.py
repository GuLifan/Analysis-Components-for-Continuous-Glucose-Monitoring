# 00_06_BatchMerger.py
# 批量合并指定规则下的Excel结果文件
# 规则说明：
# 1) 以每个工作表的每行的第1、5、6列作为唯一标识；当出现重复（1/5/6列完全相同）时，保留“非空列数更多”的那一行
# 2) 一个文件里若存在多个工作表，只有当工作表名称完全一致时，才逐个工作表进行合并
# 3) 文件名结构为：CGM_批次头_第二段_第三段_第四段_Results.xlsx
#    - 仅当“批次头”完全一致，第二段仅字母部分一致（数字部分不同可视为同组），第三、第四段完全一致时，两文件可合并
#    - 例如：CGM_DRQ260122Nova_2412EX_D1-3_Mode0_Results.xlsx 与 CGM_DRQ260122Nova_2505EX_D1-3_Mode0_Results.xlsx 可合并
#      而 CGM_DRQ260122Nova_2412IN_D7afterDischarge_Mode5_Results.xlsx 与 CGM_DRQ260122Nova_2412EX_D7afterDischarge_Mode5_Results.xlsx 不可合并
# 4) 合并后输出文件位于桌面目录（硬编码）
# 5) 输出文件名删除末尾的“_Results”，并删除第二段中的数字，仅保留字母；示例：…_EX_D1-3_Mode0.xlsx
# 6) 合并结果的第5、6列强制为文本（string），避免电话/设备号被Excel显示为科学计数法
# 7) 合并完成后按第一列（住院号）升序排列
# LAST UPDATE BY LIFANGU IN 20260325

import os
import re
import pandas as pd
import numpy as np
import yaml

# --- 配置区域 ---
# 扫描待合并结果文件的目录
OUTPUTS_DIR = r"C:\\Users\\lifan\\Desktop\\00_Outputs"
# 输出合并文件的保存目录（桌面/00_Results）
DESKTOP_DIR = r"C:\\Users\\lifan\\Desktop"
# 批次头，作为第一段共同前缀，用于筛选同组文件
PREFIX_HEAD = "CGM_DRQ260122Nova"
# 唯一标识列（1-based列号）：第1、5、6列
DEDUP_KEY_COLS_1BASED = [1, 5, 6]

def load_config(config_path):
    """读取配置文件（YAML），失败返回空dict"""
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}
    except Exception:
        return {}

def find_config():
    """尝试从当前目录或上级目录寻找config.yaml"""
    here = os.path.dirname(__file__)
    cands = [
        os.path.join(here, 'config.yaml'),
        os.path.abspath(os.path.join(here, '..', 'config.yaml'))
    ]
    for p in cands:
        cfg = load_config(p)
        if cfg:
            return cfg
    return {}

CONFIG = find_config()

def parse_filename(fname):
    """
    解析文件名结构：CGM_批次头_第二段_第三段_第四段_Results.xlsx
    返回：{'head': 批次头, 'second_letters': 第二段的字母部分, 'third': 第三段, 'fourth': 第四段}
    若不匹配返回 None
    """
    m = re.match(r'^(CGM_[^_]+)_([^_]+)_([^_]+)_([^_]+)_Results\.xlsx$', fname)
    if not m:
        return None
    head = m.group(1)
    second = m.group(2)
    third = m.group(3)
    fourth = m.group(4)
    m2 = re.match(r'^(\d+)?([A-Za-z]+)$', second)
    if not m2:
        return None
    letters = m2.group(2)
    return {'head': head, 'second_letters': letters, 'third': third, 'fourth': fourth}

def list_candidate_files():
    """列出OUTPUTS_DIR下符合命名规范且批次头匹配的候选文件"""
    files = []
    for name in os.listdir(OUTPUTS_DIR):
        if name.lower().endswith('.xlsx'):
            info = parse_filename(name)
            if info and info['head'] == PREFIX_HEAD:
                files.append((name, info))
    return files

def group_files(files):
    """
    按“批次头 + 第二段字母 + 第三段 + 第四段”分组
    第二段数字不同但字母相同视为同组（如2412EX与2505EX同组）
    """
    groups = {}
    for name, info in files:
        key = (info['head'], info['second_letters'], info['third'], info['fourth'])
        groups.setdefault(key, []).append(os.path.join(OUTPUTS_DIR, name))
    return groups

def _resolve_key_columns(df, key_cols_1based):
    """将1-based列号转换为DataFrame中的列名集合；若无效则回退为第一列"""
    cols = []
    for c in key_cols_1based:
        idx = int(c) - 1
        if 0 <= idx < df.shape[1]:
            cols.append(df.columns[idx])
    if not cols:
        cols = [df.columns[0]]
    return cols

def _completeness_score(df):
    """
    计算每行的“完整度分数”= 非缺失单元格数量
    缺失判断：NaN、空字符串/空格、'#N/A'、'#N/A N/A'、'N/A'
    """
    tmp = df.copy()
    obj = tmp.select_dtypes(include=['object', 'str']).columns
    if len(obj) > 0:
        tmp[obj] = tmp[obj].replace(r'^\s*$', np.nan, regex=True)
        tmp[obj] = tmp[obj].replace({'#N/A': np.nan, '#N/A N/A': np.nan, 'N/A': np.nan})
    return tmp.notna().sum(axis=1)

def _common_sheets(file_paths):
    """获取多个文件的工作表交集，仅交集中的表进行合并"""
    sets = []
    for fp in file_paths:
        try:
            xls = pd.ExcelFile(fp)
            sets.append(set(xls.sheet_names))
        except Exception:
            sets.append(set())
    if not sets:
        return []
    common = sets[0]
    for s in sets[1:]:
        common = common & s
    return sorted(common)

def _merge_one_sheet(sheet_name, file_paths, key_cols_1based):
    """
    合并同名工作表数据：
    1) 纵向拼接
    2) 依据唯一标识列排序并按“完整度分数”降序，保留最完整的行
    3) 第5、6列转为string文本类型
    4) 最终按第一列（住院号）升序排序
    """
    dfs = []
    for fp in file_paths:
        try:
            df = pd.read_excel(fp, sheet_name=sheet_name)
            dfs.append(df)
        except Exception:
            continue
    if not dfs:
        return pd.DataFrame()
    merged = pd.concat(dfs, ignore_index=True)
    if merged.empty:
        return merged
    key_cols = _resolve_key_columns(merged, key_cols_1based)
    merged['__score__'] = _completeness_score(merged)
    merged = merged.sort_values(by=key_cols + ['__score__'], ascending=[True]*len(key_cols) + [False])
    merged = merged.drop_duplicates(subset=key_cols, keep='first')
    merged = merged.drop(columns=['__score__'], errors='ignore')
    if merged.shape[1] >= 6:
        merged.iloc[:, 4] = merged.iloc[:, 4].astype('string')
        merged.iloc[:, 5] = merged.iloc[:, 5].astype('string')
    if merged.shape[1] >= 1:
        merged = merged.sort_values(by=merged.columns[0], ascending=True)
    return merged

def make_output_name(key):
    """
    生成输出文件名：删除“_Results”，同时第二段仅保留字母部分（去掉数字）
    结构：CGM_批次头_第二段字母_第三段_第四段.xlsx
    输出到桌面目录（DESKTOP_DIR）
    """
    head, letters, third, fourth = key
    base = f"{head}_{letters}_{third}_{fourth}.xlsx"
    return os.path.join(DESKTOP_DIR, base)

def merge_groups(groups):
    """对每个分组进行工作表交集合并并输出Excel"""
    for key, files in groups.items():
        if len(files) < 2:
            continue
        sheets = _common_sheets(files)
        if not sheets:
            continue
        out_path = make_output_name(key)
        with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
            for s in sheets:
                df = _merge_one_sheet(s, files, DEDUP_KEY_COLS_1BASED)
                df.to_excel(writer, sheet_name=s, index=False)

def main():
    """主入口：扫描候选文件、分组并执行合并"""
    files = list_candidate_files()
    groups = group_files(files)
    merge_groups(groups)

if __name__ == "__main__":
    main()

