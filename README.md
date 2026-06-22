# CGM Analysis Project (DRQ260618)

> *谷立帆 2026-06-22*

## 1. 项目简介 (Project Overview)

本项目是一个用于**持续葡萄糖监测 (CGM) 数据处理与分析**的综合 Python 工具套件。从原始 CGM 导出数据中提取临床相关的血糖指标，支持批量处理、多模式时间窗口切分、以及复杂的事件统计。

**主要用途：**

- 自动清洗和匹配患者数据（住院号 / 探头编号 / 手机号码）
- 根据临床需求（住院期间、出院后、胰岛素泵使用期间等）切分数据时间段
- 计算多种血糖指标（TIR/TAR/TBR、MAGE、LBGI/HBGI/ADRR、低血糖事件等）
- 生成汇总报表或每日详细报表
- 提取空腹血糖 (FBS) 原始值及 CGM 时长统计

**当前批次：** DRQ260618 (2025-06 ~ 2026-03 数据)

---

## 2. 文件结构 (Directory Structure)

```txt
CGM_DRQ_260618_Python/
├── 04_01_RequestBuilder.py          # 患者请求列表构建（住院号匹配患者信息）
├── 04_02_BatchRunner1.py           # 批量并行计算器（Mode 0-7，多进程）
├── 04_03_Calculation.py            # ★ 核心计算引擎（1531 行，8 模式 × 8 功能组）
├── 04_04_DailyCalculation.py       # Mode 0 每日详细输出（每 Sheet 为一天，每行一个患者）
├── 04_05_FBSExtractor_NatureDay.py # 空腹血糖提取 — 按自然日历日分组
├── 04_06_CGMDaysExtractor.py       # CGM 文件时长统计（总小时 / 天数）
├── 04_07_FBSExtractor_24H.py       # 空腹血糖提取 — 按 24h 滑动窗口分组（最多 14 天）
├── 04_08_BatchRunner2.py           # 批量串行运行器（04_04 ~ 04_07）
├── config.yaml                      # 全局配置文件
├── requirements.txt                 # Python 依赖清单
├── README.md                        # 本文件
├── Manual.md                        # 计算说明与指标定义
│
├── 00_Results/                      # ← 计算结果输出目录
├── 01_Requirements/                 # ← 需求文件 / 匹配表 / 患者信息
└── 02_OriginalData/                 # ← 原始 CGM 数据文件
    └── CGMOriginalDataAll_260331/   #    787 个患者 .xlsx 文件
```

---

## 3. 脚本详解 (Script Reference)

### 3.1 数据准备

**`04_01_RequestBuilder.py`** — 患者请求列表构建

- 以住院号为唯一匹配键，将 `MatchingRelationship.csv` 与 `患者信息.xlsx` 合并
- 左连接：匹配不到的填空白
- 输出 `requestall.csv`（含 14 列：住院号、入院/出院日期、探头编号、胰岛素泵时间、德谷/甘精时间等）
- 自动清洗日期格式与手机号码

### 3.2 核心计算

**`04_03_Calculation.py`** — ★ 核心计算引擎

- 全部 8 种模式 (Mode 0-7) 和 8 个功能组 (Group 1-8) 的计算逻辑
- 通过 argparse 接受 CLI 参数: `--mode`, `--during`, `--nametag`, `--interim`, `--datetag`, `--patient_list`, `--data_folder`, `--output_folder`, `--match_by`, `--daily_output`, `--calc_groups`
- CLI 参数优先级高于 config.yaml，用于批量运行时的参数覆盖
- 生成汇总 Excel：每行一个患者，每列一个指标

**`04_04_DailyCalculation.py`** — Mode 0 每日详细输出

- 仅支持 Mode 0（数据开始起 N 天），但按 24h 自然日逐天输出
- 输出 Excel 包含 N 个 Sheet (Day 1 ~ Day N)，每个 Sheet 列出所有患者当天的指标
- 数据不足半天的 Sheet 自动跳过
- CLI 参数: `--duringday`, `--nametag`（覆盖 config 对应值）

### 3.3 专项分析

**`04_05_FBSExtractor_NatureDay.py`** — 空腹血糖提取（自然日历日）

- 每人每天取最接近 06:30 的血糖原始值
- 按 `df['timestamp'].dt.date`（日历日）分组
- 输出列数动态（取决于数据最早和最晚日期跨度）

**`04_07_FBSExtractor_24H.py`** — 空腹血糖提取（24h 滑动窗口）

- 每人每天取最接近 06:30 的血糖原始值
- 按从数据起点 `t0` 起算的连续 24h 窗口分组，固定 14 天
- 若窗口内 06:30 早于窗口起点，则顺延至下一天
- 与 04_05 的区别：日历日 vs 连续 24h 窗口的对齐方式不同

**`04_06_CGMDaysExtractor.py`** — CGM 文件时长统计

- 计算每个 CGM 文件的总时长（小时）及按 24h 折算的天数（四舍五入）

### 3.4 批量运行器

**`04_02_BatchRunner1.py`** — 批量并行计算器

- 使用 `ProcessPoolExecutor` (max_workers=4) 并行调用 `04_03_Calculation.py`
- 覆盖全部 8 种模式 (Mode 0-7)，每种模式可配置不同的 `duringday`/`nametag`
- 显示实时进度和 ETA 预估

**`04_08_BatchRunner2.py`** — 批量串行运行器

- 顺序调用 `04_04_DailyCalculation.py`、`04_05_FBSExtractor_NatureDay.py`、`04_06_CGMDaysExtractor.py`、`04_07_FBSExtractor_24H.py`
- 通过行级字符串替换临时修改 `config.yaml`，任务完成后恢复
- TASKS 格式支持可选的 CLI 参数字符串

---

## 4. 核心计算特性 (Core Features)

### 4.1 8 种计算模式 (Modes)

用于定义"选取哪一段时间的数据"进行计算：

| 模式 | 时间范围 | 典型应用 |
| :---: | :--- | :--- |
| **Mode 0** | 数据开始 ~ 开始 + N 天 | D1-3, D1-5, D1-7, D1-14 |
| **Mode 1** | 数据开始+M 天 ~ 开始+N 天 | 中间稳定期 (D6-11, D7-14) |
| **Mode 2** | 全部数据范围 | 全程统计 |
| **Mode 3** | 出院前 N 天 ~ 出院 | 出院前 3 天 |
| **Mode 4** | 入院 ~ 出院 | 住院期间 |
| **Mode 5** | 出院 ~ 出院后 N 天 | 出院后短期随访 |
| **Mode 6** | 出院 ~ 数据结束 | 出院后全部随访 |
| **Mode 7** | 胰岛素泵开始 ~ 胰岛素泵结束 | 胰岛素泵使用期间 |

### 4.2 8 个功能组 (Groups)

| 组 | 名称 | 指标 |
| :---: | :--- | :--- |
| **1** | 基础统计 | MEAN, SD, CV, GMI |
| **2** | 风险指数 | LBGI, HBGI, ADRR, MODD |
| **3** | 波动性 | MAGE, LAGE |
| **4** | 范围指标 | TIR, TAR, TBR, TAR1, TAR2, TBR1, TBR2, TITR, GRI, TIR-TITR |
| **5** | 分时段 | 夜间 (0-6AM) 与 日间 (6AM-0) 的均值、SD、CV、最低值列表 |
| **6** | 事件统计 | 低血糖事件, 扩展低血糖, 扩展高血糖 (含时间记录) |
| **7** | 2 级低血糖 | <3.0 mmol/L, ≥15min 事件统计 |
| **8** | 条件低血糖 | 仅当全局最低值 < 阈值 (3.0/3.5) 才输出事件，否则为 #N/A |

---

## 5. 配置文件 (`config.yaml`)

```yaml
# --- 匹配设置 ---
match_by: sensor_id       # 匹配关键字: sensor_id / phone_number / hospital_id

# --- 输出设置 ---
daily_output: true         # true: 输出每日数组; false: 输出汇总均值

# --- 功能组开关 ---
calc_groups:
  1: true    # 基础指标
  2: true    # 风险指标
  3: true    # 波动指标
  4: true    # 范围指标
  5: true    # 分时段指标
  6: true    # 事件统计
  7: true    # 2级低血糖
  8: true    # 条件低血糖

# --- 模式与天数 ---
mode: 0                    # 默认计算模式 (0-7)
duringday: 14              # 持续天数
interimday: 8              # 间隔天数 (Mode 1 专用)

# --- 路径 (指向本项目或外部 R 项目) ---
patient_list_file: D:\...\requestall.csv
data_folder: D:\...\CGMOriginalDataAll_260331
output_folder: D:\...\00_Results

# --- 文件标签 ---
datetag: DRQ260618
nametag: Daily_D1-14
output_filename_template: CGM_{datetag}_{nametag}_Mode{mode}_Results.xlsx
```

---

## 6. 使用指南 (Usage)

### 6.1 环境准备

```bash
pip install -r requirements.txt
```

### 6.2 生成患者请求列表

```bash
python 04_01_RequestBuilder.py
```

### 6.3 批量汇总计算 (Mode 0-7)

直接运行 `04_02_BatchRunner1.py`，或在 IDE 中勾选需要的 TASKS 项：

```bash
python 04_02_BatchRunner1.py
```

### 6.4 每日详细输出 + 专项分析

直接运行 `04_08_BatchRunner2.py`：

```bash
python 04_08_BatchRunner2.py
```

该脚本依次执行:

1. `04_04_DailyCalculation.py` — duringday=3 和 duringday=14 两个变体
2. `04_05_FBSExtractor_NatureDay.py` — 空腹血糖（自然日）
3. `04_06_CGMDaysExtractor.py` — 时长统计
4. `04_07_FBSExtractor_24H.py` — 空腹血糖（24h 窗口）

### 6.5 单个脚本运行

每个 04_ 脚本均可独立运行，参数从 `config.yaml` 读取（部分支持 CLI 覆盖）：

```bash
python 04_03_Calculation.py --mode 7 --during 14 --nametag InsulinPump
python 04_04_DailyCalculation.py --duringday 3 --nametag Daily_D1-3
```

---

## 7. 依赖 (Dependencies)

| 包 | 用途 |
| :--- | :--- |
| `pandas` | 数据处理与表格操作 |
| `numpy` | 数值计算 |
| `openpyxl` | Excel 文件读写 |
| `PyYAML` | 配置文件解析 |
| `pyarrow` | pandas 字符串后端性能优化（推荐） |

---

## 8. 兼容性说明

- 为兼容 pandas 新的默认字符串 dtype，文本列选择同时包含 `object` 与 `str` 类型
- 推荐安装 `pyarrow` 以获得更好的字符串列性能
- `04_04_DailyCalculation.py` 的 `load_and_sample()` 使用 calamine 引擎作为首选 (fallback: openpyxl)
- 血糖值范围限制为 [1.8, 33.3] mmol/L

---

Copyright (c) 2024-2026 GuLifan, Xi'an Jiaotong University. All Rights Reserved.
