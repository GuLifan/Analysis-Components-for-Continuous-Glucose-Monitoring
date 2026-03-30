# CGM Analysis Project (DRQ260122)

#### *GuLifan* (Trae AI Assisted with Gemini-3-Pro-Preview) LAST UPDATED: 2026-02-09 00:39

## 1. 项目简介 (Project Overview)
本项目是一个用于**持续葡萄糖监测 (CGM) 数据处理与分析**的综合工具套件。旨在从原始 CGM 导出数据中提取临床相关的血糖指标，支持批量处理、多模式时间窗口切分、以及复杂的事件统计。

**主要用途：**
*   自动清洗和匹配患者数据。
*   根据临床需求（如住院期间、出院后、胰岛素泵使用期间等）切分数据时间段。
*   计算多种血糖指标（均值、变异系数、TIR/TAR/TBR、MAGE、LBGI/HBGI/ADRR、低血糖事件等）。
*   生成汇总报表或每日详细报表。

---
## 2. 文件结构与功能 (File Structure)

### 核心计算脚本
*   **`01_02_Calculation.py`**: **[核心]** 主计算引擎。读取配置文件或命令行参数，对单个或多个患者数据进行计算，生成汇总结果（每行一个患者）。
*   **`01_03_DailyCalculation.py`**: **[核心]** 每日指标计算脚本。基于 `01_02` 的逻辑，但专注于输出**每日详细数据**（每个 Sheet 为一天，每行一个患者）。适用于需要观察每日变化趋势的场景。

### 批处理与辅助脚本
*   **`01_01_BatchRunner.py`**: **[批处理]** 批量运行工具。允许定义多个数据源任务列表（Tasks），依次调用 `01_02_Calculation.py` 进行大规模数据处理。
*   **`config.yaml`**: **[配置]** 全局配置文件。存储默认的运行参数、路径设置和功能开关。
*   **`02_02_CGMDaysExtractor.py`**: **[统计]** 计算每个CGM文件的总小时数与按24h折算的天数。
*   **`00_06_BatchMerger.py`**: **[合并]** 按工作表同名逐表合并，唯一键为第1/5/6列；重复时保留非空列更多的行；第5/6列强制文本；按第一列升序。

### 预处理工具 (00系列)
*   **`00_01_DataCleaning.py`**: 数据清洗脚本。用于检查患者列表文件（Data Request），剔除信息缺失（如无住院号、无关键时间点）的无效行。
*   **`00_02_FileFetching.py`**: 数据抓取脚本。根据清洗后的患者列表，从海量原始数据文件夹中复制并重命名对应的 CGM Excel 文件到指定目录。支持基于时间重叠的有效性检查。
*   **`00_03_DataMerger.py`**: 数据合并脚本。用于将 `01_02` 或 `01_03` 生成的多个结果文件（如不同批次的输出）合并为一个总表。
*   **`00_04_FixDataFormat.py`**: 数据格式化脚本。用于检查并修复患者列表文件（Data Request）中的格式错误（如手机号格式、时间格式等）。

### 其他脚本
*   **`02_01_FBSExtractor.py`**: (如有) 专门用于提取空腹血糖（FBS）相关指标的独立脚本。

---

## 3. 核心软件详解 (Core Software)

### 3.1 `01_02_Calculation.py` (主计算程序)
这是项目的核心逻辑所在。它包含了一系列标准化的血糖指标计算函数。

**主要函数与定义：**

*   **`calc_basic_stats(df)`**: 计算基础指标。
    *   MEAN (平均血糖), SD (标准差), CV (变异系数), GMI (预估糖化血红蛋白)。
*   **`calc_lbgi_hbgi_adrr(df)`**: 计算风险指标。
    *   LBGI (低血糖风险指数), HBGI (高血糖风险指数), ADRR (平均每日风险范围)。
    *   *注：采用最新标准算法 (Sum of Maxes for ADRR)。*
*   **`calc_modd(df)`**: 计算 MODD (日间血糖平均绝对差)。需要至少 2 天数据。
*   **`calc_mage_daily(glucose_series)`** / **`calc_lage_mage(df)`**: 计算血糖波动指标。
    *   MAGE (平均血糖波动幅度): 基于 Service 1970 定义，仅计算幅度 > 1SD 的有效波动。
    *   LAGE (最大血糖波动幅度): Max - Min。
*   **`calc_range_stats(df)`**: 计算范围指标。
    *   TIR, TAR, TBR, TAR1, TAR2, TBR1, TBR2, TITR, TIR-TITR, GRI。
    *   GRI = (3.0 × TBR2) + (2.4 × TBR1) + (1.6 × TAR2) + (0.8 × TAR1)。
*   **`calc_event_stats(df)`**: 复杂的事件统计。
    *   低/高血糖事件次数及时间。支持简单事件（连续15min）及扩展事件（Extended Events，需判断恢复条件）。
*   **`process_patient_file(...)`**: 单个患者的处理流。负责读取 Excel，根据 `Mode` 切分时间段，调用上述计算函数，并返回结果字典。

### 3.2 `01_03_DailyCalculation.py` (每日计算程序)
*   **用途**: 当需要分析患者**每一天**的具体表现时使用。
*   **特点**: 输出 Excel 包含多个 Sheets (Day 1, Day 2, ...)，每个 Sheet 列出所有患者在该自然日的指标。
*   **逻辑差异**: 强制对齐到自然日（从第1个有数据日期的 00:00 开始切分），确保“Day 1”代表一个绝对的 24 小时自然日区间（即使数据可能缺失）。
*   **列映射**: 支持通过 `config.yaml -> request_columns` 指定前7列字段（hospital_id/pump_start_time/pump_end_time/discharge_time/admission_time/sensor_id/phone_number），无需改动 request 文件结构。

### 3.3 8 种计算模式 (Modes)
用于定义“选取哪一段时间的数据”进行计算：
*   **Mode 0**: 开始时间 ~ 开始时间 + `duringday` 天
*   **Mode 1**: 开始时间 + `interimday` 天 ~ 开始时间 + `duringday` 天
*   **Mode 2**: 开始时间 ~ 结束时间 (全数据)
*   **Mode 3**: 出院时间 - `duringday` 天 ~ 出院时间
*   **Mode 4**: 入院时间 ~ 出院时间 (住院期间)
*   **Mode 5**: 出院时间 ~ 出院时间 + `duringday` 天 (出院后)
*   **Mode 6**: 出院时间 ~ 结束时间 (出院后所有数据)
*   **Mode 7**: 胰岛素泵开始时间 ~ 胰岛素泵结束时间

### 3.4 8 个功能组 (Groups)
*   **Group 1 (Basic)**: 基础统计 (Mean, SD, CV, GMI)。
*   **Group 2 (Risk)**: 风险指数 (LBGI, HBGI, ADRR, MODD)。
*   **Group 3 (Variability)**: 波动性 (MAGE, LAGE)。
*   **Group 4 (Ranges)**: TIR/TAR/TBR 系列。
*   **Group 5 (Hourly)**: 夜间(0-6h)与日间(6-24h)的分段统计。
*   **Group 6 (Events)**: 低血糖/高血糖事件统计 (含扩展事件逻辑)。
*   **Group 7 (Level 2 Hypo)**: 2级低血糖 (<3.0 mmol/L, >15min) 统计。
*   **Group 8 (Conditional)**: 复合低血糖事件 (仅当全局最小值低于阈值时才输出事件，否则为 N/A)。

---

## 4. 配置文件 (`config.yaml`) 含义

```yaml
# --- 匹配设置 ---
match_by: 'sensor_id'       # 匹配关键字：'sensor_id'(探头号), 'phone_number'(手机号), 'hospital_id'(住院号)

# --- 输出设置 ---
daily_output: false         # True: 输出每日数组(不推荐在01_02用); False: 输出汇总均值

# --- 功能组配置 (开关) ---
calc_groups:
  1: true   # 基础指标 (Mean, SD, CV, GMI)
  2: true   # 风险指标 (LBGI, HBGI, ADRR, MODD)
  3: true   # 波动指标 (MAGE, LAGE)
  4: true   # 范围指标 (TIR, TAR, TBR)
  5: true   # 分时段指标 (0-6点, 6-24点)
  6: true   # 事件统计 (Hypo/Hyper Events)
  7: true   # 2级低血糖 (<3.0)
  8: true   # 条件低血糖 (基于全局最小值的统计)

# --- 日期/模式设置 ---
mode: 0           # 默认计算模式 (0-7)
duringday: 7      # 持续天数
interimday: 0     # 间隔天数 (仅Mode 1使用)

# --- 路径设置 ---
patient_list_file: "..."  # 患者列表 Excel 路径
data_folder: "..."        # CGM 数据文件夹路径
output_folder: "..."      # 结果输出路径
# --- Request 前7列映射（支持列名或1-based列号） ---
request_columns:
  hospital_id: 1
  pump_start_time: 2
  pump_end_time: 3
  discharge_time: 4
  admission_time: 5
  sensor_id: 6
  phone_number: 7

# --- 标签设置 ---
datetag: "..."    # 日期标签 (出现在文件名中)
nametag: "..."    # 名称标签 (出现在文件名中)
```

---

## 5. 如何使用 `01_01_BatchRunner.py`

此脚本用于**一键运行多个任务**，避免手动修改 Config 文件。

1.  **打开文件**: 使用 IDE 打开 `01_01_BatchRunner.py`。
2.  **配置全局参数**: 修改 `GLOBAL_MODE`, `GLOBAL_DURING`, `GLOBAL_INTERIM` 等变量，设定统一的计算规则。
3.  **配置任务列表 (`TASKS`)**:
    *   在 `TASKS` 列表中添加字典对象。
    *   每个字典代表一个数据源（例如科内数据、科外数据）。
    *   必须指定：`patient_list` (患者名单), `data_folder` (数据目录), `nametag` (输出文件名前缀)。
    *   可选指定：`match_by` (如科内用 sensor_id，科外用 phone_number)。
4.  **运行**: 直接运行该脚本。
    ```bash
    python 01_01_BatchRunner.py
    ```
5.  **查看结果**: 脚本会自动调用 `01_02` 处理每个任务，并在控制台显示进度。结果将保存在 `OUTPUT_FOLDER` 中。

---
## 6. 兼容性说明（pandas 字符串类型）
- 为兼容 pandas 新的默认字符串 dtype（str），文本列选择同时包含 'object' 与 'str' 两种类型，避免未来版本的选择警告。
- 推荐在环境中安装 pyarrow，以获得更好的字符串列性能（pandas 将在可用时默认使用 pyarrow 字符串后端）。
Copyright (c) 2024-2026 GuLifan, Xi'an Jiaotong University. All Rights Reserved.
