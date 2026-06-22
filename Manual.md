# 20260622 CGM 计算说明

> *谷立帆 2026-06-22*

**计算结果保存于 `00_Results/`**

---

## 1 数据来源与构成

本批次 (DRQ260618) 覆盖 2025 年 6 月至 2026 年 3 月 (2601-2603 批次) 的 CGM 数据。分别由董睿青在2026年3月25日和2026年6月18日发送。

**匹配方式：** 以一次性探头编号 (`sensor_id`) 作为与 CGM 记录文件关联的唯一匹配关键字。

**匹配文件：** `01_Requirements/Matching250601_260331.csv` 与 `01_Requirements/PatientInfo250601_260331.xlsx`，经 `04_01_RequestBuilder.py` 以住院号为键合并生成 `02_OriginalData\RequestList250601_260331.csv`。

**原始数据：** `02_OriginalData/CGMOriginalDataAll_260331/` 包含 787 个患者 CGM 原始文件。

| 批次范围 | 匹配关键字 | 数据目录 | CGM文件数量 | 需要计算数量 |
| :---: | :---: | :--- | :---: | :---: |
| 2025.06 ~ 2025.12 | 探头编号 (sensor_id) | CGMOriginalDataAll_260331 | 563 | / |
| 2025.06 ~ 2026.03 | 探头编号 (sensor_id) | CGMOriginalDataAll_260331 | 223 | / |
| 合计 | 探头编号 (sensor_id) | CGMOriginalDataAll_260331 | 786 | 687 |

---

## 2 文件结构

```txt
CGM_DRQ_260618_Python/
├── README.md                         # 项目说明
├── 00_Results/                       # ← 计算结果
├── 01_Requirements/                  # ← 匹配表 / 患者信息 / 需求文件
└── 02_OriginalData/                  # ← 原始 CGM 数据
    └── CGMOriginalDataAll_260331/    #    787 个患者 xlsx
├── 03_Codes/                         # ← 涉及的代码文件 (默认不发送)
```

---

## 3 计算范围

### 3.1 计算指标 (8 个功能组)

#### (1) 基础指标组 (Basic Indicators Group)

- **MEAN**: 平均血糖 (mmol/L)
- **SD**: 标准差 (mmol/L)
- **CV**: 变异系数 (SD / MEAN)
- **GMI**: 血糖管理指标，基于平均血糖估算的 HbA1c

#### (2) 风险指标组 (Risk Indicators Group)

- **LBGI**: 低血糖风险指数
- **HBGI**: 高血糖风险指数
- **ADRR**: 平均每日风险范围 (采用 Sum of Maxes 标准算法)
- **MODD**: 日间血糖平均绝对差 (需至少 2 天数据，比较相邻日同时间点差值)

#### (3) 血糖波动指标组 (Glucose Variability Indicators Group)

- **MAGE**: 平均血糖波动幅度 (Service 1970 定义，仅计算幅度 > 1SD 的有效波动)
- **LAGE**: 最大血糖波动幅度 (Max - Min)

#### (4) 血糖范围指标组 (Range Group)

- **TIR**: 目标范围内时间 (3.9–10.0 mmol/L)
- **TAR**: 高于目标范围时间 (>10.0 mmol/L)
- **TBR**: 低于目标范围时间 (<3.9 mmol/L)
- **TAR1**: 1 级高血糖 (10.0–13.9 mmol/L)
- **TAR2**: 2 级高血糖 (>13.9 mmol/L)
- **TBR1**: 1 级低血糖 (3.0–3.9 mmol/L)
- **TBR2**: 2 级低血糖 (<3.0 mmol/L)
- **TITR**: 狭窄目标范围内时间 (3.9–7.8 mmol/L)
- **GRI**: 血糖风险指数 = (3.0×TBR2) + (2.4×TBR1) + (1.6×TAR2) + (0.8×TAR1)
- **TIR-TITR**: 广义目标范围与狭窄目标范围之差

#### (5) 分时段指标组 (Hourly Stats Group)

- **MEAN-0TO6AM / MEAN-6AMTO0**: 夜间 (0–6 点) / 日间 (6–24 点) 平均血糖
- **SD-0TO6AM / SD-6AMTO0**: 夜间 / 日间标准差
- **CV-0TO6AM / CV-6AMTO0**: 夜间 / 日间变异系数
- **VV-0TO6AM**: 夜间血糖最低值列表 (每天一个值)
- **VVtime-0TO6AM**: 夜间血糖最低值出现时间列表

#### (6) 事件统计指标组 (Event Stats Group)

- **HYPO**: 低血糖事件次数 (<3.9 mmol/L, 持续 ≥15min)
- **Time-HYPO**: 低血糖事件时间段列表
- **HYPO 0TO6AM**: 夜间低血糖事件次数 (事件开始时间在 0–6 点)
- **Time-HYPO 0TO6AM**: 夜间低血糖事件时间段列表
- **EX HYPO**: 扩展低血糖事件 (<3.9 mmol/L, 持续 >120min, 且恢复需 ≥3.9 持续 15min)
- **Time-EX HYPO**: 扩展低血糖时间段
- **EX HYPO 0TO6AM**: 夜间扩展低血糖事件次数
- **EX HYPER**: 扩展高血糖事件 (>13.9 mmol/L, 持续 >120min, 且恢复需 ≤10.0 持续 15min)
- **Time-EX HYPER**: 扩展高血糖时间段

#### (7) 2 级低血糖指标组 (Level 2 Hypo Group)

- **LV2 HYPO**: 2 级低血糖事件次数 (<3.0 mmol/L, 持续 ≥15min)
- **Time-LV2 HYPO**: 2 级低血糖事件时间段列表
- **LV2 HYPO 0TO6AM**: 夜间 2 级低血糖事件次数
- **Time-LV2 HYPO 0TO6AM**: 夜间 2 级低血糖事件时间段列表

#### (8) 条件低血糖指标组 (Conditional Hypo Group)

*注：本组指标基于"全局最小值"进行判断。若全局最小值 ≥ 阈值，则输出 `#N/A`。*

- **HYPO_COND_3.0**: 低血糖事件次数 (仅当全局最小值 <3.0 时有效)
- **HYPO_COND_3.5**: 低血糖事件次数 (仅当全局最小值 <3.5 时有效)
- 各组均有夜间 (0TO6AM) 子指标和时间列表

### 3.2 附加专项输出

除上述汇总指标外，工具套件还输出：

1. **空腹血糖原始值 (04_06 / 04_07)**：每人每天最接近 06:30 的血糖读数
   - `04_06` — 按自然日历日分组
   - `04_07` — 按连续 24h 滑动窗口分组 (固定 14 天)
2. **每日详细指标 (04_04)**：Mode 0 下逐天计算全部 8 组指标，每个 Sheet 为一天
3. **CGM 时长统计 (04_06)**：每个文件的总小时数与按 24h 折算的天数

---

## 4 时间范围 (8 种计算模式)

| 模式 | 时间范围 | 典型应用 |
| :---: | :--- | :--- |
| **Mode 0** | 数据开始 ~ 开始 + N 天 | D1-3, D1-5, D1-7, D1-14 |
| **Mode 1** | 开始+M 天 ~ 开始+N 天 | 中间稳定期 (如 D6-11, D7-14) |
| **Mode 2** | 数据开始 ~ 数据结束 | 全部数据 |
| **Mode 3** | 出院前 N 天 ~ 出院 | 出院前 3 天 |
| **Mode 4** | 入院 ~ 出院 | 住院期间 |
| **Mode 5** | 出院 ~ 出院后 N 天 | 出院后 1 周 |
| **Mode 6** | 出院 ~ 数据结束 | 出院后全部随访 |
| **Mode 7** | 胰岛素泵开始 ~ 胰岛素泵结束 | 胰岛素泵使用期间 |

---

## 5 代码维护记录

| 日期 | 变更内容 |
| :--- | :--- |
| 2026-06-22 | 项目重构为 04_ 系列脚本命名；拆分 BatchRunner1 (Mode 计算) 与 BatchRunner2 (专项分析)；新增 .gitignore；更新 README/Manual |
| 2026-06-18 | 04_04 新增 argparse 支持 `--duringday` / `--nametag` CLI 参数；04_08 TASKS 格式扩展支持 CLI 参数 |
| 2026-06 | 04_05 重命名为 FBSExtractor_NatureDay 以明确其按自然日分组的特性；04_02 重命名为 BatchRunner1 |
| 2026-03-25 | 重构全部计算代码；新增 Mode 7 (胰岛素泵)；新增 FBS 提取、时长统计、Daily 功能、CGM 天数统计；维护 README 和 requirements.txt |

---

Copyright (c) 2024-2026 GuLifan, Xi'an Jiaotong University. All Rights Reserved.
