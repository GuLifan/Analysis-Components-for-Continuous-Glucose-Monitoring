
# 20260325 CGM计算说明
#### *谷立帆* *2026-03-25* *DRQ260122批次更正后，代号DRQ260122Nova*

**注意：计算结果保存在"~DRQ260122Nova\00_Results\All"**

## 1 数据来源和构成
本次原始数据仍然来自强薇老师前一次打包发送的2304例（截至2024年11月30日）及2025年6月董睿青老师发送的558例（截至2025年5月31日）。

这两个批次分别以患者登记的手机号码和佩戴的探头号码作为与CGM记录文件相关联的唯一匹配关键字。 以2024-06-28、2024-11-30为分界线，每个名录下分别计算了三个时期的结果。

由于是手工区分，在日期衔接处有所重叠，**准确数值请看第4行“合计”，已做去重处理**。


| 批次  | 日期范围         |  唯一匹配字符  |  标识  | 排除后科内病例数 IN | 排除后科外病例数 EX | 胰岛素泵 IN | 胰岛素泵 EX |
|:----:|:----------------------|:-----------:|:-------:|:-------------:|:-------------:|:-------------:|:-------------:|
|  1   |  2022-2024.11.30      |   手机号码  | “2412” |      1014     |      397      |      535      |      182      |
|  2   | 2024.12.01-2025.05.31 |   探头编号  | “2505” |      339      |      185     |      176      |      58      |
| 合计  | 2022-2025.05.31       |      /     |    /   |     1329      |      563      |      694      |      234      |




## 2 文件结构


├── DRQ260122Nova

│   ├── README.pdf 本文件

│   ├── 00_Results 包含全部计算结果

│   │   ├── All 条目相同，时间范围不同文件合并后的结果

│   │   ├── Daily_Results 科内患者1-14天每日的计算结果以及科外患者1-3天每日的计算结果

│   │   ├── FBS_Separated 科内科外每人每日06:30的血糖原始值

│   │   ├── Days_Extraction 每个CGM文件包含的总时长（小时）以及按 24h 为一自然日的天数（从数据开始时间起算）

│   │   └── ……

│   ├── 01_requirements 董老师发送的要求性文件 *(默认不发送，有需求可说明)*

│   ├── 02_OriginalData 强老师前后两批原始数据 *(默认不发送，有需求可说明)*

│   │   ├── CGMOriginalDataAll_241130

│   │   └── CGMOriginalDataAll_250531

│   ├── 03_Requests 记录了需要被抓取、计算的文件信息

│   ├── 04_SelectedData 包含了筛选后的数据

│   │   ├── DataSelected2412IN

│   │   ├── DataSelected2412EX

│   │   ├── DataSelected2505IN

│   │   ├── DataSelected2505EX

│   │   ├── DataSelected2412IN_Pump

│   │   ├── DataSelected2412EX_Pump

│   │   ├── DataSelected2505IN_Pump

│   │   └── DataSelected2505EX_Pump

│   └── 05_Codes 包含项目代码和维护历史 *(默认不发送，有需求可说明)*


## 3 计算范围
#### 3.1 计算指标 (分为8个功能组，新增7和8)

**(1) 基础指标组 (Basic Indicators Group)**
- MEAN: 平均血糖 (mmol/L)
- SD: 标准差 (mmol/L)
- CV: 变异系数 (SD/MEAN)
- GMI: 血糖管理指标 (基于平均血糖估算的HbA1c)

**(2) 风险指标组 (Risk Indicators Group)**
- LBGI: 低血糖风险指数
- HBGI: 高血糖风险指数
- ADRR: 平均每日风险范围 (采用Sum of Maxes标准算法)
- MODD: 日间血糖平均绝对差 (需至少2天数据)

**(3) 血糖波动指标组 (Glucose Variability Indicators Group)**
- MAGE: 平均血糖波动幅度 (Service 1970定义，仅计算>1SD的有效波动)
- LAGE: 最大血糖波动幅度 (Max - Min)

**(4) 血糖范围指标组 (Range Group)**
- TIR: 目标范围内时间 (3.9-10.0 mmol/L)
- TAR: 高于目标范围时间 (>10.0 mmol/L)
- TBR: 低于目标范围时间 (<3.9 mmol/L)
- TAR1: 1级高血糖 (10.0-13.9 mmol/L)
- TAR2: 2级高血糖 (>13.9 mmol/L)
- TBR1: 1级低血糖 (3.0-3.9 mmol/L)
- TBR2: 2级低血糖 (<3.0 mmol/L)
- TITR: 狭窄目标范围内时间 (3.9-7.8 mmol/L)
- TIR-TITR: 广义目标范围与狭窄目标范围之差
- GRI: GRI=(3.0×TBR2)+(2.4×TBR1)+(1.6×TAR2)+(0.8×TAR1)

**(5) 分时段指标组 (Hourly Stats Group)**
- MEAN-0TO6AM / MEAN-6AMTO0: 夜间(0-6点)/日间(6-24点)平均血糖
- SD-0TO6AM / SD-6AMTO0: 夜间/日间标准差
- CV-0TO6AM / CV-6AMTO0: 夜间/日间变异系数
- VV-0TO6AM: 夜间血糖最低值列表
- VVtime-0TO6AM: 夜间血糖最低值出现时间列表

**(6) 事件统计指标组 (Event Stats Group)**
- HYPO: 低血糖事件次数 (<3.9 mmol/L, 持续≥15min)
- Time-HYPO: 低血糖事件时间段列表
- HYPO 0TO6AM: 夜间低血糖事件次数 (开始时间在0-6点)
- Time-HYPO 0TO6AM: 夜间低血糖事件时间段列表
- EX HYPO: 扩展低血糖事件次数 (<3.9, 持续>120min, 且恢复需满足≥3.9持续15min)
- EX HYPER: 扩展高血糖事件次数 (>13.9, 持续>120min, 且恢复需满足≤10.0持续15min)

**(7) 2级低血糖指标组 (Level 2 Hypo Group)**
- LV2 HYPO: 2级低血糖事件次数 (<3.0 mmol/L, 持续≥15min)
- Time-LV2 HYPO: 2级低血糖事件时间段列表
- LV2 HYPO 0TO6AM: 夜间2级低血糖事件次数
- Time-LV2 HYPO 0TO6AM: 夜间2级低血糖事件时间段列表

**(8) 条件低血糖指标组 (Conditional Hypo Group)**
*注：本组指标基于“全局最小值”进行判断。若全局最小值高于阈值，则输出 #N/A。*
- HYPO_COND_3.0: 次数A (低血糖事件且全局最小值<3.0)
- Time-HYPO_COND_3.0: 时间A
- HYPO_COND_3.0 0TO6AM: 夜间次数A (夜间事件且全局最小值<3.0)
- Time-HYPO_COND_3.0 0TO6AM: 夜间时间A
- HYPO_COND_3.5: 次数B (低血糖事件且全局最小值<3.5)
- Time-HYPO_COND_3.5: 时间B
- HYPO_COND_3.5 0TO6AM: 夜间次数B (夜间事件且全局最小值<3.5)
- Time-HYPO_COND_3.5 0TO6AM: 夜间时间B


#### 3.2 时间范围（新增mode7）

| 模式 ID | 描述 | 应用场景 |
| :--- | :--- | :--- |
| **Mode 0** | **数据开始 + N天** | D1-3、D1-5、D1-7、D1-14|
| **Mode 1** | **中间段 (Start+M ~ Start+N)** | D1-3、D1-5、D1-7、D1-14 |
| **Mode 2** | **全程 (Full Range)** | Day alltime |
| **Mode 3** | **出院前 N天** | 对应出院前3天 |
| **Mode 4** | **入院 ~ 出院** | 对应住院期间 |
| **Mode 5** | **出院后 N天** | 对应出院后1周 |
| **Mode 6** | **出院 ~ 结束** | 对应出院后全部时间 |
| **Mode 7** | **胰岛素泵使用期间** | 对应胰岛素泵期间 |


**此外，还包括:**

(1) 新增每个患者每天06:30的血糖原始值（单列文件）;

(2) 科内患者D1-14每天分别的数值，科外患者D1-3每日分别的数值。

(3) 每个患者的血糖的天数（单列文件）


## 3 代码维护

(1) 重构了全部计算代码;

(2) 新增了mode7的计算及相关工具兼容性优化;

(3) 新增了FBS提取功能与时长提取功能；

(4) 新增了Daily功能；新增 CGM 天数统计工具；

(5) 维护了README.md以及requirements.txt文件；新增 request_columns 配置说明；完善批量合并工具文档。

---
Copyright (c) 2024-2026 GuLifan, Xi'an Jiaotong University. All Rights Reserved.
