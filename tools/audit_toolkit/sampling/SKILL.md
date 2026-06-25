---
name: audit-sampling
description: 审计抽样引擎——PPS/MUS货币单位抽样、属性抽样、分层抽样、随机抽样，符合《中国注册会计师审计准则第1314号》。触发词：抽样、抽凭、PPS、MUS、货币单位抽样、随机抽样、分层抽样、属性抽样、样本量计算、按凭证抽样。
---

# 审计抽样引擎

## 概述

基于《中国注册会计师审计准则第1314号——审计抽样》（CSA 1314）和 AICPA Audit Guide: Audit Sampling 实现的四种审计抽样方法。适用于实质性程序和控制测试中的统计抽样。

## 触发条件

用户提到以下任一关键词时触发：
- "抽样" / "抽凭" / "凭证抽样"
- "PPS" / "MUS" / "货币单位抽样"
- "属性抽样" / "控制测试抽样"
- "分层抽样" / "随机抽样"
- "样本量计算" / "样本容量"
- "审计抽样" / "统计抽样"

## 抽样方法选择指南

| 测试类型 | 推荐方法 | 适用场景 | 关键输入 |
|---------|---------|---------|---------|
| **实质性程序** | PPS/MUS | 金额测试（应收账款、存货、收入） | 可容忍错报、预期错报、误受风险 |
| **控制测试** | 属性抽样 | 内部控制有效性（审批、授权、复核） | 置信水平、可容忍偏差率 |
| **实质性+分层** | 分层抽样 | 金额波动大的总体（按金额分层） | 分层数、金额列 |
| **通用/快速** | 简单随机 | 无金额信息的凭证抽查 | 样本量 |

### 决策树

```
测试目的？
├── 控制测试 → 属性抽样 (--method attribute)
└── 实质性程序
    ├── 有金额信息？
    │   ├── 是 → PPS/MUS (--method pps)
    │   └── 金额波动大 → 分层抽样 (--method stratified)
    └── 无金额 → 简单随机 (--method random)
```

## 方法详解

### 1. PPS/MUS 货币单位抽样

**原理**：金额越大的记录被选中的概率越高。每条记录的选中概率与其金额成正比（Probability Proportional to Size）。

**样本量公式**：
```
n = (BV × RF) / (TM - EM × EF)
```
- BV = 总体账面值
- RF = 可靠性因子（基于误受风险）
- TM = 可容忍错报
- EM = 预期错报
- EF = 扩展因子（仅 EM > 0 时使用）

**可靠性因子 RF（零预期错报）**：
| 风险 | 1% | 5% | 10% | 15% | 20% | 25% | 30% |
|------|-----|-----|------|------|------|------|------|
| RF | 4.61 | 3.00 | 2.31 | 1.90 | 1.61 | 1.39 | 1.21 |

**选样方式**：系统性等距选样 — 计算抽样间隔 = BV/n，随机起点，按累计金额扫描选中。

### 2. 属性抽样

**原理**：基于二项分布，查 AICPA 标准样本量表确定样本量。用于评估总体偏差率是否在可接受范围内。

**支持的参数**：
- 置信水平：90%、95%、99%
- 可容忍偏差率：2% ~ 20%（含常用档位）
- 预期偏差率：0% ~ 5%（含常用档位）

### 3. 分层抽样

**原理**：按金额排序后分为 N 层，使用 **Neyman 最优分配** 将总样本量分配到各层（标准差大的层分配更多样本），各层内独立随机选样。

```
n_h = n × (N_h × σ_h) / Σ(N_i × σ_i)
```

### 4. 简单随机抽样

**原理**：无放回随机选样，每条记录等概率被选中。样本量按经验公式 `√N × 0.8` 计算。

## 使用示例

```bash
# ---- PPS/MUS 货币单位抽样（实质性程序）----

# 基本用法：自动计算样本量
python sampling_engine.py --input "凭证.xlsx" --method pps \
    --amount-col "借方金额" --tolerable 50000 --risk 10

# 指定关键列和自定义种子
python sampling_engine.py --input "应收账款.xlsx" --method pps \
    --sheet "明细表" --key-col "凭证号" --amount-col "余额" \
    --tolerable 100000 --expected 5000 --risk 5 --seed 2024

# 手动指定样本量（覆盖自动计算）
python sampling_engine.py --input "收入.xlsx" --method pps \
    --amount-col "金额" --tolerable 80000 --sample-size 60

# 指定输出路径
python sampling_engine.py --input "数据.xlsx" --method pps \
    --amount-col "金额" --tolerable 50000 \
    --output "D:/审计输出/抽样结果.xlsx"


# ---- 属性抽样（控制测试）----

# 95%置信水平，5%可容忍偏差率
python sampling_engine.py --input "付款审批.xlsx" --method attribute \
    --confidence 95 --tolerable-rate 0.05 --expected-rate 0.01

# 90%置信水平，10%可容忍偏差率
python sampling_engine.py --input "合同审批.xlsx" --method attribute \
    --confidence 90 --tolerable-rate 0.10


# ---- 分层抽样 ----

# 3层自动分层
python sampling_engine.py --input "应付账款.xlsx" --method stratified \
    --amount-col "余额" --strata 3

# 5层分层，指定样本量
python sampling_engine.py --input "存货.xlsx" --method stratified \
    --amount-col "金额" --strata 5 --sample-size 80


# ---- 简单随机抽样 ----

# 自动计算样本量
python sampling_engine.py --input "全部凭证.xlsx" --method random

# 指定样本量
python sampling_engine.py --input "费用报销.xlsx" --method random \
    --sample-size 50 --seed 2024
```

## 输出说明

生成的 Excel 文件含两个 Sheet：

### Sheet 1：抽样清单
- 选中的原始记录（保留所有原始列）
- 额外列：**抽样序号**（选中顺序）、**选中金额**（金额列值）

### Sheet 2：抽样参数
- 方法名称、输入文件、工作表
- 总体记录数、总体金额
- 计算样本量、实际选中数、唯一样本数
- 随机种子、执行时间
- PPS：可容忍错报、预期错报、误受风险、RF、EF、抽样间隔
- 属性抽样：置信水平、可容忍偏差率、预期偏差率
- 分层抽样：分层数、分配方式
- 方法公式说明、参考标准

**审计轨迹完整性**：抽样参数 Sheet 记录了所有输入参数和计算过程，可作为审计底稿存档依据。

## 数据处理与注意事项

### 文件格式
- 输入：Excel (.xlsx / .xlsm)
- 支持指定 Sheet 名和表头行号
- 金额列自动处理空值（视为 0）、非数值（视为 0）

### 错误处理
引擎对以下情况提供明确中文错误提示：
- 文件不存在
- Sheet 不存在
- 列名找不到（列出可用列名）
- PPS 分母 ≤ 0（建议全量检查）
- 样本量超过总体
- 不支持的参数组合

### 随机可复现
通过 `--seed` 参数固定随机种子（默认 42），确保同一输入产生完全相同的抽样结果，满足审计可复现性要求。

## 参考标准

- 《中国注册会计师审计准则第1314号——审计抽样》
- 《中国注册会计师审计准则第1314号——审计抽样》应用指南
- AICPA Audit Guide: Audit Sampling (2017)
- AICPA Audit Sampling — Appendix A: Attributes Statistical Sampling Tables
- ISA 530: Audit Sampling

## 依赖

- **openpyxl**: Excel 读写（必须）
- **numpy**: 分层标准差计算（可选，无 numpy 时使用纯 Python 实现）

安装：
```bash
pip install openpyxl
# 可选
pip install numpy
```
