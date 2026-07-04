# 税审底稿自动填充引擎 — 任务计划

## 目标
基于规则理解，构建税审计算引擎+底稿生成器。不硬解原有模板，而是再造一套结构化底稿。

## 范围
第一期：企业所得税汇缴（1-1 独立纳税企业）

## 架构
```
输入数据 → 规则引擎(50+条税法规则) → 计算结果 → 底稿生成器 → 结构化Excel
```

## 阶段

### Phase 1: 数据模型 ✅
- 定义输入结构：TB、资产台账、薪酬明细、研发项目等
- 文件：`tax_audit_engine/models.py`

### Phase 2: 规则引擎核心
- 实现50+条纳税调整规则
- 按税法公式计算：账载→税收→调增/调减
- 文件：`tax_audit_engine/rules.py`

### Phase 3: 计算器
- 规则编排、依赖解析、批量执行
- 文件：`tax_audit_engine/calculator.py`

### Phase 4: 底稿生成器
- 按底稿逻辑生成结构化Excel
- 审定表、纳税调整表、测算表、汇总表
- 文件：`tax_audit_engine/workpaper_generator.py`

### Phase 5: 集成与测试
- 示例数据、端到端运行
- 文件：`tax_audit_engine/main.py`

## 状态
Phase 1: 完成 ✅
Phase 2: 完成 ✅（25条规则覆盖收入/扣除/资产/优惠四大类）
Phase 3: 完成 ✅
Phase 4: 完成 ✅（9张sheet：封面/利润表审定/收入调整/扣除调整/资产调整/优惠/折旧测算/汇总/税额计算）
Phase 5: 完成 ✅

## 产出物
| 文件 | 说明 |
|------|------|
| `tax_audit_engine/models.py` | 输入/输出数据模型（EnterpriseInfo, TrialBalance, AssetItem, TaxAdjustment, CalculationResult） |
| `tax_audit_engine/rules.py` | 25条纳税调整规则，按税法公式计算（RuleContext + TaxRule体系） |
| `tax_audit_engine/calculator.py` | 计算引擎：规则编排→应纳税所得额→税额计算（含小型微利/高新税率） |
| `tax_audit_engine/workpaper_generator.py` | 底稿生成器：openpyxl输出结构化Excel，无合并单元格 |
| `tax_audit_engine/main.py` | 入口+示例数据（制造业企业，9张底稿表） |

## 验证结果
- 示例数据计算通过：会计利润720万→应纳税所得额535.8万→高新税率15%→应纳80.37万
- 底稿已生成至桌面：`D:\Users\12844\Desktop\税审底稿_示例.xlsx`（9张sheet）
