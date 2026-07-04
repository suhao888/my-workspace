"""
数据模型 — 税审输入/输出数据结构
"""

from dataclasses import dataclass, field
from typing import Dict, List, Optional
from enum import Enum


# ============================================================
# 枚举
# ============================================================


class AssetCategory(str, Enum):
    """固定资产类别（税法分类）"""

    BUILDING = "房屋、建筑物"
    MACHINERY = "机器设备"
    PRODUCTION_TOOLS = "生产器具、工具"
    TRANSPORT = "运输工具"
    ELECTRONIC = "电子设备"
    OTHER = "其他设备"


class AdjustmentCategory(str, Enum):
    """纳税调整类别"""

    INCOME = "收入类"
    DEDUCTION = "扣除类"
    ASSET = "资产类"
    SPECIAL = "特殊事项"
    OVERSEAS = "境外税收"
    TAX_INCENTIVE = "税收优惠"
    OTHER = "其他"


# ============================================================
# 输入数据模型
# ============================================================


@dataclass
class EnterpriseInfo:
    """企业基本信息"""

    name: str = ""  # 被审核单位名称
    uscc: str = ""  # 统一社会信用代码
    industry: str = ""  # 所属行业
    tax_year: int = 2025  # 纳税年度
    legal_rep: str = ""  # 法定代表人
    address: str = ""  # 注册地址
    registered_capital: float = 0  # 注册资本（万元）
    employee_count: int = 0  # 全年平均从业人数
    total_assets: float = 0  # 资产总额（万元）— 用于小型微利判断
    is_high_tech: bool = False  # 是否高新技术企业
    is_small_micro: Optional[bool] = None  # 是否小型微利企业（自动计算）


@dataclass
class TrialBalance:
    """
    试算平衡表（税审需要的关键科目）
    只需输入科目名称→金额的映射
    """

    items: Dict[str, float] = field(default_factory=dict)

    # 常用科目快捷访问
    def __getitem__(self, key: str) -> float:
        return self.items.get(key, 0.0)

    def __setitem__(self, key: str, value: float):
        self.items[key] = value

    def get(self, key: str, default: float = 0.0) -> float:
        """get方法兼容dict式访问"""
        return self.items.get(key, default)

    # === 收入科目 ===
    @property
    def revenue_main(self):
        return self["主营业务收入"]

    @property
    def revenue_other(self):
        return self["其他业务收入"]

    @property
    def revenue_total(self):
        return self.revenue_main + self.revenue_other

    @property
    def non_operating_income(self):
        return self["营业外收入"]

    @property
    def investment_income(self):
        return self["投资收益"]

    @property
    def fair_value_change(self):
        return self["公允价值变动收益"]

    @property
    def asset_disposal_income(self):
        return self["资产处置收益"]

    # === 成本科目 ===
    @property
    def cost_main(self):
        return self["主营业务成本"]

    @property
    def cost_other(self):
        return self["其他业务成本"]

    @property
    def tax_surcharge(self):
        return self["税金及附加"]

    # === 费用科目 ===
    @property
    def selling_expense(self):
        return self["销售费用"]

    @property
    def admin_expense(self):
        return self["管理费用"]

    @property
    def finance_expense(self):
        return self["财务费用"]

    # === 利润 ===
    @property
    def accounting_profit(self):
        return self["利润总额"]

    # === 明细科目（从费用中拆分） ===
    @property
    def salary_wages(self):
        return self["工资薪金"]

    @property
    def welfare(self):
        return self["职工福利费"]

    @property
    def education(self):
        return self["职工教育经费"]

    @property
    def union_fund(self):
        return self["工会经费"]

    @property
    def business_entertainment(self):
        return self["业务招待费"]

    @property
    def advertising(self):
        return self["广告费和业务宣传费"]

    @property
    def donation(self):
        return self["公益性捐赠支出"]

    @property
    def interest_expense(self):
        return self["利息支出"]

    @property
    def penalties(self):
        return self["罚金、罚款和被没收财物"]

    @property
    def late_fees(self):
        return self["税收滞纳金、加收利息"]

    @property
    def sponsorship(self):
        return self["赞助支出"]

    @property
    def r_and_d_expense(self):
        return self["研发费用"]

    @property
    def asset_impairment(self):
        return self["资产减值损失"]


@dataclass
class PayrollDetail:
    """
    工资薪金及三项费用明细
    CAS 9 / 企业所得税法实施条例第34-42条
    """

    total_wages: float = 0  # 工资薪金总额（税前）
    welfare_expense: float = 0  # 职工福利费发生额
    education_expense: float = 0  # 职工教育经费发生额
    union_expense: float = 0  # 工会经费发生额
    supplementary_pension: float = 0  # 补充养老保险
    supplementary_medical: float = 0  # 补充医疗保险
    housing_fund: float = 0  # 住房公积金
    social_insurance: float = 0  # 基本社会保险


@dataclass
class AssetItem:
    """
    固定资产卡片
    企业所得税法实施条例第59-64条
    """

    category: str = ""  # 资产类别
    name: str = ""  # 资产名称
    original_value: float = 0  # 原值
    accounting_life_years: int = 0  # 会计折旧年限
    tax_life_years: int = 0  # 税法折旧年限
    salvage_rate: float = 0.05  # 残值率
    current_accounting_depr: float = 0  # 本年会计折旧
    current_tax_depr: float = 0  # 本年税收折旧
    is_accelerated: bool = False  # 是否享受加速折旧


@dataclass
class TaxReturnData:
    """
    纳税申报表关键数据（A类主表+附表）
    用于与底稿勾稽核对
    """

    items: Dict[str, float] = field(default_factory=dict)


# ============================================================
# 输出数据模型
# ============================================================


@dataclass
class TaxAdjustment:
    """
    单条纳税调整结果
    """

    item_name: str  # 调整项目名称
    category: AdjustmentCategory  # 类别
    book_amount: float = 0  # 账载金额
    tax_base: float = 0  # 税收金额/计税基础
    increase: float = 0  # 纳税调增
    decrease: float = 0  # 纳税调减
    tax_law_ref: str = ""  # 税法依据
    calculation: str = ""  # 计算过程说明
    remark: str = ""  # 备注


@dataclass
class AuditConclusion:
    """
    审定结论
    """

    sheet_name: str  # 底稿表名
    items: List[Dict] = field(default_factory=list)  # 审定项目列表


@dataclass
class CalculationResult:
    """
    税审计算引擎完整输出
    """

    enterprise: EnterpriseInfo = field(default_factory=EnterpriseInfo)
    tb: TrialBalance = field(default_factory=TrialBalance)
    adjustments: List[TaxAdjustment] = field(default_factory=list)
    accounting_profit: float = 0  # 会计利润
    taxable_income: float = 0  # 应纳税所得额
    tax_payable: float = 0  # 应纳所得税额
    tax_rate: float = 0.25  # 适用税率
    deducted_tax: float = 0  # 减免所得税额
    final_tax: float = 0  # 实际应纳所得税额

    @property
    def total_increase(self) -> float:
        return sum(a.increase for a in self.adjustments)

    @property
    def total_decrease(self) -> float:
        return sum(a.decrease for a in self.adjustments)

    def summary(self) -> dict:
        """返回计算结果摘要"""
        return {
            "会计利润": self.accounting_profit,
            "纳税调增合计": self.total_increase,
            "纳税调减合计": self.total_decrease,
            "应纳税所得额": self.taxable_income,
            "适用税率": f"{self.tax_rate * 100:.1f}%",
            "应纳所得税额": self.tax_payable,
            "减免所得税额": self.deducted_tax,
            "实际应纳所得税额": self.final_tax,
            "调整项数": len(self.adjustments),
        }
