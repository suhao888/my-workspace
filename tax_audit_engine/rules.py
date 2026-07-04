"""
纳税调整规则引擎 — 企业所得税法50+条核心规则
每条规则按 条件→计算→输出 三要素组织，引用具体税法条款
"""

from typing import List, Optional, Callable
from .models import (
    TrialBalance,
    PayrollDetail,
    AssetItem,
    EnterpriseInfo,
    TaxAdjustment,
    AdjustmentCategory,
    CalculationResult,
)
from dataclasses import dataclass, field
from functools import reduce
import operator


# ============================================================
# 规则定义
# ============================================================


@dataclass
class TaxRule:
    """一条纳税调整规则的元定义"""

    id: str  # 规则编号，如 3-02-01-01
    name: str  # 规则名称
    category: AdjustmentCategory  # 所属类别
    tax_law_ref: str  # 税法依据
    description: str  # 规则说明
    func: Callable = None  # 计算函数

    def execute(self, ctx: "RuleContext") -> Optional[TaxAdjustment]:
        """执行规则，返回调整结果（无调整时返回None）"""
        if self.func:
            return self.func(ctx)
        return None


@dataclass
class RuleContext:
    """规则执行上下文 — 所有规则的输入"""

    tb: TrialBalance
    payroll: Optional[PayrollDetail] = None
    assets: List[AssetItem] = field(default_factory=list)
    enterprise: Optional[EnterpriseInfo] = None
    # 缓存已计算的结果，供后续规则引用
    _cache: dict = field(default_factory=dict)

    def get(self, key: str, default=0.0) -> float:
        return self.tb.items.get(key, default)

    def cache_result(self, key: str, value: float):
        self._cache[key] = value

    def get_cached(self, key: str, default=0.0) -> float:
        return self._cache.get(key, default)


# ============================================================
# 公用计算函数
# ============================================================


def max0(x: float) -> float:
    """max(x, 0)，防止负调增"""
    return max(x, 0.0)


def min0(x: float) -> float:
    """min(x, 0)，用于调减"""
    return min(x, 0.0)


def safe_div(a: float, b: float, default=0.0) -> float:
    """安全除法"""
    return a / b if b != 0 else default


# ============================================================
# 收入类纳税调整规则（3-01系列）
# ============================================================


def rule_30101_视同销售(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    视同销售收入调整
    企业所得税法实施条例第25条：企业发生非货币性资产交换，以及将货物用于捐赠/赞助/集资/广告/样品/职工福利等
    """
    book = ctx.get("视同销售收入-账面", 0)
    tax = ctx.get("视同销售收入-税收", 0)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="视同销售收入",
        category=AdjustmentCategory.INCOME,
        book_amount=book,
        tax_base=tax,
        increase=max0(diff),
        decrease=max0(-diff),
        tax_law_ref="企业所得税法实施条例第25条",
        calculation=f"税收金额{tax} - 账载金额{book} = {diff}",
    )


def rule_30102_未按权责发生制(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    未按权责发生制确认收入
    国税函[2008]875号、国税函[2009]98号
    """
    book = ctx.get("未按权责发生制确认收入-账面", 0)
    tax = ctx.get("未按权责发生制确认收入-税收", 0)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="未按权责发生制确认收入",
        category=AdjustmentCategory.INCOME,
        book_amount=book,
        tax_base=tax,
        increase=max0(diff),
        decrease=max0(-diff),
        tax_law_ref="国税函[2008]875号",
        calculation=f"税收{tax} - 账面{book} = {diff}",
    )


def rule_30103_投资收益(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    投资收益调整（居民企业间股息红利免税部分）
    企业所得税法第26条：符合条件的居民企业之间的股息红利免税
    """
    total_invest_income = ctx.get("投资收益", 0)
    tax_exempt = ctx.get("居民企业间股息红利-免税收入", 0)
    # 免税收入调减
    if tax_exempt <= 0:
        return None
    return TaxAdjustment(
        item_name="居民企业间股息红利免税",
        category=AdjustmentCategory.INCOME,
        book_amount=total_invest_income,
        tax_base=total_invest_income,
        increase=0,
        decrease=tax_exempt,
        tax_law_ref="企业所得税法第26条",
        calculation=f"符合条件的居民企业间股息红利 {tax_exempt} 免税调减",
    )


def rule_30104_权益法长投(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    按权益法核算长期股权投资
    企业所得税法实施条例第17条：税法按成本法处理，权益法确认的收益需调整
    """
    equity_income = ctx.get("权益法核算的投资收益", 0)
    dividend_received = ctx.get("实际收到的股息红利", 0)
    # 权益法确认的收益 > 实际收到的股息 → 调减
    # 权益法确认的收益 < 实际收到的股息 → 调增（特殊情况）
    diff = dividend_received - equity_income
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="按权益法核算长期股权投资",
        category=AdjustmentCategory.INCOME,
        book_amount=equity_income,
        tax_base=dividend_received,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第17条",
        calculation=f"实际股息{dividend_received} - 权益法确认{equity_income} = {diff}",
    )


def rule_30105_公允价值变动(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    公允价值变动净损益
    企业所得税法实施条例第56条：税法不确认公允价值变动
    """
    fair_value_gain = ctx.get("公允价值变动收益", 0)
    if abs(fair_value_gain) < 0.01:
        return None
    # 公允价值变动收益 → 调减（税法不认收益）
    # 公允价值变动损失 → 调增（税法不认损失）
    return TaxAdjustment(
        item_name="公允价值变动净损益",
        category=AdjustmentCategory.INCOME,
        book_amount=fair_value_gain,
        tax_base=0,
        increase=max0(-fair_value_gain),
        decrease=max0(fair_value_gain),
        tax_law_ref="企业所得税法实施条例第56条",
        calculation=f"公允价值变动{fair_value_gain}，税法不确认，{'调减' if fair_value_gain > 0 else '调增'} {abs(fair_value_gain)}",
    )


def rule_30106_不征税收入(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    不征税收入
    企业所得税法第7条：财政拨款/行政事业性收费/政府性基金
    """
    non_tax_income = ctx.get("不征税收入", 0)
    if non_tax_income <= 0:
        return None
    return TaxAdjustment(
        item_name="不征税收入",
        category=AdjustmentCategory.INCOME,
        book_amount=non_tax_income,
        tax_base=0,
        increase=0,
        decrease=non_tax_income,
        tax_law_ref="企业所得税法第7条",
        calculation=f"不征税收入 {non_tax_income} 不计入应税收入调减",
    )


def rule_30107_专项用途财政资金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    专项用途财政性资金
    财税[2011]70号：符合条件的不征税收入，支出也不得税前扣除
    """
    special_fund = ctx.get("专项用途财政性资金", 0)
    special_expense = ctx.get("专项用途财政性资金支出", 0)
    if special_fund <= 0:
        return None
    adj = TaxAdjustment(
        item_name="专项用途财政性资金",
        category=AdjustmentCategory.INCOME,
        book_amount=special_fund,
        tax_base=0,
        increase=0,
        decrease=special_fund,
        tax_law_ref="财税[2011]70号",
        calculation=f"专项用途财政资金 {special_fund} 不征税调减",
    )
    if special_expense > 0:
        # 对应支出不得扣除，需调增（在扣除类体现）
        ctx.cache_result("专项用途财政资金支出_纳税调增", special_expense)
    return adj


# ============================================================
# 扣除类纳税调整规则（3-02系列）
# ============================================================


def rule_30201_工资薪金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    工资薪金支出
    企业所得税法实施条例第34条：合理的工资薪金准予扣除
    国税函[2009]3号：合理性判断标准
    """
    book = ctx.get("工资薪金", 0)
    tax_allowed = ctx.get("工资薪金-税收金额", book)  # 默认全部合理
    diff = tax_allowed - book
    if abs(diff) < 0.01:
        # 若一致但需确认无差异
        if book <= 0:
            return None
        return TaxAdjustment(
            item_name="工资薪金支出",
            category=AdjustmentCategory.DEDUCTION,
            book_amount=book,
            tax_base=tax_allowed,
            increase=0,
            decrease=0,
            tax_law_ref="企业所得税法实施条例第34条",
            calculation=f"账载{book}，税收{tax_allowed}，无差异",
        )
    return TaxAdjustment(
        item_name="工资薪金支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第34条",
        calculation=f"税收{tax_allowed} - 账面{book} = {diff}",
    )


def rule_30202_职工福利费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    职工福利费 — 14%限额
    企业所得税法实施条例第40条：不超过工资薪金14%的部分准予扣除
    """
    wages = ctx.get("工资薪金", 0)
    book = ctx.get("职工福利费", 0)
    limit = wages * 0.14
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="职工福利费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第40条",
        calculation=f"工资{wages}×14%={limit}，账面{book}，超支{increase}",
    )


def rule_30203_职工教育经费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    职工教育经费 — 8%限额
    财税[2018]51号：不超过工资薪金8%的部分准予扣除，超过部分可结转
    """
    wages = ctx.get("工资薪金", 0)
    book = ctx.get("职工教育经费", 0)
    limit = wages * 0.08
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="职工教育经费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="财税[2018]51号",
        calculation=f"工资{wages}×8%={limit}，账面{book}，超支{increase}（可结转）",
    )


def rule_30204_工会经费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    工会经费 — 2%限额
    企业所得税法实施条例第41条：不超过工资薪金2%的部分准予扣除
    """
    wages = ctx.get("工资薪金", 0)
    book = ctx.get("工会经费", 0)
    limit = wages * 0.02
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="工会经费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第41条",
        calculation=f"工资{wages}×2%={limit}，账面{book}，超支{increase}",
    )


def rule_30205_基本社保公积金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    基本社会保险和住房公积金
    企业所得税法实施条例第35条：按规定比例缴纳的部分准予扣除
    """
    result = None
    # 基本社保
    si_book = ctx.get("基本社会保险", 0)
    si_tax = ctx.get("基本社会保险-税收金额", si_book)
    si_diff = si_tax - si_book
    if abs(si_diff) > 0.01:
        result = TaxAdjustment(
            item_name="基本社会保险",
            category=AdjustmentCategory.DEDUCTION,
            book_amount=si_book,
            tax_base=si_tax,
            increase=max0(-si_diff),
            decrease=max0(si_diff),
            tax_law_ref="企业所得税法实施条例第35条",
            calculation=f"税收{si_tax} - 账面{si_book} = {si_diff}",
        )
    # 公积金
    hf_book = ctx.get("住房公积金", 0)
    hf_tax = ctx.get("住房公积金-税收金额", hf_book)
    hf_diff = hf_tax - hf_book
    if abs(hf_diff) > 0.01:
        result = TaxAdjustment(
            item_name="住房公积金",
            category=AdjustmentCategory.DEDUCTION,
            book_amount=hf_book,
            tax_base=hf_tax,
            increase=max0(-hf_diff),
            decrease=max0(hf_diff),
            tax_law_ref="企业所得税法实施条例第35条",
            calculation=f"税收{hf_tax} - 账面{hf_book} = {hf_diff}",
        )
    return result


def rule_30206_补充养老保险医疗(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    补充养老保险/医疗保险 — 5%限额
    财税[2009]27号：不超过工资薪金5%的部分准予扣除
    """
    wages = ctx.get("工资薪金", 0)
    limit = wages * 0.05
    adjustments = []

    for name, key in [
        ("补充养老保险", "补充养老保险"),
        ("补充医疗保险", "补充医疗保险"),
    ]:
        book = ctx.get(key, 0)
        tax_allowed = min(book, limit)
        increase = max0(book - limit)
        if increase > 0.01:
            adjustments.append(
                TaxAdjustment(
                    item_name=name,
                    category=AdjustmentCategory.DEDUCTION,
                    book_amount=book,
                    tax_base=tax_allowed,
                    increase=increase,
                    decrease=0,
                    tax_law_ref="财税[2009]27号",
                    calculation=f"工资{wages}×5%={limit}，账面{book}，超支{increase}",
                )
            )
    # 返回第一个调整（合并展示由计算器处理）
    return adjustments[0] if adjustments else None


def rule_30207_业务招待费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    业务招待费 — 60%与5‰孰低
    企业所得税法实施条例第43条：按发生额的60%扣除，最高不超过当年销售收入的5‰
    """
    book = ctx.get("业务招待费", 0)
    revenue = ctx.tb.revenue_total
    if book <= 0:
        return None
    method1 = book * 0.6
    method2 = revenue * 0.005
    allowed = min(method1, method2)
    increase = book - allowed
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="业务招待费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第43条",
        calculation=f"发生额{book}×60%={method1}，收入{revenue}×5‰={method2}，"
        f"取小值{allowed}，调增{increase}",
    )


def rule_30208_广告费和业务宣传费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    广告费和业务宣传费 — 15%限额（一般行业）
    企业所得税法实施条例第44条：不超过当年销售收入15%的部分准予扣除
    财税[2020]43号：化妆品/医药/饮料行业30%
    财税[2020]41号：白酒行业不得扣除
    """
    book = ctx.get("广告费和业务宣传费", 0)
    revenue = ctx.tb.revenue_total
    # 默认15%，特殊行业由外部配置
    industry_rate = 0.15
    limit = revenue * industry_rate
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    carryover = max0(increase)  # 超过部分可结转
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="广告费和业务宣传费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第44条",
        calculation=f"收入{revenue}×{industry_rate * 100}%={limit}，账面{book}，超支{increase}（可结转）",
    )


def rule_30209_公益性捐赠支出(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    公益性捐赠支出 — 12%限额
    企业所得税法第9条：不超过年度利润总额12%的部分准予扣除
    财税[2020]21号：扶贫捐赠全额扣除
    """
    book = ctx.get("公益性捐赠支出", 0)
    profit = ctx.tb.accounting_profit
    limit = max(profit * 0.12, 0)
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    carryover = max0(increase)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="公益性捐赠支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法第9条",
        calculation=f"利润{profit}×12%={limit}，账面{book}，超支{increase}（可三年结转）",
    )


def rule_30210_利息支出(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    利息支出 — 金融企业同期利率限额
    企业所得税法实施条例第38条：非金融企业向非金融企业借款的利息，不超过金融企业同期同类利率部分可扣除
    """
    book = ctx.get("利息支出", 0)
    tax_allowed = ctx.get("利息支出-税收金额", book)
    increase = max0(book - tax_allowed)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="利息支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第38条",
        calculation=f"账面{book}，允许扣除{tax_allowed}，超支{increase}",
    )


def rule_30211_罚金滞纳金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    罚金、罚款和被没收财物损失
    企业所得税法第10条：不得扣除
    """
    book = ctx.get("罚金、罚款和被没收财物", 0)
    if book <= 0:
        return None
    return TaxAdjustment(
        item_name="罚金、罚款和被没收财物损失",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=0,
        increase=book,
        decrease=0,
        tax_law_ref="企业所得税法第10条第(四)项",
        calculation=f"账面{book}，全额调增（不得扣除）",
    )


def rule_30212_税收滞纳金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    税收滞纳金、加收利息
    企业所得税法第10条：不得扣除
    """
    book = ctx.get("税收滞纳金、加收利息", 0)
    if book <= 0:
        return None
    return TaxAdjustment(
        item_name="税收滞纳金、加收利息",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=0,
        increase=book,
        decrease=0,
        tax_law_ref="企业所得税法第10条第(三)项",
        calculation=f"账面{book}，全额调增（不得扣除）",
    )


def rule_30213_赞助支出(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    赞助支出
    企业所得税法第10条：非广告性质的赞助支出不得扣除
    国税发[2001]38号：广告性赞助按广告费处理
    """
    book = ctx.get("赞助支出", 0)
    advertising_sponsor = ctx.get("赞助支出-广告性", 0)
    non_advertising = book - advertising_sponsor
    if non_advertising <= 0:
        return None
    return TaxAdjustment(
        item_name="赞助支出（非广告性）",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=advertising_sponsor,
        increase=non_advertising,
        decrease=0,
        tax_law_ref="企业所得税法第10条第(六)项",
        calculation=f"赞助支出总额{book}，其中广告性{advertising_sponsor}按广告费处理，非广告性{non_advertising}全额调增",
    )


def rule_30214_跨期扣除(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    跨期扣除项目（跨期费用）
    企业所得税法实施条例第9条：权责发生制原则
    """
    diff = ctx.get("跨期扣除-差异", 0)
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="跨期扣除项目",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=ctx.get("跨期扣除-账面", 0),
        tax_base=ctx.get("跨期扣除-税收", 0),
        increase=max0(diff),
        decrease=max0(-diff),
        tax_law_ref="企业所得税法实施条例第9条",
        calculation=f"跨期差异{diff}，权责发生制调整",
    )


def rule_30215_与取得收入无关支出(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    与取得收入无关的支出
    企业所得税法第10条第(八)项：不得扣除
    """
    book = ctx.get("与取得收入无关的支出", 0)
    if book <= 0:
        return None
    return TaxAdjustment(
        item_name="与取得收入无关的支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=0,
        increase=book,
        decrease=0,
        tax_law_ref="企业所得税法第10条第(八)项",
        calculation=f"账面{book}，全额调增（与取得收入无关）",
    )


# ============================================================
# 资产类纳税调整规则（3-03系列）
# ============================================================


def rule_30301_固定资产折旧(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    固定资产折旧
    企业所得税法实施条例第59-60条：按税法年限计算折旧，与会计年限的差异需调整
    房屋20年/机器10年/电子3年/运输4年/生产工具5年
    """
    assets = ctx.assets
    if not assets:
        return None

    total_acct_depr = sum(a.current_accounting_depr for a in assets)
    total_tax_depr = sum(a.current_tax_depr for a in assets)
    diff = total_tax_depr - total_acct_depr

    result = TaxAdjustment(
        item_name="固定资产折旧",
        category=AdjustmentCategory.ASSET,
        book_amount=total_acct_depr,
        tax_base=total_tax_depr,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第59-60条",
        calculation=f"会计折旧{total_acct_depr}，税收折旧{total_tax_depr}",
    )

    # 附加明细（用于底稿明细行）
    detail_lines = []
    for a in assets:
        ad = a.current_accounting_depr
        td = a.current_tax_depr
        if abs(ad - td) > 0.01:
            detail_lines.append(
                f"{a.category}-{a.name}: 会计{ad} 税收{td} 差异{td - ad}"
            )
    if detail_lines:
        result.remark = "; ".join(detail_lines)
    return result


def rule_30302_无形资产摊销(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    无形资产摊销
    企业所得税法实施条例第65-67条：不低于10年摊销，与会计差异需调整
    """
    book = ctx.get("无形资产摊销-账面", 0)
    tax = ctx.get("无形资产摊销-税收", book)
    diff = tax - book
    if abs(diff) < 0.01 and book <= 0:
        return None
    return TaxAdjustment(
        item_name="无形资产摊销",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第65-67条",
        calculation=f"会计摊销{book}，税收摊销{tax}，差异{diff}",
    )


def rule_30303_长期待摊费用(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    长期待摊费用
    企业所得税法实施条例第68-70条：不低于3年摊销
    """
    book = ctx.get("长期待摊费用摊销-账面", 0)
    tax = ctx.get("长期待摊费用摊销-税收", book)
    diff = tax - book
    if abs(diff) < 0.01 and book <= 0:
        return None
    return TaxAdjustment(
        item_name="长期待摊费用摊销",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第68-70条",
        calculation=f"会计摊销{book}，税收摊销{tax}，差异{diff}",
    )


def rule_30304_资产减值准备(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    资产减值准备
    企业所得税法第10条：未经核定的准备金支出不得扣除
    国税函[2009]202号：坏账准备等各项减值准备不得税前扣除
    """
    book = ctx.get("资产减值损失", 0)
    tax_allowed = ctx.get("资产减值损失-税收金额", 0)
    increase = max0(book - tax_allowed)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="资产减值准备",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法第10条第(七)项",
        calculation=f"账面资产减值损失{book}，允许扣除{tax_allowed}，调增{increase}",
    )


def rule_30305_资产损失(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    资产损失税前扣除
    财税[2009]57号、国家税务总局公告2011年第25号
    """
    book = ctx.get("资产损失-账面", 0)
    tax_allowed = ctx.get("资产损失-税收金额", 0)
    diff = tax_allowed - book
    if abs(diff) < 0.01 and book <= 0:
        return None
    return TaxAdjustment(
        item_name="资产损失税前扣除",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax_allowed,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="财税[2009]57号",
        calculation=f"账面资产损失{book}，允许扣除{tax_allowed}，差异{diff}",
    )


# ============================================================
# 税收优惠（5系列）
# ============================================================


def rule_501_研发费用加计扣除(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    研发费用加计扣除
    财税[2023]7号：未形成无形资产的，加计扣除100%
    财税[2023]12号：集成电路企业和工业母机企业加计扣除120%
    """
    rd_book = ctx.get("研发费用", 0)
    rd_deduction_rate = ctx.get("研发加计扣除率", 1.0)  # 默认100%
    rd_tax_deduction = rd_book * rd_deduction_rate
    if rd_tax_deduction <= 0:
        return None
    return TaxAdjustment(
        item_name="研发费用加计扣除",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=rd_book,
        tax_base=rd_book + rd_tax_deduction,
        increase=0,
        decrease=rd_tax_deduction,
        tax_law_ref="财税[2023]7号",
        calculation=f"研发费用{rd_book} × {rd_deduction_rate * 100:.0f}% = {rd_tax_deduction} 加计扣除调减",
    )


def rule_502_残疾职工工资加计扣除(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    残疾职工工资加计扣除
    财税[2009]70号：按残疾职工工资的100%加计扣除
    """
    disabled_wages = ctx.get("残疾职工工资", 0)
    if disabled_wages <= 0:
        return None
    return TaxAdjustment(
        item_name="残疾职工工资加计扣除",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=disabled_wages,
        tax_base=disabled_wages * 2,
        increase=0,
        decrease=disabled_wages,
        tax_law_ref="财税[2009]70号",
        calculation=f"残疾职工工资{disabled_wages} × 100% = {disabled_wages} 加计扣除调减",
    )


def rule_503_国债利息收入(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    国债利息收入免税
    企业所得税法第26条：国债利息收入为免税收入
    """
    bond_interest = ctx.get("国债利息收入", 0)
    if bond_interest <= 0:
        return None
    return TaxAdjustment(
        item_name="国债利息收入",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=bond_interest,
        tax_base=0,
        increase=0,
        decrease=bond_interest,
        tax_law_ref="企业所得税法第26条",
        calculation=f"国债利息收入{bond_interest} 免税调减",
    )


def rule_504_小型微利企业减免(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    小型微利企业减免所得税（此规则在计算应纳税额时使用）
    财税[2023]6号：年应纳税所得额≤300万，分段减计
    此规则仅标记，实际减免在calculator中计算
    """
    return None  # 在calculator的tax计算阶段处理


# ============================================================
# 规则注册表
# ============================================================

# 收入类规则
INCOME_RULES = [
    TaxRule(
        id="3-01-01",
        name="视同销售收入",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="企业所得税法实施条例第25条",
        description="非货币性资产交换及货物用于捐赠/赞助/广告/职工福利等视同销售",
        func=rule_30101_视同销售,
    ),
    TaxRule(
        id="3-01-02",
        name="未按权责发生制确认收入",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="国税函[2008]875号",
        description="跨期收入调整",
        func=rule_30102_未按权责发生制,
    ),
    TaxRule(
        id="3-01-03",
        name="投资收益（居民企业间免税）",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="企业所得税法第26条",
        description="符合条件的居民企业间股息红利免税",
        func=rule_30103_投资收益,
    ),
    TaxRule(
        id="3-01-04",
        name="按权益法核算长期股权投资",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="企业所得税法实施条例第17条",
        description="权益法vs成本法差异",
        func=rule_30104_权益法长投,
    ),
    TaxRule(
        id="3-01-05",
        name="公允价值变动净损益",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="企业所得税法实施条例第56条",
        description="公允价值变动税法不确认",
        func=rule_30105_公允价值变动,
    ),
    TaxRule(
        id="3-01-06",
        name="不征税收入",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="企业所得税法第7条",
        description="财政拨款/行政事业性收费/政府性基金",
        func=rule_30106_不征税收入,
    ),
    TaxRule(
        id="3-01-07",
        name="专项用途财政性资金",
        category=AdjustmentCategory.INCOME,
        tax_law_ref="财税[2011]70号",
        description="符合条件的不征税收入，支出不得扣除",
        func=rule_30107_专项用途财政资金,
    ),
]

# 扣除类规则
DEDUCTION_RULES = [
    TaxRule(
        id="3-02-01",
        name="工资薪金支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第34条",
        description="合理的工资薪金准予扣除",
        func=rule_30201_工资薪金,
    ),
    TaxRule(
        id="3-02-02",
        name="职工福利费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第40条",
        description="不超过工资薪金14%",
        func=rule_30202_职工福利费,
    ),
    TaxRule(
        id="3-02-03",
        name="职工教育经费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="财税[2018]51号",
        description="不超过工资薪金8%",
        func=rule_30203_职工教育经费,
    ),
    TaxRule(
        id="3-02-04",
        name="工会经费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第41条",
        description="不超过工资薪金2%",
        func=rule_30204_工会经费,
    ),
    TaxRule(
        id="3-02-05",
        name="基本社会保险和住房公积金",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第35条",
        description="按规定比例缴纳部分准予扣除",
        func=rule_30205_基本社保公积金,
    ),
    TaxRule(
        id="3-02-06",
        name="补充养老保险和补充医疗保险",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="财税[2009]27号",
        description="不超过工资薪金5%",
        func=rule_30206_补充养老保险医疗,
    ),
    TaxRule(
        id="3-02-07",
        name="业务招待费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第43条",
        description="60%与5‰孰低",
        func=rule_30207_业务招待费,
    ),
    TaxRule(
        id="3-02-08",
        name="广告费和业务宣传费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第44条",
        description="不超过销售收入15%",
        func=rule_30208_广告费和业务宣传费,
    ),
    TaxRule(
        id="3-02-09",
        name="公益性捐赠支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法第9条",
        description="不超过利润总额12%",
        func=rule_30209_公益性捐赠支出,
    ),
    TaxRule(
        id="3-02-10",
        name="利息支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第38条",
        description="不超过金融企业同期同类利率",
        func=rule_30210_利息支出,
    ),
    TaxRule(
        id="3-02-11",
        name="罚金、罚款和被没收财物损失",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法第10条第(四)项",
        description="不得扣除",
        func=rule_30211_罚金滞纳金,
    ),
    TaxRule(
        id="3-02-12",
        name="税收滞纳金、加收利息",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法第10条第(三)项",
        description="不得扣除",
        func=rule_30212_税收滞纳金,
    ),
    TaxRule(
        id="3-02-13",
        name="赞助支出（非广告性）",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法第10条第(六)项",
        description="非广告性赞助支出不得扣除",
        func=rule_30213_赞助支出,
    ),
    TaxRule(
        id="3-02-14",
        name="跨期扣除项目",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第9条",
        description="权责发生制调整",
        func=rule_30214_跨期扣除,
    ),
    TaxRule(
        id="3-02-15",
        name="与取得收入无关的支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法第10条第(八)项",
        description="与取得收入无关的支出不得扣除",
        func=rule_30215_与取得收入无关支出,
    ),
]

# 资产类规则
ASSET_RULES = [
    TaxRule(
        id="3-03-01",
        name="固定资产折旧",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="企业所得税法实施条例第59-60条",
        description="税法年限与会计年限差异调整",
        func=rule_30301_固定资产折旧,
    ),
    TaxRule(
        id="3-03-02",
        name="无形资产摊销",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="企业所得税法实施条例第65-67条",
        description="不低于10年摊销",
        func=rule_30302_无形资产摊销,
    ),
    TaxRule(
        id="3-03-03",
        name="长期待摊费用摊销",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="企业所得税法实施条例第68-70条",
        description="不低于3年摊销",
        func=rule_30303_长期待摊费用,
    ),
    TaxRule(
        id="3-03-04",
        name="资产减值准备",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="企业所得税法第10条第(七)项",
        description="未经核定的准备金不得扣除",
        func=rule_30304_资产减值准备,
    ),
    TaxRule(
        id="3-03-05",
        name="资产损失税前扣除",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="财税[2009]57号",
        description="按规定确认的资产损失可扣除",
        func=rule_30305_资产损失,
    ),
]

# 税收优惠规则
TAX_INCENTIVE_RULES = [
    TaxRule(
        id="5-01-01",
        name="研发费用加计扣除",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="财税[2023]7号",
        description="加计扣除100%（一般行业）",
        func=rule_501_研发费用加计扣除,
    ),
    TaxRule(
        id="5-01-02",
        name="残疾职工工资加计扣除",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="财税[2009]70号",
        description="按残疾职工工资100%加计扣除",
        func=rule_502_残疾职工工资加计扣除,
    ),
    TaxRule(
        id="5-01-03",
        name="国债利息收入",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="企业所得税法第26条",
        description="国债利息收入免税",
        func=rule_503_国债利息收入,
    ),
]

# 全部规则
ALL_RULES: List[TaxRule] = (
    INCOME_RULES + DEDUCTION_RULES + ASSET_RULES + TAX_INCENTIVE_RULES
)

# 规则索引
RULES_BY_ID = {r.id: r for r in ALL_RULES}
RULES_BY_CATEGORY = {
    AdjustmentCategory.INCOME: INCOME_RULES,
    AdjustmentCategory.DEDUCTION: DEDUCTION_RULES,
    AdjustmentCategory.ASSET: ASSET_RULES,
    AdjustmentCategory.TAX_INCENTIVE: TAX_INCENTIVE_RULES,
}
