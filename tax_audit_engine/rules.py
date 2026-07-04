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
# 特殊事项纳税调整规则（3-04~3-06系列）
# ============================================================


def rule_30401_企业重组特殊性(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    企业重组特殊性税务处理
    企业所得税法实施条例第75条：特殊性税务处理暂不确认损益
    """
    book = ctx.get("企业重组损益-账面", 0)
    if abs(book) < 0.01:
        return None
    return TaxAdjustment(
        item_name="企业重组特殊性税务处理",
        category=AdjustmentCategory.SPECIAL,
        book_amount=book,
        tax_base=0,
        increase=max0(-book),
        decrease=max0(book),
        tax_law_ref="企业所得税法实施条例第75条",
        calculation=f"账面确认重组损益{book}，特殊性税务处理暂不确认，"
        f"{'调减' if book > 0 else '调增'}{abs(book)}",
    )


def rule_30402_企业重组一般性(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    企业重组一般性税务处理
    企业所得税法实施条例第75条：按公允价值确认损益
    """
    book = ctx.get("企业重组损益-账面", 0)
    tax = ctx.get("企业重组损益-公允价值", book)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="企业重组一般性税务处理",
        category=AdjustmentCategory.SPECIAL,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第75条",
        calculation=f"账面{book} - 公允价值{tax} = {diff}，一般性税务处理",
    )


def rule_305_政策性搬迁(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    政策性搬迁
    国家税务总局公告2012年第40号：搬迁收入扣除搬迁支出后的余额计入应纳税所得额
    """
    book = ctx.get("政策性搬迁净额-账面", 0)
    relocation_income = ctx.get("政策性搬迁收入", 0)
    relocation_expense = ctx.get("政策性搬迁支出", 0)
    tax_net = relocation_income - relocation_expense
    diff = tax_net - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="政策性搬迁",
        category=AdjustmentCategory.SPECIAL,
        book_amount=book,
        tax_base=tax_net,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="国家税务总局公告2012年第40号",
        calculation=f"搬迁收入{relocation_income} - 搬迁支出{relocation_expense} = {tax_net}，"
        f"账面{book}，差异{diff}",
    )


def rule_30601_新收入准则差异(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    新收入准则与税法收入确认时点差异
    企业所得税法实施条例第9条：权责发生制原则
    """
    book = ctx.get("新收入准则-账面确认收入", 0)
    tax = ctx.get("新收入准则-税收确认收入", book)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="新收入准则与税法差异",
        category=AdjustmentCategory.SPECIAL,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第9条",
        calculation=f"账面{book} - 税收{tax} = {diff}（收入确认时点差异）",
    )


def rule_30602_其他特殊调整(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    其他特殊纳税调整（兜底项）
    企业所得税法第6章：特别纳税调整
    """
    book = ctx.get("其他特殊调整-账面", 0)
    tax = ctx.get("其他特殊调整-税收金额", 0)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="其他特殊纳税调整",
        category=AdjustmentCategory.SPECIAL,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法第6章",
        calculation=f"其他特殊调整差异{diff}",
    )


# ============================================================
# 境外税收纳税调整规则（4系列）
# ============================================================


def rule_401_境外所得纳税调整(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    境外所得纳税调整
    企业所得税法第17条：境外税前所得并入境内应纳税所得额
    """
    book = ctx.get("境外所得-账面金额", 0)
    tax = ctx.get("境外所得-税前所得", book)
    diff = tax - book
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="境外所得纳税调整",
        category=AdjustmentCategory.OVERSEAS,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法第17条",
        calculation=f"境外税前所得{tax} - 账面{book} = {diff}",
    )


def rule_402_境外所得抵免(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    境外所得抵免（限额抵免法）
    企业所得税法第23条：境外已缴税款不超过抵免限额部分可抵免
    """
    overseas_income = ctx.get("境外所得-应纳税所得额", 0)
    foreign_tax_paid = ctx.get("境外所得-已缴税款", 0)
    domestic_tax_rate = ctx.get("境外所得-境内税率", 0.25)
    credit_limit = overseas_income * domestic_tax_rate
    actual_credit = min(foreign_tax_paid, credit_limit)
    if actual_credit <= 0:
        return None
    return TaxAdjustment(
        item_name="境外所得抵免",
        category=AdjustmentCategory.OVERSEAS,
        book_amount=foreign_tax_paid,
        tax_base=actual_credit,
        increase=0,
        decrease=actual_credit,
        tax_law_ref="企业所得税法第23条",
        calculation=f"境外应纳税所得额{overseas_income}×税率{domestic_tax_rate * 100:.0f}%"
        f"=限额{credit_limit}，已缴{foreign_tax_paid}，实际抵免{actual_credit}",
    )


def rule_403_境外亏损弥补(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    境外亏损不得抵减境内盈利
    企业所得税法第17条：境外亏损不得抵减境内应纳税所得额
    """
    overseas_loss = ctx.get("境外所得-亏损金额", 0)
    if overseas_loss >= 0:
        return None
    # 境外亏损为负数，全额调增
    increase = abs(overseas_loss)
    return TaxAdjustment(
        item_name="境外亏损弥补",
        category=AdjustmentCategory.OVERSEAS,
        book_amount=overseas_loss,
        tax_base=0,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法第17条",
        calculation=f"境外亏损{overseas_loss}不得抵减境内盈利，全额调增{increase}",
    )


# ============================================================
# 税收优惠补充规则（5系列）
# ============================================================


def rule_504_资源综合利用减计收入(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    资源综合利用减计收入
    财税[2008]47号：资源综合利用产品收入减按90%计入收入总额
    """
    income = ctx.get("资源综合利用收入", 0)
    if income <= 0:
        return None
    reduction = income * 0.1
    return TaxAdjustment(
        item_name="资源综合利用减计收入",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=income,
        tax_base=income * 0.9,
        increase=0,
        decrease=reduction,
        tax_law_ref="财税[2008]47号",
        calculation=f"资源综合利用收入{income}×10%={reduction} 减计调减",
    )


def rule_505_所得减免(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    所得减免
    企业所得税法第27条：农林牧渔/基础设施/环保节能项目所得可免征或减征
    """
    book = ctx.get("减免所得-账面金额", 0)
    tax_exempt = ctx.get("减免所得-免税金额", 0)
    if tax_exempt <= 0:
        return None
    return TaxAdjustment(
        item_name="所得减免",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=book,
        tax_base=book - tax_exempt,
        increase=0,
        decrease=tax_exempt,
        tax_law_ref="企业所得税法第27条",
        calculation=f"免税所得{tax_exempt}调减",
    )


def rule_506_抵扣应纳税所得额(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    抵扣应纳税所得额（创投企业）
    国税发[2009]87号：创业投资企业投资额70%抵扣应纳税所得额
    """
    investment = ctx.get("创投企业投资额", 0)
    if investment <= 0:
        return None
    deduction = investment * 0.7
    return TaxAdjustment(
        item_name="创投企业投资额抵扣",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=0,
        tax_base=-deduction,
        increase=0,
        decrease=deduction,
        tax_law_ref="国税发[2009]87号",
        calculation=f"创投投资额{investment}×70%={deduction} 抵扣应纳税所得额",
    )


def rule_507_税额抵免(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    专用设备投资额税额抵免
    企业所得税法第34条：环保/节能/安全生产专用设备投资额10%抵免税额
    """
    equipment_investment = ctx.get("专用设备投资额", 0)
    if equipment_investment <= 0:
        return None
    tax_credit = equipment_investment * 0.1
    return TaxAdjustment(
        item_name="专用设备投资额税额抵免",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=0,
        tax_base=-tax_credit,
        increase=0,
        decrease=tax_credit,
        tax_law_ref="企业所得税法第34条",
        calculation=f"专用设备投资额{equipment_investment}×10%={tax_credit} 抵免应纳税额",
    )


def rule_508_软件集成电路优惠(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    软件和集成电路产业所得税优惠
    财税[2020]45号：两免三减半/十年免税
    """
    book = ctx.get("软件集成电路-账面利润", 0)
    exempt_amount = ctx.get("软件集成电路-免税利润", 0)
    half_amount = ctx.get("软件集成电路-减半利润", 0)
    reduction = exempt_amount + half_amount * 0.5
    if reduction <= 0:
        return None
    return TaxAdjustment(
        item_name="软件集成电路产业优惠",
        category=AdjustmentCategory.TAX_INCENTIVE,
        book_amount=book,
        tax_base=book - reduction,
        increase=0,
        decrease=reduction,
        tax_law_ref="财税[2020]45号",
        calculation=f"免税利润{exempt_amount}+减半利润{half_amount}×50%={reduction} 调减",
    )


# ============================================================
# 缴纳情况纳税调整规则（6系列）
# ============================================================


def rule_601_所得税预缴(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    所得税预缴差异
    企业所得税法第54条：分月或分季预缴，年终汇算清缴
    """
    actual_prepaid = ctx.get("所得税-实际预缴额", 0)
    estimated_prepaid = ctx.get("所得税-应预缴额", 0)
    diff = estimated_prepaid - actual_prepaid
    if abs(diff) < 0.01:
        return None
    return TaxAdjustment(
        item_name="所得税预缴差异",
        category=AdjustmentCategory.PAYMENT,
        book_amount=actual_prepaid,
        tax_base=estimated_prepaid,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法第54条",
        calculation=f"应预缴{estimated_prepaid} - 实际预缴{actual_prepaid} = {diff}",
    )


def rule_602_汇总纳税分摊(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    汇总纳税总分支机构分摊
    国家税务总局公告2012年第57号：总机构分摊比例
    """
    total_tax = ctx.get("汇总纳税-应纳税额", 0)
    head_office_ratio = ctx.get("汇总纳税-总机构分摊比例", 0.5)
    branch_ratio = ctx.get("汇总纳税-分支机构分摊比例", 0.5)
    head_portion = total_tax * head_office_ratio
    branch_portion = total_tax * branch_ratio
    if total_tax <= 0:
        return None
    return TaxAdjustment(
        item_name="汇总纳税分摊",
        category=AdjustmentCategory.PAYMENT,
        book_amount=total_tax,
        tax_base=head_portion,
        increase=0,
        decrease=0,
        tax_law_ref="国家税务总局公告2012年第57号",
        calculation=f"应纳税额{total_tax}×总机构{head_office_ratio}={head_portion}，"
        f"分支机构{branch_ratio}={branch_portion}",
    )


# ============================================================
# 扣除类补充规则（3-02系列）
# ============================================================


def rule_30216_手续费及佣金(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    手续费及佣金支出
    财税[2009]29号：按服务协议或合同确认收入5%（一般企业）限额
    """
    book = ctx.get("手续费及佣金支出", 0)
    revenue = ctx.tb.revenue_total
    industry_rate = ctx.get("手续费及佣金-扣除比例", 0.05)
    limit = revenue * industry_rate
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="手续费及佣金支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="财税[2009]29号",
        calculation=f"收入{revenue}×{industry_rate * 100:.0f}%={limit}，账面{book}，超支{increase}",
    )


def rule_30217_党组织工作经费(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    党组织工作经费
    组通字[2017]38号：不超过工资薪金总额1%的部分据实扣除
    """
    wages = ctx.get("工资薪金", 0)
    book = ctx.get("党组织工作经费", 0)
    limit = wages * 0.01
    tax_allowed = min(book, limit)
    increase = max0(book - limit)
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="党组织工作经费",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="组通字[2017]38号",
        calculation=f"工资{wages}×1%={limit}，账面{book}，超支{increase}",
    )


def rule_30218_劳动保护支出(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    劳动保护支出
    企业所得税法实施条例第48条：合理劳动保护支出准予扣除
    """
    book = ctx.get("劳动保护支出", 0)
    unreasonable = ctx.get("劳动保护支出-不合理部分", 0)
    tax_allowed = book - unreasonable
    increase = unreasonable
    if increase < 0.01:
        return None
    return TaxAdjustment(
        item_name="劳动保护支出",
        category=AdjustmentCategory.DEDUCTION,
        book_amount=book,
        tax_base=tax_allowed,
        increase=increase,
        decrease=0,
        tax_law_ref="企业所得税法实施条例第48条",
        calculation=f"账面{book}，扣除不合理部分{unreasonable}，调增{increase}",
    )


# ============================================================
# 资产类补充规则（3-03系列）
# ============================================================


def rule_30306_生产性生物资产折旧(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    生产性生物资产折旧
    企业所得税法实施条例第62条：林木类10年/畜产类3年
    """
    book = ctx.get("生产性生物资产折旧-账面", 0)
    tax = ctx.get("生产性生物资产折旧-税收", book)
    diff = tax - book
    if abs(diff) < 0.01 and book <= 0:
        return None
    return TaxAdjustment(
        item_name="生产性生物资产折旧",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="企业所得税法实施条例第62条",
        calculation=f"会计折旧{book}，税收折旧{tax}，差异{diff}",
    )


def rule_30307_油气资产折耗(ctx: RuleContext) -> Optional[TaxAdjustment]:
    """
    油气资产折耗
    财税[2009]135号：油气资产折耗按税法规定计算
    """
    book = ctx.get("油气资产折耗-账面", 0)
    tax = ctx.get("油气资产折耗-税收", book)
    diff = tax - book
    if abs(diff) < 0.01 and book <= 0:
        return None
    return TaxAdjustment(
        item_name="油气资产折耗",
        category=AdjustmentCategory.ASSET,
        book_amount=book,
        tax_base=tax,
        increase=max0(-diff),
        decrease=max0(diff),
        tax_law_ref="财税[2009]135号",
        calculation=f"会计折耗{book}，税收折耗{tax}，差异{diff}",
    )


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
    TaxRule(
        id="3-02-16",
        name="手续费及佣金支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="财税[2009]29号",
        description="不超过服务收入5%（一般企业）",
        func=rule_30216_手续费及佣金,
    ),
    TaxRule(
        id="3-02-17",
        name="党组织工作经费",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="组通字[2017]38号",
        description="不超过工资薪金1%",
        func=rule_30217_党组织工作经费,
    ),
    TaxRule(
        id="3-02-18",
        name="劳动保护支出",
        category=AdjustmentCategory.DEDUCTION,
        tax_law_ref="企业所得税法实施条例第48条",
        description="合理劳动保护支出准予扣除",
        func=rule_30218_劳动保护支出,
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
    TaxRule(
        id="3-03-06",
        name="生产性生物资产折旧",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="企业所得税法实施条例第62条",
        description="林木10年/畜产3年",
        func=rule_30306_生产性生物资产折旧,
    ),
    TaxRule(
        id="3-03-07",
        name="油气资产折耗",
        category=AdjustmentCategory.ASSET,
        tax_law_ref="财税[2009]135号",
        description="油气资产折耗按规定计算",
        func=rule_30307_油气资产折耗,
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
    TaxRule(
        id="5-02-01",
        name="资源综合利用减计收入",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="财税[2008]47号",
        description="资源综合利用收入减按90%计入",
        func=rule_504_资源综合利用减计收入,
    ),
    TaxRule(
        id="5-02-02",
        name="所得减免",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="企业所得税法第27条",
        description="农林牧渔/基础设施/环保节能项目所得减免",
        func=rule_505_所得减免,
    ),
    TaxRule(
        id="5-02-03",
        name="抵扣应纳税所得额",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="国税发[2009]87号",
        description="创投企业投资额70%抵扣",
        func=rule_506_抵扣应纳税所得额,
    ),
    TaxRule(
        id="5-02-04",
        name="专用设备投资额税额抵免",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="企业所得税法第34条",
        description="环保/节能/安全生产专用设备投资额10%抵免",
        func=rule_507_税额抵免,
    ),
    TaxRule(
        id="5-02-05",
        name="软件集成电路产业优惠",
        category=AdjustmentCategory.TAX_INCENTIVE,
        tax_law_ref="财税[2020]45号",
        description="两免三减半等优惠",
        func=rule_508_软件集成电路优惠,
    ),
]

# 特殊事项规则
SPECIAL_RULES = [
    TaxRule(
        id="3-04-01",
        name="企业重组特殊性税务处理",
        category=AdjustmentCategory.SPECIAL,
        tax_law_ref="企业所得税法实施条例第75条",
        description="特殊性税务处理暂不确认损益",
        func=rule_30401_企业重组特殊性,
    ),
    TaxRule(
        id="3-04-02",
        name="企业重组一般性税务处理",
        category=AdjustmentCategory.SPECIAL,
        tax_law_ref="企业所得税法实施条例第75条",
        description="按公允价值确认损益",
        func=rule_30402_企业重组一般性,
    ),
    TaxRule(
        id="3-05-01",
        name="政策性搬迁",
        category=AdjustmentCategory.SPECIAL,
        tax_law_ref="国家税务总局公告2012年第40号",
        description="搬迁收入扣除支出后余额计入应纳税所得额",
        func=rule_305_政策性搬迁,
    ),
    TaxRule(
        id="3-06-01",
        name="新收入准则与税法差异",
        category=AdjustmentCategory.SPECIAL,
        tax_law_ref="企业所得税法实施条例第9条",
        description="新收入准则与税法收入确认时点差异",
        func=rule_30601_新收入准则差异,
    ),
    TaxRule(
        id="3-06-02",
        name="其他特殊纳税调整",
        category=AdjustmentCategory.SPECIAL,
        tax_law_ref="企业所得税法第6章",
        description="特别纳税调整兜底项",
        func=rule_30602_其他特殊调整,
    ),
]

# 境外税收规则
OVERSEAS_RULES = [
    TaxRule(
        id="4-01-01",
        name="境外所得纳税调整",
        category=AdjustmentCategory.OVERSEAS,
        tax_law_ref="企业所得税法第17条",
        description="境外税前所得并入境内应纳税所得额",
        func=rule_401_境外所得纳税调整,
    ),
    TaxRule(
        id="4-02-01",
        name="境外所得抵免",
        category=AdjustmentCategory.OVERSEAS,
        tax_law_ref="企业所得税法第23条",
        description="境外已缴税款限额抵免",
        func=rule_402_境外所得抵免,
    ),
    TaxRule(
        id="4-03-01",
        name="境外亏损弥补",
        category=AdjustmentCategory.OVERSEAS,
        tax_law_ref="企业所得税法第17条",
        description="境外亏损不得抵减境内盈利",
        func=rule_403_境外亏损弥补,
    ),
]

# 缴纳情况规则
PAYMENT_RULES = [
    TaxRule(
        id="6-01-01",
        name="所得税预缴差异",
        category=AdjustmentCategory.PAYMENT,
        tax_law_ref="企业所得税法第54条",
        description="实际预缴与应预缴差异",
        func=rule_601_所得税预缴,
    ),
    TaxRule(
        id="6-02-01",
        name="汇总纳税分摊",
        category=AdjustmentCategory.PAYMENT,
        tax_law_ref="国家税务总局公告2012年第57号",
        description="总分支机构分摊比例",
        func=rule_602_汇总纳税分摊,
    ),
]

# 全部规则
ALL_RULES: List[TaxRule] = (
    INCOME_RULES
    + DEDUCTION_RULES
    + ASSET_RULES
    + TAX_INCENTIVE_RULES
    + SPECIAL_RULES
    + OVERSEAS_RULES
    + PAYMENT_RULES
)

# 规则索引
RULES_BY_ID = {r.id: r for r in ALL_RULES}
RULES_BY_CATEGORY = {
    AdjustmentCategory.INCOME: INCOME_RULES,
    AdjustmentCategory.DEDUCTION: DEDUCTION_RULES,
    AdjustmentCategory.ASSET: ASSET_RULES,
    AdjustmentCategory.TAX_INCENTIVE: TAX_INCENTIVE_RULES,
    AdjustmentCategory.SPECIAL: SPECIAL_RULES,
    AdjustmentCategory.OVERSEAS: OVERSEAS_RULES,
    AdjustmentCategory.PAYMENT: PAYMENT_RULES,
}
