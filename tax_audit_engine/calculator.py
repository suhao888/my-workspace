"""
税审计算引擎 — 规则编排 + 纳税调整执行 + 税额计算
"""

from typing import List, Optional, Dict, Tuple
from .models import (
    TrialBalance,
    PayrollDetail,
    AssetItem,
    EnterpriseInfo,
    TaxAdjustment,
    AdjustmentCategory,
    CalculationResult,
    AssetCategory,
)
from .rules import (
    RuleContext,
    ALL_RULES,
    RULES_BY_CATEGORY,
    AdjustmentCategory as RuleCat,
)


class TaxCalculator:
    """
    企业所得税计算器
    核心流程：会计利润 → 全部纳税调整 → 应纳税所得额 → 应纳所得税额
    """

    def __init__(self):
        self.result = CalculationResult()
        self._context: Optional[RuleContext] = None

    def calculate(
        self,
        tb: TrialBalance,
        enterprise: Optional[EnterpriseInfo] = None,
        payroll: Optional[PayrollDetail] = None,
        assets: Optional[List[AssetItem]] = None,
    ) -> CalculationResult:
        """
        执行完整的企业所得税计算

        Parameters
        ----------
        tb : TrialBalance
            试算平衡表数据（必须）
        enterprise : EnterpriseInfo, optional
            企业基本信息
        payroll : PayrollDetail, optional
            薪酬明细
        assets : List[AssetItem], optional
            固定资产卡片列表

        Returns
        -------
        CalculationResult
            完整计算结果
        """
        # 1. 构建执行上下文
        ctx = RuleContext(
            tb=tb,
            payroll=payroll,
            assets=assets or [],
            enterprise=enterprise,
        )
        self._context = ctx
        self.result = CalculationResult(
            enterprise=enterprise or EnterpriseInfo(),
            tb=tb,
        )
        self.result.accounting_profit = tb.accounting_profit

        # 2. 按类别顺序执行规则
        categories = [
            RuleCat.INCOME,
            RuleCat.DEDUCTION,
            RuleCat.ASSET,
            RuleCat.SPECIAL,
            RuleCat.OVERSEAS,
            RuleCat.TAX_INCENTIVE,
            RuleCat.PAYMENT,
        ]
        all_adjustments = []
        for cat in categories:
            adjustments = self._run_category(ctx, cat)
            all_adjustments.extend(adjustments)

        self.result.adjustments = all_adjustments

        # 3. 计算应纳税所得额
        self._compute_taxable_income()

        # 4. 计算应纳所得税额
        self._compute_tax_payable(enterprise)

        return self.result

    def _run_category(
        self, ctx: RuleContext, category: AdjustmentCategory
    ) -> List[TaxAdjustment]:
        """执行某一类别的所有规则"""
        rules = RULES_BY_CATEGORY.get(category, [])
        results = []
        for rule in rules:
            try:
                adj = rule.execute(ctx)
                if adj is not None:
                    results.append(adj)
            except Exception as e:
                results.append(
                    TaxAdjustment(
                        item_name=f"[错误] {rule.name}",
                        category=category,
                        book_amount=0,
                        tax_base=0,
                        increase=0,
                        decrease=0,
                        tax_law_ref=rule.tax_law_ref,
                        remark=f"规则执行异常: {e}",
                    )
                )
        return results

    def _compute_taxable_income(self):
        """
        计算应纳税所得额

        应纳税所得额 = 会计利润
            + 纳税调增（收入类+扣除类+资产类+特殊事项）
            - 纳税调减（收入类+扣除类+资产类+税收优惠）
        """
        total_increase = sum(a.increase for a in self.result.adjustments)
        total_decrease = sum(a.decrease for a in self.result.adjustments)

        self.result.taxable_income = (
            self.result.accounting_profit + total_increase - total_decrease
        )
        # 防止负值（亏损）
        self.result.taxable_income = max(0, self.result.taxable_income)

    def _compute_tax_payable(self, enterprise: Optional[EnterpriseInfo] = None):
        """计算应纳所得税额"""
        taxable = self.result.taxable_income

        # 判断适用税率
        # 小型微利企业分段计税（财税[2023]6号）
        if enterprise and self._is_small_micro(enterprise, taxable):
            if taxable <= 1000000:
                # 年应纳税所得额≤100万：减按25%计入，20%税率，实际税负5%
                tax = taxable * 0.25 * 0.20
                self.result.tax_rate = 0.05
                self.result.deducted_tax = taxable * 0.25 - tax
            elif taxable <= 3000000:
                # 100万<应纳税所得额≤300万：100万按5%，超出按10%
                part1 = min(taxable, 1000000) * 0.05
                part2 = max(taxable - 1000000, 0) * 0.10
                tax = part1 + part2
                self.result.tax_rate = 0.10  # 实际综合税负
                self.result.deducted_tax = taxable * 0.25 - tax
            else:
                # 超过300万：按25%
                tax = taxable * 0.25
                self.result.tax_rate = 0.25
                self.result.deducted_tax = 0
        else:
            # 一般企业：25%
            tax = taxable * 0.25
            self.result.tax_rate = 0.25
            self.result.deducted_tax = 0
            # 高新技术企业：15%
            if enterprise and enterprise.is_high_tech:
                tax = taxable * 0.15
                self.result.tax_rate = 0.15
                self.result.deducted_tax = taxable * 0.10

        self.result.tax_payable = tax
        self.result.final_tax = tax

    @staticmethod
    def _is_small_micro(enterprise: EnterpriseInfo, taxable_income: float) -> bool:
        """判断是否小型微利企业"""
        # 条件：年应纳税所得额≤300万、从业人数≤300、资产总额≤5000万
        if enterprise.is_small_micro is not None:
            return enterprise.is_small_micro  # 显式指定
        return (
            taxable_income <= 3000000
            and enterprise.employee_count <= 300
            and enterprise.total_assets <= 50000000
        )

    # ============================================================
    # 查询接口
    # ============================================================

    def get_adjustments_by_category(
        self, cat: AdjustmentCategory
    ) -> List[TaxAdjustment]:
        """按类别查询调整结果"""
        return [a for a in self.result.adjustments if a.category == cat]

    def get_adjustment(self, name: str) -> Optional[TaxAdjustment]:
        """按名称查询某项调整"""
        for a in self.result.adjustments:
            if a.item_name == name:
                return a
        return None
