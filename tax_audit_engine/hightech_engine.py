"""
高新技术企业认定审核鉴证引擎

依据：
- 国科发火〔2016〕32号《高新技术企业认定管理办法》
- 国科发火〔2016〕195号《高新技术企业认定管理工作指引》
对应底稿：高新认定申请材料审核、研发费用归集审核、知识产权审核
"""

from dataclasses import dataclass, field
from typing import List

from tax_audit_engine.models import EnterpriseInfo


# ============================================================
# 数据模型
# ============================================================


@dataclass
class HightechInput:
    """高新企业认定审核输入"""

    enterprise: EnterpriseInfo = field(default_factory=EnterpriseInfo)

    # 核心指标
    income_last_year: float = 0  # 最近一年销售收入
    income_3year_total: float = 0  # 近3年销售收入总额
    rd_expense_3year_total: float = 0  # 近3年研发费用总额
    rd_expense_last_year: float = 0  # 最近一年研发费用

    # 人员情况
    total_staff: int = 0  # 职工总数
    tech_staff: int = 0  # 科技人员数

    # 知识产权
    ip_count: int = 0  # 核心知识产权数量（I类）
    ip_2_count: int = 0  # 其他知识产权数量（II类）

    # 高新技术产品
    hi_product_revenue: float = 0  # 高新技术产品（服务）收入（最近一年）
    hi_product_revenue_3year: float = 0  # 近3年高新产品总收入
    total_revenue_3year: float = 0  # 近3年总收入

    # 其他
    has_safety_accident: bool = False  # 是否发生重大安全事故
    has_env_violation: bool = False  # 是否发生重大环境违法
    rd_management_system: bool = True  # 是否有研发管理制度


@dataclass
class HightechResult:
    """高新认定审核结果"""

    input: HightechInput = field(default_factory=HightechInput)
    passed: bool = False

    # 各指标得分
    ip_score: float = 0  # 知识产权（≤30分）
    tech_transfer_score: float = 0  # 科技成果转化能力（≤30分）
    rd_management_score: float = 0  # 研发组织管理水平（≤20分）
    growth_score: float = 0  # 企业成长性（≤20分）
    total_score: float = 0

    # 门槛条件检查
    rd_ratio_passed: bool = False
    tech_staff_ratio_passed: bool = False
    hi_product_ratio_passed: bool = False
    safe_env_passed: bool = True

    # 明细评分
    ip_detail: dict = field(default_factory=dict)
    transfer_detail: str = ""
    management_detail: str = ""
    growth_detail: str = ""

    # 违规条款
    violations: List[str] = field(default_factory=list)


# ============================================================
# 高新认定审核引擎
# ============================================================


class HightechVerifier:
    """高新技术企业认定审核鉴证引擎"""

    # 评分上限
    MAX_IP_SCORE = 30.0
    MAX_TRANSFER_SCORE = 30.0
    MAX_MANAGEMENT_SCORE = 20.0
    MAX_GROWTH_SCORE = 20.0
    PASS_THRESHOLD = 71.0

    def verify(self, input_data: HightechInput) -> HightechResult:
        """
        主审核入口
        1. 门槛检查 → 直接FAIL
        2. 评分计算 → 总分≥71通过
        """
        result = HightechResult(input=input_data)

        self._check_thresholds(input_data, result)

        if not result.passed:
            return result

        result.ip_score = self._score_ip(input_data, result)
        result.tech_transfer_score = self._score_transfer(input_data, result)
        result.rd_management_score = self._score_management(input_data, result)
        result.growth_score = self._score_growth(input_data, result)

        result.total_score = (
            result.ip_score
            + result.tech_transfer_score
            + result.rd_management_score
            + result.growth_score
        )
        result.total_score = min(result.total_score, 100.0)
        result.passed = result.total_score >= self.PASS_THRESHOLD

        return result

    def _check_thresholds(self, input_data: HightechInput, result: HightechResult):
        """
        门槛条件检查（四项必须全部满足）
        """
        violations = []

        # 1. 研发费用占比
        if input_data.income_3year_total > 0:
            rd_ratio = input_data.rd_expense_3year_total / input_data.income_3year_total
        else:
            rd_ratio = 0

        if input_data.income_last_year <= 50_000_000:
            required_rd_ratio = 0.05
        elif input_data.income_last_year <= 200_000_000:
            required_rd_ratio = 0.04
        else:
            required_rd_ratio = 0.03

        result.rd_ratio_passed = rd_ratio >= required_rd_ratio
        if not result.rd_ratio_passed:
            violations.append(
                f"研发费用占比不达标：实际{rd_ratio * 100:.2f}%，"
                f"要求≥{required_rd_ratio * 100:.0f}%"
            )

        # 2. 科技人员占比
        if input_data.total_staff > 0:
            tech_ratio = input_data.tech_staff / input_data.total_staff
        else:
            tech_ratio = 0

        result.tech_staff_ratio_passed = tech_ratio >= 0.10
        if not result.tech_staff_ratio_passed:
            violations.append(
                f"科技人员占比不达标：实际{tech_ratio * 100:.2f}%，要求≥10%"
            )

        # 3. 高新产品收入占比
        if input_data.total_revenue_3year > 0:
            hi_ratio = (
                input_data.hi_product_revenue_3year / input_data.total_revenue_3year
            )
        else:
            hi_ratio = 0

        result.hi_product_ratio_passed = hi_ratio >= 0.60
        if not result.hi_product_ratio_passed:
            violations.append(
                f"高新技术产品收入占比不达标：实际{hi_ratio * 100:.2f}%，要求≥60%"
            )

        # 4. 安全/环境合规
        result.safe_env_passed = (
            not input_data.has_safety_accident and not input_data.has_env_violation
        )
        if not result.safe_env_passed:
            reasons = []
            if input_data.has_safety_accident:
                reasons.append("重大安全事故")
            if input_data.has_env_violation:
                reasons.append("重大环境违法")
            violations.append(f"存在合规问题：{'、'.join(reasons)}")

        result.violations = violations
        result.passed = len(violations) == 0

    def _score_ip(self, input_data: HightechInput, result: HightechResult) -> float:
        """
        知识产权评分（≤30分）

        - I类知识产权：每个5-8分，≤24分
        - II类知识产权：每个1-3分，≤6分
        - 技术先进性：0-5分
        - 核心支持作用：0-3分（自动估算）
        """
        detail = {}

        # I类知识产权评分
        if input_data.ip_count >= 5:
            ip1_score = 24.0
        elif input_data.ip_count >= 3:
            ip1_score = 15.0 + (input_data.ip_count - 3) * 4.5
        elif input_data.ip_count >= 1:
            ip1_score = 5.0 + (input_data.ip_count - 1) * 5.0
        else:
            ip1_score = 0.0
        ip1_score = min(ip1_score, 24.0)
        detail["i_class_score"] = round(ip1_score, 1)
        detail["i_class_count"] = input_data.ip_count

        # II类知识产权评分
        if input_data.ip_2_count >= 2:
            ip2_score = 6.0
        elif input_data.ip_2_count == 1:
            ip2_score = 3.0
        else:
            ip2_score = 0.0
        detail["ii_class_score"] = round(ip2_score, 1)
        detail["ii_class_count"] = input_data.ip_2_count

        # 技术先进性（基于研发投入强度估算）
        if input_data.income_3year_total > 0:
            rd_intensity = (
                input_data.rd_expense_3year_total / input_data.income_3year_total
            )
        else:
            rd_intensity = 0
        if rd_intensity >= 0.08:
            tech_advance = 5.0
        elif rd_intensity >= 0.05:
            tech_advance = 3.0
        elif rd_intensity >= 0.03:
            tech_advance = 1.0
        else:
            tech_advance = 0.0
        detail["tech_advance_score"] = round(tech_advance, 1)
        detail["rd_intensity"] = round(rd_intensity * 100, 2)

        # 核心支持作用（高新产品收入占比估算）
        if input_data.total_revenue_3year > 0:
            support_ratio = (
                input_data.hi_product_revenue_3year / input_data.total_revenue_3year
            )
        else:
            support_ratio = 0
        if support_ratio >= 0.80:
            support_score = 3.0
        elif support_ratio >= 0.60:
            support_score = 2.0
        elif support_ratio >= 0.30:
            support_score = 1.0
        else:
            support_score = 0.0
        detail["core_support_score"] = round(support_score, 1)
        detail["hi_product_revenue_ratio"] = round(support_ratio * 100, 2)

        total = ip1_score + ip2_score + tech_advance + support_score
        total = min(total, self.MAX_IP_SCORE)

        result.ip_detail = detail
        return round(total, 1)

    def _score_transfer(
        self, input_data: HightechInput, result: HightechResult
    ) -> float:
        """
        科技成果转化能力评分（≤30分）

        以高新收入占最近一年收入比作为转化能力的替代衡量
        """
        if input_data.income_last_year > 0:
            ratio = input_data.hi_product_revenue / input_data.income_last_year
        else:
            ratio = 0
        ratio_pct = ratio * 100

        if ratio_pct > 70:
            score = 20.0 + (ratio_pct - 70) / 30 * 10.0
            score = min(score, 30.0)
            result.transfer_detail = (
                f"高新产品收入占比{ratio_pct:.1f}% > 70%，转化能力强"
            )
        elif ratio_pct > 50:
            score = 10.0 + (ratio_pct - 50) / 20 * 9.0
            result.transfer_detail = (
                f"高新产品收入占比{ratio_pct:.1f}%，50%-70%区间，转化能力中等"
            )
        else:
            score = ratio_pct / 50 * 9.0
            result.transfer_detail = (
                f"高新产品收入占比{ratio_pct:.1f}% < 50%，转化能力较弱"
            )

        return round(score, 1)

    def _score_management(
        self, input_data: HightechInput, result: HightechResult
    ) -> float:
        """
        研发组织管理水平评分（≤20分）
        """
        score = 0.0
        notes = []

        if input_data.rd_management_system:
            score += 5.0
            notes.append("有研发管理制度(+5)")

        if input_data.total_staff > 0:
            tech_ratio = input_data.tech_staff / input_data.total_staff
            if tech_ratio >= 0.10:
                score += 3.0
                notes.append("科技人员占比达标(+3)")

        if input_data.hi_product_revenue > 0:
            score += 4.0
            notes.append("高新技术产品有收入(+4)")

        if input_data.rd_expense_last_year > 0:
            score += 2.0
            notes.append("有研发费用投入(+2)")

        if input_data.ip_count > 0:
            bonus = min(input_data.ip_count * 1.0, 4.0)
            score += bonus
            notes.append(f"知识产权数量加分(+{bonus:.0f})")

        score = min(score, self.MAX_MANAGEMENT_SCORE)
        result.management_detail = "；".join(notes)
        return round(score, 1)

    def _score_growth(self, input_data: HightechInput, result: HightechResult) -> float:
        """
        企业成长性评分（≤20分）

        用近3年收入增长率近似代替净资产增长率
        """
        if input_data.income_last_year > 0 and input_data.income_3year_total > 0:
            avg_income = input_data.income_3year_total / 3
            if avg_income > 0:
                growth_rate = (input_data.income_last_year - avg_income) / avg_income
            else:
                growth_rate = 0
        else:
            growth_rate = 0

        rate_pct = growth_rate * 100

        if rate_pct > 30:
            score = 6.0 + min((rate_pct - 30) / 70 * 4.0, 4.0)
            result.growth_detail = f"收入增长率{rate_pct:.1f}% > 30%，成长性强"
        elif rate_pct > 15:
            score = 3.0 + (rate_pct - 15) / 15 * 2.0
            result.growth_detail = f"收入增长率{rate_pct:.1f}%，15%-30%区间"
        elif rate_pct > 0:
            score = 1.0 + rate_pct / 15 * 1.0
            result.growth_detail = f"收入增长率{rate_pct:.1f}%，0%-15%区间"
        else:
            score = 0.0
            result.growth_detail = f"收入增长率{rate_pct:.1f}%，负增长"

        score = min(score, 10.0)
        result.growth_detail += f"，得分{score:.1f}/10"
        score = score * 2.0

        return round(score, 1)

    def report(self, result: HightechResult) -> str:
        """生成格式化审核报告"""
        lines = []
        sep = "=" * 70
        s = "-" * 70
        inp = result.input
        ent = inp.enterprise

        lines.append(sep)
        lines.append("  高新技术企业认定审核报告")
        lines.append(sep)
        lines.append(f"  被审核单位：{ent.name or '(未填写)'}")
        lines.append(f"  统一社会信用代码：{ent.uscc or '(未填写)'}")
        lines.append(f"  所属行业：{ent.industry or '(未填写)'}")
        lines.append(f"  审核年度：{ent.tax_year}")
        lines.append("")
        lines.append("  依据：国科发火〔2016〕32号《高新技术企业认定管理办法》")
        lines.append(f"       国科发火〔2016〕195号《高新技术企业认定管理工作指引》")
        lines.append(sep)
        lines.append("")

        # 一、门槛条件
        lines.append("一、门槛条件检查")
        lines.append(s)

        rd_ok = "OK" if result.rd_ratio_passed else "NG"
        if inp.income_3year_total > 0:
            rd_ratio = inp.rd_expense_3year_total / inp.income_3year_total
        else:
            rd_ratio = 0

        if inp.income_last_year <= 50_000_000:
            rd_req = "5%"
        elif inp.income_last_year <= 200_000_000:
            rd_req = "4%"
        else:
            rd_req = "3%"

        lines.append(
            f"  [{rd_ok}] 研发费用占比：{rd_ratio * 100:.2f}%（要求≥{rd_req}）"
        )

        tech_ratio = inp.tech_staff / inp.total_staff if inp.total_staff > 0 else 0
        tech_ok = "OK" if result.tech_staff_ratio_passed else "NG"
        lines.append(f"  [{tech_ok}] 科技人员占比：{tech_ratio * 100:.2f}%（要求≥10%）")

        hi_ratio = (
            inp.hi_product_revenue_3year / inp.total_revenue_3year
            if inp.total_revenue_3year > 0
            else 0
        )
        hi_ok = "OK" if result.hi_product_ratio_passed else "NG"
        lines.append(f"  [{hi_ok}] 高新产品收入占比：{hi_ratio * 100:.2f}%（要求≥60%）")

        safe_ok = "OK" if result.safe_env_passed else "NG"
        lines.append(
            f"  [{safe_ok}] 安全/环境合规：{'无重大事故/违法' if result.safe_env_passed else '存在合规问题'}"
        )
        lines.append("")

        if result.violations:
            for v in result.violations:
                lines.append(f"    [!] {v}")
            lines.append("")
            lines.append("  >> 门槛条件未通过，审核终止 <<")
            lines.append("")
            lines.append(sep)
            return "\n".join(lines)

        lines.append("  >> 全部门槛条件通过 <<")
        lines.append("")

        # 二、关键指标一览
        lines.append("二、关键指标一览")
        lines.append(s)
        lines.append(f"  {'最近一年销售收入':<30}{inp.income_last_year:>16,.2f}")
        lines.append(f"  {'近3年销售收入总额':<30}{inp.income_3year_total:>16,.2f}")
        lines.append(f"  {'近3年研发费用总额':<30}{inp.rd_expense_3year_total:>16,.2f}")
        lines.append(f"  {'科技人员/职工总数':<30}{inp.tech_staff}/{inp.total_staff}")
        lines.append(f"  {'I类知识产权数量':<30}{inp.ip_count}")
        lines.append(f"  {'II类知识产权数量':<30}{inp.ip_2_count}")
        lines.append(
            f"  {'高新技术产品收入(近1年)':<30}{inp.hi_product_revenue:>16,.2f}"
        )
        lines.append(
            f"  {'研发管理制度':<30}{'有' if inp.rd_management_system else '无'}"
        )
        lines.append("")

        # 三、评分明细
        lines.append("三、评分明细")
        lines.append(s)

        # 知识产权
        lines.append(f"  1. 知识产权（≤30分）")
        if result.ip_detail:
            d = result.ip_detail
            lines.append(
                f"     I类({d.get('i_class_count', 0)}个)：{d.get('i_class_score', 0):.1f}分"
            )
            lines.append(
                f"     II类({d.get('ii_class_count', 0)}个)：{d.get('ii_class_score', 0):.1f}分"
            )
            lines.append(f"     技术先进性：{d.get('tech_advance_score', 0):.1f}分")
            lines.append(f"     核心支持作用：{d.get('core_support_score', 0):.1f}分")
        lines.append(f"     小计：{result.ip_score:.1f}分")
        lines.append("")

        # 科技成果转化
        lines.append(f"  2. 科技成果转化能力（≤30分）")
        lines.append(f"     {result.transfer_detail}")
        lines.append(f"     小计：{result.tech_transfer_score:.1f}分")
        lines.append("")

        # 研发组织管理
        lines.append(f"  3. 研发组织管理水平（≤20分）")
        if result.management_detail:
            lines.append(f"     {result.management_detail}")
        lines.append(f"     小计：{result.rd_management_score:.1f}分")
        lines.append("")

        # 企业成长性
        lines.append(f"  4. 企业成长性（≤20分）")
        lines.append(f"     {result.growth_detail}")
        lines.append(f"     小计：{result.growth_score:.1f}分")
        lines.append("")

        # 四、审核结论
        lines.append(s)
        lines.append(
            f"  总分：{result.total_score:.1f}分（通过线：{self.PASS_THRESHOLD}分）"
        )
        if result.passed:
            lines.append(f"  >> 审核结果：通过 OK")
        else:
            lines.append(f"  >> 审核结果：不通过 NG（总分不足{self.PASS_THRESHOLD}分）")

        lines.append(sep)
        return "\n".join(lines)


# ============================================================
# 演示
# ============================================================


def demo():
    """演示：生成一组测试数据并输出完整评分报告"""
    ent = EnterpriseInfo(
        name="XX智能科技有限公司",
        uscc="91440100MA5XXXXXX",
        industry="软件和信息技术服务业",
        tax_year=2025,
    )

    inp = HightechInput(
        enterprise=ent,
        income_last_year=85_000_000,
        income_3year_total=210_000_000,
        rd_expense_3year_total=15_000_000,
        rd_expense_last_year=6_000_000,
        total_staff=200,
        tech_staff=65,
        ip_count=5,
        ip_2_count=3,
        hi_product_revenue=68_000_000,
        hi_product_revenue_3year=168_000_000,
        total_revenue_3year=210_000_000,
        has_safety_accident=False,
        has_env_violation=False,
        rd_management_system=True,
    )

    verifier = HightechVerifier()
    result = verifier.verify(inp)
    print(verifier.report(result))
    return result


if __name__ == "__main__":
    demo()
