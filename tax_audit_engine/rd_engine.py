"""
研发费用加计扣除审核引擎

依据：财税[2015]119号、国家税务总局公告2017年第40号、2021年第28号
对应底稿：2-05 对比表、2-10 可加计扣除审核汇总表、2-10-01~03 明细表
"""

from dataclasses import dataclass, field
from typing import List, Tuple, Optional
from tax_audit_engine.models import EnterpriseInfo


# ============================================================
# 输入数据模型
# ============================================================


@dataclass
class RDProject:
    """单个研发项目"""

    project_id: str = ""  # 项目编号，如 RD01
    project_name: str = ""  # 项目名称
    dev_method: str = (
        "自主研发"  # 研发方式：自主研发/合作研发/集中研发/委托研发(境内)/委托研发(境外)
    )
    is_qualifying: bool = True  # 是否符合条件的研发活动（财税[2015]119号）

    # 费用明细（账面发生额）
    personnel_costs: float = 0  # 人员人工费用
    direct_input: float = 0  # 直接投入
    depreciation: float = 0  # 折旧费用
    amortization: float = 0  # 无形资产摊销
    design_trial: float = 0  # 设计试验费
    other_costs: float = 0  # 其他相关费用

    # 委托研发补充字段（仅当 dev_method 含"委托研发"时使用）
    commissioned_amount: float = 0  # 委托研发实际支付金额
    commissioned_type: str = "境内"  # 境内/境外

    @property
    def base_costs(self) -> float:
        """五项费用合计（不含其他费用），用于计算其他费用上限"""
        return (
            self.personnel_costs
            + self.direct_input
            + self.depreciation
            + self.amortization
            + self.design_trial
        )

    @property
    def total_book(self) -> float:
        """账面研发费用合计"""
        if "委托研发" in self.dev_method:
            return self.commissioned_amount
        return self.base_costs + self.other_costs

    @property
    def is_commissioned(self) -> bool:
        return "委托研发" in self.dev_method

    @property
    def is_overseas(self) -> bool:
        return self.is_commissioned and self.commissioned_type == "境外"


@dataclass
class RDInput:
    """研发费加计扣除输入"""

    enterprise: EnterpriseInfo = field(default_factory=EnterpriseInfo)
    projects: List[RDProject] = field(default_factory=list)
    total_wages: float = 0  # 企业全部工资薪金（用于人员人工合规校验）
    deduction_rate: float = 1.0  # 加计扣除比例：1.0(一般), 1.2(集成电路/工业母机)


@dataclass
class RDResult:
    """研发费加计扣除计算结果"""

    enterprise: EnterpriseInfo = field(default_factory=EnterpriseInfo)
    projects: List[RDProject] = field(default_factory=list)

    total_book: float = 0  # 账面研发费用合计
    qualifying_amount: float = 0  # 可加计扣除金额
    deduction_amount: float = 0  # 加计扣除额（qualifying_amount × deduction_rate）

    has_other_limit_issue: bool = False  # 其他费用是否超10%上限
    other_limit_excess: float = 0  # 其他费用超限金额合计

    # 委托研发相关
    domestic_commissioned_total: float = 0  # 境内委托研发可扣除合计
    overseas_commissioned_total: float = 0  # 境外委托研发可扣除合计
    overseas_ceiling: float = 0  # 境外委托研发上限（境内×2/3）
    overseas_capped: bool = False  # 境外委托是否超限

    # 按项目明细
    project_details: List[dict] = field(default_factory=list)

    # 合规检查
    has_personnel_exceed_wages: bool = False  # 人员人工是否超过全部工资


# ============================================================
# 研发费用加计扣除计算引擎
# ============================================================


class RDCalculator:
    """研发费用加计扣除计算器"""

    # 其他费用上限比例
    OTHER_COST_RATIO = 0.10

    # 委托研发计入比例
    COMMISSIONED_RATIO = 0.80

    # 境外委托上限倍数（相对于境内委托）
    OVERSEAS_CEILING_MULTIPLE = 2.0 / 3.0

    def calculate(self, input_data: RDInput) -> RDResult:
        """
        主计算入口

        规则流程：
        1. 逐项目判定是否可加计扣除
        2. 其他费用10%上限调整
        3. 委托研发80%计入 + 境外2/3上限
        4. 汇总计算加计扣除额
        """
        result = RDResult(
            enterprise=input_data.enterprise,
            projects=input_data.projects,
        )

        total_qualifying = 0.0
        total_book = 0.0
        total_other_excess = 0.0
        domestic_commissioned = 0.0
        overseas_commissioned = 0.0
        total_personnel = 0.0

        details = []

        for proj in input_data.projects:
            book = proj.total_book
            total_book += book

            # 不符合条件的研发活动 → 不可加计扣除
            if not proj.is_qualifying:
                details.append(
                    self._make_detail(proj, book, 0, 0, "研发活动不符合条件", 0)
                )
                continue

            if proj.is_commissioned:
                # === 委托研发 ===
                raw = proj.commissioned_amount
                qualifying = raw * self.COMMISSIONED_RATIO

                if proj.is_overseas:
                    overseas_commissioned += qualifying
                else:
                    domestic_commissioned += qualifying

                details.append(
                    self._make_detail(
                        proj,
                        raw,
                        qualifying,
                        0,
                        f"委托研发({proj.commissioned_type}) × 80%",
                        0,
                    )
                )
                total_qualifying += qualifying
                continue

            # === 自主研发/合作研发/集中研发 ===
            # 其他费用10%上限检查
            qualified_other, excess = self._check_other_costs_limit(proj)
            total_other_excess += excess
            if excess > 0:
                result.has_other_limit_issue = True

            # 本项目可加计扣除金额
            qualifying = proj.base_costs + qualified_other
            total_qualifying += qualifying
            total_personnel += proj.personnel_costs

            details.append(
                self._make_detail(
                    proj,
                    book,
                    qualifying,
                    0,
                    f"其他费用合格={qualified_other:.2f}, 超限={excess:.2f}"
                    if excess > 0
                    else "其他费用未超限",
                    excess,
                )
            )

        # 境外委托研发上限检查
        if overseas_commissioned > 0:
            ceiling = domestic_commissioned * self.OVERSEAS_CEILING_MULTIPLE
            result.overseas_ceiling = ceiling
            if overseas_commissioned > ceiling:
                result.overseas_capped = True
                cut = overseas_commissioned - ceiling
                total_qualifying -= cut
                overseas_commissioned = ceiling

        result.domestic_commissioned_total = domestic_commissioned
        result.overseas_commissioned_total = overseas_commissioned
        result.total_book = total_book
        result.qualifying_amount = total_qualifying
        result.deduction_amount = total_qualifying * input_data.deduction_rate
        result.other_limit_excess = total_other_excess
        result.project_details = details

        # 合规检查：人员人工是否超过全部工资
        if total_personnel > input_data.total_wages > 0:
            result.has_personnel_exceed_wages = True

        return result

    def _check_other_costs_limit(self, project: RDProject) -> Tuple[float, float]:
        """
        检查其他相关费用10%上限

        财税[2015]119号：其他相关费用不超过可加计扣除研发费用总额的10%
        公式：上限 = 五项费用合计 × 10% / (1 - 10%)

        Returns:
            (合格其他费用, 超限金额)
        """
        limit = project.base_costs * self.OTHER_COST_RATIO / (1 - self.OTHER_COST_RATIO)

        if project.other_costs <= limit:
            return project.other_costs, 0.0

        return limit, project.other_costs - limit

    @staticmethod
    def _make_detail(
        project: RDProject,
        book: float,
        qualifying: float,
        deduction: float,
        note: str,
        other_excess: float,
    ) -> dict:
        return {
            "project_id": project.project_id,
            "project_name": project.project_name,
            "dev_method": project.dev_method,
            "is_qualifying": project.is_qualifying,
            "book_amount": round(book, 2),
            "qualifying_amount": round(qualifying, 2),
            "deduction_amount": round(deduction, 2),
            "other_cost_excess": round(other_excess, 2),
            "note": note,
        }

    def report(self, result: RDResult) -> str:
        """生成格式化报告"""
        lines = []
        sep = "=" * 70
        s = "-" * 70

        lines.append(sep)
        lines.append(f"  研发费用加计扣除审核报告")
        lines.append(f"  被审核单位：{result.enterprise.name}")
        lines.append(f"  纳税年度：{result.enterprise.tax_year}")
        lines.append(sep)
        lines.append("")

        lines.append("一、项目明细")
        lines.append(s)
        header = f"{'编号':<8}{'项目名称':<20}{'研发方式':<16}{'账面金额':>12}{'可扣除':>12}{'其他超限':>10}"
        lines.append(header)
        lines.append("-" * len(header))
        for d in result.project_details:
            lines.append(
                f"{d['project_id']:<8}"
                f"{d['project_name']:<20}"
                f"{d['dev_method']:<16}"
                f"{d['book_amount']:>12.2f}"
                f"{d['qualifying_amount']:>12.2f}"
                f"{d['other_cost_excess']:>10.2f}"
            )
        lines.append("")

        lines.append("二、汇总数据")
        lines.append(s)
        lines.append(f"  {'账面研发费用合计':<30}{result.total_book:>16.2f}")
        lines.append(f"  {'可加计扣除金额':<30}{result.qualifying_amount:>16.2f}")
        lines.append(
            f"  {'加计扣除比例':<30}{result.enterprise.deduction_rate if hasattr(result.enterprise, 'deduction_rate') else '100%':>16}"
        )
        lines.append(f"  {'加计扣除额':<30}{result.deduction_amount:>16.2f}")
        lines.append("")

        lines.append("三、其他费用10%上限检查")
        lines.append(s)
        if result.has_other_limit_issue:
            lines.append(
                f"  [!] 存在超限项目，超限金额合计：{result.other_limit_excess:.2f}"
            )
        else:
            lines.append(f"  [OK] 其他费用未超10%上限")
        lines.append("")

        if (
            result.domestic_commissioned_total > 0
            or result.overseas_commissioned_total > 0
        ):
            lines.append("四、委托研发")
            lines.append(s)
            lines.append(
                f"  {'境内委托研发可扣除':<30}{result.domestic_commissioned_total:>16.2f}"
            )
            lines.append(
                f"  {'境外委托研发可扣除':<30}{result.overseas_commissioned_total:>16.2f}"
            )
            if result.overseas_capped:
                lines.append(
                    f"  {'[超限] 境外委托超限，上限(境内×2/3)':<34}{result.overseas_ceiling:>16.2f}"
                )
            lines.append("")

        if result.has_personnel_exceed_wages:
            lines.append("  [!] 人员人工费用超过企业全部工资薪金，请核实！")

        lines.append(sep)
        return "\n".join(lines)
