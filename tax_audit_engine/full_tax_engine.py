"""
全税种核查引擎 FullTaxVerifier
=================================
参照中税网-2026全税种核查模板V1 设计，覆盖以下税种：
  增值税、消费税、企业所得税、个人所得税、房产税、城镇土地使用税、
  土地增值税、印花税、城建税、教育费附加、资源税、契税、耕地占用税、
  车船税、环境保护税

用法示例:
    >>> from tax_audit_engine.full_tax_engine import FullTaxVerifier, FullTaxInput, TaxDeclarationItem
    >>> items = [TaxDeclarationItem(tax_type="增值税", tax_period="2025年", declared_amount=100, actual_amount=110)]
    >>> inp = FullTaxInput(enterprise_name="测试公司", tax_id="91310115MAE1234567", audit_year="2025", items=items)
    >>> verifier = FullTaxVerifier(threshold=0.05)
    >>> result = verifier.verify(inp)
    >>> print(verifier.report(result))
"""

from __future__ import annotations

import csv
import io
import json
import math
import textwrap
from dataclasses import dataclass, field, asdict
from datetime import datetime
from typing import Dict, List, Optional, Tuple


# ---------------------------------------------------------------------------
# 数据模型
# ---------------------------------------------------------------------------


@dataclass
class TaxDeclarationItem:
    """单个税种的申报情况"""

    tax_type: str  # 税种名称
    tax_period: str  # 所属期，如 "2025年" 或 "2025第一季度"
    declared_base: float = 0  # 申报计税依据
    declared_amount: float = 0  # 已申报税额
    actual_amount: float = 0  # 应缴税额（审计认定）
    late_payment: float = 0  # 逾期缴纳（如有）
    remark: str = ""  # 备注


@dataclass
class FullTaxInput:
    """全税种核查输入"""

    enterprise_name: str
    tax_id: str
    audit_year: str
    items: List[TaxDeclarationItem] = field(default_factory=list)


@dataclass
class FullTaxResult:
    """全税种核查结果"""

    input: FullTaxInput
    total_declared: float = 0
    total_actual: float = 0
    total_difference: float = 0
    diff_rate: float = 0  # 总体差异率
    high_risk_taxes: list = field(default_factory=list)  # 高风险税种明细
    recommendation: str = ""
    item_details: list = field(default_factory=list)  # 逐项核查明细


# ---------------------------------------------------------------------------
# 全税种核查引擎
# ---------------------------------------------------------------------------

# 常见税种列表（排序与模板2-01表一致）
COMMON_TAX_TYPES = [
    "增值税",
    "消费税",
    "城市维护建设税",
    "教育费附加",
    "地方教育附加",
    "企业所得税",
    "个人所得税",
    "房产税",
    "城镇土地使用税",
    "土地增值税",
    "印花税",
    "资源税",
    "契税",
    "耕地占用税",
    "车船税",
    "环境保护税",
]

# 附加税基准映射
SURCHARGE_BASES = {
    "增值税": {"城市维护建设税": 0.07, "教育费附加": 0.03, "地方教育附加": 0.02},
    "消费税": {"城市维护建设税": 0.07, "教育费附加": 0.03, "地方教育附加": 0.02},
}


class FullTaxVerifier:
    """全税种核查引擎：比对申报数据与审计认定数据，识别高风险差异"""

    def __init__(self, threshold: float = 0.05):
        """
        Args:
            threshold: 差异率阈值，超过此值标记为高风险（默认 5%）
        """
        self.threshold = threshold

    # ------------------------------------------------------------------
    # 主核查流程
    # ------------------------------------------------------------------

    def verify(self, input_data: FullTaxInput) -> FullTaxResult:
        """执行全税种核查"""
        result = FullTaxResult(input=input_data)
        details = []

        for item in input_data.items:
            diff = item.actual_amount - item.declared_amount
            denom = max(abs(item.actual_amount), 0.01)
            rate = abs(diff) / denom

            detail = {
                "tax_type": item.tax_type,
                "tax_period": item.tax_period,
                "declared_base": item.declared_base,
                "declared_amount": item.declared_amount,
                "actual_amount": item.actual_amount,
                "difference": round(diff, 2),
                "diff_rate": round(rate, 4),
                "is_high_risk": rate > self.threshold,
                "late_payment": item.late_payment,
                "remark": item.remark,
                "risk_reason": "",
            }
            details.append(detail)

        result.item_details = details
        result.total_declared = sum(d["declared_amount"] for d in details)
        result.total_actual = sum(d["actual_amount"] for d in details)
        result.total_difference = round(result.total_actual - result.total_declared, 2)

        denom = max(abs(result.total_actual), 0.01)
        result.diff_rate = round(abs(result.total_difference) / denom, 4)

        self._classify_diffs(result)
        result.recommendation = self._generate_recommendation(result)

        return result

    # ------------------------------------------------------------------
    # 差异分类 & 风险原因推断
    # ------------------------------------------------------------------

    def _classify_diffs(self, result: FullTaxResult) -> None:
        """遍历明细，标记高风险项目并推断常见原因"""
        high_risk = []
        for d in result.item_details:
            if not d["is_high_risk"]:
                continue
            reason = self._infer_risk_reason(d["tax_type"], d)
            d["risk_reason"] = reason
            high_risk.append(d)
        result.high_risk_taxes = high_risk

    @staticmethod
    def _infer_risk_reason(tax_type: str, detail: dict) -> str:
        """
        根据税种和差异特征推断常见风险原因。
        参照底稿模板中的审核要点。
        """
        diff = detail["difference"]
        if diff > 0:
            direction = "少申报"
        else:
            direction = "多申报"

        reasons_map: Dict[str, List[str]] = {
            "增值税": [
                "进项税额抵扣是否合规 — 检查是否有虚开/异常抵扣凭证",
                "销项税额适用税率是否正确 — 检查是否错用低税率/免税政策",
                "价外费用是否并入销售额申报",
                "混合销售/兼营业务是否分别核算",
            ],
            "企业所得税": [
                "收入确认是否跨期 — 检查是否有提前/延迟确认收入",
                "成本费用配比是否合理 — 是否存在应计未计费用",
                "关联交易定价是否符合独立交易原则",
                "税前扣除凭证是否合规（国税总局公告2018年第28号）",
                "研发费用加计扣除是否准确归集",
                "资产损失税前扣除是否履行专项申报程序",
            ],
            "个人所得税": [
                "工资表与个税申报数据是否一致 — 检查是否全员全额申报",
                "年终奖计税方式是否最优",
                "劳务报酬/稿酬是否按规定代扣代缴",
                "股权激励/分红是否足额代扣个税",
                "外籍人员免税补贴是否有合规依据",
            ],
            "房产税": [
                "申报房产原值与固定资产台账是否一致",
                "从价计征 vs 从租计征分类是否正确",
                "新建/改建房产是否及时转入应税范围",
                "免税房产认定是否符合政策规定",
            ],
            "城镇土地使用税": [
                "申报面积与土地权证/台账面积是否一致",
                "土地等级适用税额是否正确",
                "免税面积认定是否有合规依据",
            ],
            "印花税": [
                "各类应税合同是否贴花完税（含电子合同）",
                "是否按核定征收方式足额缴纳",
                "营业账簿/权利许可证照是否贴花",
                "合同金额与账面金额是否一致",
            ],
            "城市维护建设税": [
                "计税依据（实缴增值税+消费税）是否准确",
                "适用税率（7%/5%/1%）是否正确",
            ],
            "教育费附加": [
                "计税依据（实缴增值税+消费税）是否准确",
                "是否有减免政策未正确适用",
            ],
            "土地增值税": [
                "清算单位划分是否合规",
                "扣除项目归集是否准确",
                "适用税率/速算扣除系数是否正确",
            ],
            "资源税": [
                "计税销售量/销售额与生产台账是否一致",
                "适用税率是否符合税目表",
                "减免税项目是否单独核算",
            ],
            "契税": [
                "计税价格与合同/评估价是否一致",
                "适用税率是否正确",
                "减免税政策适用是否合规",
            ],
            "耕地占用税": [
                "占用耕地面积与审批面积是否一致",
                "适用税额是否符合当地标准",
            ],
            "车船税": [
                "投保交强险时是否同步缴纳车船税",
                "单位自用车辆是否自行申报",
            ],
            "消费税": [
                "应税消费品品目是否完整申报",
                "从价/从量计税方式是否正确",
                "委托加工收回的应税消费品已纳税款抵扣是否合规",
            ],
            "环境保护税": [
                "污染物排放类型与申报是否一致",
                "计税依据（排污量/当量数）计算是否准确",
                "是否按规定安装自动监测设备",
            ],
        }

        reasons = reasons_map.get(
            tax_type, [f"{tax_type}申报数据与审计认定存在差异，请逐项核实"]
        )
        return f"{direction}：{'；'.join(reasons[:3])}"

    # ------------------------------------------------------------------
    # 建议生成
    # ------------------------------------------------------------------

    def _generate_recommendation(self, result: FullTaxResult) -> str:
        """根据核查结果生成审计建议"""
        high_count = len(result.high_risk_taxes)
        if high_count == 0:
            return "经核查，各税种申报数据与审计认定数据基本一致，无重大差异。建议保持现有申报流程，定期进行税务健康检查。"

        high_names = [d["tax_type"] for d in result.high_risk_taxes]
        total_diff_abs = abs(result.total_difference)

        lines = [
            f"经全税种核查，发现 {high_count} 个税种存在高风险差异：{'、'.join(high_names)}。",
            f"整体差异率 {result.diff_rate:.2%}，差异总额 {total_diff_abs:,.2f} 元。",
            "",
            "建议采取以下措施：",
        ]

        for d in result.high_risk_taxes:
            diff_str = f"{d['difference']:+,.2f} 元"
            lines.append(f"  [{d['tax_type']}] 差异 {diff_str}，{d['risk_reason']}")

        lines.extend(
            [
                "",
                "一般性建议：",
                "  1. 建立税企协同机制，及时获取最新税收优惠政策。",
                "  2. 完善税务台账管理，确保申报数据有据可查。",
                "  3. 对高风险税种实施季度滚动核查，降低滞纳金风险。",
            ]
        )

        if result.input.audit_year >= "2024":
            lines.append("  4. 关注数电发票全面推广后的进项税抵扣凭证管理变化。")

        return "\n".join(lines)

    # ------------------------------------------------------------------
    # 中文报告输出
    # ------------------------------------------------------------------

    def report(self, result: FullTaxResult, width: int = 100) -> str:
        """生成格式化中文核查报告"""
        sep = "=" * width
        sub = "-" * width

        now_str = datetime.now().strftime("%Y年%m月%d日")
        ent = result.input

        lines = [
            sep,
            "                     全 税 种 核 查 报 告",
            sep,
            f"  被审计单位：{ent.enterprise_name}",
            f"  统一社会信用代码：{ent.tax_id}",
            f"  核查年度：{ent.audit_year}年",
            f"  报告生成日期：{now_str}",
            f"  差异率阈值：{self.threshold:.0%}",
            sep,
            "",
            "一、核查汇总",
            sub,
        ]

        # 汇总统计
        total_items = len(result.item_details)
        high_count = len(result.high_risk_taxes)
        lines.extend(
            [
                f"  核查税种数：{total_items} 个",
                f"  高风险税种数：{high_count} 个",
                f"  申报税额合计：{result.total_declared:>14,.2f} 元",
                f"  认定税额合计：{result.total_actual:>14,.2f} 元",
                f"  差异总额：    {result.total_difference:>+14,.2f} 元",
                f"  总体差异率：  {result.diff_rate:>14.2%}",
                "",
            ]
        )

        # 逐项明细表
        lines.extend(
            [
                "二、逐项核查明细",
                sub,
            ]
        )

        header = (
            f"  {'税种':<14} {'所属期':<14} {'申报税额':>12} {'认定税额':>12}"
            f" {'差异额':>12} {'差异率':>8} {'风险':>6}"
        )
        lines.append(header)
        lines.append(
            f"  {'—' * 14} {'—' * 14} {'—' * 12} {'—' * 12} {'—' * 12} {'—' * 8} {'—' * 6}"
        )

        for d in result.item_details:
            risk_flag = "!!高风险" if d["is_high_risk"] else "  正常"
            line = (
                f"  {d['tax_type']:<14} {d['tax_period']:<14} "
                f"{d['declared_amount']:>12,.2f} {d['actual_amount']:>12,.2f} "
                f"{d['difference']:>+12,.2f} {d['diff_rate']:>8.2%} {risk_flag:>6}"
            )
            lines.append(line)

        lines.extend(["", sub, ""])

        # 高风险分析
        if result.high_risk_taxes:
            lines.extend(
                [
                    "三、高风险差异分析及原因推断",
                    sub,
                ]
            )
            for i, d in enumerate(result.high_risk_taxes, 1):
                lines.extend(
                    [
                        f"  {i}. {d['tax_type']}（{d['tax_period']}）",
                        f"      差异额：{d['difference']:+,.2f} 元",
                        f"      差异率：{d['diff_rate']:.2%}",
                        f"      可能原因：{d['risk_reason']}",
                        f"      备注：{d['remark'] or '无'}",
                        "",
                    ]
                )

        # 建议
        lines.extend(
            [
                "四、审计建议",
                sub,
                textwrap.fill(
                    result.recommendation,
                    width=width - 4,
                    initial_indent="  ",
                    subsequent_indent="  ",
                ),
                "",
                sep,
            ]
        )

        return "\n".join(lines)

    # ------------------------------------------------------------------
    # 工具方法：批量处理、导出
    # ------------------------------------------------------------------

    @classmethod
    def check_surcharge_consistency(
        cls,
        vat_items: List[TaxDeclarationItem],
        surcharge_items: List[TaxDeclarationItem],
    ) -> List[str]:
        """
        检查附加税（城建税/教育费附加）与增值税/消费税的勾稽关系。
        返回异常信息列表。
        """
        issues = []

        total_vat = sum(i.actual_amount for i in vat_items)
        total_vat_declared = sum(i.declared_amount for i in vat_items)

        for sur_item in surcharge_items:
            base_tax_key = "增值税"
            if "城建" in sur_item.tax_type or "教育" in sur_item.tax_type:
                expected_rate = SURCHARGE_BASES.get("增值税", {}).get(sur_item.tax_type)
                if expected_rate is None:
                    continue

                expected_declared = round(total_vat * expected_rate, 2)
                diff = sur_item.declared_amount - expected_declared
                if abs(diff) > max(abs(expected_declared) * 0.05, 1):
                    issues.append(
                        f"{sur_item.tax_type}：申报 {sur_item.declared_amount:,.2f} 元，"
                        f"按增值税{total_vat:,.2f}元×{expected_rate:.0%} 应为 "
                        f"{expected_declared:,.2f} 元，差异 {diff:+,.2f} 元"
                    )
        return issues

    @classmethod
    def from_dict(cls, data: dict) -> "FullTaxVerifier":
        """从字典构建"""
        return cls(threshold=data.get("threshold", 0.05))

    @classmethod
    def batch_verify(
        cls, inputs: List[FullTaxInput], threshold: float = 0.05
    ) -> List[FullTaxResult]:
        """批量核查"""
        verifier = cls(threshold=threshold)
        return [verifier.verify(inp) for inp in inputs]

    @classmethod
    def export_to_csv(cls, result: FullTaxResult, output_path: str) -> None:
        """导出核查明细到 CSV"""
        with io.open(output_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(
                [
                    "税种",
                    "所属期",
                    "申报计税依据",
                    "申报税额",
                    "认定税额",
                    "差异额",
                    "差异率",
                    "高风险",
                    "风险原因",
                    "备注",
                ]
            )
            for d in result.item_details:
                writer.writerow(
                    [
                        d["tax_type"],
                        d["tax_period"],
                        d["declared_base"],
                        d["declared_amount"],
                        d["actual_amount"],
                        d["difference"],
                        d["diff_rate"],
                        "是" if d["is_high_risk"] else "否",
                        d["risk_reason"],
                        d["remark"],
                    ]
                )

    @classmethod
    def export_to_json(cls, result: FullTaxResult, output_path: str) -> None:
        """导出核查结果到 JSON"""
        data = {
            "enterprise_name": result.input.enterprise_name,
            "tax_id": result.input.tax_id,
            "audit_year": result.input.audit_year,
            "threshold": None,
            "summary": {
                "total_declared": result.total_declared,
                "total_actual": result.total_actual,
                "total_difference": result.total_difference,
                "diff_rate": result.diff_rate,
                "high_risk_count": len(result.high_risk_taxes),
            },
            "high_risk_taxes": result.high_risk_taxes,
            "details": result.item_details,
            "recommendation": result.recommendation,
        }
        with io.open(output_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)


# ---------------------------------------------------------------------------
# 便捷工厂方法
# ---------------------------------------------------------------------------


def make_item(
    tax_type: str,
    period: str,
    declared_amount: float = 0,
    actual_amount: float = 0,
    declared_base: float = 0,
    late_payment: float = 0,
    remark: str = "",
) -> TaxDeclarationItem:
    """快速构造一个税种申报项"""
    return TaxDeclarationItem(
        tax_type=tax_type,
        tax_period=period,
        declared_base=declared_base,
        declared_amount=declared_amount,
        actual_amount=actual_amount,
        late_payment=late_payment,
        remark=remark,
    )


def make_full_input(
    name: str,
    tax_id: str,
    year: str,
    items: Optional[List[TaxDeclarationItem]] = None,
) -> FullTaxInput:
    """快速构造全税种核查输入"""
    return FullTaxInput(
        enterprise_name=name,
        tax_id=tax_id,
        audit_year=year,
        items=items or [],
    )
