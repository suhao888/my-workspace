"""
财产损失税前扣除审核引擎
依据：国家税务总局公告2011年第25号《企业资产损失所得税税前扣除管理办法》
"""

from dataclasses import dataclass, field
from typing import List, Tuple
from enum import Enum

from tax_audit_engine.models import EnterpriseInfo


# ============================================================
# 枚举
# ============================================================


class LossCategory(str, Enum):
    """财产损失类型"""

    CASH = "现金"
    BANK_DEPOSIT = "银行存款"
    RECEIVABLE = "应收款项"
    INVENTORY = "存货"
    FIXED_ASSET = "固定资产"
    INTANGIBLE = "无形资产"
    INVESTMENT = "投资"
    OTHER = "其他"


class DeclareType(str, Enum):
    """申报类型"""

    LIST = "清单申报"  # 清单申报
    SPECIAL = "专项申报"  # 专项申报


class EvidenceLevel(str, Enum):
    """证据完备性"""

    COMPLETE = "齐全"
    BASIC = "基本齐全"
    MISSING = "缺失"


# ============================================================
# 数据模型
# ============================================================


@dataclass
class LossItem:
    """单笔财产损失"""

    loss_id: str  # 损失编号
    asset_name: str  # 资产名称
    category: str = ""  # 损失类型：现金/银行存款/应收/存货/固定资产/无形资产/投资/其他
    loss_type: str = ""  # 申报类型：清单申报/专项申报
    book_value: float = 0  # 账面净值
    recoverable: float = 0  # 可收回金额（保险赔款+责任人赔偿+残值）
    loss_amount: float = 0  # 损失金额 = book_value - recoverable
    evidence_level: str = "齐全"  # 证据完备性：齐全/基本齐全/缺失
    has_external_evidence: bool = True  # 是否有外部证据
    approval_status: str = ""  # 内部审批情况
    remark: str = ""


@dataclass
class LossInput:
    """财产损失审核输入"""

    enterprise: EnterpriseInfo = field(default_factory=EnterpriseInfo)
    items: List[LossItem] = field(default_factory=list)


@dataclass
class LossResult:
    """财产损失审核结果"""

    input: LossInput = field(default_factory=lambda: LossInput())
    total_book_value: float = 0
    total_recoverable: float = 0
    total_loss: float = 0
    qualifying_loss: float = 0  # 可税前扣除损失
    disallowed_loss: float = 0  # 不得扣除损失
    item_count: int = 0
    issue_items: List[str] = field(default_factory=list)  # 有问题的损失项
    item_results: List[dict] = field(default_factory=list)  # 逐项审核明细


# ============================================================
# 损失类型与规则映射
# ============================================================

# 应税规章引用
LAW_BASE = "国家税务总局公告2011年第25号"

CLASSIFICATION_RULES = {
    "应收": (DeclareType.LIST, "应收款项正常损失适用清单申报"),
    "应收账款": (DeclareType.LIST, "应收款项正常损失适用清单申报"),
    "其他应收款": (DeclareType.LIST, "应收款项正常损失适用清单申报"),
    "存货": (DeclareType.LIST, "存货正常损失适用清单申报"),
    "原材料": (DeclareType.LIST, "存货正常损失适用清单申报"),
    "库存商品": (DeclareType.LIST, "存货正常损失适用清单申报"),
    "在产品": (DeclareType.LIST, "存货正常损失适用清单申报"),
    "固定资产": (DeclareType.LIST, "固定资产正常损失适用清单申报"),
    "机器设备": (DeclareType.LIST, "固定资产正常损失适用清单申报"),
    "房屋建筑物": (DeclareType.LIST, "固定资产正常损失适用清单申报"),
    "运输工具": (DeclareType.LIST, "固定资产正常损失适用清单申报"),
    "电子设备": (DeclareType.LIST, "固定资产正常损失适用清单申报"),
    "无形资产": (DeclareType.SPECIAL, "无形资产损失适用专项申报，依据第25号令第24条"),
    "投资": (DeclareType.SPECIAL, "投资损失适用专项申报，依据第25号令第38-43条"),
    "长期股权投资": (
        DeclareType.SPECIAL,
        "投资损失适用专项申报，依据第25号令第38-43条",
    ),
    "金融资产": (DeclareType.SPECIAL, "投资损失适用专项申报，依据第25号令第38-43条"),
    "现金": (
        DeclareType.SPECIAL,
        "货币性资产损失需公安机关报案证明，依据第25号令第19-20条",
    ),
    "银行存款": (
        DeclareType.SPECIAL,
        "货币性资产损失需公安机关报案证明，依据第25号令第19-20条",
    ),
    "其他": (DeclareType.SPECIAL, "无法明确分类的损失建议按专项申报处理"),
}

# 需公安证明的损失类型
POLICE_REQUIRED = {"现金", "银行存款"}


# ============================================================
# 核心引擎
# ============================================================


class LossCalculator:
    """财产损失税前扣除审核计算器"""

    def calculate(self, input_data: LossInput) -> LossResult:
        result = LossResult(input=input_data)
        result.item_count = len(input_data.items)

        for item in input_data.items:
            item_result = self._audit_item(item)
            result.item_results.append(item_result)

            result.total_book_value += item.book_value
            result.total_recoverable += item.recoverable
            result.total_loss += item.loss_amount

            if item_result["disallowed"]:
                result.disallowed_loss += item.loss_amount
                result.issue_items.append(item.loss_id)
            else:
                result.qualifying_loss += item.loss_amount

        return result

    def _audit_item(self, item: LossItem) -> dict:
        """逐项审核"""
        declare_type, remark = self._check_loss_type(item)
        issues = []
        disallowed = False

        # 1. 证据完备性检查
        if item.evidence_level != "齐全":
            issues.append(f"证据完备性为'{item.evidence_level}'，非'齐全'")
            if item.evidence_level == "缺失":
                disallowed = True

        # 2. 外部证据检查
        if not item.has_external_evidence:
            issues.append("缺少外部证据（司法机关/公安机关/保险公司等证明文件）")
            disallowed = True

        # 3. 可收回金额合理性检查
        if item.recoverable > item.book_value:
            issues.append(
                f"可收回金额({item.recoverable:.2f})超过账面净值({item.book_value:.2f})"
            )
        if item.recoverable < 0:
            issues.append(f"可收回金额为负数({item.recoverable:.2f})")
            disallowed = True

        # 4. 损失金额验算
        expected_loss = item.book_value - item.recoverable
        if abs(expected_loss - item.loss_amount) >= 0.01:
            issues.append(
                f"损失金额({item.loss_amount:.2f})不等于账面净值-可收回金额"
                f"({item.book_value:.2f}-{item.recoverable:.2f}={expected_loss:.2f})"
            )
            disallowed = True

        # 5. 内部审批检查
        if not item.approval_status:
            issues.append("未经企业内部审批")
            disallowed = True

        # 6. 货币性资产专项检查（需公安证明）
        if item.category in POLICE_REQUIRED:
            if not item.has_external_evidence:
                issues.append(
                    f"{item.category}损失需公安机关报案证明（第25号令第19条）"
                )
            declare_type = DeclareType.SPECIAL

        # 7. 申报类型自动推断
        auto_type, type_remark = self._check_loss_type(item)
        if item.loss_type and item.loss_type != auto_type:
            issues.append(f"申报类型'{item.loss_type}'与规则推断'{auto_type}'不一致")

        return {
            "loss_id": item.loss_id,
            "asset_name": item.asset_name,
            "declare_type": declare_type,
            "type_remark": type_remark,
            "book_value": item.book_value,
            "recoverable": item.recoverable,
            "loss_amount": item.loss_amount,
            "evidence_level": item.evidence_level,
            "has_external_evidence": item.has_external_evidence,
            "approval_status": item.approval_status or "未审批",
            "issues": issues,
            "disallowed": disallowed,
            "passed": not disallowed and len(issues) == 0,
            "remark": remark,
        }

    def _check_loss_type(self, item: LossItem) -> Tuple[str, str]:
        """根据资产名称和类别判断申报类型"""
        name = item.asset_name.strip()
        cat = item.category.strip() if item.category else ""

        # 先按名称精确匹配
        for key in sorted(CLASSIFICATION_RULES, key=len, reverse=True):
            if key in name or key in cat:
                return CLASSIFICATION_RULES[key]

        # 按类别模糊匹配
        for rule_cat, (declare, reason) in CLASSIFICATION_RULES.items():
            if rule_cat in cat:
                return declare, reason

        return DeclareType.SPECIAL, "未能识别损失类型，建议按专项申报处理"

    def report(self, result: LossResult) -> str:
        """生成审核报告"""
        lines = []
        ent = result.input.enterprise

        lines.append("=" * 70)
        lines.append("企业资产损失所得税税前扣除审核报告")
        lines.append("=" * 70)
        lines.append(f"被审核单位：{ent.name or '(未填写)'}")
        lines.append(f"统一社会信用代码：{ent.uscc or '(未填写)'}")
        lines.append(f"纳税年度：{ent.tax_year}")
        lines.append(f"审核日期：{ent.tax_year}年度")
        lines.append("")
        lines.append(f"审核依据：{LAW_BASE}《企业资产损失所得税税前扣除管理办法》")
        lines.append("-" * 70)

        # 汇总
        items_total = result.item_count
        issue_count = len(result.issue_items)
        lines.append("一、资产损失审核汇总")
        lines.append(f"  损失项数：{items_total} 项")
        lines.append(f"  账面净值合计：{result.total_book_value:>12,.2f}")
        lines.append(f"  可收回金额合计：{result.total_recoverable:>12,.2f}")
        lines.append(f"  申报损失合计：{result.total_loss:>12,.2f}")
        lines.append(f"  可税前扣除损失：{result.qualifying_loss:>12,.2f}")
        lines.append(f"  不得扣除损失：{result.disallowed_loss:>12,.2f}")
        lines.append(f"  问题项数：{issue_count}/{items_total}")
        lines.append("")

        # 逐项明细
        lines.append("二、逐项审核明细")
        lines.append("-" * 70)
        for r in result.item_results:
            status = "[通过]" if r["passed"] else "[问题]"
            lines.append(f"  [{status}] {r['loss_id']}: {r['asset_name']}")
            lines.append(f"    申报类型：{r['declare_type']}（{r['type_remark']}）")
            lines.append(
                f"    账面净值：{r['book_value']:>12,.2f}  |  "
                f"可收回：{r['recoverable']:>10,.2f}  |  "
                f"损失金额：{r['loss_amount']:>10,.2f}"
            )
            lines.append(
                f"    证据：{r['evidence_level']}  |  "
                f"外部证据：{'有' if r['has_external_evidence'] else '无'}  |  "
                f"审批：{r['approval_status']}"
            )
            if r["issues"]:
                for iss in r["issues"]:
                    lines.append(f"    [!] {iss}")
            lines.append("")

        # 处理建议
        if result.issue_items:
            lines.append("三、审核发现及处理建议")
            lines.append("-" * 70)
            for idx, r in enumerate(result.item_results, 1):
                if r["issues"]:
                    lines.append(f"  {idx}. {r['loss_id']} {r['asset_name']}")
                    for iss in r["issues"]:
                        lines.append(f"    - {iss}")
                    if r["disallowed"]:
                        lines.append(f"    => 建议：不得税前扣除，需补充证据后重新申报")
                    else:
                        lines.append(f"    => 建议：补充完善后仍可税前扣除")
                    lines.append("")

        lines.append("=" * 70)
        lines.append(f"审核结论：")
        if result.disallowed_loss > 0:
            lines.append(
                f"  部分损失不得税前扣除，需调整应纳税所得额 "
                f"{result.disallowed_loss:,.2f}"
            )
        else:
            lines.append(f"  全部损失可税前扣除，合计 {result.qualifying_loss:,.2f}")
        lines.append("=" * 70)

        return "\n".join(lines)


# ============================================================
# 演示
# ============================================================


def demo():
    """简单演示：生成一批测试数据并输出审核报告"""
    ent = EnterpriseInfo(
        name="XXX制造有限公司",
        uscc="91440101MA5XXXXXX",
        tax_year=2025,
    )

    items = [
        LossItem(
            loss_id="L001",
            asset_name="应收账款-客户A",
            category="应收款项",
            book_value=500000,
            recoverable=0,
            loss_amount=500000,
            evidence_level="齐全",
            has_external_evidence=True,
            approval_status="已审批",
        ),
        LossItem(
            loss_id="L002",
            asset_name="原材料-钢材",
            category="存货",
            book_value=120000,
            recoverable=15000,
            loss_amount=105000,
            evidence_level="齐全",
            has_external_evidence=True,
            approval_status="已审批",
        ),
        LossItem(
            loss_id="L003",
            asset_name="机器设备-冲压机",
            category="固定资产",
            book_value=800000,
            recoverable=50000,
            loss_amount=750000,
            evidence_level="基本齐全",
            has_external_evidence=True,
            approval_status="已审批",
        ),
        LossItem(
            loss_id="L004",
            asset_name="库存现金-被盗",
            category="现金",
            book_value=50000,
            recoverable=0,
            loss_amount=50000,
            evidence_level="齐全",
            has_external_evidence=False,
            approval_status="已报案未出证明",
            remark="已向公安机关报案，尚未取得报案回执",
        ),
        LossItem(
            loss_id="L005",
            asset_name="无形资产-专利权",
            category="无形资产",
            book_value=300000,
            recoverable=0,
            loss_amount=300000,
            evidence_level="齐全",
            has_external_evidence=True,
            approval_status="已审批",
        ),
        LossItem(
            loss_id="L006",
            asset_name="库存商品-电子产品",
            category="存货",
            book_value=250000,
            recoverable=50000,
            loss_amount=200000,
            evidence_level="齐全",
            has_external_evidence=True,
            approval_status="",
        ),
    ]

    inp = LossInput(enterprise=ent, items=items)
    calc = LossCalculator()
    result = calc.calculate(inp)
    print(calc.report(result))
    return result
