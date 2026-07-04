"""
税审底稿自动填充引擎 — 统一入口

支持6大业务类型：
  1. 企业所得税汇算清缴审核 (eit)
  2. 研发费用加计扣除审核 (rd)
  3. 全税种核查 (fulltax)
  4. 财产损失税前扣除审核 (loss)
  5. 高新技术企业认定审核 (hightech)
  6. 全流程整体演示 (all)
"""

import sys
from pathlib import Path
from typing import Optional

from .models import EnterpriseInfo
from .calculator import TaxCalculator
from .workpaper_generator import WorkpaperGenerator


def run_eit(output_path: Optional[str] = None):
    """
    企业所得税汇算清缴
    """
    from .main import run_demo

    result = run_demo(output_path)
    return result


def run_rd():
    """
    研发费用加计扣除审核
    """
    from .rd_engine import RDCalculator, RDInput, RDProject

    enterprise = EnterpriseInfo(
        name="XX科技公司",
        uscc="91310115MAE1234567",
        industry="制造业",
        is_high_tech=True,
    )

    projects = [
        RDProject(
            project_id="RD01",
            project_name="AI视觉识别系统研发",
            dev_method="自主研发",
            personnel_costs=500000,
            direct_input=1200000,
            depreciation=300000,
            amortization=100000,
            design_trial=200000,
            other_costs=150000,
        ),
        RDProject(
            project_id="RD02",
            project_name="云计算平台V2.0升级",
            dev_method="自主研发",
            personnel_costs=300000,
            direct_input=200000,
            depreciation=150000,
            amortization=50000,
            design_trial=80000,
            other_costs=200000,  # 超限测试
        ),
        RDProject(
            project_id="RD03",
            project_name="智能边缘计算终端",
            dev_method="委托研发",
            is_qualifying=True,
            personnel_costs=0,
            direct_input=0,
            depreciation=0,
            amortization=0,
            design_trial=0,
            other_costs=0,
        ),
    ]
    # 委托境外
    projects[2]._entrusted_domestic = 500000
    projects[2]._entrusted_overseas = 400000

    input_data = RDInput(
        enterprise=enterprise,
        projects=projects,
        total_wages=3000000,
    )

    calc = RDCalculator()
    result = calc.calculate(input_data)
    print(calc.report(result))
    return result


def run_fulltax():
    """
    全税种核查
    """
    from .full_tax_engine import (
        FullTaxVerifier,
        FullTaxInput,
        TaxDeclarationItem,
    )

    items = [
        TaxDeclarationItem(
            "增值税", "2025年", declared_amount=5000000, actual_amount=5100000
        ),
        TaxDeclarationItem(
            "企业所得税", "2025年", declared_amount=1200000, actual_amount=1361600
        ),
        TaxDeclarationItem(
            "个人所得税", "2025年", declared_amount=450000, actual_amount=445000
        ),
        TaxDeclarationItem(
            "房产税", "2025年", declared_amount=240000, actual_amount=310000
        ),
        TaxDeclarationItem(
            "土地使用税", "2025年", declared_amount=80000, actual_amount=82000
        ),
        TaxDeclarationItem(
            "印花税", "2025年", declared_amount=35000, actual_amount=33000
        ),
        TaxDeclarationItem(
            "城建税", "2025年", declared_amount=350000, actual_amount=357000
        ),
        TaxDeclarationItem(
            "教育费附加", "2025年", declared_amount=250000, actual_amount=255000
        ),
    ]

    input_data = FullTaxInput(
        enterprise_name="XX科技公司",
        tax_id="91310115MAE1234567",
        audit_year="2025",
        items=items,
    )

    result = FullTaxVerifier(threshold=0.05).verify(input_data)
    print(FullTaxVerifier().report(result))
    return result


def run_loss():
    """
    财产损失税前扣除审核
    """
    from .loss_engine import LossCalculator, LossInput, LossItem

    items = [
        LossItem(
            "L001",
            "应收A公司货款",
            "应收款项",
            book_value=500000,
            recoverable=50000,
            loss_amount=450000,
            approval_status="已审批",
        ),
        LossItem(
            "L002",
            "存货毁损-原料B",
            "存货",
            book_value=300000,
            recoverable=80000,
            loss_amount=220000,
            approval_status="已审批",
        ),
        LossItem(
            "L003",
            "固定资产C报废",
            "固定资产",
            book_value=800000,
            recoverable=100000,
            loss_amount=700000,
            evidence_level="基本齐全",
            approval_status="已审批",
        ),
        LossItem(
            "L004",
            "现金被盗",
            "现金",
            book_value=50000,
            recoverable=0,
            loss_amount=50000,
            has_external_evidence=False,
            approval_status="已审批",
        ),
        LossItem(
            "L005",
            "无形资产D减值",
            "无形资产",
            book_value=200000,
            recoverable=0,
            loss_amount=200000,
            approval_status="已审批",
        ),
        LossItem(
            "L006",
            "存货盘亏-原料E",
            "存货",
            book_value=150000,
            recoverable=30000,
            loss_amount=120000,
            evidence_level="部分缺失",
            approval_status="未审批",
        ),
    ]

    from .models import EnterpriseInfo

    input_data = LossInput(
        enterprise=EnterpriseInfo(name="XX公司", uscc="91310115MAE1234567"),
        items=items,
    )

    result = LossCalculator().calculate(input_data)
    print(LossCalculator().report(result))
    return result


def run_hightech():
    """
    高新技术企业认定审核
    """
    from .hightech_engine import HightechVerifier, HightechInput

    from .models import EnterpriseInfo

    input_data = HightechInput(
        enterprise=EnterpriseInfo(name="XX智能科技", uscc="91310115MAE1234567"),
        income_last_year=180000000,
        income_3year_total=460000000,
        rd_expense_3year_total=28000000,
        rd_expense_last_year=12000000,
        total_staff=350,
        tech_staff=120,
        ip_count=8,
        ip_2_count=15,
        hi_product_revenue=140000000,
        hi_product_revenue_3year=360000000,
        total_revenue_3year=480000000,
    )

    result = HightechVerifier().verify(input_data)
    print(HightechVerifier().report(result))
    return result


def main():
    """统一入口"""
    args = sys.argv[1:]

    engine_map = {
        "eit": ("企业所得税汇算清缴", run_eit),
        "rd": ("研发费用加计扣除", run_rd),
        "fulltax": ("全税种核查", run_fulltax),
        "loss": ("财产损失审核", run_loss),
        "hightech": ("高新认定审核", run_hightech),
    }

    if not args or args[0] == "all":
        # 全部运行
        print("=" * 60)
        print("  税审底稿自动填充引擎 — 全流程演示")
        print("=" * 60)
        for key, (name, fn) in engine_map.items():
            print(f"\n{'=' * 60}")
            print(f"  【{name}】")
            print(f"{'=' * 60}")
            try:
                fn()
            except Exception as e:
                print(f"  执行失败: {e}")
    elif args[0] in engine_map:
        name, fn = engine_map[args[0]]
        print(f"  【{name}】")
        fn()
    else:
        print(f"可用引擎: {', '.join(engine_map.keys())} 或 all")
        print(f"使用: python -m tax_audit_engine.app <引擎名>")
        sys.exit(1)


if __name__ == "__main__":
    main()
