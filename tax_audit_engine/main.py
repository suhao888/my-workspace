#!/usr/bin/env python3
"""
税审底稿自动填充引擎 — 入口
演示：用示例数据执行完整的企业所得税汇缴计算，生成工作底稿Excel
"""

import sys, json, io, os
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from .models import TrialBalance, EnterpriseInfo, AssetItem, AssetCategory
from .calculator import TaxCalculator
from .workpaper_generator import WorkpaperGenerator
from .template_filler import fill_all_templates


def build_sample_data() -> dict:
    """
    构建示例数据（典型制造业企业）
    """
    # === 企业信息 ===
    enterprise = EnterpriseInfo(
        name="XX制造有限公司",
        uscc="91440101MA5XXXXXX",
        industry="制造业",
        tax_year=2025,
        legal_rep="张三",
        employee_count=150,
        total_assets=35000000,
        is_high_tech=True,
    )

    # === 试算平衡表（关键科目） ===
    tb = TrialBalance(
        items={
            # 收入
            "主营业务收入": 50000000,
            "其他业务收入": 2000000,
            "营业外收入": 300000,
            "投资收益": 500000,
            "公允价值变动收益": 200000,
            "资产处置收益": 100000,
            # 成本
            "主营业务成本": 32000000,
            "其他业务成本": 1200000,
            "税金及附加": 800000,
            # 费用
            "销售费用": 3500000,
            "管理费用": 4200000,
            "财务费用": 600000,
            # 利润
            "利润总额": 7200000,
            # ===== 资产负债表科目（SH审定表填报用） =====
            "货币资金": 12500000,
            "应收账款": 8500000,
            "预付款项": 1200000,
            "其他应收款": 500000,
            "存货": 18500000,
            "在建工程": 3000000,
            "预收款项": 2400000,
            "应付账款": 6800000,
            "其他应付款": 900000,
            "应付股利": 500000,
            "短期借款": 5000000,
            "长期借款": 3000000,
            "长期应付款": 1000000,
            "实收资本": 10000000,
            "资本公积": 2000000,
            "盈余公积": 1800000,
            "未分配利润": 4200000,
            # ===== 明细科目（纳税调整用） =====
            "工资薪金": 8500000,
            "职工福利费": 1300000,  # 限额=850万*14%=119万 → 调增11万
            "职工教育经费": 750000,  # 限额=850万*8%=68万 → 调增7万
            "工会经费": 180000,  # 限额=850万*2%=17万 → 调增1万
            "业务招待费": 600000,  # 60%=36万, 营收5200万*5‰=26万, 取26万 → 调增34万
            "广告费和业务宣传费": 9000000,  # 限额=5200万*15%=780万 → 调增120万
            "公益性捐赠支出": 1200000,  # 利润720万*12%=86.4万 → 调增33.6万
            "利息支出": 400000,  # 假定全部可扣除
            "罚金、罚款和被没收财物": 50000,  # 全额调增
            "税收滞纳金、加收利息": 20000,  # 全额调增
            "赞助支出": 300000,  # 非广告性
            "研发费用": 3800000,  # 加计扣除100%
            "资产减值损失": 450000,  # 全额调增
            "基本社会保险": 1785000,
            "住房公积金": 680000,
            "补充养老保险": 300000,  # 限额=850万*5%=42.5万 → 不超
            "补充医疗保险": 200000,  # 限额=850万*5%=42.5万 → 不超
            # 投资类
            "居民企业间股息红利-免税收入": 300000,
            "权益法核算的投资收益": 200000,
            "实际收到的股息红利": 300000,
            "国债利息收入": 100000,
            # 其他
            "不征税收入": 0,
        }
    )

    # === 固定资产卡片 ===
    assets = [
        AssetItem(
            category=AssetCategory.BUILDING.value,
            name="办公楼",
            original_value=12000000,
            accounting_life_years=20,
            tax_life_years=20,
            current_accounting_depr=570000,
            current_tax_depr=570000,
        ),
        AssetItem(
            category=AssetCategory.MACHINERY.value,
            name="生产线A",
            original_value=8000000,
            accounting_life_years=10,
            tax_life_years=10,
            current_accounting_depr=760000,
            current_tax_depr=760000,
        ),
        AssetItem(
            category=AssetCategory.ELECTRONIC.value,
            name="服务器集群",
            original_value=1500000,
            accounting_life_years=5,
            tax_life_years=3,  # 税法3年，差异
            current_accounting_depr=285000,
            current_tax_depr=475000,  # 税收折旧更多 → 调减
        ),
        AssetItem(
            category=AssetCategory.TRANSPORT.value,
            name="运输车辆",
            original_value=800000,
            accounting_life_years=5,
            tax_life_years=4,
            current_accounting_depr=152000,
            current_tax_depr=190000,
        ),
        AssetItem(
            category=AssetCategory.PRODUCTION_TOOLS.value,
            name="模具及夹具",
            original_value=600000,
            accounting_life_years=5,
            tax_life_years=5,
            current_accounting_depr=114000,
            current_tax_depr=114000,
        ),
    ]

    return {
        "enterprise": enterprise,
        "tb": tb,
        "assets": assets,
    }


def run_demo(output_path: str = None):
    """
    执行完整演示流程：
    1. 构建示例数据
    2. 执行税审计算
    3. 打印摘要
    4. 生成底稿Excel
    """
    print("=" * 60)
    print("税审底稿自动填充引擎 — 演示")
    print("企业所得税汇算清缴 — 制造业企业")
    print("=" * 60)

    # 1. 构建数据
    print("\n>> 构建示例数据...")
    data = build_sample_data()

    # 2. 执行计算
    print(">> 执行纳税调整计算...")
    calculator = TaxCalculator()
    result = calculator.calculate(
        tb=data["tb"],
        enterprise=data["enterprise"],
        assets=data["assets"],
    )

    # 3. 打印摘要
    print("\n" + "─" * 60)
    print("计算摘要")
    print("─" * 60)
    summary = result.summary()
    for key, value in summary.items():
        if isinstance(value, float):
            print(f"  {key}: {value:>14,.2f}")
        else:
            print(f"  {key}: {value}")

    print("\n>> 调整明细（调增>0或调减>0的项目）:")
    print("─" * 60)
    print(f"  {'项目':<24} {'调增':>12} {'调减':>12}")
    print("─" * 60)
    for adj in result.adjustments:
        if adj.increase > 0 or adj.decrease > 0:
            inc_str = f"{adj.increase:>12,.2f}" if adj.increase > 0 else f"{'':>12}"
            dec_str = f"{adj.decrease:>12,.2f}" if adj.decrease > 0 else f"{'':>12}"
            print(f"  [{adj.category.value}] {adj.item_name:<18} {inc_str} {dec_str}")

    # 4. 生成底稿
    if not output_path:
        output_path = str(Path("D:/Users/12844/Desktop/税审底稿_示例.xlsx"))
    print(f"\n>> 生成工作底稿 → {output_path}")

    gen = WorkpaperGenerator(result)
    gen.set_assets(data["assets"])
    gen.generate(output_path)

    print(f"\n{'=' * 60}")
    print(f"完成！底稿已保存至: {output_path}")
    print(f"共 {len(result.adjustments)} 条纳税调整项")
    print(f"生成 {len(gen.wb.sheetnames)} 张工作表: {', '.join(gen.wb.sheetnames)}")
    print(f"{'=' * 60}")


def run_with_custom_data(
    tb_data: dict,
    enterprise_info: dict = None,
    asset_list: list = None,
    output_path: str = None,
):
    """
    用自定义数据运行税审计算

    Parameters
    ----------
    tb_data : dict
        试算平衡表，{科目名称: 金额}
    enterprise_info : dict, optional
        企业信息
    asset_list : list[dict], optional
        固定资产列表
    output_path : str, optional
        输出路径，默认桌面
    """
    tb = TrialBalance(items=tb_data)
    enterprise = None
    if enterprise_info:
        enterprise = EnterpriseInfo(**enterprise_info)

    assets = []
    if asset_list:
        for a in asset_list:
            assets.append(AssetItem(**a))

    calculator = TaxCalculator()
    result = calculator.calculate(
        tb=tb,
        enterprise=enterprise,
        assets=assets,
    )

    if not output_path:
        output_path = "D:/Users/12844/Desktop/税审底稿_自定义.xlsx"

    gen = WorkpaperGenerator(result)
    if assets:
        gen.set_assets(assets)
    gen.generate(output_path)

    return result


# ============================================================
# 模板填充演示（中税网底稿模板）
# ============================================================

TEMPLATE_BASE = (
    "D:/Users/12844/Desktop/业务工作底稿模版/"
    "2026_07_04_1-1、中税网-2026年企业所得税纳税申报审核报告及底稿模板-适用于独立纳税企业V1/"
    "1-1、中税网-2026年企业所得税纳税申报审核报告及底稿模板-适用于独立纳税企业V1/"
    "1、中税网企业所得税汇缴鉴证报告、申报表及工作底稿模板-适用独立纳税企业-必做底稿2026"
)


def run_demo_with_templates(output_dir: str = None):
    """
    演示流程：计算 → 填充中税网模板
    """
    print("=" * 60)
    print("税审底稿填充引擎 — 模板填充模式")
    print("=" * 60)

    # 1. 构建数据
    print("\n>> 构建示例数据...")
    data = build_sample_data()

    # 2. 执行计算
    print(">> 执行纳税调整计算...")
    calculator = TaxCalculator()
    result = calculator.calculate(
        tb=data["tb"],
        enterprise=data["enterprise"],
        assets=data["assets"],
    )

    # 3. 打印摘要
    print("\n" + "─" * 60)
    print("计算摘要")
    print("─" * 60)
    summary = result.summary()
    for key, value in summary.items():
        if isinstance(value, float):
            print(f"  {key}: {value:>14,.2f}")
        else:
            print(f"  {key}: {value}")

    # 4. 填充模板
    base = output_dir or "D:/Users/12844/Desktop"
    print(f"\n>> 检查模板目录: {TEMPLATE_BASE}")
    if not os.path.exists(TEMPLATE_BASE):
        print(f"  ❌ 模板目录不存在: {TEMPLATE_BASE}")
        print("  >> 回退到自生成模式...")
        # fallback: 用 WorkpaperGenerator
        gen = WorkpaperGenerator(result)
        gen.set_assets(data["assets"])
        gen.generate(os.path.join(base, "税审底稿_示例.xlsx"))
        print(f"  ✅ 自生成底稿 → {os.path.join(base, '税审底稿_示例.xlsx')}")
        return

    print("\n>> 填充模板...")
    results = fill_all_templates(
        result=result,
        template_dir=TEMPLATE_BASE,
        output_dir=base,
        assets=data["assets"],
    )

    print(f"\n{'=' * 60}")
    print(f"生成 {len(results)} 份底稿文件:")
    for name, path in results.items():
        print(f"  ✅ {name}: {path}")
    print(f"共 {len(result.adjustments)} 条纳税调整项")
    print(f"{'=' * 60}")


# ============================================================
# 带命令行参数的主入口
# ============================================================


def main():
    """带命令行参数的主入口"""
    import argparse

    parser = argparse.ArgumentParser(description="税审底稿填充引擎")
    parser.add_argument(
        "--mode",
        choices=["demo", "template"],
        default="template",
        help="运行模式: demo=WorkpaperGenerator自生成, template=填充中税网模板(默认)",
    )
    parser.add_argument(
        "--output",
        "-o",
        default=None,
        help="输出目录（默认桌面）",
    )
    args = parser.parse_args()

    if args.mode == "demo":
        out_path = args.output or "D:/Users/12844/Desktop/税审底稿_示例.xlsx"
        run_demo(out_path)
    else:
        run_demo_with_templates(args.output)


if __name__ == "__main__":
    main()
