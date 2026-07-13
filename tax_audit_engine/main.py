#!/usr/bin/env python3
"""
税审底稿自动填充引擎 v3 — 配置驱动 + 业务注册

业务类型识别流程：
  1. 用户显式指定 --business-type
  2. 自动检测（根据模板目录名/TB科目/调整项特征）
  3. 模糊时列出候选

填充流程：
  BusinessRegistry.scan() → 加载所有业务配置
  → detect_business_type() 或 用户指定 → 选中业务
  → TaxAuditEngine.fill_all() → 按 manifest 依次填充所有模板
"""

import sys, io, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

from pathlib import Path

from .models import TrialBalance, EnterpriseInfo, AssetItem, AssetCategory
from .calculator import TaxCalculator
from .core.engine import TaxAuditEngine
from .business.registry import BusinessRegistry
from .business.detector import detect_business_type, list_matching_types


# ============================================================
# 示例数据
# ============================================================

TEMPLATE_DIR = (
    "D:/Users/12844/Desktop/业务工作底稿模版/"
    "2026_07_04_1-1、中税网-2026年企业所得税纳税申报审核报告及底稿模板-适用于独立纳税企业V1/"
    "1-1、中税网-2026年企业所得税纳税申报审核报告及底稿模板-适用于独立纳税企业V1/"
    "1、中税网企业所得税汇缴鉴证报告、申报表及工作底稿模板-适用独立纳税企业-必做底稿2026"
)


def build_sample_data():
    """构建示例数据（典型制造业企业）"""
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

    tb = TrialBalance(
        items={
            "主营业务收入": 50000000,
            "其他业务收入": 2000000,
            "营业外收入": 300000,
            "投资收益": 500000,
            "公允价值变动收益": 200000,
            "资产处置收益": 100000,
            "主营业务成本": 32000000,
            "其他业务成本": 1200000,
            "税金及附加": 800000,
            "销售费用": 3500000,
            "管理费用": 4200000,
            "财务费用": 600000,
            "利润总额": 7200000,
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
            "工资薪金": 8500000,
            "职工福利费": 1300000,
            "职工教育经费": 750000,
            "工会经费": 180000,
            "业务招待费": 600000,
            "广告费和业务宣传费": 9000000,
            "公益性捐赠支出": 1200000,
            "利息支出": 400000,
            "罚金、罚款和被没收财物": 50000,
            "税收滞纳金、加收利息": 20000,
            "赞助支出": 300000,
            "研发费用": 3800000,
            "资产减值损失": 450000,
            "基本社会保险": 1785000,
            "住房公积金": 680000,
            "补充养老保险": 300000,
            "补充医疗保险": 200000,
            "居民企业间股息红利-免税收入": 300000,
            "权益法核算的投资收益": 200000,
            "实际收到的股息红利": 300000,
            "国债利息收入": 100000,
        }
    )

    assets = [
        AssetItem(
            category="房屋、建筑物",
            name="办公楼",
            original_value=12000000,
            accounting_life_years=20,
            tax_life_years=20,
            current_accounting_depr=570000,
            current_tax_depr=570000,
        ),
        AssetItem(
            category="机器设备",
            name="生产线A",
            original_value=8000000,
            accounting_life_years=10,
            tax_life_years=10,
            current_accounting_depr=760000,
            current_tax_depr=760000,
        ),
        AssetItem(
            category="电子设备",
            name="服务器集群",
            original_value=1500000,
            accounting_life_years=5,
            tax_life_years=3,
            current_accounting_depr=285000,
            current_tax_depr=475000,
        ),
        AssetItem(
            category="运输工具",
            name="运输车辆",
            original_value=800000,
            accounting_life_years=5,
            tax_life_years=4,
            current_accounting_depr=152000,
            current_tax_depr=190000,
        ),
        AssetItem(
            category="生产器具、工具",
            name="模具及夹具",
            original_value=600000,
            accounting_life_years=5,
            tax_life_years=5,
            current_accounting_depr=114000,
            current_tax_depr=114000,
        ),
    ]

    return {"enterprise": enterprise, "tb": tb, "assets": assets}


# ============================================================
# 主流程
# ============================================================


def run(output_dir: str = None, business_type: str = None, template_dir: str = None):
    """
    完整执行流程：
    1. 构建/加载数据
    2. 计算纳税调整
    3. 检测业务类型
    4. 填充模板
    5. 校验
    """
    print("=" * 60)
    print("税审底稿填充引擎 v3 — 配置驱动 + 业务注册")
    print("=" * 60)

    # 1. 数据
    print("\n>> 构建数据...")
    data = build_sample_data()
    tb = data["tb"]
    enterprise = data["enterprise"]
    assets = data["assets"]

    # 2. 计算
    print(">> 执行纳税调整计算...")
    calculator = TaxCalculator()
    result = calculator.calculate(tb=tb, enterprise=enterprise, assets=assets)

    print(f"  会计利润: {tb.accounting_profit:>12,.2f}")
    print(f"  纳税调增: {result.total_increase:>12,.2f}")
    print(f"  纳税调减: {result.total_decrease:>12,.2f}")
    print(f"  应纳税所得额: {result.taxable_income:>12,.2f}")

    # 3. 检测业务类型
    tmpl_dir = template_dir or TEMPLATE_DIR
    biz_id = business_type

    if not biz_id:
        print(f"\n>> 检测业务类型...")
        # 用 registry 检测
        registry = BusinessRegistry()
        registry.scan()

        # 快速检测
        biz_id = detect_business_type(
            template_dir=tmpl_dir,
            tb=tb,
            enterprise=enterprise,
            adjustments=result.adjustments,
        )

        if biz_id:
            print(f"  自动识别: {biz_id}")
        else:
            # 模糊时列出候选
            candidates = list_matching_types(
                template_dir=tmpl_dir,
                tb=tb,
                enterprise=enterprise,
                adjustments=result.adjustments,
            )
            if candidates:
                print("  无法唯一确定业务类型，候选:")
                for bid, name, score in candidates:
                    print(f"    • {name} ({bid}) 匹配度={score}")
                # 取最高分
                biz_id = candidates[0][0]
                print(f"  自动取最高分: {biz_id}")
            else:
                print("  ⚠ 未匹配到业务类型，使用默认: corporate_income_tax")
                biz_id = "corporate_income_tax"

    biz = registry.get(biz_id) if registry._businesses else None
    if not biz:
        # 直接从注册表获取
        registry.scan()
        biz = registry.get(biz_id)

    if not biz:
        print(f"  ❌ 业务类型 '{biz_id}' 未注册")
        return

    print(f"  业务: {biz.name}")

    # 4. 预加载模板配置
    print(f"\n>> 加载模板配置 ({len(biz.templates)} 个)...")
    import yaml

    configs_root = (
        Path(registry._configs_root or Path(__file__).parent / "configs") / biz_id
    )
    manifest_processed = {
        "templates": [],
        "checks": biz.raw_manifest.get("checks", []),
    }
    for tmpl in biz.templates:
        tmpl_entry = dict(tmpl)
        config_path = tmpl.get("config", "")
        if config_path:
            full_path = configs_root / config_path
            if full_path.exists():
                with open(full_path, "r", encoding="utf-8") as f:
                    cfg = yaml.safe_load(f) or {}
                # 将 manifest 条目属性（convert/output）合并到配置
                cfg["convert"] = tmpl.get("convert", False)
                cfg["output"] = tmpl.get("output", "")
                tmpl_entry["config"] = cfg
                print(f"  ✅ {tmpl['id']}: {config_path}")
            else:
                print(f"  ⚠ {tmpl['id']}: 配置 {config_path} 未找到")
                tmpl_entry["config"] = {}
        else:
            tmpl_entry["config"] = {}
        manifest_processed["templates"].append(tmpl_entry)

    # 5. 填充模板
    print(f"\n>> 填充模板...")
    engine = TaxAuditEngine(result, assets=assets)

    if not os.path.exists(tmpl_dir):
        print(f"  ❌ 模板目录不存在: {tmpl_dir}")
        return

    results = engine.fill_all(
        manifest=manifest_processed,
        template_dir=tmpl_dir,
        output_dir=output_dir or "D:/Users/12844/Desktop",
    )

    # 5. 结果
    print(f"\n{'=' * 60}")
    print(f"完成！生成 {len(results)} 份底稿文件:")
    for tmpl_id, info in results.items():
        issues = info.get("issues", [])
        status = "✅" if not issues else "⚠"
        print(f"  {status} {info['path']}")
        if issues:
            for iss in issues:
                print(f"    {iss}")
    print(f"共 {len(result.adjustments)} 条纳税调整项")
    print(f"{'=' * 60}")


def main():
    """命令行入口"""
    import argparse

    parser = argparse.ArgumentParser(description="税审底稿填充引擎 v3")
    parser.add_argument("--output", "-o", default=None, help="输出目录")
    parser.add_argument(
        "--business-type",
        "-b",
        default=None,
        help="业务类型（corporate_income_tax / high_tech / rd_expense / loss）",
    )
    parser.add_argument("--template-dir", "-t", default=None, help="模板目录路径")
    parser.add_argument(
        "--list-businesses", action="store_true", help="列出已注册业务类型"
    )
    args = parser.parse_args()

    if args.list_businesses:
        registry = BusinessRegistry()
        registry.scan()
        print("已注册业务类型:")
        for b in registry.list_all():
            print(f"  • {b.name} ({b.id})")
            print(f"    {b.description}")
        return

    run(
        output_dir=args.output,
        business_type=args.business_type,
        template_dir=args.template_dir,
    )


if __name__ == "__main__":
    main()
