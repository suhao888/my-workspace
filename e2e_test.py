#!/usr/bin/env python3
"""端到端测试：EIT模板填充 + 其他模板基础填充"""

import sys, os

sys.stdin.reconfigure(encoding="utf-8")
sys.stdout.reconfigure(encoding="utf-8")
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
from tax_audit_engine.main import build_sample_data
from tax_audit_engine.calculator import TaxCalculator
from tax_audit_engine.template_filler import TemplateFiller

OUT = r"D:\Users\12844\Desktop"
BASE = r"D:\Users\12844\Desktop\业务工作底稿模版"


def find_template(cat_kw, file_kw):
    for root, dirs, files in os.walk(BASE):
        for f in files:
            if f.startswith("~$") or not f.endswith(".xlsx"):
                continue
            fp = os.path.join(root, f)
            ok = all(kw in root for kw in cat_kw) and all(kw in f for kw in file_kw)
            if ok:
                return fp
    return None


data = build_sample_data()
ent = data["enterprise"]
tb = data["tb"]
result = TaxCalculator().calculate(tb=tb, enterprise=ent, assets=data["assets"])
filler = TemplateFiller(result)

# 1. EIT
print("=== 1. EIT ===")
eit = {
    k: find_template(["独立纳税", "1-1"], v)
    for k, v in [
        ("fuzhu", ["辅助底稿"]),
        ("cover", ["底稿-0"]),
        ("a100", ["纳税申报表A类"]),
    ]
}
if all(eit.values()):
    filler.fill_fuzhu_digao(eit["fuzhu"], os.path.join(OUT, "EIT_辅助底稿_已填.xlsx"))
    filler.fill_cover(eit["cover"], os.path.join(OUT, "EIT_封面_已填.xlsx"))
    filler.fill_A100000(eit["a100"], os.path.join(OUT, "EIT_A100000_已填.xlsx"))
    filler.fill_A105000(eit["a100"], os.path.join(OUT, "EIT_A105000_已填.xlsx"))
    print("  4/4 OK")
else:
    print("  EIT模板未找到")

# 2. RD
print("=== 2. RD ===")
rd_tmpl = find_template(["研发费加计扣除"], ["加计扣除工作底稿"])
if rd_tmpl:
    from tax_audit_engine.rd_engine import RDCalculator, RDInput, RDProject

    proj = [
        RDProject(
            project_id="RD01",
            project_name="AI研发",
            personnel_costs=500000,
            direct_input=800000,
            depreciation=300000,
        )
    ]
    rd_input = RDInput(enterprise=ent, projects=proj, total_wages=3000000)
    rd_res = RDCalculator().calculate(rd_input)
    filler.fill_rd_workpaper(
        rd_tmpl, os.path.join(OUT, "RD_工作底稿_已填.xlsx"), rd_result=rd_res
    )
    print("  OK (可加计扣除: {:.0f})".format(rd_res.qualifying_amount))

# 3. FullTax
print("=== 3. FullTax ===")
ft_tmpl = find_template(["全税种"], ["核查测算底稿"])
if ft_tmpl:
    from tax_audit_engine.full_tax_engine import FullTaxVerifier, FullTaxInput

    ft_in = FullTaxInput(
        enterprise_name=ent.name,
        tax_id=ent.uscc,
        audit_year=str(ent.tax_year),
        items=[],
    )
    ft_res = FullTaxVerifier(threshold=0.05).verify(ft_in)
    filler.fill_fulltax(
        ft_tmpl, os.path.join(OUT, "FT_底稿_已填.xlsx"), fulltax_result=ft_res
    )
    print("  OK")

# 4. Loss
print("=== 4. Loss ===")
loss_tmpl = find_template(["财产损失"], ["损失明细表"])
if loss_tmpl:
    from tax_audit_engine.loss_engine import LossCalculator, LossInput, LossItem

    li = [
        LossItem(
            loss_id="L01",
            asset_name="设备A",
            category="固定资产",
            loss_type="报废",
            book_value=500000,
            recoverable=50000,
            loss_amount=450000,
            evidence_level="充分",
            has_external_evidence=True,
            approval_status="已审批",
        )
    ]
    loss_in = LossInput(enterprise=ent, items=li)
    loss_res = LossCalculator().calculate(loss_in)
    filler.fill_loss(
        loss_tmpl, os.path.join(OUT, "Loss_底稿_已填.xlsx"), loss_result=loss_res
    )
    print("  OK (损失额: {:.0f})".format(loss_res.total_loss))

# 5. HighTech
print("=== 5. HighTech ===")
ht_tmpl = find_template(["高新鉴证"], ["鉴证专项审计底稿"])
if ht_tmpl:
    from tax_audit_engine.hightech_engine import HightechVerifier, HightechInput

    ht_in = HightechInput(
        enterprise=ent,
        income_last_year=50000000,
        income_3year_total=150000000,
        rd_expense_3year_total=8000000,
        rd_expense_last_year=3800000,
        total_staff=ent.employee_count,
        tech_staff=80,
        ip_count=6,
        ip_2_count=3,
        hi_product_revenue=35000000,
        hi_product_revenue_3year=100000000,
        total_revenue_3year=150000000,
    )
    ht_res = HightechVerifier().verify(ht_in)
    filler.fill_hightech(
        ht_tmpl, os.path.join(OUT, "HT_底稿_已填.xlsx"), hightech_result=ht_res
    )
    print("  OK (评分: {:.1f})".format(ht_res.total_score))

# 6. 分公司
print("=== 6. 分公司 ===")
br = {
    k: find_template(["分公司", "汇总纳税"], v)
    for k, v in [
        ("fuzhu", ["辅助底稿"]),
        ("cover", ["底稿-0"]),
        ("a100", ["纳税申报表A类"]),
    ]
}
if br["a100"]:
    filler.fill_fuzhu_digao(br["fuzhu"], os.path.join(OUT, "BR_辅助底稿_已填.xlsx"))
    filler.fill_cover(br["cover"], os.path.join(OUT, "BR_封面_已填.xlsx"))
    filler.fill_A100000(br["a100"], os.path.join(OUT, "BR_A100000_已填.xlsx"))
    filler.fill_A105000(br["a100"], os.path.join(OUT, "BR_A105000_已填.xlsx"))
    print("  4/4 OK")
else:
    print("  模板未找到")

print("\n=== 输出文件 ===")
for f in sorted(os.listdir(OUT)):
    if "底稿" in f or "底稿" in f:
        sz = os.path.getsize(os.path.join(OUT, f))
        print("  {} ({}KB)".format(f, sz // 1024))
