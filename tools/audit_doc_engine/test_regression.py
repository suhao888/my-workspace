"""
回归测试 — 检查分节感知匹配没有破坏非分节MAPPING
"""

import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, "E:/Projects/my-workspace/tools/audit_doc_engine")

import pandas as pd
from fill_notes import extract_tz_data, MAPPINGS

path = "D:/Users/12844/Desktop/保定年审资料/2、决算套表/25年决算套表 (2)(1)/25年决算套表/保定吉达电力设计有限公司-单体.xlsx"
xl = pd.ExcelFile(path)

# 测试几个非分节MAPPING
test_cases = [
    ("货币资金-明细", "货币资金_原始数据", [3, 4]),
    ("应收账款-账龄", "应收款项计提坏账准备情况表_原始数据", [3, 5, 8, 10]),
    ("营业收入/成本", "营业收入、营业成本_原始数据", [3, 4, 5, 6]),
]

for cat, sheet_kw, tz_cols in test_cases:
    sheet_raw = None
    for s in xl.sheet_names:
        try:
            d = s.encode("latin1").decode("gbk")
        except:
            d = s
        if sheet_kw in d:
            sheet_raw = s
            break
    if sheet_raw:
        df = pd.read_excel(path, sheet_name=sheet_raw, header=None)
        data = extract_tz_data(df, tz_cols)
        print(f"{cat}: {len(data) if data else 0}行")
        if data:
            for k, v in list(data.items())[:5]:
                print(f"  [{k}] = {v}")
    else:
        print(f"{cat}: Sheet未找到")
