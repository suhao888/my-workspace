import sys

sys.stdout.reconfigure(encoding="utf-8")
sys.path.insert(0, "E:/Projects/my-workspace/tools/audit_doc_engine")

import pandas as pd
from fill_notes import extract_tz_data, MAPPINGS

path = "D:/Users/12844/Desktop/保定年审资料/2、决算套表/25年决算套表 (2)(1)/25年决算套表/保定吉达电力设计有限公司-单体.xlsx"
xl = pd.ExcelFile(path)

# Find递延所得税MAPPING
for m in MAPPINGS:
    if m["cat"] == "递延所得税-资产/负债":
        print(f"tz_cols = {m['tz_cols']}")
        print(f"col_map = {m['col_map']}")

        # Find sheet
        sheet_raw = None
        for s in xl.sheet_names:
            try:
                d = s.encode("latin1").decode("gbk")
            except:
                d = s
            if "递延所得税" in d and "原始" in d:
                sheet_raw = s
                break

        df = pd.read_excel(path, sheet_name=sheet_raw, header=None)
        data = extract_tz_data(df, m["tz_cols"])

        print(f"\n提取数据 ({len(data)}行):")
        for name, vals in data.items():
            print(f"  [{name[:25]}] = {vals}")
        break
