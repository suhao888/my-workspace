"""
套表数据源清查 — 遍历所有sheet，理解每个表里有什么数据、属于哪个科目
"""

import sys

sys.stdout.reconfigure(encoding="utf-8")
import pandas as pd

path = "D:/Users/12844/Desktop/保定年审资料/2、决算套表/25年决算套表 (2)(1)/25年决算套表/保定吉达电力设计有限公司-单体.xlsx"
xl = pd.ExcelFile(path)

results = []
for raw_name in xl.sheet_names:
    try:
        decoded = raw_name.encode("latin1").decode("gbk")
    except:
        decoded = raw_name

    df = pd.read_excel(path, sheet_name=raw_name, header=None)

    # 找第一列有意义的文本作为标题
    title = ""
    for i in range(min(5, df.shape[0])):
        for j in range(min(3, df.shape[1])):
            v = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
            if len(v) > 3 and v != "nan":
                title = v[:60]
                break
        if title:
            break

    # 统计非空行数（去掉全空行）
    data_rows = 0
    has_number = False
    for i in range(df.shape[0]):
        row_has_data = False
        row_has_number = False
        for j in range(1, min(df.shape[1], 10)):
            v = df.iloc[i, j]
            if pd.notna(v):
                s = str(v).strip()
                if s and s != "nan":
                    row_has_data = True
                    try:
                        float(s.replace(",", ""))
                        row_has_number = True
                    except:
                        pass
        if row_has_data:
            data_rows += 1
            if row_has_number:
                has_number = True

    # 前几行列名
    col_headers = []
    for i in range(min(6, df.shape[0])):
        row_headers = []
        for j in range(min(8, df.shape[1])):
            v = str(df.iloc[i, j]).strip() if pd.notna(df.iloc[i, j]) else ""
            if v and v != "nan" and len(v) < 40:
                row_headers.append(v)
        if row_headers:
            col_headers.append("|".join(row_headers))

    results.append(
        {
            "sheet": decoded,
            "rows": df.shape[0],
            "cols": df.shape[1],
            "title": title,
            "data_rows": data_rows,
            "has_number": has_number,
            "headers": "; ".join(col_headers[:3]),
        }
    )

# 按是否有数据分类
print("=== 有数据的Sheet（含数值） ===")
print(f"{'Sheet名称':35s} {'行':4s} {'列':3s} {'数据行':6s} {'标题/首行列名'}")
print("-" * 120)
for r in results:
    if r["has_number"] and r["data_rows"] > 3:
        print(
            f"{r['sheet'][:35]:35s} {r['rows']:4d} {r['cols']:3d} {r['data_rows']:6d} {r['title'][:50]}"
        )
        if r["headers"]:
            print(f"{'':35s} {'':4s} {'':3s} {'':6s}  └─ {r['headers'][:70]}")

print()
print("=== 有结构但无数值/行数少的Sheet ===")
for r in results:
    if (r["data_rows"] <= 3 or not r["has_number"]) and r["data_rows"] > 0:
        print(
            f"{r['sheet'][:35]:35s} {r['rows']:4d} {r['cols']:3d} {r['data_rows']:6d} {r['title'][:50]}"
        )
