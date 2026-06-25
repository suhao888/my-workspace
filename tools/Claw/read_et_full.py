# -*- coding: utf-8 -*-
import xlrd
import json

filepath = r"D:\Users\12844\Desktop\10-梳理-----0528\基础数据\5、浙江油田-附件1：合同结算流程调查表（浙江油田分公司）.et"

wb = xlrd.open_workbook(filepath)
ws = wb.sheet_by_index(0)

print(f"Sheet名称: {ws.name}")
print(f"行数: {ws.nrows}, 列数: {ws.ncols}")
print()

# 读取合并单元格信息
print("合并单元格信息:")
for merge in ws.merged_cells:
    print(f"  行{merge[0]}-{merge[1]}, 列{merge[2]}-{merge[3]}")

print()
print("===== 全部数据 =====")
all_rows = []
for r in range(ws.nrows):
    row_data = []
    for c in range(ws.ncols):
        val = ws.cell_value(r, c)
        # 截断过长的值
        if isinstance(val, str) and len(val) > 500:
            val = val[:500] + f"...[截断，原长{len(val)}]"
        row_data.append(val)
    all_rows.append(row_data)
    # 只打印前30列的摘要
    short_row = []
    for c, v in enumerate(row_data):
        if isinstance(v, str) and len(v) > 100:
            short_row.append(f"[列{c}]{v[:100]}...")
        elif v != '':
            short_row.append(f"[列{c}]{v}")
    print(f"行{r}: {short_row}")

# 保存完整数据到JSON
output = {
    "sheet_name": ws.name,
    "rows": ws.nrows,
    "cols": ws.ncols,
    "data": all_rows
}
outpath = r"c:\Users\12844\WorkBuddy\Claw\zhejiang_youtian_data.json"
with open(outpath, 'w', encoding='utf-8') as f:
    json.dump(output, f, ensure_ascii=False, indent=2)
print(f"\n完整数据已保存到: {outpath}")
