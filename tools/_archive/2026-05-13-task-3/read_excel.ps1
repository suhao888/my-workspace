import openpyxl
from openpyxl import load_workbook

file_path = 'D:/Users/12844/Desktop/保定吉达电力设计有限公司.xlsx'
wb = load_workbook(file_path, data_only=True)

print(f"工作簿中的所有Sheet: {wb.sheetnames}")
print()

sheets_to_read = [
    '资产负债表_原始数据',
    '资产负债表续_原始数据',
    '利润表_原始数据',
    '现金流量表_原始数据',
    '所有者权益变动表_原始数据',
    '应交税费表_原始数据',
    '企业财务决算报表封面（境内）_原始数据',
    '附注报表封面_原始数据'
]

for sheet_name in sheets_to_read:
    if sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        print('=' * 80)
        print(f'Sheet名称: {sheet_name}')
        print(f'行数: {ws.max_row}, 列数: {ws.max_column}')
        print('=' * 80)
        for row_idx in range(1, ws.max_row + 1):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                cell_value = ws.cell(row=row_idx, column=col_idx).value
                row_data.append(cell_value if cell_value is not None else '')
            print(f'行{row_idx}: {row_data}')
        print()
    else:
        print(f'\nSheet "{sheet_name}" 未找到\n')
