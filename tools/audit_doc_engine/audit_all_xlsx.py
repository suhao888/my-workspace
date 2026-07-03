import sys, pandas as pd
from pathlib import Path

sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('D:/Users/12844/Desktop/10-梳理-----0528/整理基础数据')
files = sorted(OUT.glob('*.xlsx'))

print('=' * 80)
print('全面检查：所有13个xlsx输出文件')
print('=' * 80)

total_sheets = 0
total_data_rows = 0
issues = []

for fp in files:
    xl = pd.ExcelFile(fp)
    fname = fp.name
    print(f'\n### {fname} ({len(xl.sheet_names)} sheets) ###')

    for s in xl.sheet_names:
        df = pd.read_excel(fp, sheet_name=s, header=None)
        data_rows = df.iloc[3:] if df.shape[0] > 3 else pd.DataFrame()
        total_sheets += 1

        valid_rows = 0
        empty_rows = 0
        units = set()
        has_process = 0
        no_process = 0

        for i in range(len(data_rows)):
            row = data_rows.iloc[i]
            cat_vals = []
            for j in range(5):
                v = row.iloc[j] if j < len(row) else ''
                if pd.notna(v) and str(v).strip() and str(v).strip() != 'nan':
                    cat_vals.append(str(v).strip())

            if cat_vals:
                valid_rows += 1
                unit = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
                if unit and unit != 'nan':
                    units.add(unit)

                has_any = False
                for j in range(5, min(31, len(row))):
                    v = row.iloc[j]
                    if pd.notna(v) and str(v).strip() and str(v).strip() != 'nan':
                        has_any = True
                        break
                if has_any:
                    has_process += 1
                else:
                    no_process += 1
            else:
                empty_rows += 1

        total_data_rows += valid_rows

        status_parts = [f'{valid_rows}数据行']
        if empty_rows > 0:
            status_parts.append(f'{empty_rows}空行')
        if has_process > 0:
            status_parts.append(f'{has_process}有流程')
        if no_process > 0:
            status_parts.append(f'{no_process}无流程')
        if units:
            status_parts.append(f'{len(units)}单位')

        status = ' | '.join(status_parts)

        flags = []
        if valid_rows == 0:
            flags.append('零数据行')
        if no_process == valid_rows and valid_rows > 0:
            flags.append('全部无流程数据')
        if no_process > 0 and has_process > 0:
            flags.append(f'{no_process}行无流程数据')

        flag_str = ' [' + ', '.join(flags) + ']' if flags else ''
        print(f'  [{s}]: {status}{flag_str}')

        if flags:
            issues.append(f'{fname} / {s}: {status}{flag_str}')

print()
print('=' * 80)
print(f'总计: {len(files)}个文件, {total_sheets}个Sheet, {total_data_rows}条数据行')
print(f'发现问题Sheet: {len(issues)}个')
for issue in issues:
    print(f'  ! {issue}')
