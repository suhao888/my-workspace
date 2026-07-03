import sys, pickle, pandas as pd
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

with open('D:/Users/12844/Desktop/10-梳理-----0528/_all_rows_v2.pkl', 'rb') as f:
    data = pickle.load(f)
rows = data['rows']

OUT = 'D:/Users/12844/Desktop/10-梳理-----0528/整理基础数据'

# Build actual data: {filename: {sheet_name: {unit: count}}}
actual = {}
for cat1_name in sorted(set(r.get(2, '') for r in rows)):
    safe = cat1_name.replace('/', '_').replace(chr(92), '_')
    fname = safe + '.xlsx'
    if fname not in actual:
        actual[fname] = defaultdict(lambda: defaultdict(int))

    for r in rows:
        if r.get(2, '') != cat1_name: continue
        cat2 = r.get(3, '').strip()
        unit = r.get(1, '').strip()
        if not cat2: cat2 = '其它'
        actual[fname][cat2][unit] += 1

# For each output file, read the actual xlsx
for fname in sorted(actual.keys()):
    fp = OUT + '/' + fname
    try:
        xl = pd.ExcelFile(fp)
        actual_sheets = set(xl.sheet_names)
        expected_sheets = set(actual[fname].keys())

        print(f'\n=== {fname} ===')
        print(f'  实际Sheet: {sorted(actual_sheets)}')

        for s in sorted(actual_sheets):
            df = pd.read_excel(fp, sheet_name=s, header=None)
            data_count = max(0, df.shape[0] - 3)
            unit_counts = actual[fname].get(s, {})
            units_with_data = sum(1 for c in unit_counts.values() if c > 0)

            # Count rows with process data (cols 5+)
            process_count = 0
            for r in rows:
                if r.get(2, '') not in cat1_for_file(fname): continue
                if r.get(3, '').strip() != s: continue
                if s == '其它' and not r.get(3, '').strip(): continue
                if any(str(r.get(c, '')).strip() for c in range(5, 31)):
                    process_count += 1

            no_process = data_count - process_count

            status = []
            if no_process > 0:
                status.append(f'{no_process}行无流程数据')
            if units_with_data < 13:
                status.append(f'{units_with_data}/13单位')

            flag = ' ⚠️' if status else ''
            print(f'    [{s}]: {data_count}行 | {\", \".join(status)}{flag}' if status else f'    [{s}]: {data_count}行')

    except Exception as e:
        print(f'  读取失败: {e}')

def cat1_for_file(fname):
    for cat1 in set(r.get(2, '') for r in rows):
        safe = cat1.replace('/', '_').replace(chr(92), '_')
        if safe + '.xlsx' == fname:
            return cat1
    return ''
