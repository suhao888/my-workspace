import sys, pickle, pandas as pd
from pathlib import Path
from collections import defaultdict
sys.stdout.reconfigure(encoding='utf-8')

# Load original pickle data (as reference for what was extracted)
with open('D:/Users/12844/Desktop/10-梳理-----0528/_all_rows_v2.pkl', 'rb') as f:
    pkl = pickle.load(f)
pkl_rows = pkl['rows']

# Check each output file in 修改表头
OUT = Path('D:/Users/12844/Desktop/整理基础数据修改表头')

# Build a set of (unit, cat1, cat2, cat3) from output files
output_keys = set()
output_row_count = 0
empty_process_rows = []

for fp in sorted(OUT.glob('*.xlsx')):
    xl = pd.ExcelFile(fp)
    for s in xl.sheet_names:
        df = pd.read_excel(fp, sheet_name=s, header=None)
        for i in range(3, df.shape[0]):
            row = df.iloc[i]
            unit = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ''
            cat1 = str(row.iloc[2]).strip() if pd.notna(row.iloc[2]) else ''
            cat2 = str(row.iloc[3]).strip() if pd.notna(row.iloc[3]) else ''
            cat3 = str(row.iloc[4]).strip() if pd.notna(row.iloc[4]) else ''
            if not unit or unit == 'nan': continue

            key = (unit, cat1, cat2, cat3)
            output_keys.add(key)
            output_row_count += 1

            # Check if row has any process data (cols 5+)
            has_proc = False
            for j in range(5, min(31, df.shape[1])):
                v = row.iloc[j]
                if pd.notna(v) and str(v).strip() and str(v).strip() != 'nan':
                    has_proc = True
                    break
            if not has_proc:
                empty_process_rows.append(f'{fp.name}/{s} Row{i}: {unit} | {cat1}>{cat2}>{cat3}')

# Build same from pickle
pkl_keys = set()
pkl_with_data = set()
for r in pkl_rows:
    unit = r.get(1, '').strip()
    cat1 = r.get(2, '').strip()
    cat2 = r.get(3, '').strip()
    cat3 = r.get(4, '').strip()
    if not unit: continue
    key = (unit, cat1, cat2, cat3)
    pkl_keys.add(key)
    if any(str(r.get(c, '')).strip() for c in range(5, 31)):
        pkl_with_data.add(key)

print(f'输出文件: {output_row_count} 行, {len(output_keys)} 唯一键')
print(f'Pickle数据: {len(pkl_rows)} 行, {len(pkl_keys)} 唯一键')
print(f'输出缺但pickle有: {len(pkl_keys - output_keys)}')
print(f'Pickle缺但输出有: {len(output_keys - pkl_keys)}')

# Show missing entries
missing = pkl_keys - output_keys
if missing:
    print('\n=== 输出文件中缺失的数据条目 ===')
    by_unit_cat = defaultdict(list)
    for unit, cat1, cat2, cat3 in sorted(missing):
        by_unit_cat[(unit, cat1)].append((cat2, cat3))
    for (unit, cat1), items in sorted(by_unit_cat.items()):
        print(f'  [{unit}] {cat1}:')
        for cat2, cat3 in items:
            print(f'    > {cat2} > {cat3}')

# Show rows with no process data
print(f'\n=== 无流程数据行: {len(empty_process_rows)} 行 ===')
by_unit = defaultdict(int)
for row in empty_process_rows:
    unit = row.split(': ')[1].split(' | ')[0] if ': ' in row else ''
    by_unit[unit] += 1
for unit, count in sorted(by_unit.items(), key=lambda x: -x[1]):
    print(f'  {unit}: {count}行')

# Check files that exist in pickle but not in 修改表头
print('\n=== 修改表头缺少的文件 ===')
pkl_cat1s = set(r.get(2, '').strip() for r in pkl_rows)
out_files = {fp.stem for fp in OUT.glob('*.xlsx')}
for cat1 in sorted(pkl_cat1s):
    safe = cat1.replace('/', '_').replace(chr(92), '_')
    if safe not in out_files:
        print(f'  缺: {safe}.xlsx ({cat1})')
