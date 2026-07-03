import sys, pickle, pandas as pd
from pathlib import Path
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

OUT = Path('D:/Users/12844/Desktop/10-梳理-----0528/整理基础数据')
RP = 'D:/Users/12844/Desktop/优化方案问题汇总（529.xlsx'

with open('D:/Users/12844/Desktop/10-梳理-----0528/_all_rows_v2.pkl', 'rb') as f:
    data = pickle.load(f)
rows = data['rows']

# ===== 1. Check all output xlsx files =====
print('=' * 80)
print('扫描1: 输出xlsx文件名 vs 报告Sheet名 对照')
print('=' * 80)

actual_files = {f.stem: f.name for f in sorted(OUT.glob('*.xlsx'))}
report_xl = pd.ExcelFile(RP)
report_sheets = [s for s in report_xl.sheet_names if s != '总体分析']

print(f'输出文件数: {len(actual_files)}')
print(f'报告Sheet数(不含总体分析): {len(report_sheets)}')
print()

for rs in report_sheets:
    # Extract number and name
    parts = rs.split('-', 1)
    num = parts[0]
    name = parts[1] if len(parts) > 1 else rs
    print(f'  报告: [{rs}]')
    # Find matching output file
    matches = [k for k in actual_files if name in k or k in name]
    if not matches:
        # Fuzzy match
        for k in actual_files:
            common = set(name) & set(k)
            if len(common) > 3:
                matches.append(k)
    if matches:
        print(f'    输出文件: {matches}')
    else:
        print(f'    ⚠️ 无匹配输出文件!')
    print()

# ===== 2. Check data completeness per unit per category =====
print('=' * 80)
print('扫描2: 每单位×每合同类别的数据行数矩阵')
print('=' * 80)

all_units = sorted(set(r.get(1, '') for r in rows))
all_cat1 = sorted(set(r.get(2, '') for r in rows))

# Build matrix
matrix = {}
for r in rows:
    unit = r.get(1, '')
    cat1 = r.get(2, '')
    key = (unit, cat1)
    if key not in matrix:
        matrix[key] = {'total': 0, 'with_data': 0}
    matrix[key]['total'] += 1
    if any(str(r.get(c, '')).strip() for c in range(5, 31)):
        matrix[key]['with_data'] += 1

# Print header
header = f'{"单位":<10}'
for c in all_cat1:
    header += f'{c[:6]:>8}'
print(header)
print('-' * (10 + 8 * len(all_cat1)))

for unit in all_units:
    line = f'{unit:<10}'
    for cat1 in all_cat1:
        info = matrix.get((unit, cat1), {'total': 0, 'with_data': 0})
        if info['total'] == 0:
            line += f'{"-":>8}'
        elif info['with_data'] == 0:
            line += f'{str(info["total"])+"×":>8}'
        else:
            line += f'{str(info["total"]):>8}'
    print(line)

print('  (数字=有流程数据行数, ×=有分类但无流程, -=无此类)')

# ===== 3. Check report claims against actual data =====
print()
print('=' * 80)
print('扫描3: 报告中数据声称 vs 实际数据 对比')
print('=' * 80)

# Check Sheet 4-其它合同 claims
# Report says: "4个子类...在所有13家单位中无任何数据"
# Actual: 土地使用补偿合同 8行(玉门), 土地复垦方案设计 1行(新疆),
#         土地复垦施工 1行(新疆), 集团内部单位衍生品代理协议 8行(5有数据)
for cat2_name, cat1_name in [
    ('土地使用补偿合同', '其它合同'),
    ('土地复垦方案设计', '建设工程合同'),  # post-fix
    ('土地复垦施工', '建设工程合同'),       # post-fix
    ('集团内部单位衍生品代理协议', '买卖合同'), # post-fix
]:
    found = [r for r in rows if r.get(2, '') == cat1_name and r.get(3, '') == cat2_name]
    with_data = sum(1 for r in found if any(str(r.get(c, '')).strip() for c in range(5, 31)))
    units = set(r.get(1, '') for r in found)
    print(f'  [{cat1_name}]>{cat2_name}: {len(found)}行, {with_data}有流程, {len(units)}单位 {sorted(units)}')

# ===== 4. Check for orphan/unexpected rows =====
print()
print('=' * 80)
print('扫描4: 异常值检测')
print('=' * 80)

# Check for rows where 二级 or 三级 looks like a note/comment
for r in rows:
    cat2 = str(r.get(3, ''))
    cat3 = str(r.get(4, ''))
    issues = []
    for field, name in [(cat2, '二级'), (cat3, '三级')]:
        if len(field) > 30:
            issues.append(f'{name}名称过长({len(field)}字)')
        if '取消' in field:
            issues.append(f'{name}含"取消"')
        if '备注' in field or '说明' in field or '注：' in field:
            issues.append(f'{name}疑似备注文字')
    if issues:
        unit = r.get(1, '')
        cat1 = r.get(2, '')
        joined = '; '.join(issues)
        print(f'  [{unit}] {cat1}>{cat2}>{cat3}: {joined}')

# ===== 5. Check column fill rate by category =====
print()
print('=' * 80)
print('扫描5: 各一级分类的平均列填充率（31列中非空列数）')
print('=' * 80)

for cat1 in sorted(all_cat1):
    cat_rows = [r for r in rows if r.get(2, '') == cat1]
    if not cat_rows: continue
    fills = []
    for r in cat_rows:
        filled = sum(1 for c in range(31) if str(r.get(c, '')).strip())
        fills.append(filled)
    avg = sum(fills) / len(fills)
    min_f = min(fills)
    max_f = max(fills)
    print(f'  {cat1}: {len(cat_rows)}行, 平均填充{avg:.1f}列, 范围[{min_f}-{max_f}]')

# ===== 6. Check for duplicate rows =====
print()
print('=' * 80)
print('扫描6: 重复数据行检查')
print('=' * 80)

# Check for rows with identical (unit, cat1, cat2, cat3) but different process data
from collections import Counter
key_counts = Counter()
for r in rows:
    key = (r.get(1, ''), r.get(2, ''), r.get(3, ''), r.get(4, ''))
    key_counts[key] += 1

dupes = {k: v for k, v in key_counts.items() if v > 1}
if dupes:
    print(f'  发现 {len(dupes)} 组重复键 (同单位+同类+同子类+同明细):')
    for (unit, cat1, cat2, cat3), count in sorted(dupes.items()):
        print(f'    [{unit}] {cat1}>{cat2}>{cat3}: {count}行')
else:
    print('  未发现重复')

print()
print('扫描完成!')
