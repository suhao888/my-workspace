import sys, pickle
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

with open('D:/Users/12844/Desktop/10-梳理-----0528/_all_rows_v2.pkl', 'rb') as f:
    data = pickle.load(f)

rows = data['rows']

# ===== 1. Find all rows with no process data =====
print('=' * 80)
print('问题1: 无流程数据的行（只有分类名，无流程内容）')
print('=' * 80)

no_process_rows = []
for r in rows:
    has_any = False
    for c in range(5, 31):
        v = r.get(c, '')
        if v and str(v).strip():
            has_any = True
            break
    if not has_any:
        no_process_rows.append(r)

print(f'\n总计 {len(no_process_rows)} 行无流程数据\n')

# Group by unit + cat1
by_unit_cat1 = defaultdict(list)
for r in no_process_rows:
    unit = r.get(1, '')
    cat1 = r.get(2, '')
    cat2 = r.get(3, '')
    by_unit_cat1[(unit, cat1)].append(r)

print('按单位+一级类别分组:')
for (unit, cat1), items in sorted(by_unit_cat1.items()):
    cat2s = defaultdict(int)
    for r in items:
        cat2s[r.get(3, '')] += 1
    cat2_list = ', '.join(f'{k}({v})' for k, v in sorted(cat2s.items()))
    print(f'  [{unit}] {cat1}: {len(items)}行 -> {cat2_list}')

# ===== 2. Check single-row & unusual subcategories =====
print()
print('=' * 80)
print('问题2: 仅1行的二级分类（可能是异常值）')
print('=' * 80)

cat2_counts = defaultdict(list)
for r in rows:
    key = (r.get(2, ''), r.get(3, ''))
    cat2_counts[key].append(r)

singles = {k: v for k, v in cat2_counts.items() if len(v) == 1}
for (cat1, cat2), items in sorted(singles.items()):
    r = items[0]
    unit = r.get(1, '')
    has_data = any(str(r.get(c, '')).strip() for c in range(5, 31))
    print(f'  [{cat1}] > [{cat2}]: 单位={unit}, 有流程数据={has_data}')
    if cat2 in ('外购电', '非设备维修', '项目投资'):
        print(f'    ⚠️ 可疑分类! 此二级似乎不属于 [{cat1}]')

# ===== 3. Check unusual category names =====
print()
print('=' * 80)
print('问题3: 可能包含注释/非标准名称的二级/三级分类')
print('=' * 80)

for r in rows:
    cat2 = str(r.get(3, ''))
    cat3 = str(r.get(4, ''))
    suspicious = []
    if '取消' in cat2 or '取消' in cat3:
        suspicious.append(f'含"取消"')
    if '集团' in cat2 and len(cat2) > 15:
        suspicious.append(f'二级名称过长可能为备注')
    if '其它' in cat2 and cat2 != '其它' and len(cat2) > 5:
        suspicious.append(f'非标准"其它"命名')
    if suspicious:
        unit = r.get(1, '')
        cat1 = r.get(2, '')
        joined = ', '.join(suspicious)
        print(f'  [{unit}] {cat1} > [{cat2}] > [{cat3}] -> {joined}')

# ===== 4. Check units that consistently lack data in certain categories =====
print()
print('=' * 80)
print('问题4: 单位覆盖缺口分析')
print('=' * 80)

# All 13 expected units
expected_units = {'大港油田', '辽河油田', '西南油田', '塔里木油田', '玉门油田',
                  '青海油田', '大庆油田', '吉林油田', '长庆油田', '冀东油田',
                  '华北油田', '新疆油田', '浙江油田'}

all_units_in_data = set(r.get(1, '') for r in rows)
print(f'数据中单位: {sorted(all_units_in_data)}')
missing = expected_units - all_units_in_data
if missing:
    print(f'缺失单位: {missing}')

# Check per-category unit coverage
cat1_units = defaultdict(set)
for r in rows:
    cat1_units[r.get(2, '')].add(r.get(1, ''))

print()
for cat1 in sorted(cat1_units.keys()):
    units = cat1_units[cat1]
    missing_units = expected_units - units
    if missing_units:
        print(f'  [{cat1}]: 缺少 {sorted(missing_units)}')
    else:
        print(f'  [{cat1}]: 13/13 ✅')

# ===== 5. Check 玉门油田 special case =====
print()
print('=' * 80)
print('问题5: 玉门油田 "平台" 列特殊值检查')
print('=' * 80)

ymen_rows = [r for r in rows if '玉门' in r.get(1, '')]
for r in ymen_rows:
    platform = str(r.get(10, '')).strip()
    if '系统' in platform or platform:
        pass  # normal
    flow = str(r.get(6, '')).strip()
    if '①' in flow or '②' in flow or '③' in flow:
        pass  # normal numbered items
    # Check for notable patterns
    cat1 = r.get(2, '')
    cat2 = r.get(3, '')

# Actually, let's check for different structural issues
print()
print('=' * 80)
print('问题6: 按单位统计"无流程数据"行数')
print('=' * 80)

unit_no_data = defaultdict(int)
unit_total = defaultdict(int)
for r in rows:
    unit = r.get(1, '')
    unit_total[unit] += 1
    has_any = any(str(r.get(c, '')).strip() for c in range(5, 31))
    if not has_any:
        unit_no_data[unit] += 1

for unit in sorted(unit_total.keys()):
    total = unit_total[unit]
    no_data = unit_no_data[unit]
    pct = no_data / total * 100 if total > 0 else 0
    flag = ' ⚠️' if pct > 15 else ''
    print(f'  {unit}: {no_data}/{total} 无流程数据 ({pct:.1f}%){flag}')

print()
print('=' * 80)
print('问题7: 浙江油田 数据来源检查（源文件不存在）')
print('=' * 80)

zj = [r for r in rows if '浙江' in r.get(1, '')]
zj_with_data = [r for r in zj if any(str(r.get(c, '')).strip() for c in range(5, 31))]
zj_no_data = [r for r in zj if not any(str(r.get(c, '')).strip() for c in range(5, 31))]
print(f'  浙江油田总行数: {len(zj)}')
print(f'  有流程数据: {len(zj_with_data)}')
print(f'  无流程数据: {len(zj_no_data)}')
# Check if the with-data rows match another unit's data (potential copy)
if zj_with_data:
    sample = zj_with_data[0]
    cat1 = sample.get(2, '')
    cat2 = sample.get(3, '')
    cat3 = sample.get(4, '')
    print(f'  示例有数据行: {cat1} > {cat2} > {cat3}')
