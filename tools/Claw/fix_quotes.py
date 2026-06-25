with open(r'C:/Users/12844/WorkBuddy/Claw/gen_audit_report.js', 'r', encoding='utf-8') as f:
    content = f.read()

content = content.replace('\u201c', '\\u201c')
content = content.replace('\u201d', '\\u201d')
content = content.replace('\u2018', '\\u2018')
content = content.replace('\u2019', '\\u2019')

with open(r'C:/Users/12844/WorkBuddy/Claw/gen_audit_report.js', 'w', encoding='utf-8') as f:
    f.write(content)

print('Done')
