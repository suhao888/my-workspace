import sys

sys.stdout.reconfigure(encoding="utf-8")
from docx import Document

doc = Document("D:/Users/12844/Desktop/审计输出_测试/附注_单体_v11_测试.docx")

t = doc.tables[99]
print("=== T99 使用权资产-情况 ===")
for ri in range(len(t.rows)):
    texts = [
        t.rows[ri].cells[ci].text.strip() for ci in range(min(6, len(t.rows[ri].cells)))
    ]
    print(f"  R{ri:2d}: {texts}")
    # verifica勾稽
    if ri > 0 and texts[1] and texts[4]:
        try:
            beg = float(texts[1].replace(",", "").replace("—", "0"))
            end = float(texts[4].replace(",", "").replace("—", "0"))
            inc = (
                float(texts[2].replace(",", ""))
                if texts[2] and texts[2] not in ("—", "")
                else 0
            )
            dec = (
                float(texts[3].replace(",", ""))
                if texts[3] and texts[3] not in ("—", "")
                else 0
            )
            calc = beg + inc - dec
            if abs(calc - end) > 0.1:
                print(f"    ⚠ 勾稽异常: {beg} + {inc} - {dec} = {calc} != {end}")
            else:
                print(f"    ✓ {beg} + {inc} - {dec} = {end}")
        except Exception as e:
            print(f"    ? 无法验算: {e}")
