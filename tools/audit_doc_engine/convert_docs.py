"""
.doc → .docx 批量转换器
使用 win32com (Windows Word COM) 进行转换
"""
import sys, os, shutil
from pathlib import Path
sys.stdout.reconfigure(encoding='utf-8')

TEMPLATE_DIR = Path('D:/Users/12844/Desktop/已下发的文档模版/2.审计文档')
CONVERTED_DIR = Path('D:/Users/12844/Desktop/已下发的文档模版/2.审计文档_docx')


def convert_all():
    """批量转换所有.doc为.docx，保持目录结构"""
    CONVERTED_DIR.mkdir(exist_ok=True)

    import win32com.client
    word = win32com.client.Dispatch('Word.Application')
    word.Visible = False
    word.DisplayAlerts = 0

    stats = {'converted': 0, 'skipped': 0, 'errors': 0, 'existing_docx': 0}

    for doc_file in TEMPLATE_DIR.rglob('*.doc'):
        if doc_file.name.startswith('~$'):
            continue

        rel_path = doc_file.relative_to(TEMPLATE_DIR)
        out_file = CONVERTED_DIR / rel_path.with_suffix('.docx')
        out_file.parent.mkdir(parents=True, exist_ok=True)

        if out_file.exists():
            stats['existing_docx'] += 1
            print(f'  [跳过] {rel_path} (已存在)')
            continue

        try:
            print(f'  [转换] {rel_path} ...', end=' ')
            doc = word.Documents.Open(str(doc_file))
            doc.SaveAs2(str(out_file), FileFormat=16)  # 16 = wdFormatXMLDocument
            doc.Close()
            stats['converted'] += 1
            print('OK')
        except Exception as e:
            stats['errors'] += 1
            print(f'ERROR: {e}')

    word.Quit()

    print()
    print(f'转换: {stats["converted"]} | 跳过(已有): {stats["existing_docx"]} | 错误: {stats["errors"]}')
    print(f'输出: {CONVERTED_DIR}')
    return stats


def copy_existing_docx():
    """复制已有的.docx文件"""
    for docx_file in TEMPLATE_DIR.rglob('*.docx'):
        if docx_file.name.startswith('~$'):
            continue
        rel_path = docx_file.relative_to(TEMPLATE_DIR)
        out_file = CONVERTED_DIR / rel_path
        out_file.parent.mkdir(parents=True, exist_ok=True)
        if not out_file.exists():
            shutil.copy2(docx_file, out_file)
            print(f'  [复制] {rel_path}')


def copy_xls():
    """复制表格模板"""
    for ext in ['*.xls', '*.xlsx']:
        for f in TEMPLATE_DIR.rglob(ext):
            if f.name.startswith('~$'):
                continue
            rel = f.relative_to(TEMPLATE_DIR)
            out = CONVERTED_DIR / rel
            out.parent.mkdir(parents=True, exist_ok=True)
            if not out.exists():
                shutil.copy2(f, out)


if __name__ == '__main__':
    print('=' * 60)
    print('Step 1: 转换 .doc → .docx')
    print('=' * 60)
    convert_all()

    print()
    print('=' * 60)
    print('Step 2: 复制已有 .docx')
    print('=' * 60)
    copy_existing_docx()

    print()
    print('=' * 60)
    print('Step 3: 复制表格模板')
    print('=' * 60)
    copy_xls()

    # Count final files
    counts = {}
    for f in CONVERTED_DIR.rglob('*'):
        if f.is_file() and not f.name.startswith('~$'):
            ext = f.suffix.lower()
            counts[ext] = counts.get(ext, 0) + 1
    print()
    print(f'最终文件: {sum(counts.values())} 个')
    for ext, n in sorted(counts.items()):
        print(f'  {ext}: {n}')
