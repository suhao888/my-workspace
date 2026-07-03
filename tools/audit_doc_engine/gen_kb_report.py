# -*- coding: utf-8 -*-
"""
知识库标准文件更新报告生成器
============================
读取 .claude/knowledge/standards/ 下的三个标准文件，
生成结构化的 DOCX 报告。

 用法: python gen_kb_report.py
 输出: D:\\Users\\12844\\Desktop\\知识库标准文件更新报告_20260703.docx
"""

import os, sys, re, datetime, shutil, time
from pathlib import Path

_PROJECT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _PROJECT not in sys.path:
    sys.path.insert(0, _PROJECT)
from tools.audit_doc_engine.doc_builder import DocBuilder

# ── 知识库标准文件路径 ──
KB_DIR = Path(r"G:\opencode-GIT仓库\.claude\knowledge\standards")
FILES = [
    ("审计准则（CSAE）", "audit-standards.md"),
    ("企业会计准则解释（EAS Interpretation）", "accounting-standards.md"),
    ("注册会计师法规（Comprehensive Laws）", "comprehensive-laws.md"),
]

OUTPUT = r"D:\Users\12844\Desktop\知识库标准文件更新报告_20260703.docx"
TEMP = r"C:\Users\12844\AppData\Local\Temp\__kb_report_temp.docx"

# ── 文件解析 ──


def parse_md_sections(filepath: Path):
    """解析 Markdown 文件，提取一级/二级标题和概要内容。"""
    if not filepath.exists():
        return {"line_count": 0, "sections": []}
    with open(filepath, "r", encoding="utf-8") as f:
        lines = f.readlines()
    sections = []
    for line in lines:
        if line.startswith("## ") and not line.startswith("### "):
            title = line.strip("# \n")
            sections.append({"level": 2, "title": title, "lines": 1})
        elif line.startswith("# ") and not line.startswith("## "):
            title = line.strip("# \n")
            sections.append({"level": 1, "title": title, "lines": 1})
    return {"line_count": len(lines), "sections": sections}


def get_file_info(filepath: Path):
    if not filepath.exists():
        return {"size": 0, "mtime": "", "line_count": 0}
    stat = filepath.stat()
    mtime = datetime.datetime.fromtimestamp(stat.st_mtime).strftime("%Y-%m-%d")
    info = parse_md_sections(filepath)
    info["size_kb"] = round(stat.st_size / 1024, 1)
    info["mtime"] = mtime
    return info


# ── 报告生成 ──


def generate(output_path: str):
    b = DocBuilder()
    b.setup_page().setup_styles(body_size=11, line_spacing=1.35)

    # 非审计报告——不加粗、无封面、无目录
    bold = False

    b.add_heading_1("知识库标准文件更新报告", bold=bold)
    b.add_body_no_indent("编制日期：2026年7月3日")

    b.add_heading_1("一、概述", bold=bold)
    b.add_body(
        "本次更新涉及知识库标准文件目录下的三份核心标准文件，"
        "涵盖审计准则、企业会计准则解释和注册会计师法规三大领域。"
        "以下按文件逐一说明各部分内容结构和关键要点。"
    )

    # 汇总表
    rows = []
    for label, fname in FILES:
        info = get_file_info(KB_DIR / fname)
        rows.append(
            [
                label,
                fname,
                str(info.get("size_kb", 0)) + " KB",
                str(info.get("line_count", 0)) + " 行",
                info.get("mtime", ""),
            ]
        )
    b.add_table(
        headers=["文件", "文件名", "大小", "篇幅", "更新日期"],
        rows=rows,
        col_widths=[4, 3.5, 1.5, 1.5, 2],
        caption="表：本次覆盖的标准文件",
    )

    # 各文件详情
    for idx, (label, fname) in enumerate(FILES, 1):
        sections_text = ("一", "二", "三")[idx - 1]
        b.add_heading_1(f"{sections_text}、{label}", bold=bold)

        filepath = KB_DIR / fname
        if not filepath.exists():
            b.add_body(f"文件 {fname} 不存在。")
            continue

        info = get_file_info(filepath)
        b.add_kv_table(
            [
                ("文件路径", str(filepath)),
                ("文件大小", f"{info['size_kb']} KB"),
                ("总行数", f"{info['line_count']} 行"),
                ("最后修改", info["mtime"]),
            ],
            label_width=3.0,
            total_width=15.0,
        )

        b.add_heading_2("（一）内容结构", bold=bold)
        sections = info["sections"]
        if sections:
            for sec in sections:
                b.add_body_no_indent(sec["title"])
        else:
            b.add_body("（无结构化章节信息）")

        with open(filepath, "r", encoding="utf-8") as f:
            first_lines = [f.readline().strip() for _ in range(5)]
        desc_lines = [
            l for l in first_lines if l and not l.startswith(">") and l != "---"
        ]
        if desc_lines:
            b.add_heading_2("（二）文件说明", bold=bold)
            for dl in desc_lines[:3]:
                if dl.startswith("#"):
                    dl = dl.lstrip("# ").strip()
                b.add_body_no_indent(dl)

    # 签名
    b.add_space(30)
    b.add_signature(
        [
            ("编制人", "知识库管理系统"),
            ("编制日期", "2026年7月3日"),
            ("审核人", "（待填写）"),
        ]
    )

    b.save(output_path)
    size = os.path.getsize(output_path) / 1024
    print(f"文档已生成: {output_path} ({size:.1f} KB)")


def main():
    # 先保存到临时路径（避免被占用文件的锁）
    generate(TEMP)

    # 尝试覆盖目标文件
    for attempt in range(3):
        try:
            if os.path.exists(OUTPUT):
                os.remove(OUTPUT)
            shutil.copy2(TEMP, OUTPUT)
            print(f"已覆盖: {OUTPUT}")
            os.remove(TEMP)
            return
        except PermissionError:
            print(f"尝试 {attempt + 1}/3: 目标文件被占用，等待 2 秒...")
            time.sleep(2)

    print(f"覆盖失败，临时文件保留在: {TEMP}")
    print(f'请关闭 Word 中的文件后再执行: copy /y "{TEMP}" "{OUTPUT}"')


if __name__ == "__main__":
    main()
