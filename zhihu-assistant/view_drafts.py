# -*- coding: utf-8 -*-
"""Generate draft viewer HTML"""

import os, html
from datetime import datetime

DRAFTS_DIR = os.path.join(os.path.dirname(__file__), "drafts")
OUTPUT_FILE = os.path.join(os.path.dirname(__file__), "drafts.html")


def generate_html():
    drafts = []
    for dir_name in sorted(os.listdir(DRAFTS_DIR), reverse=True):
        dir_path = os.path.join(DRAFTS_DIR, dir_name)
        if not os.path.isdir(dir_path):
            continue
        for fname in os.listdir(dir_path):
            if not fname.endswith(".md"):
                continue
            with open(os.path.join(dir_path, fname), "r", encoding="utf-8") as f:
                content = f.read()

            lines = content.split("\n")
            meta = {}
            body_lines = []
            in_meta = False
            meta_closed = False
            for line in lines:
                if line.strip() == "---":
                    if not in_meta:
                        in_meta = True
                    elif not meta_closed:
                        meta_closed = True
                    continue
                if in_meta and not meta_closed and ":" in line:
                    k, _, v = line.partition(":")
                    meta[k.strip()] = v.strip()
                elif meta_closed:
                    body_lines.append(line)

            body = "\n".join(body_lines).strip()
            if body.startswith("# "):
                body = "\n".join(body.split("\n")[1:]).strip()

            qid = meta.get("问题ID", "")
            url = meta.get("问题链接", "")
            if not url and qid:
                url = f"https://www.zhihu.com/question/{qid}"

            drafts.append(
                {
                    "title": meta.get("问题标题", fname),
                    "url": url,
                    "date": dir_name,
                    "words": str(len(body)),
                    "body": body,
                }
            )

    cards = ""
    for d in drafts:
        body_html = html.escape(d["body"]).replace("\n", "<br>")
        title_html = html.escape(d["title"])
        link_html = ""
        if d["url"]:
            link_html = f'<a class="qlink" href="{html.escape(d["url"])}" target="_blank">打开问题 &rarr;</a>'

        cards += f"""
        <div class="draft-card">
            <div class="draft-header">
                <h2>{title_html}</h2>
                <div class="draft-meta">
                    <span>{d["date"]}</span>
                    <span>{d["words"]} 字</span>
                    {link_html}
                </div>
            </div>
            <div class="draft-body">{body_html}</div>
        </div>"""

    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>知乎草稿</title>
<style>
* {{ margin: 0; padding: 0; box-sizing: border-box; }}
body {{ background: #f5f5f5; font-family: -apple-system, 'Microsoft YaHei', sans-serif; padding: 20px; }}
.header {{ max-width: 900px; margin: 0 auto 20px; }}
.header h1 {{ font-size: 22px; color: #333; }}
.header p {{ color: #888; font-size: 14px; margin-top: 4px; }}
.draft-card {{ background: #fff; border-radius: 8px; box-shadow: 0 1px 4px rgba(0,0,0,0.08); margin: 0 auto 20px; max-width: 900px; overflow: hidden; }}
.draft-header {{ padding: 16px 20px; border-bottom: 1px solid #f0f0f0; }}
.draft-header h2 {{ font-size: 17px; color: #1a1a1a; font-weight: 600; }}
.draft-meta {{ display: flex; gap: 16px; margin-top: 6px; font-size: 12px; color: #999; align-items: center; }}
.qlink {{ margin-left: auto; background: #1a73e8; color: #fff !important; text-decoration: none; padding: 4px 12px; border-radius: 4px; font-size: 12px; }}
.qlink:hover {{ background: #1557b0; }}
.draft-body {{ padding: 20px; line-height: 1.8; font-size: 15px; color: #333; white-space: pre-wrap; }}
</style>
</head>
<body>
<div class="header">
    <h1>知乎回答草稿</h1>
    <p>共 {len(drafts)} 篇 | AI 辅助生成，发布前请审核修改</p>
</div>
{cards}
</body>
</html>"""

    with open(OUTPUT_FILE, "w", encoding="utf-8") as f:
        f.write(html_content)
    print(f"Done: {OUTPUT_FILE}")


if __name__ == "__main__":
    generate_html()
