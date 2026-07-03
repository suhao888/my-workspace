# -*- coding: utf-8 -*-
"""
知乎自动化助手 - 主入口
"""

import sys
import os
from datetime import datetime

sys.path.insert(0, os.path.dirname(__file__))

import config
import db
import scraper
import generator


def print_banner():
    print("=" * 50)
    print("  知乎自动化助手 v1.0 (MVP)")
    print(f"  运行时间: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 50)


def check_config():
    issues = []
    if not config.ZHIHU_COOKIE:
        issues.append("ZHIHU_COOKIE 未设置")
    api_key = config.DEEPSEEK_API_KEY or os.environ.get("DEEPSEEK_API_KEY")
    if not api_key:
        issues.append("DEEPSEEK_API_KEY 未设置")
    return issues


def run():
    print_banner()

    issues = check_config()
    if issues:
        print("\n[配置] 请先修复以下问题：")
        for i, issue in enumerate(issues, 1):
            print(f"  {i}. {issue}")
        print()
        return

    db.init_db()

    # 步骤1: 热点监控
    print("\n[步骤1/3] 热点监控...")
    questions = scraper.fetch_and_filter()

    if not questions:
        print("  未获取到任何热点，请检查 Cookie 是否有效。")
        return

    # 过滤已处理的问题
    new_questions = []
    for q in questions:
        if not db.is_processed(q["question_id"]):
            if q.get("topic_score", 0) > 0 or q["source"] == "hotlist":
                new_questions.append(q)

    print(f"  新增未处理问题: {len(new_questions)} 条")

    if not new_questions:
        print("\n[步骤2/3] 跳过 - 没有新问题需要处理")
        print("[步骤3/3] 跳过")
        print("\n--- 完成 ---")
        stats = db.get_stats()
        print(f"  累计收录问题: {stats['total']}")
        print(f"  已生成草稿: {stats['generated']}")
        print(f"  待处理: {stats['pending']}")
        print(f"  已发布: {stats['posted']}")
        return

    # 保存到数据库
    for q in new_questions:
        db.save_question(
            q["question_id"],
            q["title"],
            q["url"],
            q["source"],
            q.get("topic", ""),
            q.get("hot_score", 0),
        )

    # 步骤1.5: 曝光度评分
    print(f"\n[步骤1.5/3] 曝光度评分 (获取前 {len(new_questions)} 条曝光数据)...")
    top_n = min(len(new_questions), 30)  # 最多查 30 条详情
    enriched = scraper.enrich_with_exposure(new_questions[:top_n], max_fetch=top_n)

    # 高曝光标准：曝光分 > 1000
    high_exposure = [q for q in enriched if q.get("exposure_score", 0) > 1000]
    low_exposure = [q for q in enriched if q.get("exposure_score", 0) <= 1000]

    print(f"\n  高曝光问题: {len(high_exposure)} 条")
    if high_exposure:
        for q in high_exposure[:5]:
            d = q.get("detail", {}) or {}
            print(
                f"    曝光={q['exposure_score']:.0f} | "
                f"浏览={d.get('visit_count', '?')} "
                f"回答={d.get('answer_count', '?')} "
                f"关注={d.get('follower_count', '?')} | "
                f"{q['title'][:40]}"
            )
    print(f"  低曝光问题: {len(low_exposure)} 条（跳过不生成）")

    # 保存曝光数据到数据库
    for q in enriched:
        d = q.get("detail", {}) or {}
        if d:
            db._save_exposure(
                q["question_id"],
                d.get("visit_count", 0),
                d.get("answer_count", 0),
                d.get("follower_count", 0),
                q["exposure_score"],
            )

    # 步骤2: 内容生成（只生成高曝光的问题）
    to_generate = [q for q in high_exposure[: config.MAX_DRAFTS_PER_RUN]]
    if not to_generate:
        print("\n[步骤2/3] 跳过 - 没有高曝光问题需要生草稿")
        print("\n--- 完成 ---")
        stats = db.get_stats()
        print(f"  累计收录问题: {stats['total']}")
        print(f"  已生成草稿: {stats['generated']}")
        return

    print(f"\n[步骤2/3] 内容生成 (高曝光问题 {len(to_generate)} 篇)...")
    generated_count = 0

    for i, q in enumerate(to_generate, 1):
        print(f"\n  [{i}/{len(to_generate)}] 正在生成...")
        print(f"    问题: {q['title'][:60]}")

        content = generator.generate_answer(q["title"])

        if content:
            filepath = generator.save_draft_file(
                q["title"], q["question_id"], content, q.get("url", "")
            )
            db.update_status(q["question_id"], "generated", filepath)
            db.save_draft(q["question_id"], content)

            print(
                f"    [OK] 草稿已保存: {os.path.basename(filepath)} ({len(content)} 字)"
            )
            generated_count += 1
        else:
            print(f"    [X] 生成失败，跳过")
            db.update_status(q["question_id"], "skipped")

    # 步骤3: 汇总
    print(f"\n[步骤3/3] 汇总")
    stats = db.get_stats()
    print(f"  本运行生成: {generated_count} 篇")
    print(f"  累计: {stats['total']} 条问题, {stats['generated']} 篇草稿")

    today = datetime.now().strftime("%Y-%m-%d")
    draft_dir = os.path.join(os.path.dirname(__file__), "drafts", today)
    if os.path.exists(draft_dir):
        print(f"\n[草稿目录] {draft_dir}")
        print("   审核修改后手动发布到知乎")

    print("\n--- 完成 ---")


if __name__ == "__main__":
    run()
