# -*- coding: utf-8 -*-
"""
热点监控模块 - 从知乎获取热榜和搜索结果
两种数据源：
1. 知乎热榜 API
2. 知乎搜索 API

需要 Cookie 认证，从 config.py 读取
"""

import requests
import re
import json
import time
import config


def _headers():
    """构造请求头，携带 Cookie"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36",
        "Accept": "application/json, text/plain, */*",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8",
        "Referer": "https://www.zhihu.com/",
        "Origin": "https://www.zhihu.com",
    }
    if config.ZHIHU_COOKIE:
        headers["Cookie"] = config.ZHIHU_COOKIE
    return headers


def fetch_hotlist(limit=50):
    """获取知乎热榜"""
    url = f"https://www.zhihu.com/api/v3/feed/topstory/hot-lists/total?limit={limit}"

    try:
        resp = requests.get(url, headers=_headers(), timeout=15)
        if resp.status_code != 200:
            print(f"  [热榜] 请求失败: {resp.status_code}")
            if resp.status_code == 401:
                print(
                    "  [热榜] Cookie 无效或已过期，请更新 config.py 中的 ZHIHU_COOKIE"
                )
            return []

        data = resp.json()
        items = data.get("data", [])
        results = []

        for item in items:
            target = item.get("target", {})
            q_id = target.get("id", "")
            title = target.get("title", "")

            # 过滤掉没有标题的数据
            if not q_id or not title:
                continue

            # 热度分
            detail = item.get("detail", item.get("feed_specific", {}))
            hot_score = 0
            if isinstance(detail, dict):
                hot_score = detail.get("hot", 0) or detail.get("score", 0) or 0

            results.append(
                {
                    "question_id": str(q_id),
                    "title": title.strip(),
                    "url": f"https://www.zhihu.com/question/{q_id}",
                    "source": "hotlist",
                    "topic": "",
                    "hot_score": hot_score,
                }
            )

        print(f"  [热榜] 获取 {len(results)} 条")
        return results

    except requests.exceptions.Timeout:
        print("  [热榜] 请求超时")
        return []
    except Exception as e:
        print(f"  [热榜] 异常: {e}")
        return []


def _clean_title(text):
    """去掉 <em> 等 HTML 标签"""
    if not text:
        return ""
    return re.sub(r"<[^>]+>", "", text).strip()


def search_questions(keyword, limit=10):
    """搜索知乎问题"""
    url = (
        f"https://www.zhihu.com/api/v4/search_v3"
        f"?t=general&q={keyword}&correction=1&limit={limit}"
    )

    try:
        resp = requests.get(url, headers=_headers(), timeout=15)
        if resp.status_code != 200:
            print(f"  [搜索] '{keyword}' 失败: {resp.status_code}")
            return []

        data = resp.json()
        items = data.get("data", [])
        results = []

        for item in items:
            if item.get("type") != "search_result":
                continue

            obj = item.get("object") or {}
            obj_type = obj.get("type")

            q_id = None
            title = None

            if obj_type == "answer":
                # 答案对象的问题信息在 question 字段里
                question = obj.get("question") or {}
                q_id = question.get("id")
                title = _clean_title(question.get("name", ""))
            elif obj_type == "question":
                # 直接的问题对象
                q_id = obj.get("id")
                title = _clean_title(obj.get("name", obj.get("title", "")))
            else:
                # article 或其他类型，跳过
                continue

            if not q_id or not title:
                continue

            # 解析热度
            label = item.get("label", {})
            hot_score = 0
            if isinstance(label, dict):
                hot_score = label.get("score", label.get("hot", 0)) or 0

            results.append(
                {
                    "question_id": str(q_id),
                    "title": title,
                    "url": f"https://www.zhihu.com/question/{q_id}",
                    "source": "search",
                    "topic": keyword,
                    "hot_score": hot_score,
                }
            )

        print(f"  [搜索] '{keyword}' 找到 {len(results)} 条")
        return results

    except requests.exceptions.Timeout:
        print(f"  [搜索] '{keyword}' 超时")
        return []
    except Exception as e:
        print(f"  [搜索] '{keyword}' 异常: {e}")
        return []


def topic_match_score(title, topics):
    """计算标题与内容方向的匹配度"""
    score = 0
    title_lower = title.lower()
    for topic in topics:
        if topic.lower() in title_lower:
            score += 10
        # 检查话题中的关键词
        for word in topic:
            if word in title_lower:
                score += 1
    return score


def fetch_and_filter():
    """获取热点并过滤"""
    all_questions = []

    # 1. 热榜
    if config.HOT_SOURCE in ("hotlist", "both"):
        print("[热点] 正在获取知乎热榜...")
        all_questions.extend(fetch_hotlist(config.HOTLIST_LIMIT))

    # 2. 关键词搜索
    if config.HOT_SOURCE in ("search", "both"):
        print("[热点] 正在搜索相关关键词...")
        for keyword in config.SEARCH_KEYWORDS:
            all_questions.extend(search_questions(keyword, config.SEARCH_LIMIT))
            time.sleep(1)  # 避免请求过快

    # 3. 去重（按 question_id）
    seen = set()
    unique = []
    for q in all_questions:
        if q["question_id"] not in seen:
            seen.add(q["question_id"])
            unique.append(q)

    # 4. 计算话题匹配度
    for q in unique:
        q["topic_score"] = topic_match_score(q["title"], config.CONTENT_TOPICS)

    # 5. 排序：匹配度优先，热度次之
    unique.sort(key=lambda x: (x["topic_score"], x["hot_score"]), reverse=True)

    print(f"[热点] 总计获取 {len(all_questions)} 条，去重后 {len(unique)} 条")

    return unique


# ======== 曝光度评分 ========


def _extract_int(text, key):
    """从 HTML 中提取 key:value 形式的数字"""
    m = re.search(rf'"{key}":(\d+)', text)
    return int(m.group(1)) if m else 0


def get_question_detail(session, qid):
    """获取问题详情：回答数、浏览量、关注数"""
    url = f"https://www.zhihu.com/question/{qid}"
    try:
        r = session.get(url, timeout=15)
        if r.status_code != 200:
            return None
        return {
            "answer_count": _extract_int(r.text, "answerCount"),
            "visit_count": _extract_int(r.text, "visitCount"),
            "follower_count": _extract_int(r.text, "followerCount"),
        }
    except Exception:
        return None


def calc_exposure_score(detail):
    """计算曝光度分数

    核心逻辑：浏览量越高越好，回答数越少越好（竞争小）
    score = visit_count * (1 + follower_count/100) / max(1, answer_count)
    """
    if not detail:
        return 0
    ac = max(1, detail["answer_count"])
    return detail["visit_count"] * (1 + detail["follower_count"] / 100) / ac


def enrich_with_exposure(questions, max_fetch=30):
    """获取问题的曝光数据并打分

    Args:
        questions: 问题列表
        max_fetch: 最多获取前 N 条问题的详情

    Returns:
        添加了 exposure_score, detail 字段的问题列表，按分数降序
    """
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/125.0.0.0 Safari/537.36",
        "Cookie": config.ZHIHU_COOKIE,
    }
    s = requests.Session()
    s.headers.update(headers)

    # 先访问首页建立会话
    try:
        s.get("https://www.zhihu.com", timeout=10)
    except:
        pass

    scored = []
    for i, q in enumerate(questions):
        if i >= max_fetch:
            q["exposure_score"] = 0
            q["detail"] = None
            scored.append(q)
            continue

        detail = get_question_detail(s, q["question_id"])
        q["detail"] = detail
        q["exposure_score"] = calc_exposure_score(detail)
        scored.append(q)

        # 打印日志
        if detail:
            flag = " [高曝光]" if q["exposure_score"] > 5000 else ""
            print(
                f"  [{i + 1}/{min(max_fetch, len(questions))}] "
                f"曝光={q['exposure_score']:.0f} "
                f"浏览={detail['visit_count']} "
                f"回答={detail['answer_count']} "
                f"关注={detail['follower_count']}{flag}"
            )
        else:
            print(f"  [{i + 1}/{min(max_fetch, len(questions))}] 获取详情失败")

        time.sleep(1)  # 避免请求过快

    # 按曝光度降序
    scored.sort(key=lambda x: x["exposure_score"], reverse=True)
    return scored


if __name__ == "__main__":
    if not config.ZHIHU_COOKIE:
        print("请先在 config.py 中设置 ZHIHU_COOKIE")
    else:
        questions = fetch_and_filter()
        for q in questions[:10]:
            print(f"  [{q['topic_score']}] {q['title'][:50]}")
