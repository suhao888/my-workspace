# -*- coding: utf-8 -*-
"""
内容生成模块 - 使用 DeepSeek API 生成回答草稿
"""

import os
import json
import requests
from datetime import datetime
import config


def _get_api_key():
    """获取 DeepSeek API Key"""
    if config.DEEPSEEK_API_KEY:
        return config.DEEPSEEK_API_KEY
    key = os.environ.get("DEEPSEEK_API_KEY")
    if not key:
        raise ValueError(
            "未找到 DeepSeek API Key。\n"
            "请在 config.py 中设置 DEEPSEEK_API_KEY，\n"
            "或设置环境变量 DEEPSEEK_API_KEY"
        )
    return key


def _build_prompt(question_title, topics, style_hint=""):
    """构建生成回答的提示词"""
    topics_str = "、".join(topics)

    prompt = f"""你模仿一位知乎答主，身份是干了十来年审计的实务派，写过会计考证经验也带过团队。

要求你写一篇回答，必须遵守以下规则：

**格式规则：**
- 第一句就扔观点或经历，别铺垫
- 段落不超过 4 行，大部分 2-3 行就换段
- 能用短句就别用长句。一句话能说清的事不拆成三句
- 可以偶尔用破折号或括号加补充说明
- 从头到尾不要出现：值得注意的是、综上所述、首先其次再次、众所周知、毋庸置疑、换言之、不可否认、从这个角度来看

**内容规则：**
- 必须有具体数字或案例。"我去年接了个IPO项目"比"审计工作很忙"强一百倍
- 可以说"我认识的CPA"、"我之前在八大"、"上个月有个客户"这类个人视角的话
- 观点可以偏激，但要圆回来。比如"考证没啥用——但没证你连门都进不去"
- 不写正确的废话。"审计需要专业能力"这种话删掉，换成"没在年审通宵过三个月的项目经理不是好项目经理"
- 该用术语用术语（存货监盘、截止测试、审计调整、实质性程序），但用完后用大白话解释一下

**结构规则：**
- 不要用"第一第二第三"或"首先其次再次"来组织内容
- 自然的过渡：讲完一个点，扔个转折或者反问，再讲下一个
- 如果要用列举，最多列两条或四条，不要列三条
- 结尾利索，不用总结性段落。讲完最后一个观点就停

**语言风格：**
- 像在跟同行/后辈聊天，不是在写报告
- 可以用点口语："说白了"、"讲真"、"你信不信"
- 可以有情绪：吐槽、无奈、兴奋都行，别端着
- 字数 600-1200 字，讲清楚就行不硬凑

问题：{question_title}
{style_hint}
直接输出回答正文。"""

    return prompt


def generate_answer(question_title, topics=None, style_hint=""):
    """调用 DeepSeek API 生成回答"""
    if topics is None:
        topics = config.CONTENT_TOPICS

    api_key = _get_api_key()
    prompt = _build_prompt(question_title, topics, style_hint)

    payload = {
        "model": config.DEEPSEEK_MODEL,
        "messages": [
            {
                "role": "system",
                "content": "你是在知乎写回答的审计老人。不讲套话，不端架子，用故事和案例说事，像跟同行聊天一样写。",
            },
            {"role": "user", "content": prompt},
        ],
        "temperature": 0.7,
        "max_tokens": 3000,
    }

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    try:
        resp = requests.post(
            config.DEEPSEEK_API_URL,
            headers=headers,
            json=payload,
            timeout=60,
        )

        if resp.status_code != 200:
            print(f"  [生成] API 请求失败: {resp.status_code}")
            print(f"  [生成] 响应: {resp.text[:200]}")
            return None

        result = resp.json()
        content = result.get("choices", [{}])[0].get("message", {}).get("content", "")

        if not content:
            print("  [生成] 返回内容为空")
            return None

        return content.strip()

    except requests.exceptions.Timeout:
        print("  [生成] API 请求超时")
        return None
    except Exception as e:
        print(f"  [生成] 异常: {e}")
        return None


def save_draft_file(question_title, question_id, content, question_url=""):
    """将草稿保存为 Markdown 文件"""
    today = datetime.now().strftime("%Y-%m-%d")
    draft_dir = os.path.join(os.path.dirname(__file__), "drafts", today)
    os.makedirs(draft_dir, exist_ok=True)

    safe_title = "".join(c for c in question_title if c.isalnum() or c in " _-")[:20]
    filename = f"{question_id}_{safe_title}.md"
    filepath = os.path.join(draft_dir, filename)

    if not question_url:
        question_url = f"https://www.zhihu.com/question/{question_id}"

    content_full = f"""---
问题ID: {question_id}
问题标题: {question_title}
问题链接: {question_url}
生成时间: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")}
状态: 待审核
---

# {question_title}

{content}

---
> 此草稿由 AI 辅助生成，发布前请审核修改。
"""

    with open(filepath, "w", encoding="utf-8") as f:
        f.write(content_full)

    return filepath


if __name__ == "__main__":
    # 测试
    content = generate_answer("审计行业有哪些不为人知的真相？")
    if content:
        path = save_draft_file("审计行业有哪些不为人知的真相？", "test", content)
        print(f"草稿已保存: {path}")
        print(f"内容长度: {len(content)} 字")
