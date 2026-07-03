#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
自动获取知乎 Cookie
使用 Edge 的现有用户数据，继承登录态
"""

import os
import json
import sys
import time
import subprocess

EDGE_USER_DATA = os.path.expandvars(r"%LOCALAPPDATA%\Microsoft\Edge\User Data")
PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))
COOKIE_FILE = os.path.join(PROJECT_DIR, "zhihu_cookie.txt")


def main():
    print("=" * 50)
    print("知乎 Cookie 自动获取工具")
    print("=" * 50)

    # 检查 Edge 是否在运行
    result = subprocess.run(
        ["tasklist", "/FI", "IMAGENAME eq msedge.exe"],
        capture_output=True,
        text=True,
        timeout=10,
    )
    if "msedge.exe" in result.stdout:
        print("\n[!] 检测到 Edge 正在运行")
        print("  正在关闭 Edge...")
        subprocess.run(
            ["taskkill", "/F", "/IM", "msedge.exe"], capture_output=True, timeout=10
        )
        time.sleep(2)
        print("  Edge 已关闭")

    from playwright.sync_api import sync_playwright

    with sync_playwright() as p:
        # 用 Edge 的用户数据启动浏览器
        print("\n[1/4] 启动浏览器（使用 Edge 配置）...")
        context = p.chromium.launch_persistent_context(
            user_data_dir=EDGE_USER_DATA,
            channel="msedge",
            headless=False,  # 可见窗口，方便用户操作
            viewport={"width": 1200, "height": 800},
            no_viewport=False,
        )

        page = context.pages[0] if context.pages else context.new_page()

        # 跳转知乎
        print("[2/4] 访问知乎...")
        page.goto("https://www.zhihu.com/hot", timeout=30000)
        page.wait_for_load_state("networkidle")
        time.sleep(2)

        # 检查登录状态
        url = page.url
        print(f"  当前页面: {url[:60]}")

        # 用户控件 - 检查是否显示头像/用户名
        is_logged_in = False
        try:
            # 找用户头像或用户名元素
            user_elem = page.query_selector(
                '[class*="UserLink"], [class*="Avatar"], [class*="Profile"]'
            )
            if user_elem:
                is_logged_in = True
        except:
            pass

        if not is_logged_in and (
            "passport" in url or "login" in url or "signin" in url
        ):
            print("\n[3/4] [!] 检测到登录页面")
            print("  请在浏览器窗口中手动登录知乎")
            print("  （扫码或账号密码都行）")
            print("  登录后脚本会自动继续...\n")

            # 等待用户登录，最多等 3 分钟
            try:
                page.wait_for_url(
                    lambda u: "passport" not in u and "login" not in u, timeout=180000
                )
                print("  [OK] 登录检测成功")
            except:
                print("  [超时] 等待超时，请重新运行脚本")
                context.close()
                return
        else:
            print("[3/4] [OK] 已登录知乎")

        # 等待页面加载完全
        time.sleep(2)

        # 提取 cookie
        print("[4/4] 提取 Cookie...")
        cookies = context.cookies()
        zhihu_cookies = [c for c in cookies if "zhihu.com" in c.get("domain", "")]

        if not zhihu_cookies:
            print("  [X] 未找到知乎 Cookie")
            # 尝试刷新
            page.reload()
            page.wait_for_load_state("networkidle")
            time.sleep(2)
            cookies = context.cookies()
            zhihu_cookies = [c for c in cookies if "zhihu.com" in c.get("domain", "")]

        if zhihu_cookies:
            # 生成 Cookie 字符串
            cookie_str = "; ".join(f"{c['name']}={c['value']}" for c in zhihu_cookies)

            # 保存到文件
            with open(COOKIE_FILE, "w", encoding="utf-8") as f:
                f.write(cookie_str)

            print(f"  [OK] 获取到 {len(zhihu_cookies)} 条 Cookie")
            print(f"  [OK] 已保存到: {COOKIE_FILE}")

            # 自动更新 config.py
            config_path = os.path.join(PROJECT_DIR, "config.py")
            with open(config_path, "r", encoding="utf-8") as f:
                content = f.read()

            old = 'ZHIHU_COOKIE = ""'
            new = f'ZHIHU_COOKIE = """{cookie_str}"""'

            if old in content:
                content = content.replace(old, new)
                with open(config_path, "w", encoding="utf-8") as f:
                    f.write(content)
                print("  [OK] 已自动写入 config.py")
            else:
                print("  [!] 请手动将 Cookie 粘贴到 config.py 的 ZHIHU_COOKIE 位置")

            # 测试一下
            print("\n--- 测试 Cookie 有效性 ---")
            import requests

            headers = {
                "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
                "Cookie": cookie_str,
                "Accept": "application/json",
            }
            test_url = (
                "https://www.zhihu.com/api/v3/feed/topstory/hot-lists/total?limit=3"
            )
            resp = requests.get(test_url, headers=headers, timeout=15)

            if resp.status_code == 200:
                data = resp.json()
                items = data.get("data", [])
                print(f"  [OK] Cookie 有效！热榜数据获取成功")
                for item in items[:3]:
                    t = item.get("target", {}).get("title", "N/A")
                    print(f"     - {t}")
            else:
                print(f"  [X] Cookie 无效: HTTP {resp.status_code}")
                if resp.status_code == 401:
                    print("  Cookie 已过期，请重新运行本脚本")
        else:
            print("  [X] 未能获取到知乎 Cookie")

        # 关闭浏览器
        context.close()
        print("\n浏览器窗口已关闭")


if __name__ == "__main__":
    main()
