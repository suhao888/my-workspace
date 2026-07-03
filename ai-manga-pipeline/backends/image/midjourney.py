"""
Midjourney 图片后端 — 通过 API 调用 Midjourney 生成图片
需要 MJ_API_KEY 环境变量
注意：MJ API 一般有独立的触发词格式，需自行适配
"""

import os
import requests
import time
from typing import Optional

from generators.image_gen import ImageResult
from config import config


# MJ API 各家服务商接口差异较大，这里提供一个通用模板
# 实际使用时需替换为你的 MJ API 供应商的具体调用方式


class MidjourneyBackend:
    """Midjourney API 文生图（模板）"""

    def __init__(self):
        self.api_key = config.image.midjourney_api_key or os.environ.get(
            "MJ_API_KEY", ""
        )
        if not self.api_key:
            raise ValueError("Midjourney API Key 未设置")

    def generate(self, prompt: str, output_path: str) -> Optional[ImageResult]:
        """
        通用 MJ API 调用模板
        实际使用时根据你的 API 供应商修改此方法
        """
        # --- 替换为你的 MJ API 调用逻辑 ---
        # 示例：调用 Imagine API → 获取图片 URL → 下载保存
        # 不同供应商（Midjourney-Proxy、Goapi、ChatGPT 等）接口各异
        print(f"  [Midjourney] 提示: {prompt[:60]}...")
        print(f"  [Midjourney] 请实现 generate() 方法以调用你的 MJ API")
        print(f"  [Midjourney] 图片将保存到: {output_path}")
        # --- 示例结束 ---
        return None
