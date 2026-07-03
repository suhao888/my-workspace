"""
Runway 视频后端 — 调用 Runway Gen-3 API 图生视频
需要 RUNWAY_API_KEY 环境变量
"""

import os
import requests
import time
from typing import Optional

from generators.video_gen import VideoClipResult
from config import config


class RunwayBackend:
    """Runway ML API 图生视频"""

    def __init__(self):
        self.api_key = config.video.runway_api_key or os.environ.get(
            "RUNWAY_API_KEY", ""
        )
        if not self.api_key:
            raise ValueError("Runway API Key 未设置")

    def generate(
        self, image_path: str, duration: float, output_path: str
    ) -> Optional[VideoClipResult]:
        """
        调用 Runway Gen-3 API
        注意：Runway API 接口可能更新，请以最新文档为准
        """
        # --- 模板代码，需替换为实际 Runway API 调用 ---
        print(f"  [Runway] 图片: {image_path}, 时长: {duration}s")
        print(f"  [Runway] 请实现 generate() 方法")
        print(f"  [Runway] 视频将保存到: {output_path}")
        # --- 模板结束 ---
        return None
