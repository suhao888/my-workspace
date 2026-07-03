"""
SiliconFlow 图片后端 — 调用 SiliconFlow API 生成图片
"""

import os
import requests
import base64
from io import BytesIO
from typing import Optional
from PIL import Image

from generators.image_gen import ImageResult
from config import config


class SiliconFlowBackend:
    """SiliconFlow API 文生图"""

    def __init__(self):
        # 从 .env 文件读 key
        self.api_key = self._read_key()
        if not self.api_key:
            raise ValueError("SILICONFLOW_API_KEY 未找到，请先设置")
        self.base_url = "https://api.siliconflow.cn/v1"
        self.model = config.image.siliconflow_model

    def _read_key(self) -> str:
        """从 .env 文件读取 API Key"""
        env_paths = [
            r"E:\Projects\my-workspace\ai-manga-pipeline\.env",
            os.path.expanduser("~/.env"),
        ]
        for env_path in env_paths:
            if os.path.exists(env_path):
                with open(env_path, "r", encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("SILICONFLOW_API_KEY="):
                            return line.split("=", 1)[1].strip().strip('"').strip("'")
        return os.environ.get("SILICONFLOW_API_KEY", "")

    def generate(self, prompt: str, output_path: str) -> Optional[ImageResult]:
        """调用 SiliconFlow API 生成图片（含限流重试）"""
        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": self.model,
            "prompt": prompt,
            "width": config.image.width,
            "height": config.image.height,
            "n": 1,
        }

        max_retries = 5
        for attempt in range(max_retries):
            if attempt > 0:
                wait = 2**attempt  # 指数退避: 2, 4, 8, 16秒
                print(f"  [SiliconFlow] 限流，等待 {wait}s 后重试...")
                import time

                time.sleep(wait)

            resp = requests.post(
                f"{self.base_url}/images/generations",
                headers=headers,
                json=payload,
                timeout=120,
            )

            if resp.status_code == 429:
                continue  # 限流，重试
            elif resp.status_code != 200:
                print(f"  [SiliconFlow] API 错误: {resp.status_code} {resp.text[:200]}")
                return None
            else:
                break  # 成功
        else:
            print(f"  [SiliconFlow] 重试耗尽，放弃: {prompt[:40]}")
            return None

        data = resp.json()
        item = data["data"][0]

        # SiliconFlow 可能返回 b64_json 或 url
        if "b64_json" in item:
            img = Image.open(BytesIO(base64.b64decode(item["b64_json"])))
            img.save(output_path, "PNG")
        elif "url" in item:
            img_resp = requests.get(item["url"], timeout=30)
            img = Image.open(BytesIO(img_resp.content))
            img.save(output_path, "PNG")
        else:
            print(f"  [SiliconFlow] 未知响应格式: {str(item)[:200]}")
            return None

        return ImageResult(path=output_path, prompt=prompt)
