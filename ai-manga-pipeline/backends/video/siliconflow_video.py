"""
SiliconFlow 视频后端 — 调用 Wan2.2 图生视频 API
API 文档: https://api-docs.siliconflow.cn/docs/api/video-submit-post
"""

import os
import time
import requests
import base64
from typing import Optional

from generators.video_gen import VideoClipResult
from config import config


class SiliconFlowVideoBackend:
    """SiliconFlow Wan2.2 图生视频"""

    def __init__(self):
        self.api_key = self._read_key()
        if not self.api_key:
            raise ValueError("SILICONFLOW_API_KEY 未找到")
        self.base_url = "https://api.siliconflow.cn/v1"
        self.model = "Wan-AI/Wan2.2-I2V-A14B"

    def _read_key(self) -> str:
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

    def generate(
        self, image_path: str, duration: float, output_path: str
    ) -> Optional[VideoClipResult]:
        """
        提交图生视频任务，等待完成，下载视频
        """
        # Step 1: 读取图片并转 base64
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }

        # 根据图片尺寸选择分辨率
        from PIL import Image

        with Image.open(image_path) as img:
            w, h = img.size
        # 竖屏用 720x1280，横屏用 1280x720，接近方形用 960x960
        if h > w * 1.2:
            image_size = "720x1280"
        elif w > h * 1.2:
            image_size = "1280x720"
        else:
            image_size = "960x960"

        # Step 2: 提交视频生成任务
        submit_payload = {
            "model": self.model,
            "prompt": "人物自然轻微动作，画面流畅运动",
            "image_size": image_size,
            "image": f"data:image/png;base64,{img_b64}",
        }

        try:
            resp = requests.post(
                f"{self.base_url}/video/submit",
                headers=headers,
                json=submit_payload,
                timeout=30,
            )
        except Exception as e:
            print(f"  [SiliconFlow视频] 提交失败: {e}")
            return None

        if resp.status_code != 200:
            print(f"  [SiliconFlow视频] 提交错误: {resp.status_code} {resp.text[:200]}")
            return None

        request_id = resp.json().get("requestId")
        if not request_id:
            print(f"  [SiliconFlow视频] 未获取到 requestId: {resp.text[:200]}")
            return None

        print(f"  [视频任务已提交] requestId={request_id[:12]}... 等待生成...", end="")

        # Step 3: 轮询任务状态
        max_wait = 600  # 最长等 10 分钟
        poll_interval = 5
        waited = 0
        while waited < max_wait:
            time.sleep(poll_interval)
            waited += poll_interval

            status_resp = requests.post(
                f"{self.base_url}/video/status",
                headers=headers,
                json={"requestId": request_id},
                timeout=15,
            )
            if status_resp.status_code != 200:
                continue

            data = status_resp.json()
            status = data.get("status")

            if status == "Succeed":
                videos = data.get("results", {}).get("videos", [])
                if videos and videos[0].get("url"):
                    video_url = videos[0]["url"]
                    print(f" 完成! 下载中...")
                    # 下载视频
                    dl = requests.get(video_url, timeout=60)
                    with open(output_path, "wb") as f:
                        f.write(dl.content)
                    return VideoClipResult(path=output_path, duration=duration)
                else:
                    print(f"\n  [SiliconFlow视频] 成功但无视频URL: {str(data)[:200]}")
                    return None

            elif status == "Failed":
                reason = data.get("reason", "未知错误")
                print(f"\n  [SiliconFlow视频] 生成失败: {reason}")
                return None

            elif status in ("InQueue", "InProgress"):
                print(".", end="")

            else:
                print(f"\n  [SiliconFlow视频] 未知状态: {status}")
                time.sleep(10)  # 等久一点再试

        print(f"\n  [SiliconFlow视频] 超时 (>{max_wait}s)")
        return None
