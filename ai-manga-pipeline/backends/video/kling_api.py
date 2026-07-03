"""
可灵 Kling 视频后端 — 调用 Kling API 图生视频
文档: https://docs.klingai.com
"""

import os
import time
import requests
import base64
import json
import hashlib
import hmac
from typing import Optional
from urllib.parse import urlencode

from generators.video_gen import VideoClipResult
from config import config


class KlingBackend:
    """可灵 AI 图生视频"""

    def __init__(self):
        self.api_key = self._read_key()
        if not self.api_key:
            raise ValueError("KLING_API_KEY 未设置")
        self.base_url = "https://api.klingai.com"

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
                        if line.startswith("KLING_API_KEY="):
                            return line.split("=", 1)[1].strip().strip('"').strip("'")
        return os.environ.get("KLING_API_KEY", "")

    def _sign(self, method: str, path: str, params: dict, body: str = "") -> dict:
        """可灵 HMAC-SHA256 签名认证"""
        ak, sk = self.api_key, ""
        if "|" in self.api_key:
            ak, sk = self.api_key.split("|", 1)
        else:
            # api-key-kling-xxx 格式：直接用 Bearer
            return {"Authorization": f"Bearer {self.api_key}"}

        now = str(int(time.time()))
        # 构建签名串
        header_str = (
            f"host:{self.base_url.replace('https://', '')}\nx-api-timestamp:{now}\n"
        )
        sorted_params = sorted(params.items(), key=lambda x: x[0])
        param_str = urlencode(sorted_params) if params else ""
        sign_str = f"{method}\n{path}\n{param_str}\n{header_str}\n{body}"

        # HMAC-SHA256
        signature = hmac.new(sk.encode(), sign_str.encode(), hashlib.sha256).hexdigest()

        return {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {ak}",
            "X-Api-Timestamp": now,
            "X-Api-Signature": signature,
        }

    def generate(
        self, image_path: str, duration: float, output_path: str
    ) -> Optional[VideoClipResult]:
        """提交图生视频任务，轮询至完成，下载视频"""
        # 读取图片转 base64
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        headers = self._sign("POST", "/v1/images/image2video", {})
        headers["Content-Type"] = "application/json"

        # 可灵图生视频参数
        payload = {
            "model": "kling-v1.5",
            "image": f"data:image/png;base64,{img_b64}",
            "prompt": "人物自然轻微动作，画面流畅",
            "cfg_scale": 0.5,
        }

        # Step 1: 提交任务
        try:
            resp = requests.post(
                f"{self.base_url}/v1/images/image2video",
                headers=headers,
                json=payload,
                timeout=30,
            )
        except Exception as e:
            print(f"  [可灵] 提交失败: {e}")
            return None

        if resp.status_code != 200:
            # 可能 Bearer 不行，试试另一种 auth
            if resp.status_code == 401 or resp.status_code == 403:
                print(f"  [可灵] 认证失败({resp.status_code})，尝试备用方案...")
                return self._generate_bearer(image_path, duration, output_path)
            print(f"  [可灵] 提交错误: {resp.status_code} {resp.text[:200]}")
            return None

        result = resp.json()
        if result.get("code") != 0:
            print(f"  [可灵] API 错误: {result.get('message', str(result)[:200])}")
            return None

        task_id = result.get("data", {}).get("task_id")
        if not task_id:
            print(f"  [可灵] 无 task_id: {str(result)[:200]}")
            return None

        print(f"  [可灵任务已提交] task_id={task_id[:12]}... 等待生成...", end="")

        # Step 2: 轮询
        max_wait = 600
        poll_interval = 5
        waited = 0
        while waited < max_wait:
            time.sleep(poll_interval)
            waited += poll_interval

            poll_headers = self._sign("GET", f"/v1/images/image2video/{task_id}", {})
            poll_resp = requests.get(
                f"{self.base_url}/v1/images/image2video/{task_id}",
                headers=poll_headers,
                timeout=15,
            )

            if poll_resp.status_code != 200:
                continue

            data = poll_resp.json()
            if data.get("code") != 0:
                continue

            task_data = data.get("data", {})
            status = task_data.get("task_status", task_data.get("status", ""))

            if status == "succeed":
                result_data = task_data.get("task_result", {})
                videos = result_data.get("videos", [])
                if not videos:
                    videos = result_data.get("images", [])
                if videos and videos[0].get("url"):
                    url = videos[0]["url"]
                    print(f" 完成! 下载中...")
                    dl = requests.get(url, timeout=120)
                    with open(output_path, "wb") as f:
                        f.write(dl.content)
                    return VideoClipResult(path=output_path, duration=duration)
                print(f"\n  [可灵] 成功但无视频: {str(data)[:200]}")
                return None

            elif status in ("failed", "error"):
                print(f"\n  [可灵] 失败: {task_data.get('message', '')}")
                return None

            elif status in ("submitted", "processing", "in_queue", "in_progress"):
                print(".", end="")
            else:
                print(f"[{status}]", end="")

        print(f"\n  [可灵] 超时")
        return None

    def _generate_bearer(
        self, image_path: str, duration: float, output_path: str
    ) -> Optional[VideoClipResult]:
        """备用方案：Bearer token 方式"""
        with open(image_path, "rb") as f:
            img_b64 = base64.b64encode(f.read()).decode()

        headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
        }
        payload = {
            "model": "kling-v1.5",
            "image": f"data:image/png;base64,{img_b64}",
            "prompt": "人物自然轻微动作",
            "cfg_scale": 0.5,
        }

        try:
            resp = requests.post(
                f"{self.base_url}/v1/images/image2video",
                headers=headers,
                json=payload,
                timeout=30,
            )
        except Exception as e:
            print(f"  [可灵-Bearer] 提交失败: {e}")
            return None

        print(f"  [可灵-Bearer] 提交结果: {resp.status_code} {resp.text[:200]}")
        return None
