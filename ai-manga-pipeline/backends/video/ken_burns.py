"""
Ken Burns 视频后端 — 用 FFmpeg 给静态图片做推拉摇移
零成本，无需 API
"""

import subprocess
import os
from typing import Optional

from generators.video_gen import VideoClipResult
from config import config


class KenBurnsBackend:
    """用 FFmpeg 实现 Ken Burns 效果：图片缓慢缩放/平移"""

    def __init__(self):
        self.fps = config.video.fps

    def generate(
        self, image_path: str, duration: float, output_path: str
    ) -> Optional[VideoClipResult]:
        """
        用 FFmpeg 对静态图片做 Ken Burns（镜头缓慢放大）
        ffmpeg -loop 1 -i input.png -vf "zoompan=z='min(zoom+0.002,1.1)':d=120" -c:v libx264 -t 5 -pix_fmt yuv420p output.mp4
        """
        # 计算 zoom 参数：duration 秒 × fps = 总帧数
        total_frames = int(duration * self.fps)
        # 每帧放大倍率增量
        zoom_step = config.video.zoom_ratio

        zoom_expr = (
            f"zoompan=z='min(zoom+{zoom_step},1.15)':d={total_frames}:fps={self.fps}"
        )

        cmd = [
            "ffmpeg",
            "-y",
            "-loop",
            "1",
            "-i",
            image_path,
            "-vf",
            zoom_expr,
            "-c:v",
            "libx264",
            "-t",
            str(duration),
            "-pix_fmt",
            "yuv420p",
            "-preset",
            "ultrafast",
            output_path,
        ]

        try:
            subprocess.run(cmd, capture_output=True, text=True, timeout=300)
            if os.path.exists(output_path):
                return VideoClipResult(path=output_path, duration=duration)
        except subprocess.TimeoutExpired:
            print(f"  [KenBurns] 超时: {image_path}")
        except Exception as e:
            print(f"  [KenBurns] 错误: {e}")

        return None
