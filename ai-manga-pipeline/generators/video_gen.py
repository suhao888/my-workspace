"""
视频生成模块 — 图片 → 视频片段
"""

import os
from typing import Optional
from config import config


class VideoClipResult:
    """视频片段结果"""

    def __init__(self, path: str, duration: float):
        self.path = path
        self.duration = duration


class VideoGenerator:
    """视频生成器，按配置路由到不同后端"""

    def __init__(self):
        self.backend_name = config.video.backend
        self._backend = self._load_backend()

    def _load_backend(self):
        name = self.backend_name
        if name == "ken_burns":
            from backends.video.ken_burns import KenBurnsBackend

            return KenBurnsBackend()
        elif name == "siliconflow":
            from backends.video.siliconflow_video import SiliconFlowVideoBackend

            return SiliconFlowVideoBackend()
        elif name == "runway":
            from backends.video.runway_api import RunwayBackend

            return RunwayBackend()
        elif name == "kling":
            from backends.video.kling_api import KlingBackend

            return KlingBackend()
        else:
            raise ValueError(
                f"未知视频后端: {name}，可选: ken_burns, siliconflow, runway, kling"
            )

    def generate(
        self, image_path: str, duration: float, output_dir: str, filename: str
    ) -> Optional[VideoClipResult]:
        """将一张图片转为视频片段"""
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        return self._backend.generate(image_path, duration, path)

    def batch_generate(
        self, image_paths: list, durations: list, output_dir: str, prefix: str = "clip"
    ) -> list:
        """批量生成视频片段"""
        results = []
        for i, (img_path, dur) in enumerate(zip(image_paths, durations)):
            filename = f"{prefix}_{i + 1:03d}.mp4"
            result = self.generate(img_path, dur, output_dir, filename)
            if result:
                results.append(result)
        return results
