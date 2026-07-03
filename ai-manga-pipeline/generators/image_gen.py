"""
图片生成模块 — 统一接口 + 多后端路由
"""

import os
from typing import Optional
from config import config


class ImageResult:
    """图片生成结果"""

    def __init__(self, path: str, prompt: str):
        self.path = path
        self.prompt = prompt


class ImageGenerator:
    """图片生成器，按配置路由到不同后端"""

    def __init__(self):
        self.backend_name = config.image.backend
        self._backend = self._load_backend()

    def _load_backend(self):
        name = self.backend_name
        if name == "demo":
            from backends.image.demo import DemoImageBackend

            return DemoImageBackend()
        elif name == "siliconflow":
            from backends.image.siliconflow import SiliconFlowBackend

            return SiliconFlowBackend()
        elif name == "midjourney":
            from backends.image.midjourney import MidjourneyBackend

            return MidjourneyBackend()
        else:
            raise ValueError(
                f"未知图片后端: {name}，可选: demo, siliconflow, midjourney"
            )

    def generate(
        self, prompt: str, output_dir: str, filename: str
    ) -> Optional[ImageResult]:
        """生成一张图片"""
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)
        return self._backend.generate(prompt, path)

    def batch_generate(
        self, prompts: list, output_dir: str, prefix: str = "scene"
    ) -> list:
        """批量生成图片"""
        results = []
        for i, prompt in enumerate(prompts):
            filename = f"{prefix}_{i + 1:03d}.png"
            result = self.generate(prompt, output_dir, filename)
            results.append(result)
        return results
