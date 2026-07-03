"""
Demo 图片后端 — 用 PIL 生成占位图（零成本，适合流程测试）
生成渐变背景 + 文字标注的图片，代替真实 AI 出图
"""

import os
from typing import Optional
from PIL import Image, ImageDraw, ImageFont

from generators.image_gen import ImageResult


class DemoImageBackend:
    """PIL 占位图生成器"""

    def __init__(self):
        self.width = 1080
        self.height = 1920

    def generate(self, prompt: str, output_path: str) -> Optional[ImageResult]:
        """生成一张带文字标注的渐变色占位图"""
        img = Image.new("RGB", (self.width, self.height), self._gradient(prompt))
        draw = ImageDraw.Draw(img)

        # 尝试加载字体，失败则用默认
        font_large = None
        font_small = None
        font_paths = [
            "C:/Windows/Fonts/msyh.ttc",
            "C:/Windows/Fonts/simhei.ttf",
            "C:/Windows/Fonts/yahei.ttf",
        ]
        for fp in font_paths:
            if os.path.exists(fp):
                try:
                    font_large = ImageFont.truetype(fp, 40)
                    font_small = ImageFont.truetype(fp, 24)
                except Exception:
                    pass
                break

        # 画面中央显示 "AI漫剧 demo" 和水印
        center_y = self.height // 2
        draw.text(
            (self.width // 2 - 150, center_y - 100),
            "🎬 AI 漫剧",
            fill=(255, 255, 255),
            font=font_large,
        )

        # 换行显示 prompt 摘要
        max_chars_per_line = 30
        lines = []
        current = prompt
        while len(current) > max_chars_per_line:
            cut = current[:max_chars_per_line]
            lines.append(cut)
            current = current[max_chars_per_line:]
        lines.append(current)

        y_offset = center_y
        for line in lines:
            draw.text(
                (self.width // 2 - len(line) * 12, y_offset),
                line,
                fill=(200, 200, 200),
                font=font_small,
            )
            y_offset += 35

        # 底部标注
        draw.text(
            (50, self.height - 100),
            f"[Demo模式] {prompt[:60]}...",
            fill=(150, 150, 150),
            font=font_small,
        )

        img.save(output_path, "PNG")
        return ImageResult(path=output_path, prompt=prompt)

    def _gradient(self, seed: str) -> tuple:
        """根据 prompt 生成不同颜色"""
        hue = (hash(seed) % 360) / 360.0
        import colorsys

        r, g, b = colorsys.hsv_to_rgb(hue, 0.6, 0.3)
        return (int(r * 255), int(g * 255), int(b * 255))
