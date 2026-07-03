"""
AI漫剧管线 — 配置管理
"""

import os
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class LLMConfig:
    """大模型配置"""

    api_key: str = field(default_factory=lambda: os.environ.get("DEEPSEEK_API_KEY", ""))
    base_url: str = "https://api.deepseek.com"
    model: str = "deepseek-chat"
    temperature: float = 0.8
    max_tokens: int = 4096


@dataclass
class ImageConfig:
    """图片生成配置"""

    backend: str = "demo"  # demo | siliconflow | midjourney
    # SiliconFlow
    siliconflow_api_key: str = ""
    siliconflow_model: str = "Kwai-Kolors/Kolors"
    # Midjourney
    midjourney_api_key: str = ""
    midjourney_api_base: str = ""
    # 通用
    width: int = 1080
    height: int = 1920  # 竖屏 9:16


@dataclass
class VideoConfig:
    """视频生成配置"""

    backend: str = "ken_burns"  # ken_burns | siliconflow | runway | kling
    fps: int = 24
    # Ken Burns
    zoom_ratio: float = 0.02
    # SiliconFlow (Wan2.2)
    siliconflow_api_key: str = ""
    # Runway
    runway_api_key: str = ""
    # 可灵
    kling_api_key: str = ""


@dataclass
class TTSConfig:
    """配音配置"""

    voice: str = "zh-CN-XiaoxiaoNeural"  # edge-tts 中文女声
    rate: str = "+0%"
    volume: str = "+0%"


@dataclass
class PipelineConfig:
    """主管线配置"""

    output_dir: str = "output"
    temp_dir: str = "output/temp"
    subtitle_font: str = "C:/Windows/Fonts/msyh.ttc"  # 微软雅黑
    subtitle_font_size: int = 36


@dataclass
class AppConfig:
    """全局配置"""

    llm: LLMConfig = field(default_factory=LLMConfig)
    image: ImageConfig = field(default_factory=ImageConfig)
    video: VideoConfig = field(default_factory=VideoConfig)
    tts: TTSConfig = field(default_factory=TTSConfig)
    pipeline: PipelineConfig = field(default_factory=PipelineConfig)


# 全局单例
config = AppConfig()
