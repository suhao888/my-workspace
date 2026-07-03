"""
配音生成模块 — edge-tts 实现中文配音
"""

import os
import asyncio
from typing import List, Optional
from dataclasses import dataclass

import edge_tts

from config import config


@dataclass
class AudioClipResult:
    text: str
    speaker: str
    path: str
    duration: float  # 秒


class TTSGenerator:
    """用 edge-tts 生成配音（免费，无需 API Key）"""

    def __init__(self):
        self.voice = config.tts.voice
        self.rate = config.tts.rate
        self.volume = config.tts.volume

    def generate(
        self, text: str, speaker: str, output_dir: str, filename: str
    ) -> Optional[AudioClipResult]:
        """生成一句配音"""
        os.makedirs(output_dir, exist_ok=True)
        path = os.path.join(output_dir, filename)

        # 旁白用女声，不同角色用不同音色
        voice = self._pick_voice(speaker)
        communicate = edge_tts.Communicate(
            text, voice, rate=self.rate, volume=self.volume
        )
        # edge-tts 是异步的，sync 包装
        loop = asyncio.new_event_loop()
        asyncio.set_event_loop(loop)
        try:
            loop.run_until_complete(communicate.save(path))
        finally:
            loop.close()

        # 获取音频时长
        duration = self._get_duration(path)
        return AudioClipResult(text=text, speaker=speaker, path=path, duration=duration)

    def batch_generate(
        self, dialogue_list: list, output_dir: str, prefix: str = "audio"
    ) -> List[AudioClipResult]:
        """批量生成配音"""
        results = []
        for i, (speaker, text) in enumerate(dialogue_list):
            filename = f"{prefix}_{i + 1:03d}.mp3"
            result = self.generate(text, speaker, output_dir, filename)
            if result:
                results.append(result)
        return results

    def _pick_voice(self, speaker: str) -> str:
        """根据角色选择音色"""
        if speaker == "旁白":
            return "zh-CN-XiaoxiaoNeural"  # 女声 温柔叙述
        elif speaker in ["男", "霸总", "皇帝", "王爷", "帝王", "师尊", "师父", "父亲"]:
            return "zh-CN-YunxiNeural"  # 男声
        elif speaker in ["系统", "旁白音"]:
            return "zh-CN-XiaoyiNeural"  # 女声 活泼
        else:
            return self.voice  # 默认

    def _get_duration(self, audio_path: str) -> float:
        """用 ffprobe 获取音频时长"""
        import subprocess

        try:
            result = subprocess.run(
                [
                    "ffprobe",
                    "-v",
                    "error",
                    "-show_entries",
                    "format=duration",
                    "-of",
                    "default=noprint_wrappers=1:nokey=1",
                    audio_path,
                ],
                capture_output=True,
                text=True,
                timeout=10,
            )
            return float(result.stdout.strip())
        except Exception:
            return 0.0
