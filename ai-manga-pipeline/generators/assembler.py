"""
视频组装模块 — 将视频片段 + 音频 + 字幕合成最终成片
"""

import os
import subprocess
import json
from typing import List

from config import config
from generators.video_gen import VideoClipResult


class Assembler:
    """用 FFmpeg 合成最终视频"""

    def __init__(self):
        self.font = config.pipeline.subtitle_font
        self.font_size = config.pipeline.subtitle_font_size

    def assemble(
        self,
        video_clips: List[VideoClipResult],
        audio_clips: list,
        output_path: str,
        title: str = "AI漫剧",
        add_subtitles: bool = True,
    ) -> str:
        """
        合成最终视频
        video_clips: 视频片段列表
        audio_clips: 音频片段列表 (每个有 .path 和 .text, .speaker 属性)
        output_path: 输出文件路径
        """
        temp_dir = config.pipeline.temp_dir
        os.makedirs(temp_dir, exist_ok=True)
        os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)

        # Step 1: 拼接所有视频片段为一个完整视频
        concat_video = os.path.join(temp_dir, "concat_video.mp4")
        self._concat_videos(video_clips, concat_video)

        # Step 2: 拼接所有音频片段为一个完整音轨
        concat_audio = os.path.join(temp_dir, "concat_audio.mp3")
        self._concat_audios(audio_clips, concat_audio)

        # Step 3: 合成视频+音频
        mixed = os.path.join(temp_dir, "mixed.mp4")
        self._mix_audio_video(concat_video, concat_audio, mixed)

        # Step 4: (可选) 压入字幕
        if add_subtitles and audio_clips:
            # 生成字幕文件（ASS 格式，带说话人标注）
            subtitle_file = os.path.join(temp_dir, "subtitles.ass")
            self._generate_subtitles(audio_clips, subtitle_file)
            final = self._burn_subtitles(mixed, subtitle_file, output_path)
        else:
            final = mixed
            if output_path != mixed:
                import shutil

                shutil.copy2(mixed, output_path)

        return output_path

    def _concat_videos(self, clips: List[VideoClipResult], output: str):
        """用 concat demuxer 拼接视频片段（不重新编码）"""
        # 确保所有片段编码一致（都需要 yuv420p）
        list_file = os.path.join(config.pipeline.temp_dir, "video_list.txt")
        with open(list_file, "w", encoding="utf-8") as f:
            for clip in clips:
                f.write(f"file '{os.path.abspath(clip.path)}'\n")

        cmd = [
            "ffmpeg",
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            list_file,
            "-c",
            "copy",
            output,
        ]
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            encoding="utf-8",
            errors="replace",
        )

    def _concat_audios(self, audio_clips: list, output: str):
        """用 concat demuxer 拼接音频"""
        list_file = os.path.join(config.pipeline.temp_dir, "audio_list.txt")
        with open(list_file, "w", encoding="utf-8") as f:
            for clip in audio_clips:
                apath = os.path.abspath(clip.path)
                # 确保 mp3 文件存在
                if os.path.exists(apath):
                    f.write(f"file '{apath}'\n")

        if not os.path.exists(list_file):
            # 无音频，生成静音轨道
            cmd = [
                "ffmpeg",
                "-y",
                "-f",
                "lavfi",
                "-i",
                "anullsrc=r=44100:cl=mono",
                "-t",
                "30",
                output,
            ]
            subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=60,
                encoding="utf-8",
                errors="replace",
            )
            return

        cmd = [
            "ffmpeg",
            "-y",
            "-f",
            "concat",
            "-safe",
            "0",
            "-i",
            list_file,
            "-c",
            "copy",
            output,
        ]
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            encoding="utf-8",
            errors="replace",
        )

    def _mix_audio_video(self, video_path: str, audio_path: str, output: str):
        """将视频和音频合成"""
        cmd = [
            "ffmpeg",
            "-y",
            "-i",
            video_path,
            "-i",
            audio_path,
            "-c:v",
            "copy",
            "-c:a",
            "aac",
            "-map",
            "0:v:0",
            "-map",
            "1:a:0",
            "-shortest",
            output,
        ]
        subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=300,
            encoding="utf-8",
            errors="replace",
        )

    def _generate_subtitles(self, audio_clips: list, output_path: str):
        """生成 ASS 字幕文件，包含说话人标注"""
        lines = [
            "[Script Info]",
            "Title: AI漫剧字幕",
            "ScriptType: v4.00+",
            "PlayResX: 1080",
            "PlayResY: 1920",
            "",
            "[V4+ Styles]",
            f"Format: Name, Fontname, Fontsize, PrimaryColour, SecondaryColour, OutlineColour, BackColour, Bold, Italic, Underline, StrikeOut, ScaleX, ScaleY, Spacing, Angle, BorderStyle, Outline, Shadow, Alignment, MarginL, MarginR, MarginV, Encoding",
            f"Style: Default,{self._font_name()},{self.font_size},&H00FFFFFF,&H000000FF,&H80000000,&H00000000,0,0,0,0,100,100,0,0,1,2,1,2,50,50,60,1",
            f"Style: Speaker,{self._font_name()},{self.font_size + 4},&H00FFD700,&H000000FF,&H80000000,&H00000000,1,0,0,0,100,100,0,0,1,2,1,2,50,50,120,1",
            "",
            "[Events]",
            "Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text",
        ]

        current_time = 0.0
        for clip in audio_clips:
            start = current_time
            end = current_time + clip.duration

            start_str = self._time_str(start)
            end_str = self._time_str(end)

            # 说话人标签（金色）
            speaker_text = clip.speaker.replace("旁白", "[旁白]")
            lines.append(
                f"Dialogue: 1,{start_str},{end_str},Speaker,,0,0,0,,{speaker_text}"
            )
            # 台词正文（白色）
            lines.append(
                f"Dialogue: 0,{start_str},{end_str},Default,,0,0,0,,{clip.text}"
            )

            current_time = end

        with open(output_path, "w", encoding="utf-8") as f:
            f.write("\n".join(lines))

    def _burn_subtitles(
        self, video_path: str, subtitle_path: str, output_path: str
    ) -> str:
        """将字幕嵌入视频（用 subtitles 滤镜，ass 在部分 ffmpeg 版本上不兼容）"""
        # Windows 路径反斜杠需转义，冒号也要处理
        safe_path = subtitle_path.replace("\\", "/").replace(":", "\\\\:")
        cmd = [
            "ffmpeg",
            "-y",
            "-i",
            video_path,
            "-vf",
            f"subtitles={safe_path}",
            "-c:a",
            "copy",
            "-preset",
            "ultrafast",
            output_path,
        ]
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=600,
            encoding="utf-8",
            errors="replace",
        )
        if result.returncode != 0 or not os.path.exists(output_path):
            print("  [字幕] 烧录失败，回退到无字幕版本")
            import shutil

            if video_path != output_path:
                shutil.copy2(video_path, output_path)
            else:
                return video_path
        return output_path

    def _font_name(self) -> str:
        """返回字体名称（用 Windows 标准字体名，无需文件路径）"""
        # msyh.ttc → "Microsoft YaHei"
        font_map = {
            "msyh": "Microsoft YaHei",
            "msyhbd": "Microsoft YaHei Bold",
            "msyhl": "Microsoft YaHei Light",
            "simhei": "SimHei",
            "yahei": "Microsoft YaHei",
        }
        base = os.path.splitext(os.path.basename(self.font))[0]
        return font_map.get(base.lower(), base)

    def _time_str(self, seconds: float) -> str:
        """seconds → ASS 时间格式 H:MM:SS.cc"""
        h = int(seconds // 3600)
        m = int((seconds % 3600) // 60)
        s = int(seconds % 60)
        cs = int((seconds - int(seconds)) * 100)
        return f"{h}:{m:02d}:{s:02d}.{cs:02d}"
