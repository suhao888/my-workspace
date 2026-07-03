"""
AI漫剧主管线 — 编排全流程
"""

import os
import time
import json
from pathlib import Path

from config import config
from generators.script_gen import ScriptGenerator
from generators.image_gen import ImageGenerator
from generators.video_gen import VideoGenerator
from generators.tts_gen import TTSGenerator
from generators.assembler import Assembler


class AIPipeline:
    """AI漫剧全自动管线"""

    def __init__(self, output_dir: str = None):
        self.output_dir = output_dir or config.pipeline.output_dir
        self.temp_dir = config.pipeline.temp_dir
        os.makedirs(self.output_dir, exist_ok=True)
        os.makedirs(self.temp_dir, exist_ok=True)

        self.script_gen = ScriptGenerator()
        self.image_gen = ImageGenerator()
        self.video_gen = VideoGenerator()
        self.tts_gen = TTSGenerator()
        self.assembler = Assembler()

    def run(
        self,
        genre: str,
        premise: str,
        num_scenes: int = 15,
        skip_script: bool = False,
        script_path: str = None,
        output_name: str = None,
    ) -> str:
        """
        全自动管线：
        1. 生成剧本 → 2. 出图 → 3. 图转视频 → 4. 配音 → 5. 合成成片
        """
        pipeline_start = time.time()

        # === Step 1: 生成/加载剧本 ===
        if skip_script and script_path:
            print("📖 加载已有剧本...")
            script = self.script_gen.load(script_path)
        else:
            print("📖 用 DeepSeek 生成剧本...")
            script = self.script_gen.generate(genre, premise, num_scenes)
            # 保存剧本
            script_out = os.path.join(self.output_dir, "script.json")
            self.script_gen.save(script, script_out)
            print(f"   标题: {script.title}")
            print(f"   分镜: {len(script.scenes)} 个")
            print(f"   角色: {', '.join(c.name for c in script.characters)}")

        # === Step 2: 生成图片 ===
        print("\n🖼️ 生成场景图片...")
        img_dir = os.path.join(self.temp_dir, "images")
        image_results = {}
        import time as time_module

        for i, scene in enumerate(script.scenes):
            print(
                f"   [{i + 1}/{len(script.scenes)}] {scene.location}: {scene.image_prompt[:50]}..."
            )
            result = self.image_gen.generate(
                scene.image_prompt, img_dir, f"scene_{scene.id:03d}.png"
            )
            if result:
                image_results[scene.id] = result
                scene.image_path = result.path
                # API 限流保护：硅基流动免费/低余额账户 IPM 限制较低
                if config.image.backend == "siliconflow" and i < len(script.scenes) - 1:
                    print(f"  [限流保护] 等待 20s...")
                    time_module.sleep(20)

        # === Step 3: 图转视频 ===
        print("\n🎬 生成视频片段...")
        video_dir = os.path.join(self.temp_dir, "clips")
        video_results = []
        for i, scene in enumerate(script.scenes):
            if scene.id not in image_results:
                print(f"   [{i + 1}/{len(script.scenes)}] 跳过 (无图片)")
                continue
            print(
                f"   [{i + 1}/{len(script.scenes)}] {scene.location} ({scene.duration}s)..."
            )
            result = self.video_gen.generate(
                image_results[scene.id].path,
                scene.duration,
                video_dir,
                f"clip_{scene.id:03d}.mp4",
            )
            if result:
                video_results.append(result)

        # === Step 4: 生成配音 ===
        print("\n🎙️ 生成配音...")
        audio_dir = os.path.join(self.temp_dir, "audio")
        audio_results = []
        for i, scene in enumerate(script.scenes):
            for j, line in enumerate(scene.dialogue):
                print(f"   [{i + 1}.{j + 1}] {line.speaker}: {line.text[:40]}...")
                result = self.tts_gen.generate(
                    line.text,
                    line.speaker,
                    audio_dir,
                    f"audio_{scene.id:03d}_{j + 1:02d}.mp3",
                )
                if result:
                    audio_results.append(result)

        # === Step 5: 合成成片 ===
        print("\n🎞️ 合成最终视频...")
        output_name = output_name or f"{script.title}.mp4"
        output_path = os.path.join(self.output_dir, output_name)
        final_path = self.assembler.assemble(
            video_results,
            audio_results,
            output_path,
            title=script.title,
            add_subtitles=True,
        )

        elapsed = time.time() - pipeline_start
        print(f"\n✅ 完成! 总耗时: {elapsed:.1f}秒")
        print(f"   输出: {final_path}")
        print(f"   大小: {self._file_size(final_path):.1f} MB")

        return final_path

    def _file_size(self, path: str) -> float:
        return os.path.getsize(path) / (1024 * 1024)
