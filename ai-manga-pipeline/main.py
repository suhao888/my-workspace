#!/usr/bin/env python
"""
AI漫剧全自动生成管线 — CLI 入口

用法：
  # 完整流程（生成剧本→出图→视频→配音→合成）
  python main.py --genre "穿越" --premise "女医生穿越成古代王妃"

  # 使用已有剧本（跳过剧本生成）
  python main.py --script output/script.json

  # 只生成剧本，不做后面的步骤
  python main.py --genre "末世" --premise "..." --script-only

  # 指定输出位置
  python main.py --genre "重生" --premise "..." --output-dir ./my_work

  # Demo 模式（默认）：PIL 占位图 + Ken Burns 效果，零 API 成本跑通流程
  # 生产模式：设置图片后端和视频后端，详见 config.py
"""

import argparse
import sys
import os

# 确保项目根目录在路径中
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from pipeline import AIPipeline
from generators.script_gen import ScriptGenerator
from config import config


def main():
    parser = argparse.ArgumentParser(
        description="AI漫剧全自动生成管线",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  python main.py --genre "穿越" --premise "女医生穿越成古代王妃，用现代医术逆袭"
  python main.py --script output/script.json
  python main.py --genre "末世" --premise "外卖员在丧尸末世靠外卖箱生存" --script-only
  python main.py --genre "重生" --premise "..." --scenes 10 --output-dir ./my_episode
        """,
    )

    parser.add_argument(
        "--genre", "-g", default="穿越", help="题材：穿越/重生/霸总/末世/兽世/娱乐圈等"
    )
    parser.add_argument("--premise", "-p", default="", help="故事梗概/核心设定")
    parser.add_argument(
        "--scenes", "-n", type=int, default=15, help="分镜数量（默认 15）"
    )
    parser.add_argument("--script", "-s", help="已有剧本 JSON 路径，跳过剧本生成")
    parser.add_argument(
        "--script-only", action="store_true", help="只生成剧本，不执行后续步骤"
    )
    parser.add_argument("--output-dir", "-o", help="输出目录（默认 output/）")
    parser.add_argument("--output-name", help="输出文件名（默认 剧本标题.mp4）")
    parser.add_argument(
        "--image-backend",
        choices=["demo", "siliconflow", "midjourney"],
        help="图片后端（默认 demo）",
    )
    parser.add_argument(
        "--video-backend",
        choices=["ken_burns", "siliconflow", "runway", "kling"],
        help="视频后端（默认 ken_burns）",
    )

    args = parser.parse_args()

    # 应用命令行配置覆盖
    if args.image_backend:
        config.image.backend = args.image_backend
    if args.video_backend:
        config.video.backend = args.video_backend

    # 只生成剧本
    if args.script_only:
        if not args.premise:
            parser.error("--script-only 需要 --premise 参数")
        gen = ScriptGenerator()
        print("📖 生成剧本中...")
        script = gen.generate(args.genre, args.premise, args.scenes)
        out_dir = args.output_dir or config.pipeline.output_dir
        os.makedirs(out_dir, exist_ok=True)
        path = os.path.join(out_dir, "script.json")
        gen.save(script, path)
        print(f"   标题: {script.title}")
        print(f"   分镜: {len(script.scenes)} 个")
        print(f"   角色: {', '.join(c.name for c in script.characters)}")
        print(f"   已保存: {path}")
        return

    # 完整管线
    pipeline = AIPipeline(output_dir=args.output_dir)

    # 交互式输入梗概（如果没提供）
    premise = args.premise
    if not premise and not args.script:
        print("请输入故事梗概（支持多行，空行结束）：")
        lines = []
        while True:
            line = input()
            if line.strip() == "":
                break
            lines.append(line)
        premise = "\n".join(lines)

    print(f"\n{'=' * 60}")
    print(f"  题材：{args.genre}")
    print(f"  梗概：{premise[:100]}{'...' if len(premise) > 100 else ''}")
    print(f"  分镜：{args.scenes}")
    print(f"  图片后端：{config.image.backend}")
    print(f"  视频后端：{config.video.backend}")
    print(f"{'=' * 60}\n")

    try:
        final_path = pipeline.run(
            genre=args.genre,
            premise=premise,
            num_scenes=args.scenes,
            skip_script=bool(args.script),
            script_path=args.script,
            output_name=args.output_name,
        )
        print(f"\n🎉 漫剧已生成: {final_path}")
    except KeyboardInterrupt:
        print("\n\n已取消")
        sys.exit(1)
    except Exception as e:
        print(f"\n❌ 错误: {e}")
        import traceback

        traceback.print_exc()
        sys.exit(1)


if __name__ == "__main__":
    main()
