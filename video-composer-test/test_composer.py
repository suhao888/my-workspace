"""Test video_composer with generated audit slides."""

import sys, os

sys.path.insert(0, os.path.expanduser("~/.claude/skills/video-composer"))

from video_composer import VideoComposer
from pathlib import Path

vc = VideoComposer()
out_dir = os.path.dirname(os.path.abspath(__file__))

# 1. Single image → video
print("\n=== Test 1: Image to static video ===")
slide1 = os.path.join(out_dir, "slide-01.png")
clip1 = vc.create_static_video(
    slide1,
    duration=8,
    output_path=os.path.join(out_dir, "clip-01.mp4"),
    resolution="1080x1920",
    fps=30,
)
if clip1 and os.path.exists(clip1):
    sz = os.path.getsize(clip1)
    print(f"  PASS: {clip1} ({sz / 1024:.1f} KB)")
else:
    print(f"  FAIL")

# 2. Image → video with different duration
print("\n=== Test 2: Second slide ===")
slide2 = os.path.join(out_dir, "slide-02.png")
clip2 = vc.create_static_video(
    slide2,
    duration=5,
    output_path=os.path.join(out_dir, "clip-02.mp4"),
    resolution="1080x1920",
    fps=30,
)
if clip2 and os.path.exists(clip2):
    sz = os.path.getsize(clip2)
    print(f"  PASS: {clip2} ({sz / 1024:.1f} KB)")
else:
    print(f"  FAIL")

# 3. Concat two clips (normalized)
print("\n=== Test 3: Concat two clips (normalized) ===")
concat_out = os.path.join(out_dir, "concat-test.mp4")
ok = vc.concat_videos(
    [clip1, clip2],
    concat_out,
    normalize_params=True,
    target_resolution="1080x1920",
    target_fps=30,
)
if ok and os.path.exists(concat_out):
    sz = os.path.getsize(concat_out)
    import subprocess

    r = subprocess.run(
        [
            "ffprobe",
            "-v",
            "error",
            "-show_entries",
            "format=duration",
            "-of",
            "csv=p=0",
            concat_out,
        ],
        capture_output=True,
        text=True,
    )
    dur = r.stdout.strip() if r.returncode == 0 else "?"
    print(f"  PASS: {concat_out} ({sz / 1024:.1f} KB, {dur}s)")
else:
    print(f"  FAIL")

# 4. Concat fast mode (no re-encode)
print("\n=== Test 4: Concat fast mode ===")
concat_fast = os.path.join(out_dir, "concat-fast.mp4")
ok = vc.concat_videos([clip1, clip2], concat_fast, normalize_params=False)
if ok and os.path.exists(concat_fast):
    sz = os.path.getsize(concat_fast)
    print(f"  PASS: {concat_fast} ({sz / 1024:.1f} KB)")
else:
    print(f"  FAIL")

print("\n=== All tests complete ===")
