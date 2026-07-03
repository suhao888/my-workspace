from PIL import Image, ImageDraw, ImageFont
import os

out_dir = os.path.dirname(os.path.abspath(__file__))
colors = [(20, 20, 40), (30, 20, 50), (15, 15, 35), (40, 30, 60), (25, 20, 45)]
labels = [
    "审计转行数据概览",
    "CPA报考年龄分布",
    "每日学习投入",
    "薪资增长曲线",
    "转行路径规划",
]

for i, (color, label) in enumerate(zip(colors, labels)):
    img = Image.new("RGB", (1080, 1920), color)
    dr = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 80)
    except:
        font = ImageFont.load_default()
    bbox = dr.textbbox((0, 0), label, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    dr.text(
        ((1080 - tw) // 2, (1920 - th) // 2), label, fill=(255, 255, 255), font=font
    )
    fp = os.path.join(out_dir, f"slide-{i + 1:02d}.png")
    img.save(fp)
    sz = os.path.getsize(fp)
    print(f"slide-{i + 1:02d}.png: {sz} bytes, {img.size}")

print("Done - 5 slides generated")
