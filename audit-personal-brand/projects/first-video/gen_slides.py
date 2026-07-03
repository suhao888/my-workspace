"""Generate 5 intro slides for the first personal brand video."""

from PIL import Image, ImageDraw, ImageFont
import os, math

OUT = os.path.dirname(os.path.abspath(__file__))
W, H = 1080, 1920

# ---- colors ----
BG_TOP = (15, 15, 35)
BG_BOT = (22, 22, 50)
ACCENT = (200, 170, 110)  # muted gold accent line
WHITE = (255, 255, 255)
LIGHT = (200, 200, 210)  # secondary text


def make_gradient(w, h, c1, c2):
    img = Image.new("RGB", (w, h))
    for y in range(h):
        r = int(c1[0] + (c2[0] - c1[0]) * y / h)
        g = int(c1[1] + (c2[1] - c1[1]) * y / h)
        b = int(c1[2] + (c2[2] - c1[2]) * y / h)
        for x in range(w):
            img.putpixel((x, y), (r, g, b))
    return img


def draw_slide(text_lines, font_large, font_small, accent_offset=0):
    img = make_gradient(W, H, BG_TOP, BG_BOT)
    dr = ImageDraw.Draw(img)

    # subtle accent line (horizontal)
    line_y = H // 2 - 200 + accent_offset
    for x in range(W // 2 - 200, W // 2 + 200):
        alpha = 1.0 - abs(x - W // 2) / 200
        c = tuple(int(v * alpha) for v in ACCENT)
        dr.point((x, line_y), fill=c)

    # draw text lines
    total_h = 0
    line_hs = []
    for i, line in enumerate(text_lines):
        font = (
            font_large
            if isinstance(font_large, ImageFont.FreeTypeFont) or i == 0
            else font_small
        )
        bbox = dr.textbbox((0, 0), line, font=font)
        lh = bbox[3] - bbox[1]
        total_h += lh + 20
        line_hs.append(lh)

    start_y = (H - total_h) // 2
    y = start_y
    for i, line in enumerate(text_lines):
        font = font_small if (len(text_lines) > 2 and i > 0) else font_large
        bbox = dr.textbbox((0, 0), line, font=font)
        lh = bbox[3] - bbox[1]
        tw = bbox[2] - bbox[0]
        x = (W - tw) // 2
        fill = WHITE if i == 0 else LIGHT
        dr.text((x, y), line, fill=fill, font=font)
        y += lh + 20

    # accent line below text
    line_y2 = y + 40
    for x in range(W // 2 - 150, W // 2 + 150):
        alpha = 1.0 - abs(x - W // 2) / 150
        c = tuple(int(v * alpha) for v in ACCENT)
        dr.point((x, line_y2), fill=c)

    # very subtle noise overlay (barely visible grain)
    import random

    rand = random.Random(42)
    for _ in range(800):
        px = rand.randint(0, W - 1)
        py = rand.randint(0, H - 1)
        v = rand.randint(0, 5)
        r, g, b = img.getpixel((px, py))
        r = min(255, max(0, r + v))
        g = min(255, max(0, g + v))
        b = min(255, max(0, b + v))
        img.putpixel((px, py), (r, g, b))

    return img


# ---- fonts ----
try:
    f_large = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 96)  # 微软雅黑 Bold
    f_small = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 72)  # 微软雅黑 Regular
except:
    f_large = f_small = ImageFont.load_default()

# ---- slides data ----
slides = [
    ["做了7年施工企业财务，", "2年会计师事务所审计。"],
    ["懂一些财务、税务、审计", "方面的知识。"],
    ["谈不上资深，", "就是干得久了，", "踩过的坑比较多。"],
    ["以后想在这，", "交一些朋友，", "互相探讨。"],
    ["很高兴认识你。"],
]

# ---- generate ----
for idx, lines in enumerate(slides):
    offset = (idx - 2) * 30  # slight vertical variation per slide
    img = draw_slide(lines, f_large, f_small, accent_offset=offset)
    fp = os.path.join(OUT, f"slide-{idx + 1:02d}.png")
    img.save(fp)
    sz = os.path.getsize(fp)
    print(f"{fp}: {sz / 1024:.1f} KB")

print("Done - 5 slides generated")
