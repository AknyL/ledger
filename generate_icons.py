# -*- coding: utf-8 -*-
import os
from PIL import Image, ImageDraw

BASE = os.path.dirname(os.path.abspath(__file__))
ICON_DIR = os.path.join(BASE, "icons")
os.makedirs(ICON_DIR, exist_ok=True)

def lerp(a, b, t):
    return tuple(int(a[i] + (b[i] - a[i]) * t) for i in range(3))

def make(size):
    S = 1024  # 超采样
    img = Image.new("RGB", (S, S))
    d = ImageDraw.Draw(img)
    top = (79, 110, 247)     # #4f6ef7
    bottom = (123, 92, 240)  # #7b5cf0
    for y in range(S):
        d.line([(0, y), (S, y)], fill=lerp(top, bottom, y / S))
    # 三根上升的圆角柱
    bar_w = int(S * 0.135)
    radius = int(S * 0.05)
    gap = int(S * 0.075)
    total_w = 3 * bar_w + 2 * gap
    x0 = (S - total_w) // 2
    baseline = int(S * 0.68)
    heights = [int(S * 0.20), int(S * 0.32), int(S * 0.44)]
    for i, h in enumerate(heights):
        x = x0 + i * (bar_w + gap)
        y = baseline - h
        d.rounded_rectangle([x, y, x + bar_w, baseline], radius=radius, fill=(255, 255, 255))
    # 上升趋势折线 + 箭头（半透明白）
    line = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    dl = ImageDraw.Draw(line)
    pts = [(int(S*0.30), int(S*0.46)), (int(S*0.48), int(S*0.34)), (int(S*0.72), int(S*0.24))]
    dl.line(pts, fill=(255, 255, 255, 215), width=int(S*0.028), joint="curve")
    ax, ay = pts[-1]
    px, py = pts[-2]
    import math
    ang = math.atan2(ay - py, ax - px)
    L = int(S * 0.06)
    spread = 0.55
    tip = (ax + L * math.cos(ang), ay + L * math.sin(ang))
    left = (ax - L * math.cos(ang - spread), ay - L * math.sin(ang - spread))
    right = (ax - L * math.cos(ang + spread), ay - L * math.sin(ang + spread))
    dl.polygon([tip, left, right], fill=(255, 255, 255, 215))
    img = Image.alpha_composite(img.convert("RGBA"), line).convert("RGB")
    # 降采样到目标尺寸
    return img.resize((size, size), Image.LANCZOS)

for name, s in [("icon-512", 512), ("icon-192", 192), ("apple-touch-icon", 180)]:
    make(s).save(os.path.join(ICON_DIR, name + ".png"))
    print("saved", name + ".png", s)
print("done")
