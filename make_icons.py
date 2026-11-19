"""Generates the five PNG icons used by the add-in (needs: pip install pillow)."""
import os
from PIL import Image, ImageDraw

os.makedirs("docs/assets", exist_ok=True)


def make_icon(size):
    scale = 4
    s = size * scale
    img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle([0, 0, s - 1, s - 1], radius=int(s * 0.22), fill=(0, 120, 212, 255))
    m = int(s * 0.2)
    top, bottom = int(s * 0.3), int(s * 0.7)
    width = max(2, int(s * 0.05))
    d.rectangle([m, top, s - m, bottom], outline=(255, 255, 255, 255), width=width)
    d.line([m, top, s // 2, int(s * 0.55), s - m, top], fill=(255, 255, 255, 255), width=width)
    return img.resize((size, size), Image.LANCZOS)


for size in (16, 32, 64, 80, 128):
    make_icon(size).save(f"docs/assets/icon-{size}.png")
print("Icons written to docs/assets/")
