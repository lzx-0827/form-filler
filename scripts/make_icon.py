# -*- coding: utf-8 -*-
"""生成应用图标：墨绿圆角方块 + 白色「表」字。"""
import os

from PIL import Image, ImageDraw, ImageFont

SIZE = 256
img = Image.new("RGBA", (SIZE, SIZE), (0, 0, 0, 0))
draw = ImageDraw.Draw(img)

# 圆角底
draw.rounded_rectangle([6, 6, SIZE - 6, SIZE - 6], radius=58, fill=(15, 118, 110, 255))

font = ImageFont.truetype("C:/Windows/Fonts/msyhbd.ttc", 156)
text = "表"
box = draw.textbbox((0, 0), text, font=font)
w, h = box[2] - box[0], box[3] - box[1]
draw.text(((SIZE - w) / 2 - box[0], (SIZE - h) / 2 - box[1]), text, font=font, fill=(255, 255, 255, 255))

out = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "app.ico")
img.save(out, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("written", out, os.path.getsize(out))
