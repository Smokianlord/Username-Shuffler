from PIL import Image, ImageDraw, ImageFont, ImageFilter
from pathlib import Path

out = Path(__file__).resolve().parent
size = 256
img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
shadow = Image.new('RGBA', (size, size), (0, 0, 0, 0))
sd = ImageDraw.Draw(shadow)
sd.rounded_rectangle((26, 30, 234, 238), radius=46, fill=(15, 23, 42, 90))
shadow = shadow.filter(ImageFilter.GaussianBlur(12))
img.alpha_composite(shadow)

base = Image.new('RGBA', (size, size), (0, 0, 0, 0))
d = ImageDraw.Draw(base)
# gradient rounded square mask
mask = Image.new('L', (size, size), 0)
md = ImageDraw.Draw(mask)
md.rounded_rectangle((22, 18, 234, 230), radius=46, fill=255)
grad = Image.new('RGBA', (size, size), (0, 0, 0, 0))
gd = ImageDraw.Draw(grad)
for y in range(size):
    t = y / (size - 1)
    r = int(37 * (1 - t) + 8 * t)
    g = int(99 * (1 - t) + 145 * t)
    b = int(235 * (1 - t) + 178 * t)
    gd.line((0, y, size, y), fill=(r, g, b, 255))
grad.putalpha(mask)
img.alpha_composite(grad)

d = ImageDraw.Draw(img)
# Decorative shuffle arrows
white = (255, 255, 255, 235)
light = (191, 219, 254, 235)
d.line((58, 82, 104, 82, 148, 174, 192, 174), fill=light, width=12)
d.polygon([(192, 150), (224, 174), (192, 198)], fill=light)
d.line((58, 174, 102, 174, 146, 82, 192, 82), fill=white, width=12)
d.polygon([(192, 58), (224, 82), (192, 106)], fill=white)
# user dots
for xy, fill in [((55, 56, 83, 84), white), ((55, 148, 83, 176), light), ((171, 56, 199, 84), light), ((171, 148, 199, 176), white)]:
    d.ellipse(xy, fill=fill)
# US text
try:
    font = ImageFont.truetype('/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf', 54)
except Exception:
    font = ImageFont.load_default()
text = 'US'
bbox = d.textbbox((0, 0), text, font=font)
tw = bbox[2] - bbox[0]
th = bbox[3] - bbox[1]
d.text(((size - tw) / 2, 184 - th / 2), text, font=font, fill=(255, 255, 255, 245))

img.save(out / 'icon.png')
img.save(out / 'icon.ico', sizes=[(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)])
