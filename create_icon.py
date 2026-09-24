"""
Username Shuffler v3.0.0
Icon Generator Script
Creates icon.png, icon.ico, and titlebar.ico with multi-resolution layers
using modern Windows 11 Fluent style with squircle, indigo-azure-cyan gradient,
smooth bezier shuffle arrows, and user avatar emblem.
"""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFilter


def create_v3_icons() -> None:
    out = Path(__file__).resolve().parent
    size = 512
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))

    # 1. Soft Ambient Drop Shadow
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    sd.rounded_rectangle((44, 52, 468, 476), radius=100, fill=(10, 18, 38, 140))
    shadow = shadow.filter(ImageFilter.GaussianBlur(26))
    img.alpha_composite(shadow)

    # 2. Main Squircle Body with Gradient
    mask = Image.new("L", (size, size), 0)
    md = ImageDraw.Draw(mask)
    md.rounded_rectangle((40, 36, 472, 468), radius=96, fill=255)

    grad = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    gd = ImageDraw.Draw(grad)
    for y in range(size):
        t = y / (size - 1)
        # Gradient: Royal Indigo (#4338CA) -> Electric Azure (#2563EB) -> Bright Cyan (#06B6D4)
        if t < 0.55:
            nt = t / 0.55
            r = int(67 * (1 - nt) + 37 * nt)
            g = int(56 * (1 - nt) + 99 * nt)
            b = int(202 * (1 - nt) + 235 * nt)
        else:
            nt = (t - 0.55) / 0.45
            r = int(37 * (1 - nt) + 6 * nt)
            g = int(99 * (1 - nt) + 182 * nt)
            b = int(235 * (1 - nt) + 212 * nt)
        gd.line((0, y, size, y), fill=(r, g, b, 255))
    grad.putalpha(mask)
    img.alpha_composite(grad)

    # 3. Glassmorphic Inner Edge Highlight
    border = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    bd = ImageDraw.Draw(border)
    bd.rounded_rectangle((41, 37, 471, 467), radius=95, outline=(255, 255, 255, 75), width=3)
    img.alpha_composite(border)

    draw = ImageDraw.Draw(img)

    def bezier_pts(p0, p1, p2, p3, steps=24):
        pts = []
        for i in range(steps + 1):
            t = i / steps
            u = 1 - t
            x = u * u * u * p0[0] + 3 * u * u * t * p1[0] + 3 * u * t * t * p2[0] + t * t * t * p3[0]
            y = u * u * u * p0[1] + 3 * u * u * t * p1[1] + 3 * u * t * t * p2[1] + t * t * t * p3[1]
            pts.append((x, y))
        return pts

    white = (255, 255, 255, 245)
    cyan_glow = (186, 230, 253, 245)

    # Shuffle Arrow 1 (Top-Left to Bottom-Right)
    p_start1 = [(100, 168), (170, 168)]
    curve1 = bezier_pts((170, 168), (256, 168), (256, 336), (342, 336))
    p_end1 = [(342, 336), (370, 336)]
    draw.line(p_start1 + curve1 + p_end1, fill=white, width=32, joint="curve")
    # Arrowhead 1
    draw.polygon([(365, 298), (426, 336), (365, 374)], fill=white)

    # Shuffle Arrow 2 (Bottom-Left to Top-Right)
    p_start2 = [(100, 336), (170, 336)]
    curve2_a = bezier_pts((170, 336), (220, 336), (226, 295), (236, 276))
    curve2_b = bezier_pts((276, 228), (286, 209), (292, 168), (342, 168))
    p_end2 = [(342, 168), (370, 168)]
    draw.line(p_start2 + curve2_a, fill=cyan_glow, width=32, joint="curve")
    draw.line(curve2_b + p_end2, fill=cyan_glow, width=32, joint="curve")
    # Arrowhead 2
    draw.polygon([(365, 130), (426, 168), (365, 206)], fill=cyan_glow)

    # Central User Avatar Badge (Floating Modern Pill)
    badge = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    bb = ImageDraw.Draw(badge)
    bb.ellipse((216, 212, 296, 292), fill=(15, 23, 42, 220), outline=(255, 255, 255, 200), width=3)
    bb.ellipse((242, 228, 270, 256), fill=white)
    bb.chord((232, 256, 280, 290), start=180, end=0, fill=white)
    img.alpha_composite(badge)

    # Generate 256x256 master PNG
    icon_256 = img.resize((256, 256), Image.Resampling.LANCZOS)
    icon_256.save(out / "icon.png")

    # Multi-resolution ICO for application & desktop
    sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (16, 16)]
    images = [img.resize(s, Image.Resampling.LANCZOS) for s in sizes]
    images[0].save(out / "icon.ico", format="ICO", sizes=sizes, append_images=images[1:])

    # Multi-resolution ICO for titlebar
    title_sizes = [(64, 64), (48, 48), (32, 32), (16, 16)]
    title_images = [img.resize(s, Image.Resampling.LANCZOS) for s in title_sizes]
    title_images[0].save(out / "titlebar.ico", format="ICO", sizes=title_sizes, append_images=title_images[1:])

    print("Generated new v3.0.0 icon.png, icon.ico, and titlebar.ico!")


if __name__ == "__main__":
    create_v3_icons()
