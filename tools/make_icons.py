#!/usr/bin/env python3
"""Draw the ribbon and add-in icons (PNG) into assets/icons/. Needs Pillow."""
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

OUT = Path(__file__).resolve().parent.parent / "assets" / "icons"
TEAL = (2, 80, 67, 255)
MID = (91, 145, 136, 255)
LIGHT = (179, 203, 199, 255)
GREY = (140, 140, 140, 255)
YEL = (255, 214, 0, 255)
WHITE = (255, 255, 255, 255)
S = 320  # draw size, downsampled later


def canvas():
    im = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    return im, ImageDraw.Draw(im)


def r(x):
    return int(x * S / 32)


def library():
    im, d = canvas()
    for (x, y) in [(3, 4), (17, 4), (3, 18), (17, 18)]:
        d.rounded_rectangle([r(x), r(y), r(x + 12), r(y + 10)], radius=r(1.5), fill=TEAL if (x, y) == (3, 4) else MID)
    return im


def starters():
    im, d = canvas()
    d.rounded_rectangle([r(9), r(3), r(29), r(17)], radius=r(1.5), fill=LIGHT)
    d.rounded_rectangle([r(6), r(8), r(26), r(22)], radius=r(1.5), fill=MID)
    d.rounded_rectangle([r(3), r(13), r(23), r(27)], radius=r(1.5), fill=TEAL)
    d.rectangle([r(6), r(17), r(16), r(18.5)], fill=WHITE)
    d.rectangle([r(6), r(21), r(19), r(22.5)], fill=WHITE)
    return im


def sticker():
    im, d = canvas()
    d.polygon([(r(3), r(9)), (r(20), r(9)), (r(28), r(16)), (r(20), r(23)), (r(3), r(23))], fill=TEAL)
    d.ellipse([r(18), r(14), r(22), r(18)], fill=WHITE)
    return im


def source():
    im, d = canvas()
    d.rounded_rectangle([r(3), r(4), r(29), r(28)], radius=r(2), outline=TEAL, width=r(2))
    for i, w in enumerate([18, 14, 18]):
        d.rectangle([r(7), r(9 + i * 4), r(7 + w), r(10.5 + i * 4)], fill=GREY)
    d.rectangle([r(7), r(22), r(19), r(24)], fill=TEAL)
    return im


def needed():
    im, d = canvas()
    d.rounded_rectangle([r(2), r(8), r(30), r(24)], radius=r(2), fill=YEL)
    f = ImageFont.truetype("DejaVuSans-Bold.ttf", r(12))
    d.text((r(16), r(16)), "[ ! ]", font=f, fill=(38, 38, 38, 255), anchor="mm")
    return im


def tools():
    im, d = canvas()
    for i, (y, kx) in enumerate([(8, 20), (16, 10), (24, 16)]):
        d.rectangle([r(4), r(y - 1), r(28), r(y + 1)], fill=GREY)
        d.ellipse([r(kx - 3.5), r(y - 3.5), r(kx + 3.5), r(y + 3.5)], fill=TEAL if i != 1 else MID)
    return im


def check():
    im, d = canvas()
    d.ellipse([r(2), r(2), r(30), r(30)], fill=TEAL)
    d.line([(r(9), r(16.5)), (r(14), r(21.5)), (r(23.5), r(11))], fill=WHITE, width=r(3.2), joint="curve")
    return im


def logo():
    im, d = canvas()
    d.rounded_rectangle([0, 0, S - 1, S - 1], radius=r(5), fill=TEAL)
    try:
        f = ImageFont.truetype("/usr/share/fonts/truetype/google-fonts/Lora-Variable.ttf", r(12.5))
        f.set_variation_by_name("Bold")
    except Exception:
        f = ImageFont.truetype("DejaVuSerif-Bold.ttf", r(11))
    d.text((S / 2, S / 2 + r(0.5)), "GNA", font=f, fill=WHITE, anchor="mm")
    return im


def save(name, im, sizes):
    for s in sizes:
        im.resize((s, s), Image.LANCZOS).save(OUT / f"{name}-{s}.png")


if __name__ == "__main__":
    OUT.mkdir(parents=True, exist_ok=True)
    ribbon = [16, 32, 80]
    for name, fn in [("library", library), ("starters", starters), ("sticker", sticker), ("source", source),
                     ("needed", needed), ("tools", tools), ("check", check)]:
        save(name, fn(), ribbon)
    save("logo", logo(), [16, 32, 64, 80, 128])
    print("icons written to", OUT)
