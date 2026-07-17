"""生成 PyInstaller bootloader 用的启动图（Fluent 深色风格）。
构建时先跑一次这个脚本，把 PNG 写到 assets/splash.png，然后被 spec 引用打进 exe。
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

WIDTH, HEIGHT = 360, 148
RADIUS = 10
BG = (43, 45, 49, 255)         # #2b2d31，Fluent 深色主背景
BORDER = (67, 69, 75, 255)     # #43454b
TITLE_COLOR = (230, 232, 236, 255)  # #e6e8ec
SUB_COLOR = (160, 164, 172, 255)    # #a0a4ac

TITLE = "音频响度标准化工具"
SUBTITLE = "正在启动…"


def _load_font(candidates, size):
    for name in candidates:
        try:
            return ImageFont.truetype(name, size)
        except OSError:
            continue
    return ImageFont.load_default()


def _text_center(draw, text, font):
    bbox = draw.textbbox((0, 0), text, font=font)
    return bbox[2] - bbox[0], bbox[3] - bbox[1]


def main():
    img = Image.new("RGBA", (WIDTH, HEIGHT), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    d.rounded_rectangle(
        [0, 0, WIDTH - 1, HEIGHT - 1],
        radius=RADIUS, fill=BG, outline=BORDER, width=1,
    )

    title_font = _load_font(["msyhbd.ttc", "msyh.ttc", "arial.ttf"], 18)
    sub_font = _load_font(["msyh.ttc", "arial.ttf"], 12)

    tw, th = _text_center(d, TITLE, title_font)
    d.text(((WIDTH - tw) / 2, 46), TITLE, font=title_font, fill=TITLE_COLOR)

    sw, sh = _text_center(d, SUBTITLE, sub_font)
    d.text(((WIDTH - sw) / 2, 92), SUBTITLE, font=sub_font, fill=SUB_COLOR)

    out = Path(__file__).resolve().parent / "assets" / "splash.png"
    out.parent.mkdir(exist_ok=True)
    img.save(out)
    print(f"Wrote {out}")


if __name__ == "__main__":
    main()
