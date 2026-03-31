#!/usr/bin/env python3
from pathlib import Path
from PIL import Image, ImageDraw


def draw_icon(size: int) -> Image.Image:
    scale = size / 128.0
    image = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(image)

    def s(value: float) -> int:
        return round(value * scale)

    # Approximate the selected protocol-node brand mark.
    bg = (8, 13, 26, 255)
    green = (16, 185, 129, 255)
    green_dark = (4, 120, 87, 255)
    light = (243, 244, 246, 255)
    stroke = (75, 85, 99, 255)

    draw.rounded_rectangle((0, 0, size - 1, size - 1), radius=s(28), fill=bg)

    draw.rounded_rectangle((s(32), s(32), s(58), s(58)), radius=s(6), fill=light)
    draw.rounded_rectangle((s(70), s(32), s(96), s(58)), radius=s(6), fill=green)
    draw.rounded_rectangle((s(32), s(70), s(58), s(96)), radius=s(6), fill=green_dark)
    draw.rounded_rectangle((s(70), s(70), s(96), s(96)), radius=s(6), fill=light)

    line_width = max(2, s(4))
    draw.line((s(45), s(58), s(45), s(70)), fill=stroke, width=line_width)
    draw.line((s(83), s(58), s(83), s(70)), fill=stroke, width=line_width)
    draw.line((s(58), s(45), s(70), s(45)), fill=stroke, width=line_width)
    draw.line((s(58), s(83), s(70), s(83)), fill=stroke, width=line_width)

    return image


def main() -> None:
    root = Path(__file__).resolve().parents[1]
    iconset = root / "AppIcon.iconset"
    iconset.mkdir(parents=True, exist_ok=True)

    sizes = {
        "icon_16x16.png": 16,
        "icon_16x16@2x.png": 32,
        "icon_32x32.png": 32,
        "icon_32x32@2x.png": 64,
        "icon_128x128.png": 128,
        "icon_128x128@2x.png": 256,
        "icon_256x256.png": 256,
        "icon_256x256@2x.png": 512,
        "icon_512x512.png": 512,
        "icon_512x512@2x.png": 1024,
    }

    for filename, size in sizes.items():
        draw_icon(size).save(iconset / filename)


if __name__ == "__main__":
    main()
