from __future__ import annotations

import argparse
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("directory", type=Path)
    parser.add_argument("output", type=Path)
    parser.add_argument("patterns", nargs="+")
    parser.add_argument("--columns", type=int, default=5)
    args = parser.parse_args()

    paths: list[Path] = []
    for pattern in args.patterns:
        paths.extend(args.directory.glob(pattern))
    paths = sorted(set(paths), key=lambda path: path.name)
    if not paths:
        raise SystemExit("No images matched")

    thumb_w, thumb_h, label_h = 240, 320, 42
    rows = (len(paths) + args.columns - 1) // args.columns
    sheet = Image.new("RGB", (args.columns * thumb_w, rows * (thumb_h + label_h)), "white")
    draw = ImageDraw.Draw(sheet)
    font = ImageFont.load_default(size=18)

    for index, path in enumerate(paths):
        with Image.open(path) as source:
            image = ImageOps.exif_transpose(source).convert("RGB")
            image.thumbnail((thumb_w - 12, thumb_h - 12), Image.Resampling.LANCZOS)
            x = (index % args.columns) * thumb_w
            y = (index // args.columns) * (thumb_h + label_h)
            px = x + (thumb_w - image.width) // 2
            py = y + (thumb_h - image.height) // 2
            sheet.paste(image, (px, py))
            draw.rectangle((x, y, x + thumb_w - 1, y + thumb_h + label_h - 1), outline="#888888")
            draw.text((x + 6, y + thumb_h + 8), path.stem, fill="black", font=font)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    sheet.save(args.output, quality=92)


if __name__ == "__main__":
    main()
