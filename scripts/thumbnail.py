from __future__ import annotations

import argparse
import math
import sys
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[1]
if str(PROJECT_ROOT) not in sys.path:
    sys.path.insert(0, str(PROJECT_ROOT))


def _build_grid(images: list[Path], output_path: Path, cols: int = 3) -> None:
    from PIL import Image, ImageDraw, ImageFont

    opened = [Image.open(path).convert("RGB") for path in images]
    try:
        max_w = max(img.width for img in opened)
        max_h = max(img.height for img in opened)
        cols = max(1, cols)
        rows = math.ceil(len(opened) / cols)
        label_h = 36
        gap = 16

        canvas_w = cols * max_w + (cols + 1) * gap
        canvas_h = rows * (max_h + label_h) + (rows + 1) * gap
        canvas = Image.new("RGB", (canvas_w, canvas_h), color=(245, 245, 245))
        draw = ImageDraw.Draw(canvas)

        try:
            font = ImageFont.truetype("arial.ttf", 18)
        except Exception:
            font = ImageFont.load_default()

        for idx, (img, path) in enumerate(zip(opened, images, strict=False)):
            row = idx // cols
            col = idx % cols
            x = gap + col * (max_w + gap)
            y = gap + row * (max_h + label_h + gap)

            canvas.paste(img, (x, y))
            draw.text((x, y + max_h + 8), path.name, fill=(20, 20, 20), font=font)

        output_path.parent.mkdir(parents=True, exist_ok=True)
        canvas.save(output_path, format="JPEG", quality=92)
    finally:
        for image in opened:
            image.close()


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate PPTX thumbnail grid")
    parser.add_argument("input_pptx", help="Input .pptx")
    parser.add_argument(
        "output_image", nargs="?", default="thumbnails.jpg", help="Output image path"
    )
    parser.add_argument("--cols", type=int, default=3, help="Grid columns")
    args = parser.parse_args()

    from pptx_agent.qa import render_slides_to_images

    input_pptx = Path(args.input_pptx).resolve()
    output_image = Path(args.output_image).resolve()
    temp_dir = output_image.parent / f"{output_image.stem}_slides"

    images = render_slides_to_images(input_pptx, temp_dir, prefix="thumb")
    _build_grid(images, output_image, cols=args.cols)
    print(f"Thumbnail grid written to {output_image}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
