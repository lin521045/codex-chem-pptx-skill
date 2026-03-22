#!/usr/bin/env python3
"""Render a deck and compose slide thumbnails into a contact sheet."""

from __future__ import annotations

import argparse
import math
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont, ImageOps


def run_render(pptx: Path, render_dir: Path, dpi: int) -> list[Path]:
    render_script = Path(__file__).resolve().parent / "office" / "render.py"
    command = [sys.executable, str(render_script), str(pptx), str(render_dir), "--dpi", str(dpi)]
    subprocess.run(command, check=True)
    return sorted(render_dir.glob("slide-*.png"))


def build_contact_sheet(images: list[Path], output_path: Path, cols: int) -> None:
    if not images:
        raise ValueError("No rendered slide images were found.")

    thumb_w = 320
    thumb_h = 180
    caption_h = 28
    gutter = 24
    rows = math.ceil(len(images) / cols)
    canvas = Image.new(
        "RGB",
        (cols * thumb_w + (cols + 1) * gutter, rows * (thumb_h + caption_h) + (rows + 1) * gutter),
        color=(248, 248, 248),
    )
    draw = ImageDraw.Draw(canvas)
    font = ImageFont.load_default()

    for idx, image_path in enumerate(images):
        row = idx // cols
        col = idx % cols
        x = gutter + col * (thumb_w + gutter)
        y = gutter + row * (thumb_h + caption_h + gutter)

        with Image.open(image_path) as slide_img:
            thumb = ImageOps.contain(slide_img.convert("RGB"), (thumb_w, thumb_h))
            paste_x = x + (thumb_w - thumb.width) // 2
            paste_y = y + (thumb_h - thumb.height) // 2
            canvas.paste(thumb, (paste_x, paste_y))

        draw.rectangle((x, y, x + thumb_w, y + thumb_h), outline=(200, 200, 200), width=1)
        draw.text((x, y + thumb_h + 8), image_path.stem, fill=(55, 55, 55), font=font)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    canvas.save(output_path, quality=92)


def main() -> int:
    parser = argparse.ArgumentParser(description="Create a slide thumbnail sheet from a PPTX deck.")
    parser.add_argument("pptx", type=Path, help="Path to the .pptx file")
    parser.add_argument("output_prefix", nargs="?", default=None, help="Output prefix without extension")
    parser.add_argument("--cols", type=int, default=3, help="Number of columns in the contact sheet")
    parser.add_argument("--dpi", type=int, default=160, help="Render DPI")
    parser.add_argument("--render-dir", type=Path, default=None, help="Keep individual rendered slides here")
    args = parser.parse_args()

    output_prefix = args.output_prefix or args.pptx.stem
    output_path = Path(f"{output_prefix}-thumbnails.jpg").resolve()

    if args.render_dir:
        render_dir = args.render_dir.resolve()
        render_dir.mkdir(parents=True, exist_ok=True)
        images = run_render(args.pptx.resolve(), render_dir, args.dpi)
    else:
        with tempfile.TemporaryDirectory(prefix="chem-pptx-codex-") as tmp_dir:
            render_dir = Path(tmp_dir)
            images = run_render(args.pptx.resolve(), render_dir, args.dpi)
            build_contact_sheet(images, output_path, args.cols)
        print(f"Wrote {output_path}")
        return 0

    build_contact_sheet(images, output_path, args.cols)
    print(f"Wrote {output_path}")
    print(f"Rendered slides kept in {render_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
