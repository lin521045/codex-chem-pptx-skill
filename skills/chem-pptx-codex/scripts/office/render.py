#!/usr/bin/env python3
"""Render PPTX slides to individual PNG files."""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

import fitz


class BackendUnavailable(RuntimeError):
    """Raised when a render backend cannot be used."""


def render_with_powerpoint(pptx: Path, output_dir: Path, dpi: int) -> list[Path]:
    try:
        import pythoncom
        import win32com.client
    except ImportError as exc:
        raise BackendUnavailable("pywin32 is not available for PowerPoint automation.") from exc

    pythoncom.CoInitialize()
    app = None
    deck = None
    images: list[Path] = []
    try:
        app = win32com.client.Dispatch("PowerPoint.Application")
        app.Visible = 1
        deck = app.Presentations.Open(str(pptx.resolve()), ReadOnly=True, Untitled=False, WithWindow=False)
        slide_width_pt = deck.PageSetup.SlideWidth
        slide_height_pt = deck.PageSetup.SlideHeight
        width_px = max(int(slide_width_pt / 72 * dpi), 1)
        height_px = max(int(slide_height_pt / 72 * dpi), 1)
        for idx in range(1, deck.Slides.Count + 1):
            slide = deck.Slides(idx)
            output = output_dir / f"slide-{idx:02d}.png"
            slide.Export(str(output), "PNG", width_px, height_px)
            images.append(output)
    except Exception as exc:
        raise BackendUnavailable(f"PowerPoint automation failed: {exc}") from exc
    finally:
        if deck is not None:
            deck.Close()
        if app is not None:
            app.Quit()
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return images


def render_pdf_to_pngs(pdf_path: Path, output_dir: Path, dpi: int) -> list[Path]:
    images: list[Path] = []
    document = fitz.open(pdf_path)
    try:
        for page_index in range(document.page_count):
            page = document.load_page(page_index)
            pixmap = page.get_pixmap(dpi=dpi, alpha=False)
            output = output_dir / f"slide-{page_index + 1:02d}.png"
            pixmap.save(output)
            images.append(output)
    finally:
        document.close()
    return images


def render_with_libreoffice(pptx: Path, output_dir: Path, dpi: int) -> list[Path]:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise BackendUnavailable("LibreOffice was not found in PATH.")

    command = [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(pptx)]
    try:
        subprocess.run(command, check=True, capture_output=True, text=True)
    except subprocess.CalledProcessError as exc:
        raise BackendUnavailable(f"LibreOffice conversion failed: {exc.stderr.strip()}") from exc

    pdf_path = output_dir / f"{pptx.stem}.pdf"
    if not pdf_path.exists():
        raise BackendUnavailable("LibreOffice did not produce the expected PDF file.")
    return render_pdf_to_pngs(pdf_path, output_dir, dpi)


def main() -> int:
    parser = argparse.ArgumentParser(description="Render a PPTX presentation to PNG slides.")
    parser.add_argument("pptx", type=Path, help="Path to the .pptx file")
    parser.add_argument("output_dir", type=Path, help="Directory for rendered PNG files")
    parser.add_argument("--dpi", type=int, default=160, help="Rasterization DPI")
    parser.add_argument(
        "--backend",
        choices=["auto", "powerpoint", "libreoffice"],
        default="auto",
        help="Preferred render backend",
    )
    args = parser.parse_args()

    args.output_dir.mkdir(parents=True, exist_ok=True)
    errors: list[str] = []

    strategies = {
        "powerpoint": [render_with_powerpoint],
        "libreoffice": [render_with_libreoffice],
        "auto": [render_with_powerpoint, render_with_libreoffice],
    }[args.backend]

    for strategy in strategies:
        try:
            images = strategy(args.pptx.resolve(), args.output_dir.resolve(), args.dpi)
            for image in images:
                print(image)
            return 0
        except BackendUnavailable as exc:
            errors.append(str(exc))

    print("Could not render slides with any available backend.", file=sys.stderr)
    for message in errors:
        print(f"- {message}", file=sys.stderr)
    return 1


if __name__ == "__main__":
    raise SystemExit(main())

