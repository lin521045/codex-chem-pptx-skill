#!/usr/bin/env python3
"""Compatibility wrapper for converting PPTX files to PDF/images.

This mirrors the reference skill's `scripts/office/soffice.py` entrypoint,
but is adapted for Codex environments. It prefers LibreOffice when available
and falls back to Microsoft PowerPoint automation on Windows for PDF export.
"""

from __future__ import annotations

import argparse
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

from PIL import Image

from render import render_with_powerpoint


def run_libreoffice(raw_args: list[str]) -> int:
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        return -1
    completed = subprocess.run([soffice, *raw_args], check=False)
    return completed.returncode


def export_pdf_with_powerpoint(input_path: Path, outdir: Path) -> int:
    try:
        outdir.mkdir(parents=True, exist_ok=True)
        output_pdf = outdir / f"{input_path.stem}.pdf"
        with tempfile.TemporaryDirectory(prefix="chem-pptx-codex-pdf-") as tmp_dir:
            images = render_with_powerpoint(input_path.resolve(), Path(tmp_dir), dpi=150)
            if not images:
                print("PowerPoint render produced no slide images.", file=sys.stderr)
                return 1
            pil_images = [Image.open(path).convert("RGB") for path in images]
            first = pil_images[0]
            rest = pil_images[1:]
            first.save(output_pdf, save_all=True, append_images=rest)
            for image in pil_images:
                image.close()
        print(output_pdf)
        return 0
    except Exception as exc:
        print(f"PowerPoint PDF export failed: {exc}", file=sys.stderr)
        return 1


def main() -> int:
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--convert-to")
    parser.add_argument("--outdir")
    parser.add_argument("input", nargs="?")
    args, raw_args = parser.parse_known_args()

    rc = run_libreoffice(sys.argv[1:])
    if rc != -1:
        return rc

    if args.convert_to == "pdf" and args.input:
        input_path = Path(args.input)
        outdir = Path(args.outdir) if args.outdir else input_path.parent
        rc = export_pdf_with_powerpoint(input_path, outdir)
        if rc != -1:
            return rc

    print(
        "No compatible office converter found. Install LibreOffice (`soffice`) "
        "or run on Windows with Microsoft PowerPoint installed.",
        file=sys.stderr,
    )
    return 1


if __name__ == "__main__":
    raise SystemExit(main())
