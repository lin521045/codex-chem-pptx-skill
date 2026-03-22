#!/usr/bin/env python3
"""Alias entrypoint to export a PPTX deck into slide images."""

from __future__ import annotations

import argparse
import subprocess
import sys
from pathlib import Path


def main() -> int:
    parser = argparse.ArgumentParser(description="Convert a PPTX deck into slide images.")
    parser.add_argument("pptx", type=Path, help="Path to the PPTX file")
    parser.add_argument("output_dir", type=Path, help="Directory for rendered slides")
    parser.add_argument("--dpi", type=int, default=160, help="Rasterization DPI")
    args = parser.parse_args()

    render_script = Path(__file__).resolve().parent / "office" / "render.py"
    command = [
        sys.executable,
        str(render_script),
        str(args.pptx.resolve()),
        str(args.output_dir.resolve()),
        "--dpi",
        str(args.dpi),
    ]
    completed = subprocess.run(command, check=False)
    return completed.returncode


if __name__ == "__main__":
    raise SystemExit(main())
