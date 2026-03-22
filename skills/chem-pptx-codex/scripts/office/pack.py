#!/usr/bin/env python3
"""Pack an unpacked PPTX directory back into a .pptx file."""

from __future__ import annotations

import argparse
import zipfile
from pathlib import Path

from validate import validate


def main() -> int:
    parser = argparse.ArgumentParser(description="Repack an unpacked PPTX directory.")
    parser.add_argument("unpacked_dir", type=Path, help="Path to the unpacked PPTX directory")
    parser.add_argument("output_pptx", type=Path, help="Path to the output .pptx file")
    parser.add_argument("--skip-validation", action="store_true", help="Skip structural validation")
    args = parser.parse_args()

    unpacked_dir = args.unpacked_dir.resolve()
    output_pptx = args.output_pptx.resolve()

    if not args.skip_validation:
        errors = validate(unpacked_dir)
        if errors:
            for error in errors:
                print(error)
            return 1

    output_pptx.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output_pptx, "w", zipfile.ZIP_DEFLATED) as archive:
        for file_path in sorted(unpacked_dir.rglob("*")):
            if file_path.is_file():
                archive.write(file_path, file_path.relative_to(unpacked_dir).as_posix())

    print(f"Packed {output_pptx}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
