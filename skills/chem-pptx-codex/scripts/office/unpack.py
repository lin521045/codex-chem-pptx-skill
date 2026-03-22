#!/usr/bin/env python3
"""Unpack a PPTX file into an editable directory tree."""

from __future__ import annotations

import argparse
import shutil
import zipfile
from pathlib import Path
from xml.dom import minidom


def prettify_xml(path: Path) -> None:
    raw = path.read_bytes()
    try:
        pretty = minidom.parseString(raw).toprettyxml(indent="  ", encoding="utf-8")
    except Exception:
        return
    lines = [line for line in pretty.decode("utf-8").splitlines() if line.strip()]
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def main() -> int:
    parser = argparse.ArgumentParser(description="Unpack a PPTX file to a directory.")
    parser.add_argument("pptx", type=Path, help="Path to the .pptx file")
    parser.add_argument("output_dir", type=Path, help="Directory to create")
    args = parser.parse_args()

    if args.output_dir.exists():
        shutil.rmtree(args.output_dir)
    args.output_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(args.pptx, "r") as archive:
        archive.extractall(args.output_dir)

    for xml_path in args.output_dir.rglob("*.xml"):
        prettify_xml(xml_path)
    for rel_path in args.output_dir.rglob("*.rels"):
        prettify_xml(rel_path)

    print(f"Unpacked {args.pptx} to {args.output_dir}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

