#!/usr/bin/env python3
"""Validate the minimum structure of an unpacked PPTX directory."""

from __future__ import annotations

import argparse
from pathlib import Path

from common import parse_xml, presentation_paths, qn, register_namespaces


def validate(unpacked_dir: Path) -> list[str]:
    errors: list[str] = []
    paths = presentation_paths(unpacked_dir)
    for name, path in paths.items():
        if name.endswith("_dir"):
            continue
        if not path.exists():
            errors.append(f"Missing required file: {path}")

    if errors:
        return errors

    presentation_tree = parse_xml(paths["presentation_xml"])
    rels_tree = parse_xml(paths["presentation_rels"])
    used_rids = {
        node.attrib[qn("r", "id")]
        for node in presentation_tree.getroot().findall(f".//{qn('p', 'sldId')}")
    }

    rid_to_target = {}
    for rel in rels_tree.getroot():
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/slide"):
            rid_to_target[rel.attrib["Id"]] = rel.attrib["Target"]

    for rid in used_rids:
        target = rid_to_target.get(rid)
        if not target:
            errors.append(f"Slide relationship not found for {rid}")
            continue
        slide_path = unpacked_dir / "ppt" / Path(target)
        if not slide_path.exists():
            errors.append(f"Slide target missing on disk: {slide_path}")

    return errors


def main() -> int:
    register_namespaces()
    parser = argparse.ArgumentParser(description="Validate an unpacked PPTX directory.")
    parser.add_argument("unpacked_dir", type=Path)
    args = parser.parse_args()

    errors = validate(args.unpacked_dir.resolve())
    if errors:
        for error in errors:
            print(error)
        return 1
    print("Unpacked PPTX looks structurally valid.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

