#!/usr/bin/env python3
"""Remove orphaned slide files from an unpacked PPTX folder."""

from __future__ import annotations

import argparse
from pathlib import Path

from office.common import NS, parse_xml, presentation_paths, qn, register_namespaces, write_xml


def main() -> int:
    register_namespaces()
    parser = argparse.ArgumentParser(description="Clean orphaned slide files in an unpacked PPTX folder.")
    parser.add_argument("unpacked_dir", type=Path, help="Path to the unpacked PPTX directory")
    args = parser.parse_args()

    paths = presentation_paths(args.unpacked_dir.resolve())
    presentation_tree = parse_xml(paths["presentation_xml"])
    rels_tree = parse_xml(paths["presentation_rels"])
    rels_root = rels_tree.getroot()

    used_rids = {
        node.attrib[qn("r", "id")]
        for node in presentation_tree.getroot().findall(f".//{qn('p', 'sldId')}")
    }

    used_targets: set[str] = set()
    for rel in list(rels_root):
        rel_type = rel.attrib.get("Type", "")
        if rel_type.endswith("/slide"):
            if rel.attrib.get("Id") in used_rids:
                used_targets.add(rel.attrib["Target"])
            else:
                rels_root.remove(rel)
    write_xml(rels_tree, paths["presentation_rels"])

    slides_dir = paths["slides_dir"]
    slide_rels_dir = paths["slide_rels_dir"]
    target_names = {Path(target).name for target in used_targets}

    removed = []
    for slide_path in slides_dir.glob("slide*.xml"):
        if slide_path.name not in target_names:
            removed.append(slide_path.name)
            slide_path.unlink()
            rel_path = slide_rels_dir / f"{slide_path.name}.rels"
            if rel_path.exists():
                rel_path.unlink()

    content_tree = parse_xml(paths["content_types"])
    content_root = content_tree.getroot()
    for override in list(content_root.findall(f"{{{NS['ct']}}}Override")):
        part_name = override.attrib.get("PartName", "")
        if part_name.startswith("/ppt/slides/") and Path(part_name).name not in target_names:
            content_root.remove(override)
    write_xml(content_tree, paths["content_types"])

    if removed:
        print("Removed orphan slides:")
        for name in removed:
            print(f"- {name}")
    else:
        print("No orphan slide files found.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

