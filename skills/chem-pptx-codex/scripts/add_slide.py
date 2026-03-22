#!/usr/bin/env python3
"""Duplicate a slide inside an unpacked PPTX directory."""

from __future__ import annotations

import argparse
import shutil
from pathlib import Path

from office.common import (
    NS,
    next_relationship_id,
    parse_xml,
    presentation_paths,
    qn,
    register_namespaces,
    write_xml,
)


def resolve_source(slides_dir: Path, source: str) -> tuple[Path, int]:
    source_path = Path(source)
    if source_path.name.startswith("slide") and source_path.suffix == ".xml":
        slide_path = slides_dir / source_path.name
    else:
        slide_path = slides_dir / source
    if not slide_path.exists():
        raise FileNotFoundError(f"Source slide not found: {slide_path}")
    number = int(slide_path.stem.replace("slide", ""))
    return slide_path, number


def next_slide_number(slides_dir: Path) -> int:
    existing = [int(path.stem.replace("slide", "")) for path in slides_dir.glob("slide*.xml")]
    return max(existing, default=0) + 1


def strip_notes_relationships(rel_path: Path) -> None:
    if not rel_path.exists():
        return
    tree = parse_xml(rel_path)
    root = tree.getroot()
    for rel in list(root):
        rel_type = rel.attrib.get("Type", "")
        target = rel.attrib.get("Target", "")
        if rel_type.endswith("/notesSlide") or "notesSlides" in target:
            root.remove(rel)
    write_xml(tree, rel_path)


def ensure_content_type(content_types_path: Path, slide_number: int) -> None:
    tree = parse_xml(content_types_path)
    root = tree.getroot()
    target = f"/ppt/slides/slide{slide_number}.xml"
    existing = {
        node.attrib.get("PartName")
        for node in root.findall(f"{{{NS['ct']}}}Override")
    }
    if target not in existing:
        from xml.etree.ElementTree import Element

        root.append(
            Element(
            qn("ct", "Override"),
            {
                "PartName": target,
                "ContentType": "application/vnd.openxmlformats-officedocument.presentationml.slide+xml",
            },
        )
        )
    write_xml(tree, content_types_path)


def main() -> int:
    register_namespaces()
    parser = argparse.ArgumentParser(description="Duplicate a slide inside an unpacked PPTX folder.")
    parser.add_argument("unpacked_dir", type=Path, help="Path to the unpacked PPTX directory")
    parser.add_argument("source_slide", help="Source slide name such as slide2.xml")
    args = parser.parse_args()

    paths = presentation_paths(args.unpacked_dir.resolve())
    slides_dir = paths["slides_dir"]
    slide_rels_dir = paths["slide_rels_dir"]
    source_slide_path, source_number = resolve_source(slides_dir, args.source_slide)
    target_number = next_slide_number(slides_dir)

    new_slide_path = slides_dir / f"slide{target_number}.xml"
    shutil.copy2(source_slide_path, new_slide_path)

    source_rel_path = slide_rels_dir / f"slide{source_number}.xml.rels"
    target_rel_path = slide_rels_dir / f"slide{target_number}.xml.rels"
    if source_rel_path.exists():
        shutil.copy2(source_rel_path, target_rel_path)
        strip_notes_relationships(target_rel_path)

    rels_tree = parse_xml(paths["presentation_rels"])
    rels_root = rels_tree.getroot()
    new_rid = next_relationship_id(rels_root)
    from xml.etree.ElementTree import Element

    rels_root.append(
        Element(
            qn("pr", "Relationship"),
            {
                "Id": new_rid,
                "Type": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide",
                "Target": f"slides/slide{target_number}.xml",
            },
        )
    )
    write_xml(rels_tree, paths["presentation_rels"])

    presentation_tree = parse_xml(paths["presentation_xml"])
    presentation_root = presentation_tree.getroot()
    sld_id_list = presentation_root.find(f".//{qn('p', 'sldIdLst')}")
    if sld_id_list is None:
        raise RuntimeError("Could not find p:sldIdLst in presentation.xml")
    existing_ids = [int(node.attrib["id"]) for node in sld_id_list.findall(qn("p", "sldId"))]
    next_id = str(max(existing_ids, default=255) + 1)
    sld_id_list.append(Element(qn("p", "sldId"), {"id": next_id, qn("r", "id"): new_rid}))
    write_xml(presentation_tree, paths["presentation_xml"])

    ensure_content_type(paths["content_types"], target_number)

    print(f"Created slide{target_number}.xml with relationship {new_rid}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
