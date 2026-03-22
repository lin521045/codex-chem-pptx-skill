"""Shared XML helpers for PPTX unpack/pack utilities."""

from __future__ import annotations

from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "ct": "http://schemas.openxmlformats.org/package/2006/content-types",
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


def register_namespaces() -> None:
    ET.register_namespace("a", NS["a"])
    ET.register_namespace("p", NS["p"])
    ET.register_namespace("r", NS["r"])
    ET.register_namespace("", NS["pr"])


def qn(prefix: str, tag: str) -> str:
    return f"{{{NS[prefix]}}}{tag}"


def parse_xml(path: Path) -> ET.ElementTree:
    return ET.parse(path)


def write_xml(tree: ET.ElementTree, path: Path) -> None:
    tree.write(path, encoding="utf-8", xml_declaration=True)


def next_relationship_id(rel_root: ET.Element) -> str:
    maximum = 0
    for rel in rel_root:
        rel_id = rel.attrib.get("Id", "")
        if rel_id.startswith("rId"):
            suffix = rel_id[3:]
            if suffix.isdigit():
                maximum = max(maximum, int(suffix))
    return f"rId{maximum + 1}"


def presentation_paths(unpacked_dir: Path) -> dict[str, Path]:
    return {
        "content_types": unpacked_dir / "[Content_Types].xml",
        "presentation_xml": unpacked_dir / "ppt" / "presentation.xml",
        "presentation_rels": unpacked_dir / "ppt" / "_rels" / "presentation.xml.rels",
        "slides_dir": unpacked_dir / "ppt" / "slides",
        "slide_rels_dir": unpacked_dir / "ppt" / "slides" / "_rels",
    }

