#!/usr/bin/env python3
"""Minimal skill validator for CI."""

from __future__ import annotations

import re
import sys
from pathlib import Path

import yaml


def main() -> int:
    if len(sys.argv) != 2:
        print("Usage: python tools/validate_skill.py <skill_dir>")
        return 1

    skill_dir = Path(sys.argv[1])
    skill_md = skill_dir / "SKILL.md"
    if not skill_md.exists():
        print("SKILL.md not found")
        return 1

    content = skill_md.read_text(encoding="utf-8")
    match = re.match(r"^---\n(.*?)\n---", content, re.DOTALL)
    if not match:
        print("Missing or invalid YAML frontmatter")
        return 1

    data = yaml.safe_load(match.group(1))
    if not isinstance(data, dict):
        print("Frontmatter must be a mapping")
        return 1

    name = data.get("name")
    description = data.get("description")
    if not isinstance(name, str) or not re.fullmatch(r"[a-z0-9-]+", name):
        print("Invalid skill name")
        return 1
    if not isinstance(description, str) or not description.strip():
        print("Missing skill description")
        return 1

    print("Skill metadata looks valid.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

