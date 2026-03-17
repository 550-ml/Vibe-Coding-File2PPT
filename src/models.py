from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path


@dataclass
class LeafItem:
    source_path: Path
    display_name: str
    kind: str
    preview_path: Path | None = None


@dataclass
class Topic:
    name: str
    items: list[LeafItem] = field(default_factory=list)


@dataclass
class Section:
    name: str
    topics: list[Topic] = field(default_factory=list)


@dataclass
class ProjectTree:
    root_name: str
    root_path: Path
    sections: list[Section] = field(default_factory=list)
    skipped_files: list[Path] = field(default_factory=list)


@dataclass
class BuildResult:
    output_path: Path
    slide_count: int
    skipped_files: list[Path] = field(default_factory=list)
