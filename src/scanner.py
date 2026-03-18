from __future__ import annotations

import os
from collections import OrderedDict
from collections.abc import Callable
from pathlib import Path

from .models import LeafItem, ProjectTree, Section, Topic

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
PDF_EXTENSIONS = {".pdf"}
SUPPORTED_EXTENSIONS = IMAGE_EXTENSIONS | PDF_EXTENSIONS


class ScanError(ValueError):
    pass


def scan_project(
    root_dir: str | Path,
    progress_callback: Callable[[int, str], None] | None = None,
) -> ProjectTree:
    root_path = Path(root_dir).expanduser().resolve()
    if not root_path.exists():
        raise ScanError(f"目录不存在: {root_path}")
    if not root_path.is_dir():
        raise ScanError(f"选择的路径不是目录: {root_path}")

    project = ProjectTree(root_name=root_path.name, root_path=root_path)
    sections_by_name: OrderedDict[str, Section] = OrderedDict()
    scanned_dirs = 0

    def _on_walk_error(error: OSError) -> None:
        if error.filename:
            project.skipped_files.append(Path(error.filename))

    for current_root, dirnames, filenames in os.walk(root_path, topdown=True, onerror=_on_walk_error):
        dirnames.sort()
        filenames.sort()
        current_dir = Path(current_root)
        scanned_dirs += 1
        if progress_callback is not None and (scanned_dirs == 1 or scanned_dirs % 20 == 0):
            progress_callback(scanned_dirs, str(current_dir))

        files = [current_dir / name for name in filenames]
        if not files:
            continue

        supported_items: list[LeafItem] = []
        for file_path in files:
            suffix = file_path.suffix.lower()
            if suffix not in SUPPORTED_EXTENSIONS:
                project.skipped_files.append(file_path)
                continue

            kind = "pdf" if suffix in PDF_EXTENSIONS else "image"
            supported_items.append(
                LeafItem(
                    source_path=file_path,
                    display_name=file_path.stem,
                    kind=kind,
                )
            )

        if not supported_items:
            continue

        section_name, topic_name = _resolve_section_and_topic_names(root_path, current_dir)
        section = sections_by_name.setdefault(section_name, Section(name=section_name))
        section.topics.append(Topic(name=topic_name, items=supported_items))

    project.sections = list(sections_by_name.values())
    if not project.sections:
        raise ScanError("所选目录及其子目录中没有可用的图片或 PDF 文件。")

    return project


def _resolve_section_and_topic_names(root_path: Path, current_dir: Path) -> tuple[str, str]:
    relative_parts = current_dir.relative_to(root_path).parts

    if not relative_parts:
        return "当前文件夹", root_path.name
    if len(relative_parts) == 1:
        return relative_parts[0], "文件列表"
    return relative_parts[0], " / ".join(relative_parts[1:])
