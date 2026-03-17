from __future__ import annotations

from pathlib import Path

from .models import LeafItem, ProjectTree, Section, Topic

IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".webp"}
PDF_EXTENSIONS = {".pdf"}
SUPPORTED_EXTENSIONS = IMAGE_EXTENSIONS | PDF_EXTENSIONS


class ScanError(ValueError):
    pass


def scan_project(root_dir: str | Path) -> ProjectTree:
    root_path = Path(root_dir).expanduser().resolve()
    if not root_path.exists():
        raise ScanError(f"目录不存在: {root_path}")
    if not root_path.is_dir():
        raise ScanError(f"选择的路径不是目录: {root_path}")

    section_dirs = [path for path in sorted(root_path.iterdir()) if path.is_dir()]
    if not section_dirs:
        raise ScanError("总目录下没有一级子目录。")

    project = ProjectTree(root_name=root_path.name, root_path=root_path)

    for section_dir in section_dirs:
        topic_dirs = [path for path in sorted(section_dir.iterdir()) if path.is_dir()]
        if not topic_dirs:
            raise ScanError(f"一级子目录为空: {section_dir.name}")

        section = Section(name=section_dir.name)

        for topic_dir in topic_dirs:
            files = [path for path in sorted(topic_dir.iterdir()) if path.is_file()]
            if not files:
                raise ScanError(f"二级子目录为空: {section_dir.name}/{topic_dir.name}")

            topic = Topic(name=topic_dir.name)

            for file_path in files:
                suffix = file_path.suffix.lower()
                if suffix not in SUPPORTED_EXTENSIONS:
                    project.skipped_files.append(file_path)
                    continue

                kind = "pdf" if suffix in PDF_EXTENSIONS else "image"
                topic.items.append(
                    LeafItem(
                        source_path=file_path,
                        display_name=file_path.stem,
                        kind=kind,
                    )
                )

            if not topic.items:
                raise ScanError(
                    f"二级子目录中没有可用文件: {section_dir.name}/{topic_dir.name}"
                )

            section.topics.append(topic)

        project.sections.append(section)

    return project
