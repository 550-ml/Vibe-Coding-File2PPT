from __future__ import annotations

from collections.abc import Callable
from pathlib import Path

import fitz
from PIL import Image, ImageOps

from .models import ProjectTree

THUMBNAIL_SIZE = (900, 600)


def generate_previews(
    project: ProjectTree,
    preview_dir: str | Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> ProjectTree:
    preview_root = Path(preview_dir).expanduser().resolve()
    preview_root.mkdir(parents=True, exist_ok=True)
    total_items = sum(len(topic.items) for section in project.sections for topic in section.topics)
    processed = 0

    for section in project.sections:
        for topic in section.topics:
            for item in topic.items:
                processed += 1
                if progress_callback is not None:
                    progress_callback(processed, total_items, item.source_path.name)
                safe_name = (
                    _safe_filename(f"{section.name}_{topic.name}_{item.source_path.stem}_{item.kind}")
                    + ".png"
                )
                output_path = preview_root / safe_name
                if item.kind == "pdf":
                    _render_pdf_first_page(item.source_path, output_path)
                else:
                    _create_image_preview(item.source_path, output_path)
                item.preview_path = output_path

    return project


def _render_pdf_first_page(pdf_path: Path, output_path: Path) -> None:
    with fitz.open(pdf_path) as document:
        if document.page_count < 1:
            raise ValueError(f"PDF 没有可渲染页面: {pdf_path}")
        page = document.load_page(0)
        pixmap = page.get_pixmap(matrix=fitz.Matrix(1.8, 1.8), alpha=False)
        pixmap.save(str(output_path))
    _normalize_preview(output_path)


def _create_image_preview(image_path: Path, output_path: Path) -> None:
    with Image.open(image_path) as image:
        image = ImageOps.exif_transpose(image)
        image.thumbnail(THUMBNAIL_SIZE)
        canvas = Image.new("RGB", THUMBNAIL_SIZE, "white")
        x = (THUMBNAIL_SIZE[0] - image.width) // 2
        y = (THUMBNAIL_SIZE[1] - image.height) // 2
        canvas.paste(image.convert("RGB"), (x, y))
        canvas.save(output_path, format="PNG")


def _normalize_preview(image_path: Path) -> None:
    with Image.open(image_path) as image:
        image.thumbnail(THUMBNAIL_SIZE)
        canvas = Image.new("RGB", THUMBNAIL_SIZE, "white")
        x = (THUMBNAIL_SIZE[0] - image.width) // 2
        y = (THUMBNAIL_SIZE[1] - image.height) // 2
        canvas.paste(image.convert("RGB"), (x, y))
        canvas.save(image_path, format="PNG")


def _safe_filename(value: str) -> str:
    keep = []
    for char in value:
        if char.isalnum() or char in {"-", "_"}:
            keep.append(char)
        else:
            keep.append("_")
    return "".join(keep)
