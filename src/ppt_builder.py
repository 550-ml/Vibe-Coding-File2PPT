from __future__ import annotations

import math
from collections.abc import Callable
from pathlib import Path

from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_AUTO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN, MSO_VERTICAL_ANCHOR
from pptx.util import Inches, Pt

from .models import BuildResult, ProjectTree, Topic

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)


def build_ppt(
    project_tree: ProjectTree,
    output_path: str | Path,
    progress_callback: Callable[[int, int, str], None] | None = None,
) -> BuildResult:
    output = Path(output_path).expanduser().resolve()
    output.parent.mkdir(parents=True, exist_ok=True)

    presentation = Presentation()
    presentation.slide_width = SLIDE_WIDTH
    presentation.slide_height = SLIDE_HEIGHT

    tree_pages = _paginate_tree_lines(_build_directory_tree_lines(project_tree))
    total_slides = _estimate_total_slides(project_tree, len(tree_pages))
    built_slides = 0

    _add_cover_slide(presentation, project_tree.root_name)
    built_slides += 1
    if progress_callback is not None:
        progress_callback(built_slides, total_slides, project_tree.root_name)
    built_slides += _add_directory_tree_slides(presentation, tree_pages)
    if progress_callback is not None:
        progress_callback(built_slides, total_slides, "目录结构")

    for section in project_tree.sections:
        for topic in section.topics:
            added = _add_topic_slides(presentation, section.name, topic)
            built_slides += added
            if progress_callback is not None:
                progress_callback(built_slides, total_slides, f"{section.name} -- {topic.name}")

    presentation.save(output)
    return BuildResult(
        output_path=output,
        slide_count=len(presentation.slides),
        skipped_files=project_tree.skipped_files,
    )


def _estimate_total_slides(project_tree: ProjectTree, tree_slide_count: int) -> int:
    total = 1 + tree_slide_count
    for section in project_tree.sections:
        for topic in section.topics:
            total += math.ceil(len(topic.items) / _items_per_slide(len(topic.items)))
    return total


def _add_cover_slide(presentation: Presentation, root_name: str) -> None:
    slide = presentation.slides.add_slide(presentation.slide_layouts[6])
    background = slide.background.fill
    background.solid()
    background.fore_color.rgb = RGBColor(245, 247, 250)

    box = slide.shapes.add_textbox(Inches(1.2), Inches(2.1), Inches(10.9), Inches(1.8))
    frame = box.text_frame
    frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    paragraph = frame.paragraphs[0]
    paragraph.alignment = PP_ALIGN.CENTER
    run = paragraph.add_run()
    run.text = root_name
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(35, 46, 61)

    sub = slide.shapes.add_textbox(Inches(3.8), Inches(4.25), Inches(5.8), Inches(0.7))
    sub_frame = sub.text_frame
    sub_frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
    sub_p = sub_frame.paragraphs[0]
    sub_p.alignment = PP_ALIGN.CENTER
    sub_run = sub_p.add_run()
    sub_run.text = "Folder2PPT 自动生成"
    sub_run.font.name = "Microsoft YaHei"
    sub_run.font.size = Pt(14)
    sub_run.font.color.rgb = RGBColor(104, 117, 130)


def _add_directory_tree_slides(presentation: Presentation, tree_pages: list[list[str]]) -> int:
    for index, lines in enumerate(tree_pages, start=1):
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        title = "目录结构"
        if len(tree_pages) > 1:
            title = f"{title} ({index}/{len(tree_pages)})"
        _add_page_title(slide, title)

        box = slide.shapes.add_textbox(Inches(0.7), Inches(1.25), Inches(12.0), Inches(5.9))
        frame = box.text_frame
        frame.word_wrap = False
        frame.clear()
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.LEFT
        run = paragraph.add_run()
        run.text = "\n".join(lines)
        run.font.name = "Consolas"
        run.font.size = Pt(12)
        run.font.color.rgb = RGBColor(52, 60, 69)
    return len(tree_pages)


def _build_directory_tree_lines(project_tree: ProjectTree) -> list[str]:
    tree: dict[str, dict] = {}
    for section in project_tree.sections:
        for topic in section.topics:
            for item in topic.items:
                parts = item.source_path.relative_to(project_tree.root_path).parts
                node = tree
                for part in parts[:-1]:
                    node = node.setdefault(part, {})
                node[item.display_name] = None

    lines = [f"{project_tree.root_name}/"]
    lines.extend(_render_tree_nodes(tree, prefix=""))
    return lines


def _render_tree_nodes(tree: dict[str, dict | None], prefix: str) -> list[str]:
    lines: list[str] = []
    items = list(tree.items())
    for index, (name, child) in enumerate(items):
        is_last = index == len(items) - 1
        branch = "└── " if is_last else "├── "
        suffix = "/" if isinstance(child, dict) else ""
        lines.append(f"{prefix}{branch}{name}{suffix}")
        if isinstance(child, dict):
            child_prefix = f"{prefix}{'    ' if is_last else '│   '}"
            lines.extend(_render_tree_nodes(child, child_prefix))
    return lines


def _paginate_tree_lines(lines: list[str], max_lines_per_slide: int = 24) -> list[list[str]]:
    return [
        lines[index : index + max_lines_per_slide]
        for index in range(0, len(lines), max_lines_per_slide)
    ]


def _add_topic_slides(presentation: Presentation, section_name: str, topic: Topic) -> int:
    items = topic.items
    per_slide = _items_per_slide(len(items))
    total_pages = math.ceil(len(items) / per_slide)
    for page_index in range(total_pages):
        chunk = items[page_index * per_slide : (page_index + 1) * per_slide]
        slide = presentation.slides.add_slide(presentation.slide_layouts[6])
        title = f"{section_name} -- {topic.name}"
        if total_pages > 1:
            title = f"{title} ({page_index + 1}/{total_pages})"
        _add_page_title(slide, title)
        _add_items_grid(slide, chunk)
    return total_pages


def _items_per_slide(item_count: int) -> int:
    if item_count <= 4:
        return 4
    if item_count <= 6:
        return 6
    return 8


def _add_page_title(slide, title: str) -> None:
    shape = slide.shapes.add_textbox(Inches(0.7), Inches(0.35), Inches(12), Inches(0.6))
    frame = shape.text_frame
    paragraph = frame.paragraphs[0]
    run = paragraph.add_run()
    run.text = title
    run.font.name = "Microsoft YaHei"
    run.font.size = Pt(24)
    run.font.bold = True
    run.font.color.rgb = RGBColor(31, 44, 58)

    line = slide.shapes.add_shape(
        MSO_AUTO_SHAPE_TYPE.RECTANGLE,
        Inches(0.7),
        Inches(0.98),
        Inches(11.95),
        Inches(0.03),
    )
    line.fill.solid()
    line.fill.fore_color.rgb = RGBColor(210, 220, 229)
    line.line.fill.background()


def _add_items_grid(slide, items) -> None:
    count = len(items)
    columns = 2 if count <= 4 else (3 if count <= 6 else 4)
    box_width = 11.85 / columns
    image_width = box_width - 0.3
    image_height = 2.0 if columns == 4 else 2.25
    rows = math.ceil(count / columns)
    start_left = 0.7
    start_top = 1.35
    row_step = 2.75 if rows <= 2 else 2.55

    for index, item in enumerate(items):
        row = index // columns
        col = index % columns
        left = Inches(start_left + col * box_width)
        top = Inches(start_top + row * row_step)

        if item.preview_path is None:
            raise ValueError(f"缺少预览图: {item.source_path}")

        slide.shapes.add_picture(
            str(item.preview_path),
            left,
            top,
            width=Inches(image_width),
            height=Inches(image_height),
        )

        caption = slide.shapes.add_textbox(
            left,
            top + Inches(image_height + 0.06),
            Inches(image_width),
            Inches(0.46),
        )
        frame = caption.text_frame
        frame.vertical_anchor = MSO_VERTICAL_ANCHOR.MIDDLE
        paragraph = frame.paragraphs[0]
        paragraph.alignment = PP_ALIGN.CENTER
        run = paragraph.add_run()
        run.text = item.display_name
        run.font.name = "Microsoft YaHei"
        run.font.size = Pt(11 if columns == 4 else 12)
        run.font.color.rgb = RGBColor(54, 64, 74)
