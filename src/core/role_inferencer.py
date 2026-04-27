"""Rule-based role inference for shape elements."""

from __future__ import annotations

from statistics import median

from src.schema import BBox, ShapeRole, ShapeType, TextSubtype

EMU_PER_PT = 12700


def estimate_char_capacity(
    bbox: BBox,
    font_size_pt: float,
    is_title: bool = False,
) -> tuple[int, int]:
    """Estimate (min_chars, max_chars) a text box can hold based on physical dimensions."""
    bbox_width_pt = bbox.width / EMU_PER_PT
    bbox_height_pt = bbox.height / EMU_PER_PT

    if font_size_pt <= 0:
        font_size_pt = 12.0

    chars_per_line = bbox_width_pt / font_size_pt
    line_height = font_size_pt * 1.25
    lines = bbox_height_pt / line_height

    if is_title:
        lines = 1.0

    theoretical_max = chars_per_line * lines
    max_chars = max(int(theoretical_max * 0.90), 1)
    min_chars = max(int(theoretical_max * 0.70), 1)

    return (min_chars, max_chars)


def _resolve_font_size(el: ShapeRole, fallback_pt: float) -> float:
    """Get font size: element's own font > fallback median > 12pt."""
    if el.first_run_font and el.first_run_font.size_pt:
        return el.first_run_font.size_pt
    return fallback_pt


def infer_roles(
    elements: list[ShapeRole],
    slide_w_emu: int,
    slide_h_emu: int,
) -> list[ShapeRole]:
    title_counter = 0
    body_counter = 0
    image_counter = 0

    body_font_sizes = []
    for el in elements:
        if el.type == ShapeType.TEXT and el.first_run_font and el.first_run_font.size_pt:
            if el.first_run_font.size_pt < 24:
                body_font_sizes.append(el.first_run_font.size_pt)
    fallback_font_pt = median(body_font_sizes) if body_font_sizes else 12.0

    for el in elements:
        if not el.is_editable:
            continue

        if el.type == ShapeType.IMAGE:
            image_counter += 1
            el.role_key = f"image_{image_counter:02d}"
            el.role_zh = f"图片{image_counter}"
            el.role_confirmed = False
            continue

        if el.type != ShapeType.TEXT:
            continue

        font_size = el.first_run_font.size_pt if el.first_run_font else None
        area = el.bbox.width * el.bbox.height
        slide_area = slide_w_emu * slide_h_emu
        area_ratio = area / slide_area if slide_area > 0 else 0

        if area_ratio < 0.02:
            el.type = ShapeType.DECORATION
            el.is_editable = False
            continue

        in_top_quarter = el.bbox.top < slide_h_emu * 0.25

        if font_size and font_size >= 24 and in_top_quarter:
            title_counter += 1
            el.role_key = "title_main" if title_counter == 1 else f"title_{title_counter:02d}"
            el.role_zh = "主标题" if title_counter == 1 else f"标题{title_counter}"
            el.text_subtype = TextSubtype.TITLE
        elif el.paragraph_count and el.paragraph_count > 1:
            body_counter += 1
            el.role_key = f"bullet_{body_counter:02d}"
            el.role_zh = f"要点{body_counter}"
            el.text_subtype = TextSubtype.BULLET
        else:
            body_counter += 1
            el.role_key = f"body_{body_counter:02d}"
            el.role_zh = f"正文{body_counter}"
            el.text_subtype = TextSubtype.BODY

        el.role_confirmed = False

        is_title = el.text_subtype == TextSubtype.TITLE
        resolved_font = _resolve_font_size(el, fallback_font_pt)
        el.min_chars, el.max_chars = estimate_char_capacity(el.bbox, resolved_font, is_title=is_title)
        el.max_lines = el.paragraph_count or 1

    return elements
