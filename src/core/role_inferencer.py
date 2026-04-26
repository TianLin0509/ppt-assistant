"""Rule-based role inference for shape elements."""

from __future__ import annotations

from src.schema import ShapeRole, ShapeType, TextSubtype


def infer_roles(
    elements: list[ShapeRole],
    slide_w_emu: int,
    slide_h_emu: int,
) -> list[ShapeRole]:
    title_counter = 0
    body_counter = 0
    image_counter = 0

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

        if el.current_content:
            el.max_chars = max(len(el.current_content) * 2, 50)
            el.max_lines = el.paragraph_count or 1

    return elements
