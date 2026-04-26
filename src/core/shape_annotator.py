"""Draw red rectangles and numbered labels on a preview PNG."""

from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

from src.schema import ShapeRole, ShapeType


def annotate_preview(
    preview_png: str,
    elements: list[ShapeRole],
    slide_w_emu: int,
    slide_h_emu: int,
) -> str:
    img = Image.open(preview_png).convert("RGB")
    draw = ImageDraw.Draw(img)
    img_w, img_h = img.size

    try:
        label_font = ImageFont.truetype("arial.ttf", 16)
    except OSError:
        label_font = ImageFont.load_default()

    idx = 0
    for el in elements:
        if el.type in (ShapeType.DECORATION, ShapeType.GROUP):
            continue

        x1 = _emu_to_pixel(el.bbox.left, slide_w_emu, img_w)
        y1 = _emu_to_pixel(el.bbox.top, slide_h_emu, img_h)
        x2 = x1 + _emu_to_pixel(el.bbox.width, slide_w_emu, img_w)
        y2 = y1 + _emu_to_pixel(el.bbox.height, slide_h_emu, img_h)

        draw.rectangle([x1, y1, x2, y2], outline="red", width=2)

        label = str(idx)
        bbox = label_font.getbbox(label)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        draw.rectangle([x1, y1, x1 + tw + 6, y1 + th + 4], fill="red")
        draw.text((x1 + 3, y1 + 1), label, fill="white", font=label_font)

        idx += 1

    out_path = str(Path(preview_png).with_stem(Path(preview_png).stem + "_annotated"))
    img.save(out_path)
    return out_path


def _emu_to_pixel(emu_val: int, emu_total: int, pixel_total: int) -> int:
    if emu_total == 0:
        return 0
    return int(emu_val * pixel_total / emu_total)
