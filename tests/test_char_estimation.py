"""Tests for bbox-based character count estimation."""

from src.schema import BBox, FontInfo, ShapeRole, ShapeType, TextSubtype
from src.core.role_inferencer import estimate_char_capacity, infer_roles


def test_estimate_basic():
    """A 200pt x 100pt box at 12pt: ~16 chars/line x ~6.6 lines ≈ 111 theoretical."""
    bbox = BBox(left=0, top=0, width=200 * 12700, height=100 * 12700)
    min_c, max_c = estimate_char_capacity(bbox, font_size_pt=12.0)
    assert max_c > min_c > 0
    assert 70 <= min_c <= 100
    assert 90 <= max_c <= 120


def test_estimate_title_single_line():
    """Title should be capped to single line."""
    bbox = BBox(left=0, top=0, width=400 * 12700, height=50 * 12700)
    min_c, max_c = estimate_char_capacity(bbox, font_size_pt=24.0, is_title=True)
    assert max_c <= 20


def test_estimate_tiny_box():
    """Very small box should return at least min_chars=1."""
    bbox = BBox(left=0, top=0, width=20 * 12700, height=20 * 12700)
    min_c, max_c = estimate_char_capacity(bbox, font_size_pt=12.0)
    assert min_c >= 1
    assert max_c >= min_c


def test_estimate_large_font():
    """Large font means fewer chars."""
    bbox = BBox(left=0, top=0, width=300 * 12700, height=60 * 12700)
    min_small, max_small = estimate_char_capacity(bbox, font_size_pt=10.0)
    min_large, max_large = estimate_char_capacity(bbox, font_size_pt=24.0)
    assert max_small > max_large


def test_infer_roles_uses_bbox_estimation():
    """infer_roles should populate min_chars from bbox, not 2x heuristic."""
    elements = [
        ShapeRole(
            shape_id=1,
            shape_name_original="Title 1",
            type=ShapeType.TEXT,
            bbox=BBox(left=0, top=0, width=400 * 12700, height=50 * 12700),
            first_run_font=FontInfo(size_pt=24.0),
            current_content="Short",
            paragraph_count=1,
        ),
    ]
    slide_w = 9144000
    slide_h = 6858000
    result = infer_roles(elements, slide_w, slide_h)
    el = result[0]
    assert el.min_chars is not None
    assert el.min_chars > 0
    assert el.max_chars != max(len("Short") * 2, 50)
