"""Tests for pptx_parser — shape extraction and metadata."""

import pytest

from src.core.pptx_parser import parse_template
from src.schema import ShapeType


class TestParseTemplate:
    def test_returns_template_meta(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        assert meta.template_id == "test_template"
        assert meta.file_mtime > 0
        assert meta.slide_width_emu > 0
        assert meta.slide_height_emu > 0

    def test_finds_all_shapes(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        assert len(meta.elements) == 4  # title + body + bullet + rectangle

    def test_text_shapes_have_content(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        text_elements = [e for e in meta.elements if e.type == ShapeType.TEXT]
        assert len(text_elements) == 3
        contents = [e.current_content for e in text_elements]
        assert "Sample Title" in contents

    def test_shape_has_bbox(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        for el in meta.elements:
            assert el.bbox.width > 0
            assert el.bbox.height > 0

    def test_font_snapshot_captured(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        title = [e for e in meta.elements if e.current_content and "Title" in e.current_content][0]
        assert title.first_run_font is not None
        assert title.first_run_font.bold is True
        assert title.first_run_font.size_pt == 28.0

    def test_bullet_has_paragraph_fonts(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        bullet = [e for e in meta.elements if e.current_content and "Point one" in e.current_content][0]
        assert bullet.paragraph_count == 3
        assert bullet.paragraph_fonts is not None
        assert len(bullet.paragraph_fonts) == 3

    def test_text_hash_computed(self, sample_pptx):
        meta = parse_template(str(sample_pptx))
        text_elements = [e for e in meta.elements if e.type == ShapeType.TEXT]
        for el in text_elements:
            assert el.text_hash is not None
            assert len(el.text_hash) == 32  # MD5 hex digest
