# PPT Assistant Phase 1 Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the minimum viable text-fill pipeline: parse a wild .pptx template, generate AI prompts, accept pasted JSON, fill text safely (preserving formatting), and export .pptx.

**Architecture:** Local Streamlit app bridging AI Web via clipboard. python-pptx + lxml for PPT manipulation, pywin32 COM for rendering previews. Pydantic v2 for data models, Jinja2 for prompt templates.

**Tech Stack:** Python 3.12, Streamlit, python-pptx, lxml, Pillow, pywin32, Pydantic v2, Jinja2

**Project root:** `C:\Users\lintian\ppt-assistant`

---

## Task 1: Project Skeleton + Schema + Config

**Files:**
- Create: `requirements.txt`
- Create: `.gitignore`
- Create: `config.yaml`
- Create: `src/__init__.py`
- Create: `src/schema.py`
- Create: `src/utils/__init__.py`
- Create: `src/utils/config.py`
- Create: `src/core/__init__.py`
- Create: `src/prompt/__init__.py`
- Create: `src/ui/__init__.py`
- Create: `src/ui/pages/` (empty dir)
- Create: `src/ui/components/__init__.py`
- Create: `tests/__init__.py`
- Create: `tests/conftest.py`
- Create: `templates/.gitkeep`
- Create: `templates_meta/.gitkeep`
- Create: `runs/.gitkeep`
- Create: `prompts/text_generation.md.j2`
- Create: `tests/samples/.gitkeep`

- [ ] **Step 1: Create requirements.txt**

```
streamlit>=1.33.0
python-pptx>=1.0.0
lxml>=5.0.0
Pillow>=10.0.0
pywin32>=306
pydantic>=2.0.0
Jinja2>=3.1.0
st-copy-to-clipboard>=0.1.0
pyyaml>=6.0.0
pytest>=8.0.0
```

- [ ] **Step 2: Create .gitignore**

```
__pycache__/
*.pyc
.venv/
*.egg-info/
dist/
build/
.streamlit/
runs/*/
templates_meta/*.png
templates_meta/*.json
!templates_meta/.gitkeep
.superpowers/
```

- [ ] **Step 3: Create config.yaml**

```yaml
templates_dir: "templates"
templates_meta_dir: "templates_meta"
runs_dir: "runs"
prompts_dir: "prompts"
powerpoint_path: null
libreoffice_path: null
auto_save_interval_sec: 10
candidates_per_element: 3
```

- [ ] **Step 4: Create src/schema.py with all Pydantic models**

```python
"""ppt-assistant data models — Pydantic v2 strict types."""

from __future__ import annotations

from datetime import datetime
from enum import Enum
from typing import Optional

from pydantic import BaseModel, Field


class ShapeType(str, Enum):
    TEXT = "text"
    IMAGE = "image"
    DECORATION = "decoration"
    GROUP = "group"
    SMARTART = "smartart"
    CHART = "chart"
    TABLE = "table"
    UNKNOWN = "unknown"


class TextSubtype(str, Enum):
    TITLE = "title"
    BULLET = "bullet"
    BODY = "body"


class BBox(BaseModel):
    left: int
    top: int
    width: int
    height: int


class FontInfo(BaseModel):
    name: Optional[str] = None
    name_east_asian: Optional[str] = None
    size_pt: Optional[float] = None
    bold: Optional[bool] = None
    italic: Optional[bool] = None
    color_rgb: Optional[str] = None
    underline: Optional[bool] = None


class ImageSlotInfo(BaseModel):
    rotation: float = 0.0
    crop_left: float = 0.0
    crop_right: float = 0.0
    crop_top: float = 0.0
    crop_bottom: float = 0.0
    aspect_ratio: Optional[str] = None


class ShapeRole(BaseModel):
    shape_id: int
    shape_name_original: str
    shape_name_with_role: Optional[str] = None
    role_key: Optional[str] = None
    role_zh: Optional[str] = None
    role_confirmed: bool = False

    type: ShapeType
    is_editable: bool = True

    bbox: BBox
    text_hash: Optional[str] = None

    text_subtype: Optional[TextSubtype] = None
    max_chars: Optional[int] = None
    max_lines: Optional[int] = None
    current_content: Optional[str] = None
    first_run_font: Optional[FontInfo] = None
    paragraph_fonts: Optional[list[FontInfo]] = None
    paragraph_count: Optional[int] = None

    image_slot: Optional[ImageSlotInfo] = None

    is_in_group: bool = False
    group_path: Optional[str] = None
    z_order_index: Optional[int] = None

    @property
    def display_label(self) -> str:
        if self.role_zh and self.role_key:
            return f"{self.role_zh} ({self.role_key})"
        return self.shape_name_original


class TemplateMeta(BaseModel):
    template_id: str
    file_path: str
    file_mtime: float

    preview_image: Optional[str] = None
    annotated_image: Optional[str] = None

    slide_width_emu: int = 0
    slide_height_emu: int = 0

    elements: list[ShapeRole] = Field(default_factory=list)

    @property
    def editable_text_elements(self) -> list[ShapeRole]:
        return [e for e in self.elements
                if e.type == ShapeType.TEXT and e.is_editable]

    @property
    def editable_image_elements(self) -> list[ShapeRole]:
        return [e for e in self.elements
                if e.type == ShapeType.IMAGE and e.is_editable]


class TextCandidates(BaseModel):
    template_id: str
    task_description: str
    candidates: dict[str, list[str]]


class TaskRun(BaseModel):
    task_id: str
    created_at: str = Field(default_factory=lambda: datetime.now().isoformat())
    status: str = "created"
    current_step: int = 0
    task_description: str
    template_id: str
    template_mtime: float = 0.0

    text_prompt: Optional[str] = None
    image_prompt: Optional[str] = None
    ai_response_raw: Optional[str] = None
    text_candidates: Optional[TextCandidates] = None

    text_choices: dict[str, str] = Field(default_factory=dict)
    image_choices: dict[str, str] = Field(default_factory=dict)

    output_pptx: Optional[str] = None
    preview_image: Optional[str] = None

    @property
    def run_dir(self) -> str:
        return f"runs/{self.task_id}"


class AppConfig(BaseModel):
    templates_dir: str = "templates"
    templates_meta_dir: str = "templates_meta"
    runs_dir: str = "runs"
    prompts_dir: str = "prompts"
    powerpoint_path: Optional[str] = None
    libreoffice_path: Optional[str] = None
    auto_save_interval_sec: int = 10
    candidates_per_element: int = 3
```

- [ ] **Step 5: Create src/utils/config.py**

```python
"""Load and resolve application configuration."""

from pathlib import Path

import yaml

from src.schema import AppConfig

_PROJECT_ROOT = Path(__file__).resolve().parent.parent.parent


def load_config(config_path: str | None = None) -> AppConfig:
    path = Path(config_path) if config_path else _PROJECT_ROOT / "config.yaml"
    if path.exists():
        with open(path, encoding="utf-8") as f:
            data = yaml.safe_load(f) or {}
        return AppConfig(**data)
    return AppConfig()


def resolve_path(relative: str) -> Path:
    p = Path(relative)
    if p.is_absolute():
        return p
    return _PROJECT_ROOT / p
```

- [ ] **Step 6: Create all __init__.py files, .gitkeep files, and empty directories**

Create these files with empty content:
- `src/__init__.py`
- `src/core/__init__.py`
- `src/prompt/__init__.py`
- `src/ui/__init__.py`
- `src/ui/components/__init__.py`
- `src/utils/__init__.py`
- `tests/__init__.py`
- `templates/.gitkeep`
- `templates_meta/.gitkeep`
- `runs/.gitkeep`
- `tests/samples/.gitkeep`

- [ ] **Step 7: Create tests/conftest.py**

```python
"""Shared test fixtures."""

from pathlib import Path

import pytest
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor

SAMPLES_DIR = Path(__file__).parent / "samples"


@pytest.fixture
def sample_pptx(tmp_path) -> Path:
    """Create a minimal .pptx with known shapes for testing."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank layout

    # Shape 0: title text box (top area, large font)
    txbox = slide.shapes.add_textbox(Inches(1), Inches(0.3), Inches(11), Inches(1))
    txbox.name = "Title Box"
    tf = txbox.text_frame
    tf.paragraphs[0].text = ""
    run = tf.paragraphs[0].add_run()
    run.text = "Sample Title"
    run.font.size = Pt(28)
    run.font.bold = True
    run.font.color.rgb = RGBColor(0xC8, 0x31, 0x3A)
    run.font.name = "Arial"

    # Shape 1: body text box (left half, medium font)
    txbox2 = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(5), Inches(4))
    txbox2.name = "Body Left"
    tf2 = txbox2.text_frame
    tf2.paragraphs[0].text = ""
    run2 = tf2.paragraphs[0].add_run()
    run2.text = "Body content here"
    run2.font.size = Pt(14)
    run2.font.name = "Arial"

    # Shape 2: bullet text box (right half, multi-paragraph)
    txbox3 = slide.shapes.add_textbox(Inches(7), Inches(2), Inches(5), Inches(4))
    txbox3.name = "Bullet Right"
    tf3 = txbox3.text_frame
    p1 = tf3.paragraphs[0]
    p1.text = ""
    r1 = p1.add_run()
    r1.text = "Point one"
    r1.font.size = Pt(12)
    r1.font.bold = True
    p2 = tf3.add_paragraph()
    r2 = p2.add_run()
    r2.text = "Point two"
    r2.font.size = Pt(12)
    p3 = tf3.add_paragraph()
    r3 = p3.add_run()
    r3.text = "Point three"
    r3.font.size = Pt(12)

    # Shape 3: image placeholder (add a simple rectangle as stand-in)
    from pptx.enum.shapes import MSO_SHAPE
    slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(1), Inches(6.5), Inches(2), Inches(0.5))

    out = tmp_path / "test_template.pptx"
    prs.save(str(out))
    return out
```

- [ ] **Step 8: Create Jinja2 prompt template prompts/text_generation.md.j2**

```
你的任务：{{ task_description }}

请为以下 PPT 元素各生成 {{ n_candidates }} 个备选文案，严格按 JSON 格式输出。

元素列表：
{% for e in elements %}
- "{{ e.role_key }}": {{ e.role_zh }}{% if e.text_subtype %}（{{ e.text_subtype.value }}类型）{% endif %}{% if e.max_chars %}，最多{{ e.max_chars }}字{% endif %}

  原文参考：{{ e.current_content or "无" }}
{% endfor %}

输出格式要求：
1. 严格 JSON，不要有注释或额外解释
2. 每个 key 对应一个列表，列表内 {{ n_candidates }} 个字符串

```json
{
{% for e in elements %}  "{{ e.role_key }}": ["备选A", "备选B", "备选C"]{% if not loop.last %},{% endif %}

{% endfor %}}
```
```

- [ ] **Step 9: Verify schema loads correctly**

Run: `cd C:\Users\lintian\ppt-assistant && python -c "from src.schema import *; print('Schema OK:', [c.__name__ for c in [ShapeType, TextSubtype, BBox, FontInfo, ShapeRole, TemplateMeta, TextCandidates, TaskRun, AppConfig]])"`

Expected: `Schema OK: ['ShapeType', 'TextSubtype', 'BBox', 'FontInfo', 'ShapeRole', 'TemplateMeta', 'TextCandidates', 'TaskRun', 'AppConfig']`

- [ ] **Step 10: Commit**

```bash
git add -A
git commit -m "feat: project skeleton with schema, config, and test fixtures"
```

---

## Task 2: JSON Parser (TDD)

**Files:**
- Create: `src/prompt/json_parser.py`
- Create: `tests/test_json_parser.py`

- [ ] **Step 1: Write failing tests**

```python
"""Tests for TD-5 tolerant JSON parsing."""

import pytest

from src.prompt.json_parser import parse_ai_json


class TestCleanJson:
    def test_valid_json(self):
        raw = '{"title_main": ["A", "B", "C"]}'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A", "B", "C"]}

    def test_nested_candidates(self):
        raw = '{"title_main": ["Option A", "Option B", "Option C"], "body_left": ["X", "Y", "Z"]}'
        result = parse_ai_json(raw)
        assert len(result) == 2
        assert len(result["body_left"]) == 3


class TestCodeBlockExtraction:
    def test_json_code_block(self):
        raw = 'Here is the result:\n```json\n{"title_main": ["A", "B", "C"]}\n```\nHope this helps!'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A", "B", "C"]}

    def test_plain_code_block(self):
        raw = '```\n{"title_main": ["A", "B", "C"]}\n```'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A", "B", "C"]}


class TestTrailingCommas:
    def test_trailing_comma_in_object(self):
        raw = '{"title_main": ["A", "B", "C"],}'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A", "B", "C"]}

    def test_trailing_comma_in_array(self):
        raw = '{"title_main": ["A", "B", "C",]}'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A", "B", "C"]}


class TestBraceExtraction:
    def test_text_before_and_after(self):
        raw = 'The output is: {"title_main": ["A"]} end of response.'
        result = parse_ai_json(raw)
        assert result == {"title_main": ["A"]}


class TestFailure:
    def test_garbage_returns_none(self):
        result = parse_ai_json("This is not JSON at all")
        assert result is None

    def test_empty_string_returns_none(self):
        result = parse_ai_json("")
        assert result is None
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_json_parser.py -v`

Expected: ImportError — `src.prompt.json_parser` does not exist yet.

- [ ] **Step 3: Implement json_parser.py**

```python
"""TD-5: Tolerant parsing of AI-generated JSON responses."""

from __future__ import annotations

import ast
import json
import re


def parse_ai_json(raw_text: str) -> dict | None:
    if not raw_text or not raw_text.strip():
        return None

    text = raw_text.strip()

    code_block = _extract_code_block(text)
    if code_block:
        text = code_block

    text = _extract_braces(text)
    if not text:
        return None

    text = _fix_trailing_commas(text)

    try:
        return json.loads(text)
    except (json.JSONDecodeError, ValueError):
        pass

    try:
        result = ast.literal_eval(text)
        if isinstance(result, dict):
            return result
    except (ValueError, SyntaxError):
        pass

    return None


def _extract_code_block(text: str) -> str | None:
    pattern = r"```(?:json)?\s*\n?(.*?)\n?\s*```"
    match = re.search(pattern, text, re.DOTALL)
    if match:
        return match.group(1).strip()
    return None


def _extract_braces(text: str) -> str:
    first = text.find("{")
    last = text.rfind("}")
    if first == -1 or last == -1 or first >= last:
        return ""
    return text[first : last + 1]


def _fix_trailing_commas(text: str) -> str:
    text = re.sub(r",\s*}", "}", text)
    text = re.sub(r",\s*]", "]", text)
    return text
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_json_parser.py -v`

Expected: All 8 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/prompt/json_parser.py tests/test_json_parser.py
git commit -m "feat: TD-5 tolerant JSON parser with TDD"
```

---

## Task 3: PPTX Parser

**Files:**
- Create: `src/core/pptx_parser.py`
- Create: `tests/test_pptx_parser.py`

- [ ] **Step 1: Write failing tests**

```python
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_pptx_parser.py -v`

Expected: ImportError.

- [ ] **Step 3: Implement pptx_parser.py**

```python
"""Parse .pptx templates into TemplateMeta with recursive shape traversal."""

from __future__ import annotations

import hashlib
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.oxml.ns import qn
from pptx.util import Pt

from src.schema import (
    BBox, FontInfo, ImageSlotInfo, ShapeRole, ShapeType, TemplateMeta,
)


def parse_template(pptx_path: str) -> TemplateMeta:
    path = Path(pptx_path)
    prs = Presentation(str(path))
    slide = prs.slides[0]

    elements: list[ShapeRole] = []
    for z_index, shape in enumerate(slide.shapes):
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            elements.extend(_parse_group(shape, z_index, None))
        else:
            elements.append(_parse_shape(shape, z_index, None))

    template_id = path.stem.replace(" ", "_")

    return TemplateMeta(
        template_id=template_id,
        file_path=str(path.resolve()),
        file_mtime=path.stat().st_mtime,
        slide_width_emu=prs.slide_width,
        slide_height_emu=prs.slide_height,
        elements=elements,
    )


def _parse_shape(shape, z_index: int, group_path: str | None) -> ShapeRole:
    shape_type = _classify_shape_type(shape)
    bbox = BBox(left=shape.left, top=shape.top, width=shape.width, height=shape.height)

    current_content = None
    first_run_font = None
    paragraph_fonts = None
    paragraph_count = None
    text_hash = None

    if shape.has_text_frame:
        tf = shape.text_frame
        paragraphs = tf.paragraphs
        paragraph_count = len(paragraphs)

        lines = []
        p_fonts = []
        for para in paragraphs:
            lines.append(para.text)
            if para.runs:
                p_fonts.append(_snapshot_font(para.runs[0]))
            else:
                p_fonts.append(FontInfo())

        current_content = "\n".join(lines)
        text_hash = _compute_text_hash(current_content)
        first_run_font = p_fonts[0] if p_fonts else None
        paragraph_fonts = p_fonts if len(p_fonts) > 1 else None

    return ShapeRole(
        shape_id=shape.shape_id,
        shape_name_original=shape.name,
        type=shape_type,
        is_editable=shape_type in (ShapeType.TEXT, ShapeType.IMAGE),
        bbox=bbox,
        text_hash=text_hash,
        current_content=current_content,
        first_run_font=first_run_font,
        paragraph_fonts=paragraph_fonts,
        paragraph_count=paragraph_count,
        is_in_group=group_path is not None,
        group_path=group_path,
        z_order_index=z_index,
    )


def _parse_group(group_shape, z_index: int, parent_path: str | None) -> list[ShapeRole]:
    results = []
    prefix = f"{z_index:02d}" if parent_path is None else f"{parent_path}-{z_index:02d}"
    for i, shape in enumerate(group_shape.shapes):
        child_path = f"{prefix}-{i:02d}"
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            results.extend(_parse_group(shape, i, prefix))
        else:
            results.append(_parse_shape(shape, i, child_path))
    return results


def _classify_shape_type(shape) -> ShapeType:
    st = shape.shape_type
    if st == MSO_SHAPE_TYPE.GROUP:
        return ShapeType.GROUP
    if st == MSO_SHAPE_TYPE.TABLE:
        return ShapeType.TABLE
    if st == MSO_SHAPE_TYPE.CHART:
        return ShapeType.CHART
    if st in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.LINKED_PICTURE):
        return ShapeType.IMAGE
    if shape.has_text_frame and shape.text_frame.text.strip():
        return ShapeType.TEXT
    if st == MSO_SHAPE_TYPE.PLACEHOLDER and shape.has_text_frame:
        return ShapeType.TEXT
    return ShapeType.DECORATION


def _snapshot_font(run) -> FontInfo:
    font = run.font
    color_rgb = None
    if font.color and font.color.rgb:
        color_rgb = f"#{font.color.rgb}"

    name_ea = None
    rPr = run._r.find(qn("a:rPr"))
    if rPr is not None:
        ea = rPr.find(qn("a:ea"))
        if ea is not None:
            name_ea = ea.get("typeface")

    size_pt = None
    if font.size is not None:
        size_pt = font.size.pt

    return FontInfo(
        name=font.name,
        name_east_asian=name_ea,
        size_pt=size_pt,
        bold=font.bold,
        italic=font.italic,
        color_rgb=color_rgb,
        underline=font.underline,
    )


def _compute_text_hash(text: str) -> str:
    return hashlib.md5(text.strip().encode("utf-8")).hexdigest()
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_pptx_parser.py -v`

Expected: All 7 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/pptx_parser.py tests/test_pptx_parser.py
git commit -m "feat: pptx parser with recursive group traversal and font snapshot"
```

---

## Task 4: PPTX Filler (TDD) — TD-1 Core

**Files:**
- Create: `src/core/pptx_filler.py`
- Create: `tests/test_pptx_filler.py`

- [ ] **Step 1: Write failing tests**

```python
"""Tests for TD-1 run-level safe text replacement."""

import pytest
from pptx import Presentation
from pptx.util import Pt

from src.core.pptx_parser import parse_template
from src.core.pptx_filler import fill_template
from src.schema import TextSubtype


class TestTitleReplacement:
    def test_title_text_replaced(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        title = [e for e in meta.elements if e.current_content and "Title" in e.current_content][0]
        title.role_key = "title_main"
        title.text_subtype = TextSubtype.TITLE

        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"title_main": "New Title"}, str(output))

        prs = Presentation(str(output))
        slide = prs.slides[0]
        found = False
        for shape in slide.shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.text == "New Title":
                        found = True
                        assert para.runs[0].font.bold is True
                        assert para.runs[0].font.size == Pt(28)
        assert found, "Replaced title not found"

    def test_title_extra_runs_removed(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        title = [e for e in meta.elements if e.current_content and "Title" in e.current_content][0]
        title.role_key = "title_main"
        title.text_subtype = TextSubtype.TITLE

        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"title_main": "Clean"}, str(output))

        prs = Presentation(str(output))
        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    if para.text == "Clean":
                        assert len(para.runs) == 1


class TestBodyReplacement:
    def test_body_text_replaced(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        body = [e for e in meta.elements if e.current_content and "Body content" in e.current_content][0]
        body.role_key = "body_left"
        body.text_subtype = TextSubtype.BODY

        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"body_left": "Replaced body"}, str(output))

        prs = Presentation(str(output))
        texts = []
        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                for para in shape.text_frame.paragraphs:
                    texts.append(para.text)
        assert "Replaced body" in texts


class TestBulletReplacement:
    def test_bullet_lines_replaced(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        bullet = [e for e in meta.elements if e.current_content and "Point one" in e.current_content][0]
        bullet.role_key = "bullet_right"
        bullet.text_subtype = TextSubtype.BULLET

        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"bullet_right": "Line A\nLine B\nLine C"}, str(output))

        prs = Presentation(str(output))
        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                paras = [p.text for p in shape.text_frame.paragraphs]
                if "Line A" in paras:
                    assert paras == ["Line A", "Line B", "Line C"]
                    assert shape.text_frame.paragraphs[0].runs[0].font.bold is True

    def test_bullet_fewer_lines(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        bullet = [e for e in meta.elements if e.current_content and "Point one" in e.current_content][0]
        bullet.role_key = "bullet_right"
        bullet.text_subtype = TextSubtype.BULLET

        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"bullet_right": "Only one"}, str(output))

        prs = Presentation(str(output))
        for shape in prs.slides[0].shapes:
            if shape.has_text_frame:
                paras = [p.text for p in shape.text_frame.paragraphs if p.text]
                if "Only one" in paras:
                    assert len(paras) == 1


class TestNoMatchSkipped:
    def test_unknown_role_key_skipped(self, sample_pptx, tmp_path):
        meta = parse_template(str(sample_pptx))
        output = tmp_path / "filled.pptx"
        fill_template(str(sample_pptx), meta, {"nonexistent": "value"}, str(output))
        assert output.exists()
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_pptx_filler.py -v`

Expected: ImportError.

- [ ] **Step 3: Implement pptx_filler.py**

```python
"""TD-1 run-level safe text replacement + TD-4 shape multi-location."""

from __future__ import annotations

from pathlib import Path

from lxml import etree
from pptx import Presentation
from pptx.oxml.ns import qn
from pptx.util import Pt

from src.schema import FontInfo, ShapeRole, ShapeType, TemplateMeta, TextSubtype


def fill_template(
    pptx_path: str,
    meta: TemplateMeta,
    text_choices: dict[str, str],
    output_path: str,
) -> str:
    prs = Presentation(pptx_path)
    slide = prs.slides[0]

    for element in meta.elements:
        if element.type != ShapeType.TEXT or not element.role_key:
            continue
        if element.role_key not in text_choices:
            continue

        shape = _find_shape(slide, element)
        if shape is None:
            continue

        new_text = text_choices[element.role_key]
        subtype = element.text_subtype or TextSubtype.BODY

        if subtype == TextSubtype.BULLET:
            _replace_bullet(
                shape.text_frame, new_text,
                element.paragraph_fonts or ([element.first_run_font] if element.first_run_font else []),
            )
        else:
            _replace_title(shape.text_frame, new_text, element.first_run_font)

    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    prs.save(output_path)
    return output_path


def _find_shape(slide, element: ShapeRole):
    suffix = f"[{element.role_key}]"
    for shape in slide.shapes:
        if suffix in shape.name:
            return shape

    for shape in slide.shapes:
        if shape.shape_id == element.shape_id:
            return shape

    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        if (shape.left == element.bbox.left and shape.top == element.bbox.top
                and shape.width == element.bbox.width):
            return shape

    for shape in slide.shapes:
        if shape.name == element.shape_name_original:
            return shape

    return None


def _replace_title(text_frame, new_text: str, font_info: FontInfo | None) -> None:
    for para in text_frame.paragraphs:
        _clear_paragraph_runs(para)
        break

    first_para = text_frame.paragraphs[0]
    if not first_para.runs:
        first_para.add_run()
    first_para.runs[0].text = new_text
    if font_info:
        _apply_font(first_para.runs[0], font_info)

    while len(text_frame.paragraphs) > 1:
        p_elem = text_frame.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)


def _replace_bullet(text_frame, new_text: str, para_fonts: list[FontInfo]) -> None:
    lines = new_text.split("\n")
    existing = list(text_frame.paragraphs)
    txBody = text_frame._txBody

    for i, line in enumerate(lines):
        if i < len(existing):
            para = existing[i]
            _clear_paragraph_runs(para)
            if not para.runs:
                para.add_run()
            para.runs[0].text = line
        else:
            from copy import deepcopy
            template_p = existing[-1]._p
            new_p = deepcopy(template_p)
            for r in new_p.findall(qn("a:r")):
                new_p.remove(r)
            new_r = etree.SubElement(new_p, qn("a:r"))
            new_t = etree.SubElement(new_r, qn("a:t"))
            new_t.text = line
            txBody.append(new_p)

        current_para = text_frame.paragraphs[min(i, len(text_frame.paragraphs) - 1)]
        font = para_fonts[min(i, len(para_fonts) - 1)] if para_fonts else None
        if font and current_para.runs:
            _apply_font(current_para.runs[0], font)

    while len(text_frame.paragraphs) > len(lines):
        p_elem = text_frame.paragraphs[-1]._p
        p_elem.getparent().remove(p_elem)


def _clear_paragraph_runs(para) -> None:
    runs = list(para.runs)
    if not runs:
        return
    for run in runs[1:]:
        run._r.getparent().remove(run._r)
    runs[0].text = ""


def _apply_font(run, font_info: FontInfo) -> None:
    font = run.font
    if font_info.name is not None:
        font.name = font_info.name
    if font_info.size_pt is not None:
        font.size = Pt(font_info.size_pt)
    if font_info.bold is not None:
        font.bold = font_info.bold
    if font_info.italic is not None:
        font.italic = font_info.italic
    if font_info.underline is not None:
        font.underline = font_info.underline
    if font_info.color_rgb is not None:
        from pptx.dml.color import RGBColor
        hex_str = font_info.color_rgb.lstrip("#")
        font.color.rgb = RGBColor(int(hex_str[0:2], 16), int(hex_str[2:4], 16), int(hex_str[4:6], 16))

    if font_info.name_east_asian:
        rPr = run._r.get_or_add_rPr()
        ea = rPr.find(qn("a:ea"))
        if ea is None:
            ea = etree.SubElement(rPr, qn("a:ea"))
        ea.set("typeface", font_info.name_east_asian)
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/test_pptx_filler.py -v`

Expected: All 5 tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/pptx_filler.py tests/test_pptx_filler.py
git commit -m "feat: TD-1 run-level safe text filler with title/body/bullet strategies"
```

---

## Task 5: pywin32 Renderer

**Files:**
- Create: `src/core/pptx_renderer.py`

- [ ] **Step 1: Implement pptx_renderer.py**

```python
"""Render .pptx slides to PNG via PowerPoint COM automation."""

from __future__ import annotations

import atexit
from pathlib import Path

import pythoncom
import win32com.client

_ppt_app = None


def render_slide_to_png(
    pptx_path: str,
    output_png: str,
    slide_index: int = 0,
    width: int = 1920,
) -> str:
    pptx_abs = str(Path(pptx_path).resolve())
    out_abs = str(Path(output_png).resolve())
    Path(out_abs).parent.mkdir(parents=True, exist_ok=True)

    app = _get_powerpoint_app()
    pres = app.Presentations.Open(pptx_abs, WithWindow=False)
    try:
        slide = pres.Slides[slide_index + 1]  # COM is 1-indexed
        slide.Export(out_abs, "PNG", width)
    finally:
        pres.Close()

    if not Path(out_abs).exists():
        raise RuntimeError(f"PowerPoint failed to export: {out_abs}")

    return out_abs


def _get_powerpoint_app():
    global _ppt_app
    if _ppt_app is not None:
        return _ppt_app

    pythoncom.CoInitialize()
    _ppt_app = win32com.client.Dispatch("PowerPoint.Application")
    atexit.register(_ensure_powerpoint_closed)
    return _ppt_app


def _ensure_powerpoint_closed() -> None:
    global _ppt_app
    if _ppt_app is not None:
        try:
            _ppt_app.Quit()
        except Exception:
            pass
        _ppt_app = None
        pythoncom.CoUninitialize()
```

- [ ] **Step 2: Smoke test the renderer**

Run: `cd C:\Users\lintian\ppt-assistant && python -c "from src.core.pptx_renderer import render_slide_to_png; print('Renderer importable')"` — verify no import errors.

Then test with a real pptx (if one exists in templates/):
```bash
python -c "
from src.core.pptx_renderer import render_slide_to_png
import tempfile, os
tpl = r'C:\Users\lintian\ppt-templates\huawei-anchors\templates\01_two_column_compare.pptx'
out = os.path.join(tempfile.gettempdir(), 'test_render.png')
result = render_slide_to_png(tpl, out)
print(f'Rendered to: {result}')
print(f'Size: {os.path.getsize(result)} bytes')
"
```

Expected: PNG file created, size > 0.

- [ ] **Step 3: Commit**

```bash
git add src/core/pptx_renderer.py
git commit -m "feat: pywin32 COM renderer for slide-to-PNG export"
```

---

## Task 6: Shape Annotator + Role Inferencer

**Files:**
- Create: `src/core/shape_annotator.py`
- Create: `src/core/role_inferencer.py`

- [ ] **Step 1: Implement shape_annotator.py**

```python
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
```

- [ ] **Step 2: Implement role_inferencer.py**

```python
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
```

- [ ] **Step 3: Smoke test annotator with a real template**

```bash
cd C:\Users\lintian\ppt-assistant && python -c "
from src.core.pptx_parser import parse_template
from src.core.pptx_renderer import render_slide_to_png
from src.core.shape_annotator import annotate_preview
from src.core.role_inferencer import infer_roles
import tempfile, os

tpl = r'C:\Users\lintian\ppt-templates\huawei-anchors\templates\01_two_column_compare.pptx'
meta = parse_template(tpl)
meta.elements = infer_roles(meta.elements, meta.slide_width_emu, meta.slide_height_emu)

preview = os.path.join(tempfile.gettempdir(), 'preview.png')
render_slide_to_png(tpl, preview)
annotated = annotate_preview(preview, meta.elements, meta.slide_width_emu, meta.slide_height_emu)
print(f'Annotated: {annotated}')
print(f'Elements: {len(meta.elements)}')
for e in meta.elements:
    print(f'  {e.shape_name_original}: type={e.type.value}, role={e.role_key}, subtype={e.text_subtype}')
"
```

Expected: Annotated PNG created, elements listed with inferred roles.

- [ ] **Step 4: Commit**

```bash
git add src/core/shape_annotator.py src/core/role_inferencer.py
git commit -m "feat: shape annotator (PIL red boxes) and rule-based role inferencer"
```

---

## Task 7: Text Prompt Builder

**Files:**
- Create: `src/prompt/text_prompt_builder.py`

- [ ] **Step 1: Implement text_prompt_builder.py**

```python
"""Generate text prompts from template metadata using Jinja2."""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from src.schema import TemplateMeta
from src.utils.config import resolve_path, load_config


def build_text_prompt(
    meta: TemplateMeta,
    task_description: str,
    n_candidates: int = 3,
) -> str:
    config = load_config()
    prompts_dir = resolve_path(config.prompts_dir)

    env = Environment(
        loader=FileSystemLoader(str(prompts_dir)),
        keep_trailing_newline=True,
    )
    template = env.get_template("text_generation.md.j2")

    elements = meta.editable_text_elements
    elements_with_roles = [e for e in elements if e.role_key]

    return template.render(
        task_description=task_description,
        n_candidates=n_candidates,
        elements=elements_with_roles,
    )
```

- [ ] **Step 2: Test prompt generation**

```bash
cd C:\Users\lintian\ppt-assistant && python -c "
from src.core.pptx_parser import parse_template
from src.core.role_inferencer import infer_roles
from src.prompt.text_prompt_builder import build_text_prompt

tpl = r'C:\Users\lintian\ppt-templates\huawei-anchors\templates\01_two_column_compare.pptx'
meta = parse_template(tpl)
meta.elements = infer_roles(meta.elements, meta.slide_width_emu, meta.slide_height_emu)

prompt = build_text_prompt(meta, '介绍最新的 L2O 优化算法')
print(prompt[:500])
print('...')
print(f'Total length: {len(prompt)} chars')
"
```

Expected: Rendered prompt with element list and JSON format example.

- [ ] **Step 3: Commit**

```bash
git add src/prompt/text_prompt_builder.py
git commit -m "feat: Jinja2-based text prompt builder"
```

---

## Task 8: State Manager

**Files:**
- Create: `src/utils/state_manager.py`

- [ ] **Step 1: Implement state_manager.py**

```python
"""TD-6/TD-7: Task state persistence and auto-save support."""

from __future__ import annotations

import json
from pathlib import Path

from src.schema import TaskRun, TemplateMeta
from src.utils.config import resolve_path, load_config


def save_task(task: TaskRun) -> Path:
    config = load_config()
    run_dir = resolve_path(config.runs_dir) / task.task_id
    run_dir.mkdir(parents=True, exist_ok=True)

    task_file = run_dir / "task.json"
    task_file.write_text(
        task.model_dump_json(indent=2),
        encoding="utf-8",
    )
    return task_file


def load_task(task_id: str) -> TaskRun | None:
    config = load_config()
    task_file = resolve_path(config.runs_dir) / task_id / "task.json"
    if not task_file.exists():
        return None
    data = json.loads(task_file.read_text(encoding="utf-8"))
    return TaskRun(**data)


def save_raw_response(task_id: str, raw: str) -> Path:
    config = load_config()
    run_dir = resolve_path(config.runs_dir) / task_id
    run_dir.mkdir(parents=True, exist_ok=True)
    out = run_dir / "ai_response_raw.txt"
    out.write_text(raw, encoding="utf-8")
    return out


def save_template_meta(meta: TemplateMeta) -> Path:
    config = load_config()
    meta_dir = resolve_path(config.templates_meta_dir)
    meta_dir.mkdir(parents=True, exist_ok=True)
    out = meta_dir / f"{meta.template_id}.json"
    out.write_text(
        meta.model_dump_json(indent=2),
        encoding="utf-8",
    )
    return out


def load_template_meta(template_id: str) -> TemplateMeta | None:
    config = load_config()
    meta_file = resolve_path(config.templates_meta_dir) / f"{template_id}.json"
    if not meta_file.exists():
        return None
    data = json.loads(meta_file.read_text(encoding="utf-8"))
    return TemplateMeta(**data)


def list_template_metas() -> list[TemplateMeta]:
    config = load_config()
    meta_dir = resolve_path(config.templates_meta_dir)
    if not meta_dir.exists():
        return []
    results = []
    for f in sorted(meta_dir.glob("*.json")):
        data = json.loads(f.read_text(encoding="utf-8"))
        results.append(TemplateMeta(**data))
    return results
```

- [ ] **Step 2: Commit**

```bash
git add src/utils/state_manager.py
git commit -m "feat: state manager for task/template persistence (TD-6/TD-7)"
```

---

## Task 9: Streamlit App Shell + Template Library Page

**Files:**
- Create: `src/ui/app.py`
- Create: `src/ui/pages/1_template_library.py`
- Create: `src/ui/components/template_card.py`

- [ ] **Step 1: Create src/ui/app.py**

```python
"""Streamlit main entry point."""

import streamlit as st

st.set_page_config(
    page_title="PPT Assistant",
    page_icon="📊",
    layout="wide",
)

st.title("PPT Assistant — 单页填充助手")
st.markdown("从左侧导航选择功能页面。")
st.markdown("""
- **模板管理**: 导入模板、预处理、标注元素角色
- **新建任务**: 选模板 → 生成 prompt → 粘贴 AI 结果 → 挑选 → 导出
- **历史任务**: 查看过往任务记录
""")
```

- [ ] **Step 2: Create src/ui/components/template_card.py**

```python
"""Template thumbnail card component."""

from __future__ import annotations

from pathlib import Path

import streamlit as st

from src.schema import TemplateMeta


def render_template_card(meta: TemplateMeta, col) -> bool:
    with col:
        if meta.annotated_image and Path(meta.annotated_image).exists():
            st.image(meta.annotated_image, use_container_width=True)
        elif meta.preview_image and Path(meta.preview_image).exists():
            st.image(meta.preview_image, use_container_width=True)
        else:
            st.info("无预览图")

        st.caption(meta.template_id)
        n_text = len(meta.editable_text_elements)
        n_img = len(meta.editable_image_elements)
        st.caption(f"文本: {n_text} | 图片: {n_img}")

        return st.button("选择", key=f"select_{meta.template_id}")
```

- [ ] **Step 3: Create src/ui/pages/1_template_library.py**

```python
"""Template library: scan, preprocess, review roles."""

import streamlit as st
from pathlib import Path

from src.utils.config import load_config, resolve_path
from src.utils.state_manager import save_template_meta, load_template_meta, list_template_metas
from src.core.pptx_parser import parse_template
from src.core.pptx_renderer import render_slide_to_png
from src.core.shape_annotator import annotate_preview
from src.core.role_inferencer import infer_roles
from src.schema import ShapeType, TextSubtype

st.set_page_config(page_title="模板管理", layout="wide")
st.header("模板管理")

config = load_config()
tpl_dir = resolve_path(config.templates_dir)
meta_dir = resolve_path(config.templates_meta_dir)

if not tpl_dir.exists():
    st.warning(f"模板目录不存在: {tpl_dir}")
    st.stop()

pptx_files = sorted(tpl_dir.glob("*.pptx"))
pptx_files = [f for f in pptx_files if not f.name.startswith("~$")]

if not pptx_files:
    st.info("templates/ 目录下没有 .pptx 文件。请先放入模板。")
    st.stop()

st.subheader(f"找到 {len(pptx_files)} 个模板")

cols = st.columns(3)
for i, pptx_file in enumerate(pptx_files):
    template_id = pptx_file.stem.replace(" ", "_")
    meta = load_template_meta(template_id)

    with cols[i % 3]:
        st.markdown(f"**{pptx_file.name}**")

        if meta and meta.annotated_image and Path(meta.annotated_image).exists():
            st.image(meta.annotated_image, use_container_width=True)
            st.success(f"已预处理 ({len(meta.editable_text_elements)} 文本, {len(meta.editable_image_elements)} 图片)")
        else:
            st.info("未预处理")

        if st.button("预处理", key=f"preprocess_{template_id}"):
            with st.spinner("解析中..."):
                meta = parse_template(str(pptx_file))

                preview_path = str(meta_dir / f"{template_id}_preview.png")
                render_slide_to_png(str(pptx_file), preview_path)
                meta.preview_image = preview_path

                meta.elements = infer_roles(meta.elements, meta.slide_width_emu, meta.slide_height_emu)

                annotated_path = annotate_preview(
                    preview_path, meta.elements,
                    meta.slide_width_emu, meta.slide_height_emu,
                )
                meta.annotated_image = annotated_path

                save_template_meta(meta)
                st.success("预处理完成!")
                st.rerun()

st.divider()

st.subheader("角色修正")
all_metas = list_template_metas()
if not all_metas:
    st.info("暂无已预处理的模板。请先点击上方「预处理」。")
    st.stop()

selected_id = st.selectbox("选择模板", [m.template_id for m in all_metas])
meta = load_template_meta(selected_id)
if meta is None:
    st.stop()

if meta.annotated_image and Path(meta.annotated_image).exists():
    st.image(meta.annotated_image, use_container_width=True, caption="标注图（红框编号 = 元素序号）")

ROLE_OPTIONS = [
    "", "title_main", "subtitle", "body_01", "body_02", "body_03",
    "bullet_01", "bullet_02", "bullet_03",
    "image_01", "image_02", "image_03", "decoration",
]
SUBTYPE_OPTIONS = ["", "title", "bullet", "body"]

changed = False
for idx, el in enumerate(meta.elements):
    if el.type in (ShapeType.GROUP, ShapeType.SMARTART, ShapeType.CHART, ShapeType.TABLE):
        st.markdown(f"**#{idx}** {el.shape_name_original} — :red[{el.type.value}，不可编辑]")
        continue

    c1, c2, c3, c4 = st.columns([2, 2, 1, 1])
    with c1:
        st.text(f"#{idx} {el.shape_name_original}")
        if el.current_content:
            st.caption(el.current_content[:60])
    with c2:
        current_role = el.role_key or ""
        new_role = st.selectbox(
            "角色", ROLE_OPTIONS,
            index=ROLE_OPTIONS.index(current_role) if current_role in ROLE_OPTIONS else 0,
            key=f"role_{idx}",
        )
        if new_role != current_role:
            el.role_key = new_role or None
            changed = True
    with c3:
        current_sub = el.text_subtype.value if el.text_subtype else ""
        new_sub = st.selectbox(
            "子类型", SUBTYPE_OPTIONS,
            index=SUBTYPE_OPTIONS.index(current_sub) if current_sub in SUBTYPE_OPTIONS else 0,
            key=f"subtype_{idx}",
        )
        if new_sub != current_sub:
            el.text_subtype = TextSubtype(new_sub) if new_sub else None
            changed = True
    with c4:
        el.role_confirmed = st.checkbox("确认", value=el.role_confirmed, key=f"confirm_{idx}")

if st.button("保存角色配置"):
    save_template_meta(meta)
    st.success("已保存!")
```

- [ ] **Step 4: Verify Streamlit app launches**

Run: `cd C:\Users\lintian\ppt-assistant && python -m streamlit run src/ui/app.py --server.headless true`

Expected: Server starts on localhost:8501. Open in browser, see landing page. Navigate to "模板管理" page.

- [ ] **Step 5: Commit**

```bash
git add src/ui/app.py src/ui/pages/1_template_library.py src/ui/components/template_card.py
git commit -m "feat: Streamlit app shell and template library page with preprocessing"
```

---

## Task 10: Candidate Picker + New Task Page

**Files:**
- Create: `src/ui/components/candidate_picker.py`
- Create: `src/ui/pages/2_new_task.py`

- [ ] **Step 1: Create src/ui/components/candidate_picker.py**

```python
"""Candidate selection component with radio + inline edit."""

from __future__ import annotations

import streamlit as st

from src.schema import TextCandidates, TemplateMeta


def render_candidate_picker(
    candidates: TextCandidates,
    meta: TemplateMeta,
) -> dict[str, str]:
    elements = meta.editable_text_elements
    role_elements = {e.role_key: e for e in elements if e.role_key}

    if st.button("一键全选 A"):
        for role_key in candidates.candidates:
            st.session_state[f"pick_{role_key}"] = "A"
        st.rerun()

    choices: dict[str, str] = {}

    for role_key, options in candidates.candidates.items():
        el = role_elements.get(role_key)
        label = el.display_label if el else role_key

        st.markdown(f"#### {label}")

        option_labels = [chr(65 + i) for i in range(len(options))]

        if f"pick_{role_key}" not in st.session_state:
            st.session_state[f"pick_{role_key}"] = "A"

        choice = st.radio(
            f"选择 {role_key}",
            option_labels,
            key=f"pick_{role_key}",
            horizontal=True,
            label_visibility="collapsed",
        )

        choice_idx = ord(choice) - 65
        selected_text = options[choice_idx] if choice_idx < len(options) else options[0]

        edit_key = f"text_{role_key}_{choice}"
        if edit_key not in st.session_state:
            st.session_state[edit_key] = selected_text

        edited = st.text_area(
            f"编辑 {role_key}",
            value=st.session_state[edit_key],
            key=f"ta_{role_key}_{choice}",
            label_visibility="collapsed",
            height=80,
        )
        st.session_state[edit_key] = edited
        choices[role_key] = edited

        st.divider()

    return choices
```

- [ ] **Step 2: Create src/ui/pages/2_new_task.py**

```python
"""New task workflow: select template -> generate prompt -> paste JSON -> pick candidates -> export."""

import json
from datetime import datetime
from pathlib import Path

import streamlit as st

from src.utils.config import load_config, resolve_path
from src.utils.state_manager import (
    list_template_metas, load_template_meta, save_task, save_raw_response,
)
from src.prompt.text_prompt_builder import build_text_prompt
from src.prompt.json_parser import parse_ai_json
from src.core.pptx_filler import fill_template
from src.core.pptx_renderer import render_slide_to_png
from src.schema import TaskRun, TextCandidates
from src.ui.components.candidate_picker import render_candidate_picker

try:
    from st_copy_to_clipboard import st_copy_to_clipboard
except ImportError:
    st_copy_to_clipboard = None

st.set_page_config(page_title="新建任务", layout="wide")
st.header("新建任务")

config = load_config()
all_metas = list_template_metas()

if not all_metas:
    st.warning("暂无已预处理的模板。请先在「模板管理」页面预处理模板。")
    st.stop()

# --- Step 1: Select template ---
st.subheader("Step 1: 选择模板")
template_ids = [m.template_id for m in all_metas]
selected_id = st.selectbox("模板", template_ids)
meta = load_template_meta(selected_id)

if meta and meta.annotated_image and Path(meta.annotated_image).exists():
    st.image(meta.annotated_image, use_container_width=True)

# --- Step 2: Task description ---
st.subheader("Step 2: 任务描述")
task_desc = st.text_area("描述你的 PPT 主题", placeholder="例如：介绍最新的 L2O 优化算法")

if not task_desc:
    st.stop()

# --- Initialize task ---
if "current_task" not in st.session_state:
    now = datetime.now()
    task_id = f"{now.strftime('%Y-%m-%d')}_{task_desc[:20].replace(' ', '_')}"
    st.session_state["current_task"] = TaskRun(
        task_id=task_id,
        task_description=task_desc,
        template_id=selected_id,
        template_mtime=meta.file_mtime if meta else 0,
    )

task: TaskRun = st.session_state["current_task"]

# --- Step 3: Generate prompt ---
st.subheader("Step 3: 生成 Prompt → 复制到 AI Web")
if meta:
    prompt = build_text_prompt(meta, task_desc, config.candidates_per_element)
    task.text_prompt = prompt
    task.status = "prompting"
    task.current_step = 3

    st.code(prompt, language="markdown")
    if st_copy_to_clipboard:
        st_copy_to_clipboard(prompt, "复制 Prompt")
    else:
        st.info("安装 st-copy-to-clipboard 后可一键复制。手动选中上方文本复制。")

# --- Step 4: Paste JSON ---
st.subheader("Step 4: 粘贴 AI 返回的 JSON")
raw_json = st.text_area("粘贴 AI 返回结果", height=200, key="raw_json_input")

if not raw_json:
    st.stop()

parsed = parse_ai_json(raw_json)
if parsed is None:
    st.error("无法解析 JSON。请检查格式后重试。原文已保存。")
    task.ai_response_raw = raw_json
    save_raw_response(task.task_id, raw_json)
    st.stop()

task.ai_response_raw = raw_json
task.text_candidates = TextCandidates(
    template_id=selected_id,
    task_description=task_desc,
    candidates=parsed,
)
task.status = "selecting"
task.current_step = 5

# --- Step 5: Pick candidates ---
st.subheader("Step 5: 挑选备选")
choices = render_candidate_picker(task.text_candidates, meta)
task.text_choices = choices

# --- Step 6: Generate preview ---
st.subheader("Step 6: 生成预览")
if st.button("生成预览"):
    task.status = "rendering"
    task.current_step = 6

    runs_dir = resolve_path(config.runs_dir) / task.task_id
    runs_dir.mkdir(parents=True, exist_ok=True)

    output_pptx = str(runs_dir / "output.pptx")
    fill_template(meta.file_path, meta, task.text_choices, output_pptx)
    task.output_pptx = output_pptx

    preview_png = str(runs_dir / "preview.png")
    render_slide_to_png(output_pptx, preview_png)
    task.preview_image = preview_png

    task.status = "completed"
    task.current_step = 7
    save_task(task)

    st.image(preview_png, caption="预览", use_container_width=True)

# --- Step 7: Download ---
if task.output_pptx and Path(task.output_pptx).exists():
    st.subheader("Step 7: 导出")
    with open(task.output_pptx, "rb") as f:
        st.download_button(
            "下载 .pptx",
            data=f,
            file_name=f"{task.task_id}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )
```

- [ ] **Step 3: Create stub page for task history**

Create `src/ui/pages/3_task_history.py`:

```python
"""Task history viewer (Phase 1 minimal stub)."""

import streamlit as st
from pathlib import Path

from src.utils.config import load_config, resolve_path

st.set_page_config(page_title="历史任务", layout="wide")
st.header("历史任务")

config = load_config()
runs_dir = resolve_path(config.runs_dir)

if not runs_dir.exists():
    st.info("暂无历史任务。")
    st.stop()

task_dirs = sorted([d for d in runs_dir.iterdir() if d.is_dir()], reverse=True)
if not task_dirs:
    st.info("暂无历史任务。")
    st.stop()

for d in task_dirs:
    task_json = d / "task.json"
    preview = d / "preview.png"
    output = d / "output.pptx"

    with st.expander(d.name):
        if preview.exists():
            st.image(str(preview), use_container_width=True)
        if output.exists():
            with open(str(output), "rb") as f:
                st.download_button("下载 .pptx", data=f, file_name=f"{d.name}.pptx", key=f"dl_{d.name}")
```

- [ ] **Step 4: Test the full flow in browser**

Run: `cd C:\Users\lintian\ppt-assistant && python -m streamlit run src/ui/app.py`

Manual test checklist:
1. Go to "模板管理" → click "预处理" on a template → see annotated image
2. Review/modify roles → click "保存角色配置"
3. Go to "新建任务" → select template → enter task description
4. Copy generated prompt → manually write JSON matching the template's role_keys → paste
5. Pick candidates → click "生成预览" → see preview image
6. Click "下载 .pptx" → open in PowerPoint → verify formatting preserved

- [ ] **Step 5: Commit**

```bash
git add src/ui/components/candidate_picker.py src/ui/pages/2_new_task.py src/ui/pages/3_task_history.py
git commit -m "feat: new task workflow with candidate picker and task history stub"
```

---

## Task 11: Auto-Save Fragment

**Files:**
- Modify: `src/ui/pages/2_new_task.py`

- [ ] **Step 1: Add auto-save fragment to new_task page**

Add at the end of `src/ui/pages/2_new_task.py`, before the download section:

```python
# --- Auto-save ---
@st.fragment(run_every=config.auto_save_interval_sec)
def auto_save():
    if "current_task" in st.session_state:
        save_task(st.session_state["current_task"])

auto_save()
```

- [ ] **Step 2: Commit**

```bash
git add src/ui/pages/2_new_task.py
git commit -m "feat: TD-7 auto-save via st.fragment(run_every=10)"
```

---

## Task 12: End-to-End Smoke Test

**Files:**
- No new files — manual verification

- [ ] **Step 1: Install dependencies**

Run: `cd C:\Users\lintian\ppt-assistant && pip install -r requirements.txt`

- [ ] **Step 2: Run all unit tests**

Run: `cd C:\Users\lintian\ppt-assistant && python -m pytest tests/ -v`

Expected: All tests pass (json_parser: 8, pptx_parser: 7, pptx_filler: 5 = 20 total).

- [ ] **Step 3: Copy a wild template into templates/**

```bash
cp "C:/Users/lintian/ppt-templates/huawei-anchors/templates/01_two_column_compare.pptx" "C:/Users/lintian/ppt-assistant/templates/"
```

- [ ] **Step 4: Launch Streamlit and run full flow**

Run: `cd C:\Users\lintian\ppt-assistant && python -m streamlit run src/ui/app.py`

Execute the complete workflow:
1. 模板管理 → 预处理 01_two_column_compare
2. Review roles, confirm all
3. 新建任务 → select template → enter "介绍 L2O 优化算法"
4. Copy prompt, manually create matching JSON, paste
5. Pick candidates, generate preview, download .pptx
6. Open in PowerPoint, verify: text replaced, fonts/colors/sizes preserved, no visual breakage

- [ ] **Step 5: Final commit**

```bash
git add -A
git commit -m "chore: Phase 1 complete — minimum viable text-fill pipeline"
```
