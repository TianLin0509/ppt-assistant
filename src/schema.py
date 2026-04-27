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
    min_chars: Optional[int] = None
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

    layout_type: Optional[str] = None

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

    style_id: Optional[str] = None
    revision_count: int = 0
    revision_notes: Optional[str] = None

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
    default_style: str = "huawei"
