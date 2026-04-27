"""Generate text prompts from template metadata using Jinja2."""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from src.schema import BBox, TemplateMeta
from src.utils.config import resolve_path, load_config


def _load_style_content(style_id: str | None) -> str | None:
    """Read a style preset file from prompts/styles/{style_id}.md. Returns None if not found."""
    if not style_id:
        return None
    config = load_config()
    style_path = resolve_path(config.prompts_dir) / "styles" / f"{style_id}.md"
    if not style_path.exists():
        return None
    return style_path.read_text(encoding="utf-8")


def list_available_styles() -> list[str]:
    """Return list of style IDs (filenames without .md) from prompts/styles/."""
    config = load_config()
    styles_dir = resolve_path(config.prompts_dir) / "styles"
    if not styles_dir.exists():
        return []
    return sorted(p.stem for p in styles_dir.glob("*.md"))


def build_text_prompt(
    meta: TemplateMeta,
    task_description: str,
    n_candidates: int = 3,
    style_id: str | None = None,
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

    style_content = _load_style_content(style_id)

    return template.render(
        task_description=task_description,
        n_candidates=n_candidates,
        elements=elements_with_roles,
        style_content=style_content,
        layout_type=meta.layout_type,
    )


def build_revision_prompt(
    previous_response: str,
    revision_notes: str,
    n_candidates: int = 3,
) -> str:
    config = load_config()
    prompts_dir = resolve_path(config.prompts_dir)

    env = Environment(
        loader=FileSystemLoader(str(prompts_dir)),
        keep_trailing_newline=True,
    )
    template = env.get_template("text_revision.md.j2")

    return template.render(
        previous_response=previous_response,
        revision_notes=revision_notes,
        n_candidates=n_candidates,
    )


def _compute_aspect_ratio(bbox: BBox) -> str:
    """Map bbox width/height to a human-readable aspect ratio description."""
    w = bbox.width
    h = bbox.height
    if h == 0:
        return "未知"
    ratio = w / h
    if ratio > 1.7:
        return "16:9 横版"
    elif ratio > 1.2:
        return "4:3 横版"
    elif ratio > 0.8:
        return "1:1 正方形"
    elif ratio > 0.6:
        return "3:4 竖版"
    else:
        return "9:16 竖版"


def build_image_prompt(
    meta: TemplateMeta,
    task_description: str,
) -> str | None:
    """Generate image description prompt for image placeholders. Returns None if no image elements."""
    images = meta.editable_image_elements
    images_with_roles = [img for img in images if img.role_key]

    if not images_with_roles:
        return None

    config = load_config()
    prompts_dir = resolve_path(config.prompts_dir)

    env = Environment(
        loader=FileSystemLoader(str(prompts_dir)),
        keep_trailing_newline=True,
    )
    template = env.get_template("image_description.md.j2")

    image_data = []
    for img in images_with_roles:
        image_data.append({
            "role_key": img.role_key,
            "role_zh": img.role_zh or img.shape_name_original,
            "aspect_ratio": _compute_aspect_ratio(img.bbox),
            "current_content": img.current_content,
        })

    return template.render(
        task_description=task_description,
        images=image_data,
    )
