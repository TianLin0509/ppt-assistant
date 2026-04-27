"""Generate text prompts from template metadata using Jinja2."""

from __future__ import annotations

from pathlib import Path

from jinja2 import Environment, FileSystemLoader

from src.schema import TemplateMeta
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
