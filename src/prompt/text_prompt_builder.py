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
