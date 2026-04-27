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


def check_template_consistency(meta: TemplateMeta) -> bool:
    """Return True if template file mtime matches metadata snapshot. False if stale or file missing."""
    template_path = Path(meta.file_path)
    if not template_path.exists():
        return False
    return abs(template_path.stat().st_mtime - meta.file_mtime) < 0.01
