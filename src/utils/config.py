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
