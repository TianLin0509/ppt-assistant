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
