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
