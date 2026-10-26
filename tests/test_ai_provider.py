"""Synthetic provider tests: never send real keys or documents in CI."""
from types import SimpleNamespace
from unittest.mock import MagicMock

import pytest

from deallens.ai_provider import (
    PROVIDERS, DEFAULT_MODELS, ENVIRONMENT_KEYS, parse_json, generate_json
)


def test_all_three_providers_declared():
    assert PROVIDERS == ("OpenAI", "Claude", "Gemini")
    assert all(DEFAULT_MODELS[p] and ENVIRONMENT_KEYS[p] for p in PROVIDERS)


def test_json_parser_handles_fences_and_rejects_invalid():
    assert parse_json('{"company": "Demo"}') == {"company": "Demo"}
    assert parse_json('```json\n{"value": 1}\n```') == {"value": 1}
    for raw in ("", "garbage", "[1,2]", "null"):
        with pytest.raises(ValueError):
            parse_json(raw)


def test_openai_route(monkeypatch):
    from openai import OpenAI
    seen = {}
    def build_client(**kwargs):
        seen.update(kwargs)
        return SimpleNamespace(chat=SimpleNamespace(
            completions=SimpleNamespace(create=lambda **kw: fake_response(kw))
        ))
    def fake_response(kwargs):
        seen["request"] = kwargs
        return SimpleNamespace(choices=[SimpleNamespace(message=SimpleNamespace(
            content='{"source":"openai"}'
        ))])
    monkeypatch.setattr("openai.OpenAI", build_client)
    result = generate_json(provider="OpenAI", api_key="fake-oai-key",
                           system_prompt="System", user_prompt="CIM text")
    assert result == {"source":"openai"}
    assert seen["api_key"] == "fake-oai-key"
    assert seen["request"]["model"] == DEFAULT_MODELS["OpenAI"]
    assert "CIM text" in seen["request"]["messages"][-1]["content"]


def test_claude_route(monkeypatch):
    seen = {}
    def build_client(**kwargs):
        seen.update(kwargs)
        return SimpleNamespace(messages=SimpleNamespace(create=lambda **kw: fake_response(kw)))
    def fake_response(kwargs):
        seen["request"] = kwargs
        return SimpleNamespace(stop_reason="end_turn",
            content=[SimpleNamespace(type="text",text='{"source":"claude"}')])
    monkeypatch.setattr("anthropic.Anthropic", build_client)
    result = generate_json(provider="Claude",api_key="fake-anthropic",
                           system_prompt="System",user_prompt="CIM text")
    assert result == {"source":"claude"}
    assert seen["request"]["model"] == DEFAULT_MODELS["Claude"]
    assert seen["request"]["messages"][0]["content"] == "CIM text"
    assert seen["api_key"] == "fake-anthropic"


def test_gemini_route(monkeypatch):
    from google import genai
    from google.genai import types
    seen={}
    def build_client(**kwargs):
        seen.update(kwargs)
        return SimpleNamespace(models=SimpleNamespace(generate_content=lambda **kw: fake_response(kw)))
    def fake_response(kwargs):
        seen["request"]=kwargs
        return SimpleNamespace(text='{"source":"gemini"}')
    monkeypatch.setattr(genai,"Client",build_client)
    monkeypatch.setattr(types,"GenerateContentConfig",lambda **kwargs:kwargs)
    result=generate_json(provider="Gemini",api_key="fake-google",
                         system_prompt="System",user_prompt="CIM text")
    assert result=={"source":"gemini"}
    assert seen["api_key"]=="fake-google"
    assert seen["request"]["model"]==DEFAULT_MODELS["Gemini"]
    assert seen["request"]["config"]["response_mime_type"]=="application/json"


def test_unsupported_provider_or_empty_key_is_rejected_without_request():
    with pytest.raises(ValueError, match="Unsupported"):
        generate_json(provider="arbitrary-host",api_key="secret",
                      system_prompt="x",user_prompt="y")
    with pytest.raises(ValueError, match="API key"):
        generate_json(provider="Claude",api_key="",system_prompt="x",user_prompt="y")


def test_claude_truncated_output_rejected(monkeypatch):
    def fake_client(**kwargs):
        return SimpleNamespace(messages=SimpleNamespace(create=lambda **kw:
            SimpleNamespace(stop_reason="max_tokens",content=[])))
    monkeypatch.setattr("anthropic.Anthropic",fake_client)
    with pytest.raises(ValueError,match="truncated"):
        generate_json(provider="Claude",api_key="fake",system_prompt="x",user_prompt="y")
