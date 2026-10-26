"""Provider-neutral JSON extraction for DealLens.

Supported providers: OpenAI, Anthropic Claude, Google Gemini.
Keys remain runtime-only and are never written to the model, GitHub or output.
This module sends the CIM text only to the provider explicitly selected by
the user. It does not confer sandboxing or zero-data-retention guarantees.
"""
import json
import re

PROVIDERS = ("OpenAI", "Claude", "Gemini")
DEFAULT_MODELS = {
    "OpenAI": "gpt-4.1-mini",
    "Claude": "claude-sonnet-5-5",
    "Gemini": "gemini-3.8-flash",
}
ENVIRONMENT_KEYS = {
    "OpenAI": "OPENAI_API_KEY",
    "Claude": "ANTHROPIC_API_KEY",
    "Gemini": "GEMINI_API_KEY",
}


def parse_json(text):
    """Parse a JSON object with a narrow fence-only fallback, never execute data."""
    if not isinstance(text, str) or not text.strip():
        raise ValueError("AI provider returned no structured response.")
    content = text.strip()
    if content.startswith("```"):
        content = re.sub(r"^\x60{3}(?:json)?\s*", "", content, flags=re.I)
        content = re.sub(r"\s*\x60{3}$", "", content)
    try:
        result = json.loads(content)
    except json.JSONDecodeError as exc:
        raise ValueError("AI provider returned invalid JSON. Retry or use another model.") from exc
    if not isinstance(result, dict):
        raise ValueError("AI provider returned a JSON value instead of an object.")
    return result


def generate_json(*, provider, api_key, system_prompt, user_prompt,
                  model=None, max_output_tokens=12000):
    """Make one provider-specific request and normalize its JSON response.

    Only a predeclared provider may be selected, so an injected API URL cannot
    redirect sensitive CIM text or user keys to arbitrary third-party hosts.
    """
    if provider not in PROVIDERS:
        raise ValueError("Unsupported AI provider. Choose OpenAI, Claude or Gemini.")
    if not isinstance(api_key, str) or not api_key.strip():
        raise ValueError(f"A {provider} API key is required.")
    chosen_model = (model or DEFAULT_MODELS[provider]).strip()
    if not chosen_model or len(chosen_model) > 128:
        raise ValueError("Model ID must be a nonempty provider-supported identifier.")
    if not isinstance(system_prompt, str) or not isinstance(user_prompt, str):
        raise ValueError("AI prompts must be text.")
    if provider == "OpenAI":
        from openai import OpenAI
        client = OpenAI(api_key=api_key.strip(), timeout=180.0, max_retries=1)
        response = client.chat.completions.create(
            model=chosen_model,
            temperature=0,
            response_format={"type": "json_object"},
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt},
            ],
        )
        content = response.choices[0].message.content
    elif provider == "Claude":
        from anthropic import Anthropic
        client = Anthropic(api_key=api_key.strip(), timeout=180.0, max_retries=1)
        response = client.messages.create(
            model=chosen_model,
            max_tokens=max_output_tokens,
            system=system_prompt + "\nRespond ONLY with one valid JSON object.",
            messages=[{"role": "user", "content": user_prompt}],
        )
        if getattr(response, "stop_reason", None) == "max_tokens":
            raise ValueError("Claude response was truncated. Try smaller CIM input or a higher output token allowance.")
        content = "".join(
            block.text for block in response.content
            if getattr(block, "type", None) == "text" and isinstance(getattr(block, "text", None), str)
        )
    else:
        from google import genai
        from google.genai import types
        client = genai.Client(api_key=api_key.strip(), http_options={"timeout": 180000})
        response = client.models.generate_content(
            model=chosen_model,
            contents=user_prompt,
            config=types.GenerateContentConfig(
                system_instruction=system_prompt,
                response_mime_type="application/json",
                max_output_tokens=max_output_tokens,
            ),
        )
        content = response.text
    return parse_json(content)
