"""Free AI provider adapters for factual ATS CV suggestions.

Cloud keys are supplied by the caller and are never persisted by this module.
The deterministic CV engine remains authoritative: AI may rewrite wording and
reorder already verified information, but it cannot silently add JD claims.
"""
from __future__ import annotations

import json
import re
import socket
from dataclasses import dataclass
from typing import Any
from urllib.error import HTTPError, URLError
from urllib.request import Request, urlopen


DEFAULT_SETTINGS: dict[str, Any] = {
    "selected": "ollama",
    "fallback": True,
    "ollama_url": "http://localhost:11434",
    "ollama_model": "llama3.2:3b",
    "gemini_model": "gemini-2.5-flash-lite",
    "gemini_key": "",
    "groq_model": "llama-3.1-8b-instant",
    "groq_key": "",
    "timeout": 45,
}

PROVIDER_LABELS = {
    "ollama": "Ollama (Free / Local)",
    "gemini": "Gemini (Free Tier)",
    "groq": "Groq (Free Tier)",
}


class AIProviderError(RuntimeError):
    """A provider could not return a usable response."""


@dataclass
class ProviderAttempt:
    provider: str
    status: str
    detail: str = ""

    def as_dict(self) -> dict[str, str]:
        return {
            "provider": self.provider,
            "status": self.status,
            "detail": self.detail,
        }


def provider_settings(values: dict[str, Any] | None = None) -> dict[str, Any]:
    """Return complete settings without mutating the shared defaults."""
    settings = dict(DEFAULT_SETTINGS)
    if values:
        settings.update(values)
    return settings


def _error_detail(error: Exception) -> str:
    """Short error safe to display; never include request headers or API keys."""
    if isinstance(error, HTTPError):
        return f"HTTP {error.code}: provider rejected the request"
    if isinstance(error, (TimeoutError, socket.timeout)):
        return "request timed out"
    if isinstance(error, URLError):
        reason = str(error.reason)
        return "provider is unavailable" if not reason else reason[:160]
    return str(error)[:180] or error.__class__.__name__


def _post_json(
    url: str,
    payload: dict[str, Any],
    headers: dict[str, str] | None = None,
    timeout: int = 45,
) -> dict[str, Any]:
    body = json.dumps(payload).encode("utf-8")
    request_headers = {"Content-Type": "application/json"}
    request_headers.update(headers or {})
    request = Request(url, data=body, headers=request_headers, method="POST")
    try:
        with urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except (HTTPError, URLError, TimeoutError, socket.timeout, OSError) as exc:
        raise AIProviderError(_error_detail(exc)) from exc
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AIProviderError("provider returned invalid JSON") from exc
    if not isinstance(value, dict):
        raise AIProviderError("provider returned an unexpected response")
    return value


def _get_json(
    url: str,
    headers: dict[str, str] | None = None,
    timeout: int = 45,
) -> dict[str, Any]:
    request = Request(url, headers=headers or {}, method="GET")
    try:
        with urlopen(request, timeout=timeout) as response:
            raw = response.read().decode("utf-8", errors="replace")
    except (HTTPError, URLError, TimeoutError, socket.timeout, OSError) as exc:
        raise AIProviderError(_error_detail(exc)) from exc
    try:
        value = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise AIProviderError("provider returned invalid JSON") from exc
    if not isinstance(value, dict):
        raise AIProviderError("provider returned an unexpected response")
    return value


class AIProvider:
    """Common provider interface."""

    name = ""

    def __init__(self, settings: dict[str, Any]):
        self.settings = settings

    @property
    def configured(self) -> bool:
        raise NotImplementedError

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        raise NotImplementedError


class OllamaProvider(AIProvider):
    name = "ollama"

    @property
    def configured(self) -> bool:
        return bool(
            str(self.settings.get("ollama_url", "")).strip()
            and str(self.settings.get("ollama_model", "")).strip()
        )

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        base = str(self.settings["ollama_url"]).rstrip("/")
        response = _post_json(
            base + "/api/chat",
            {
                "model": self.settings["ollama_model"],
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "stream": False,
                "format": "json",
                "options": {"temperature": 0.1},
            },
            timeout=int(self.settings.get("timeout", 45)),
        )
        try:
            return str(response["message"]["content"])
        except (KeyError, TypeError) as exc:
            raise AIProviderError("Ollama response has no message content") from exc


class GeminiProvider(AIProvider):
    name = "gemini"

    @property
    def configured(self) -> bool:
        return bool(
            str(self.settings.get("gemini_key", "")).strip()
            and str(self.settings.get("gemini_model", "")).strip()
        )

    def available_models(self) -> list[str]:
        """Models this key may call, so a renamed default does not break use."""
        data = _get_json(
            "https://generativelanguage.googleapis.com/v1beta/models",
            headers={"x-goog-api-key": str(self.settings["gemini_key"])},
            timeout=int(self.settings.get("timeout", 45)),
        )
        models: list[str] = []
        for item in data.get("models", []):
            if not isinstance(item, dict):
                continue
            name = str(item.get("name", "")).split("/")[-1]
            methods = item.get("supportedGenerationMethods") or []
            if not name or (methods and "generateContent" not in methods):
                continue
            if any(
                word in name.lower()
                for word in ("embedding", "image", "tts", "vision", "aqa")
            ):
                continue
            models.append(name)
        # Small, free-tier friendly models first.
        models.sort(key=lambda name: ("flash" not in name.lower(), name))
        return models

    def _generate_with(
        self, model: str, system_prompt: str, user_prompt: str
    ) -> str:
        response = _post_json(
            (
                "https://generativelanguage.googleapis.com/v1beta/models/"
                + model
                + ":generateContent"
            ),
            {
                "systemInstruction": {
                    "parts": [{"text": system_prompt}],
                },
                "contents": [
                    {
                        "role": "user",
                        "parts": [{"text": user_prompt}],
                    }
                ],
                "generationConfig": {
                    "temperature": 0.1,
                    "responseMimeType": "application/json",
                },
            },
            headers={"x-goog-api-key": str(self.settings["gemini_key"])},
            timeout=int(self.settings.get("timeout", 45)),
        )
        try:
            parts = response["candidates"][0]["content"]["parts"]
            return "".join(str(part.get("text", "")) for part in parts)
        except (KeyError, IndexError, TypeError) as exc:
            raise AIProviderError("Gemini response has no generated content") from exc

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        model = str(self.settings["gemini_model"]).strip()
        try:
            return self._generate_with(model, system_prompt, user_prompt)
        except AIProviderError as exc:
            if "404" not in str(exc):
                raise
            try:
                candidates = self.available_models()
            except AIProviderError:
                raise exc
            for candidate in candidates[:3]:
                if candidate == model:
                    continue
                try:
                    return self._generate_with(
                        candidate, system_prompt, user_prompt
                    )
                except AIProviderError:
                    continue
            raise AIProviderError(
                f"model '{model}' was not found for this key. Available: "
                + (", ".join(candidates[:6]) or "none")
            ) from exc


class GroqProvider(AIProvider):
    name = "groq"

    @property
    def configured(self) -> bool:
        return bool(
            str(self.settings.get("groq_key", "")).strip()
            and str(self.settings.get("groq_model", "")).strip()
        )

    def generate(self, system_prompt: str, user_prompt: str) -> str:
        response = _post_json(
            "https://api.groq.com/openai/v1/chat/completions",
            {
                "model": self.settings["groq_model"],
                "messages": [
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": user_prompt},
                ],
                "temperature": 0.1,
                "response_format": {"type": "json_object"},
            },
            headers={
                "Authorization": "Bearer " + str(self.settings["groq_key"]),
            },
            timeout=int(self.settings.get("timeout", 45)),
        )
        try:
            return str(response["choices"][0]["message"]["content"])
        except (KeyError, IndexError, TypeError) as exc:
            raise AIProviderError("Groq response has no generated content") from exc


PROVIDER_CLASSES = {
    "ollama": OllamaProvider,
    "gemini": GeminiProvider,
    "groq": GroqProvider,
}


def _provider_order(settings: dict[str, Any]) -> list[str]:
    selected = str(settings.get("selected", "ollama")).lower()
    if selected not in PROVIDER_CLASSES:
        selected = "ollama"
    if not settings.get("fallback", True):
        return [selected]
    return [selected] + [
        name for name in ("ollama", "gemini", "groq") if name != selected
    ]


SYSTEM_PROMPT = """You are an ATS resume editor.
Use only facts already present in the supplied CV and skills explicitly
confirmed by the user. Never invent a skill, employer, role, date, degree,
certification, metric, responsibility, or achievement.

Return only one JSON object with this exact shape:
{
  "rewritten_summary": "concise factual summary or empty string",
  "prioritized_skills": ["verified skill"],
  "bullet_rewrites": [
    {"original": "exact existing bullet text", "rewritten": "truthful stronger wording"}
  ],
  "suggestions": ["short actionable suggestion"]
}

For bullet_rewrites, copy original text exactly without its bullet symbol.
Rewrite wording only; preserve every number and named technology. Do not add
missing JD skills. Keep all text ATS-readable and avoid keyword stuffing."""


def _analysis_prompt(
    cv_text: str,
    jd_text: str,
    ats: dict[str, Any],
    allowed_skills: list[str],
    missing_skills: list[str],
) -> str:
    report = {
        key: value
        for key, value in ats.items()
        if key
        in {
            "overall",
            "skills_score",
            "title",
            "required_years",
            "cv_years",
            "format_score",
            "suggestions",
        }
    }
    return (
        "ALLOWED SKILLS (only these may appear as skill claims):\n"
        + json.dumps(allowed_skills, ensure_ascii=False)
        + "\n\nFORBIDDEN / UNCONFIRMED JD SKILLS:\n"
        + json.dumps(missing_skills, ensure_ascii=False)
        + "\n\nDETERMINISTIC ATS REPORT:\n"
        + json.dumps(report, ensure_ascii=False)
        + "\n\nCURRENT ATS CV:\n"
        + cv_text[:16000]
        + "\n\nJOB DESCRIPTION:\n"
        + jd_text[:12000]
    )


def _parse_model_json(text: str) -> dict[str, Any]:
    cleaned = text.strip()
    cleaned = re.sub(r"^```(?:json)?\s*", "", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"\s*```$", "", cleaned)
    if not cleaned.startswith("{"):
        start, end = cleaned.find("{"), cleaned.rfind("}")
        if start >= 0 and end > start:
            cleaned = cleaned[start : end + 1]
    try:
        value = json.loads(cleaned)
    except json.JSONDecodeError as exc:
        raise AIProviderError("AI returned malformed structured output") from exc
    if not isinstance(value, dict):
        raise AIProviderError("AI output must be a JSON object")
    return value


def _normal(text: str) -> str:
    return re.sub(r"\s+", " ", text.lstrip("-*\u2022 ").strip()).casefold()


def _contains_term(text: str, term: str) -> bool:
    return bool(
        re.search(r"(?<!\w)" + re.escape(term) + r"(?!\w)", text, re.IGNORECASE)
    )


def _safe_generated_text(
    text: str,
    source_text: str,
    forbidden_terms: list[str],
) -> tuple[bool, str]:
    for term in forbidden_terms:
        if term and _contains_term(text, term):
            return False, f"contains unconfirmed skill: {term}"
    source_numbers = set(re.findall(r"\b\d+(?:\.\d+)?%?\b", source_text))
    new_numbers = set(re.findall(r"\b\d+(?:\.\d+)?%?\b", text)) - source_numbers
    if new_numbers:
        return False, "introduces an unsupported number"
    allowed_new_words = {
        "built", "developed", "designed", "implemented", "automated",
        "managed", "led", "created", "improved", "optimized", "deployed",
        "maintained", "configured", "collaborated", "delivered", "supported",
        "monitored", "streamlined", "integrated", "migrated", "administered",
        "worked", "conducted", "performed", "ensured", "utilized", "leveraged",
        "experienced", "skilled", "professional", "results", "proven",
    }
    source_lower = source_text.casefold()
    for token in re.findall(r"\b[A-Z][A-Za-z0-9+.#/-]{2,}\b", text):
        if token.casefold() not in source_lower and token.casefold() not in allowed_new_words:
            return False, f"introduces an unsupported named term: {token}"
    return True, ""


def validate_ai_response(
    payload: dict[str, Any],
    cv_text: str,
    allowed_skills: list[str],
    missing_skills: list[str],
) -> dict[str, Any]:
    """Drop unsafe fields while retaining safe provider suggestions."""
    warnings: list[str] = []
    source_bullets = {
        _normal(line): line.lstrip("-*\u2022 ").strip()
        for line in cv_text.splitlines()
        if line.strip().startswith(("-", "*", "\u2022"))
    }
    allowed = {skill.casefold(): skill for skill in allowed_skills}

    summary = payload.get("rewritten_summary", "")
    if not isinstance(summary, str):
        summary = ""
    summary = re.sub(r"\s+", " ", summary).strip()[:1200]
    summary_safe, reason = _safe_generated_text(
        summary, cv_text, missing_skills
    )
    if summary and not summary_safe:
        warnings.append("AI summary ignored because it " + reason + ".")
        summary = ""

    prioritized: list[str] = []
    raw_skills = payload.get("prioritized_skills", [])
    if isinstance(raw_skills, list):
        for raw in raw_skills:
            key = str(raw).strip().casefold()
            if key in allowed and allowed[key] not in prioritized:
                prioritized.append(allowed[key])

    rewrites: list[dict[str, str]] = []
    raw_rewrites = payload.get("bullet_rewrites", [])
    if isinstance(raw_rewrites, list):
        for item in raw_rewrites[:12]:
            if not isinstance(item, dict):
                continue
            original = re.sub(
                r"\s+", " ", str(item.get("original", "")).strip()
            )
            rewritten = re.sub(
                r"\s+", " ", str(item.get("rewritten", "")).strip()
            )[:500]
            source_original = source_bullets.get(_normal(original))
            if not source_original or not rewritten:
                continue
            safe, rewrite_reason = _safe_generated_text(
                rewritten, source_original, missing_skills
            )
            if safe:
                rewrites.append(
                    {"original": source_original, "rewritten": rewritten}
                )
            else:
                warnings.append(
                    "One AI bullet rewrite was ignored because it "
                    + rewrite_reason
                    + "."
                )

    suggestions: list[str] = []
    raw_suggestions = payload.get("suggestions", [])
    if isinstance(raw_suggestions, list):
        for suggestion in raw_suggestions[:8]:
            clean = re.sub(r"\s+", " ", str(suggestion)).strip()[:300]
            if clean and clean not in suggestions:
                suggestions.append(clean)

    return {
        "rewritten_summary": summary,
        "prioritized_skills": prioritized,
        "bullet_rewrites": rewrites,
        "suggestions": suggestions,
        "validation_warnings": warnings,
    }


def apply_ai_response(cv_text: str, response: dict[str, Any]) -> str:
    """Apply only fields that passed validation to the deterministic CV text."""
    updated = cv_text
    summary = str(response.get("rewritten_summary", "")).strip()
    if summary:
        summary_pattern = re.compile(
            r"(^SUMMARY\s*$)(.*?)(?=^[A-Z][A-Z &/-]{2,}\s*$|\Z)",
            re.MULTILINE | re.DOTALL,
        )
        updated = summary_pattern.sub(
            lambda match: match.group(1) + "\n" + summary + "\n\n",
            updated,
            count=1,
        )

    for item in response.get("bullet_rewrites", []):
        original = str(item["original"])
        rewritten = str(item["rewritten"])
        pattern = re.compile(
            r"^(\s*[-*\u2022]\s*)" + re.escape(original) + r"\s*$",
            re.MULTILINE,
        )
        updated = pattern.sub(
            lambda match: match.group(1) + rewritten, updated, count=1
        )
    return updated


def run_ai_analysis(
    cv_text: str,
    jd_text: str,
    ats: dict[str, Any],
    allowed_skills: list[str],
    missing_skills: list[str],
    settings: dict[str, Any] | None,
) -> dict[str, Any]:
    """Try the selected provider and fall back without blocking CV generation."""
    config = provider_settings(settings)
    attempts: list[ProviderAttempt] = []
    prompt = _analysis_prompt(
        cv_text, jd_text, ats, allowed_skills, missing_skills
    )

    for name in _provider_order(config):
        provider = PROVIDER_CLASSES[name](config)
        if not provider.configured:
            attempts.append(
                ProviderAttempt(name, "skipped", "API key is not configured")
            )
            continue
        try:
            raw = provider.generate(SYSTEM_PROMPT, prompt)
            parsed = _parse_model_json(raw)
            validated = validate_ai_response(
                parsed, cv_text, allowed_skills, missing_skills
            )
            attempts.append(ProviderAttempt(name, "success"))
            return {
                "ok": True,
                "provider": name,
                "response": validated,
                "attempts": [attempt.as_dict() for attempt in attempts],
            }
        except Exception as exc:
            attempts.append(
                ProviderAttempt(name, "failed", _error_detail(exc))
            )

    return {
        "ok": False,
        "provider": "",
        "response": None,
        "attempts": [attempt.as_dict() for attempt in attempts],
    }


def test_provider(
    name: str,
    settings: dict[str, Any] | None,
) -> tuple[bool, str]:
    """Small live connectivity check used by the protected settings dialog."""
    config = provider_settings(settings)
    provider_class = PROVIDER_CLASSES.get(name)
    if provider_class is None:
        return False, "Unknown provider"
    provider = provider_class(config)
    if not provider.configured:
        return False, "Provider is not configured"
    try:
        raw = provider.generate(
            "Return only valid JSON.",
            '{"rewritten_summary":"","prioritized_skills":[],"bullet_rewrites":[],"suggestions":[]}',
        )
        _parse_model_json(raw)
        return True, "Connection successful"
    except Exception as exc:
        message = _error_detail(exc)
        if isinstance(provider, GeminiProvider) and "404" in message:
            try:
                models = provider.available_models()
            except AIProviderError:
                models = []
            if models:
                message += ". Models available for this key: " + ", ".join(
                    models[:6]
                )
        return False, message
