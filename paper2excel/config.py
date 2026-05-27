"""Portable configuration helpers for the desktop app."""

from __future__ import annotations

import json
import re
import sys
from pathlib import Path
from typing import Any


DEFAULT_OUTPUT_FIELDS = [
    {
        "name": "Summary",
        "description": "Summarize the paper's relevance and key information in English.",
        "optimized_prompt": (
            "Write 2-3 concise English sentences summarizing the paper's core topic, method or evidence, "
            "and relevance to the user's task. Use only the provided row content and PDF text. "
            "If the available evidence is insufficient, write \"Not supported\"."
        ),
        "field_type": "string",
        "required": True,
    },
]


DEFAULT_RETAINED_COLUMNS = [
    "Title",
    "Publication Title",
    "Publication Year",
    "Author",
    "DOI",
    "Abstract Note",
    "File Attachments",
]


DEFAULT_CONFIG: dict[str, Any] = {
    "provider": "DeepSeek",
    "base_url": "https://api.deepseek.com",
    "model": "deepseek-v4-flash",
    "api_key": "",
    "remember_api_key": False,
    "proxy_url": "",
    "temperature": 0.2,
    "timeout_seconds": 180,
    "max_tokens": 3000,
    "input_path": "",
    "output_path": "",
    "retained_columns": DEFAULT_RETAINED_COLUMNS,
    "selected_input_columns": DEFAULT_RETAINED_COLUMNS,
    "pdf_path_column": "File Attachments",
    "task_description": "Analyze the Excel row and fill the requested output fields.",
    "row_start": "",
    "row_end": "",
    "only_empty_outputs": False,
    "autosave_every": 3,
    "request_delay_seconds": 0,
    "pdf_text_limit": 10000,
    "output_fields": DEFAULT_OUTPUT_FIELDS,
}


PROVIDER_PRESETS = {
    "OpenAI": {
        "base_url": "https://api.openai.com/v1",
        "model": "gpt-4o-mini",
        "requires_api_key": True,
    },
    "DeepSeek": {
        "base_url": "https://api.deepseek.com",
        "model": "deepseek-v4-flash",
        "requires_api_key": True,
    },
    "Kimi / Moonshot Intl": {
        "base_url": "https://api.moonshot.ai/v1",
        "model": "kimi-k2.6",
        "requires_api_key": True,
    },
    "Kimi / Moonshot CN": {
        "base_url": "https://api.moonshot.cn/v1",
        "model": "kimi-k2.6",
        "requires_api_key": True,
    },
    "Qwen / DashScope CN": {
        "base_url": "https://dashscope.aliyuncs.com/compatible-mode/v1",
        "model": "qwen3-max-2026-01-23",
        "requires_api_key": True,
    },
    "Qwen / DashScope Intl": {
        "base_url": "https://dashscope-intl.aliyuncs.com/compatible-mode/v1",
        "model": "qwen3.6-plus",
        "requires_api_key": True,
    },
    "Zhipu GLM / BigModel": {
        "base_url": "https://open.bigmodel.cn/api/paas/v4",
        "model": "glm-5.1",
        "requires_api_key": True,
    },
    "Baidu Qianfan": {
        "base_url": "https://api.baiduqianfan.ai/v1",
        "model": "deepseek-v3.1-250821",
        "requires_api_key": True,
    },
    "Tencent Hunyuan": {
        "base_url": "https://api.hunyuan.cloud.tencent.com/v1",
        "model": "hunyuan-turbos-latest",
        "requires_api_key": True,
    },
    "Volcengine Ark": {
        "base_url": "https://ark.cn-beijing.volces.com/api/v3",
        "model": "ep-xxxxxxxx",
        "requires_api_key": True,
    },
    "Gemini": {
        "base_url": "https://generativelanguage.googleapis.com/v1beta/openai",
        "model": "gemini-3.5-flash",
        "requires_api_key": True,
    },
    "OpenRouter": {
        "base_url": "https://openrouter.ai/api/v1",
        "model": "openai/gpt-4o-mini",
        "requires_api_key": True,
    },
    "SiliconFlow": {
        "base_url": "https://api.siliconflow.cn/v1",
        "model": "Qwen/Qwen2.5-7B-Instruct",
        "requires_api_key": True,
    },
    "Groq": {
        "base_url": "https://api.groq.com/openai/v1",
        "model": "llama-3.3-70b-versatile",
        "requires_api_key": True,
    },
    "Mistral": {
        "base_url": "https://api.mistral.ai/v1",
        "model": "mistral-small-latest",
        "requires_api_key": True,
    },
    "xAI / Grok": {
        "base_url": "https://api.x.ai/v1",
        "model": "grok-4.3",
        "requires_api_key": True,
    },
    "Together AI": {
        "base_url": "https://api.together.xyz/v1",
        "model": "meta-llama/Meta-Llama-3.1-8B-Instruct-Turbo",
        "requires_api_key": True,
    },
    "Ollama Local": {
        "base_url": "http://localhost:11434/v1",
        "model": "llama3.1",
        "requires_api_key": False,
    },
    "LM Studio Local": {
        "base_url": "http://localhost:1234/v1",
        "model": "local-model",
        "requires_api_key": False,
    },
    "vLLM Local": {
        "base_url": "http://localhost:8000/v1",
        "model": "NousResearch/Meta-Llama-3-8B-Instruct",
        "requires_api_key": False,
    },
    "OpenAI-compatible": {
        "base_url": "http://127.0.0.1:11434/v1",
        "model": "local-model",
        "requires_api_key": False,
    },
}


def app_dir() -> Path:
    """Return the writable directory that travels with the one-folder EXE."""

    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parents[1]


def resource_path(relative_path: str) -> Path:
    """Return a bundled read-only resource path in source or PyInstaller builds."""

    if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS"):
        return Path(getattr(sys, "_MEIPASS")) / relative_path
    return app_dir() / relative_path


def config_path() -> Path:
    return app_dir() / "user_config.json"


def field_templates_dir() -> Path:
    return app_dir() / "templates"


def safe_template_filename(name: str) -> str:
    base = re.sub(r"[^A-Za-z0-9._\-\u4e00-\u9fff]+", "_", name.strip()).strip("._")
    return f"{base or 'paper2excel_template'}.json"


def save_field_template(
    path: str | Path,
    *,
    task_description: str,
    output_fields: list[dict[str, Any]],
) -> Path:
    target = Path(path)
    target.parent.mkdir(parents=True, exist_ok=True)
    template = {
        "schema_version": 1,
        "kind": "paper2excel_field_template",
        "task_description": task_description,
        "output_fields": _normalized_output_fields(output_fields),
    }
    target.write_text(json.dumps(template, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def load_field_template(path: str | Path) -> dict[str, Any]:
    target = Path(path)
    loaded = json.loads(target.read_text(encoding="utf-8"))
    if not isinstance(loaded, dict):
        raise ValueError("模板文件必须是 JSON 对象。")
    fields = loaded.get("output_fields")
    if not isinstance(fields, list):
        raise ValueError("模板缺少 output_fields。")
    return {
        "task_description": str(loaded.get("task_description", "") or ""),
        "output_fields": _normalized_output_fields(fields),
    }


def load_config(path: str | Path | None = None) -> dict[str, Any]:
    target = Path(path) if path else config_path()
    if not target.exists():
        return json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))
    try:
        loaded = json.loads(target.read_text(encoding="utf-8"))
    except Exception:
        return json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))
    merged = json.loads(json.dumps(DEFAULT_CONFIG, ensure_ascii=False))
    merged.update(loaded)
    if "output_fields" not in loaded:
        merged["output_fields"] = json.loads(json.dumps(DEFAULT_OUTPUT_FIELDS, ensure_ascii=False))
    if not merged.get("remember_api_key", False):
        merged["api_key"] = ""
    return merged


def save_config(config: dict[str, Any], path: str | Path | None = None) -> Path:
    target = Path(path) if path else config_path()
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(config, ensure_ascii=False, indent=2), encoding="utf-8")
    return target


def _normalized_output_fields(fields: list[dict[str, Any]]) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for item in fields:
        if not isinstance(item, dict):
            raise ValueError("输出字段模板中存在无效字段。")
        name = str(item.get("name", "") or "").strip()
        if not name:
            raise ValueError("输出字段模板中存在空列名。")
        field_type = str(item.get("field_type", "string") or "string").strip()
        if field_type not in {"string", "number", "boolean"}:
            field_type = "string"
        normalized.append(
            {
                "name": name,
                "description": str(item.get("description", "") or "").strip(),
                "optimized_prompt": str(item.get("optimized_prompt", "") or "").strip(),
                "field_type": field_type,
                "required": bool(item.get("required", True)),
            }
        )
    return normalized
