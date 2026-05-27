"""Core package for Paper2Excel."""

from .core import (
    ChatCompletionError,
    ChatCompletionSettings,
    ColumnMapping,
    JsonParseError,
    OutputField,
    Paper2ExcelError,
    build_chat_completions_url,
    build_json_prompt,
    call_chat_completion,
    normalize_proxy_url,
    optimize_field_prompt,
    parse_json_object,
    process_excel,
    read_excel_rows,
    write_excel_rows,
)
from .config import load_field_template, save_field_template
from .runner import BatchOptions, BatchResult, default_output_path, extract_pdf_text, read_excel_columns, run_batch

__all__ = [
    "ChatCompletionError",
    "ChatCompletionSettings",
    "ColumnMapping",
    "JsonParseError",
    "OutputField",
    "Paper2ExcelError",
    "build_chat_completions_url",
    "build_json_prompt",
    "call_chat_completion",
    "normalize_proxy_url",
    "optimize_field_prompt",
    "parse_json_object",
    "process_excel",
    "read_excel_rows",
    "load_field_template",
    "save_field_template",
    "write_excel_rows",
    "BatchOptions",
    "BatchResult",
    "default_output_path",
    "extract_pdf_text",
    "read_excel_columns",
    "run_batch",
]
