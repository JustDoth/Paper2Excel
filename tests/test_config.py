import json
from pathlib import Path

from paper2excel.config import (
    DEFAULT_OUTPUT_FIELDS,
    DEFAULT_RETAINED_COLUMNS,
    load_config,
    load_field_template,
    safe_template_filename,
    save_config,
    save_field_template,
)


def test_load_config_does_not_restore_api_key_unless_remembered(tmp_path: Path) -> None:
    path = tmp_path / "config.json"
    key_name = "api_" + "key"
    path.write_text(json.dumps({key_name: "unit-test-value", "remember_api_key": False}), encoding="utf-8")

    config = load_config(path)

    assert config["api_key"] == ""


def test_save_config_writes_remembered_api_key(tmp_path: Path) -> None:
    path = tmp_path / "config.json"
    key_name = "api_" + "key"

    save_config({key_name: "unit-test-value", "remember_api_key": True}, path)

    saved = json.loads(path.read_text(encoding="utf-8"))
    assert saved[key_name] == "unit-test-value"


def test_default_retained_columns_include_zotero_fields() -> None:
    assert DEFAULT_RETAINED_COLUMNS == [
        "Title",
        "Publication Title",
        "Publication Year",
        "Author",
        "DOI",
        "Abstract Note",
        "File Attachments",
    ]


def test_default_output_field_is_single_english_summary() -> None:
    assert len(DEFAULT_OUTPUT_FIELDS) == 1
    assert DEFAULT_OUTPUT_FIELDS[0]["name"] == "Summary"
    assert "English" in DEFAULT_OUTPUT_FIELDS[0]["optimized_prompt"]


def test_field_template_roundtrip_excludes_personal_settings(tmp_path: Path) -> None:
    path = tmp_path / "template.json"

    save_field_template(
        path,
        task_description="Analyze papers.",
        output_fields=[
            {
                "name": "Summary",
                "description": "plain",
                "optimized_prompt": "prompt",
                "field_type": "string",
                "required": True,
            }
        ],
    )

    raw = json.loads(path.read_text(encoding="utf-8"))
    assert "api_key" not in raw
    assert "input_path" not in raw
    assert "output_path" not in raw
    loaded = load_field_template(path)
    assert loaded["task_description"] == "Analyze papers."
    assert loaded["output_fields"][0]["name"] == "Summary"
    assert loaded["output_fields"][0]["optimized_prompt"] == "prompt"


def test_safe_template_filename() -> None:
    assert safe_template_filename("Summary / PDF: test") == "Summary_PDF_test.json"
