"""Tkinter desktop interface for Paper2Excel."""

from __future__ import annotations

import os
import threading
from pathlib import Path
from typing import Any

import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

from .config import (
    DEFAULT_OUTPUT_FIELDS,
    DEFAULT_RETAINED_COLUMNS,
    PROVIDER_PRESETS,
    config_path,
    field_templates_dir,
    load_config,
    load_field_template,
    resource_path,
    safe_template_filename,
    save_config,
    save_field_template,
)
from .core import ChatCompletionSettings, ColumnMapping, OutputField, optimize_field_prompt, parse_json_object
from .runner import BatchOptions, default_output_path, read_excel_columns, run_batch


class Paper2ExcelApp(tk.Tk):
    """Main desktop window."""

    def __init__(self) -> None:
        super().__init__()
        self.title("Paper2Excel")
        self._set_window_icon()
        self.geometry("1180x760")
        self.minsize(1040, 680)

        self.config_data = load_config()
        self.headers: list[str] = []
        self.output_fields: list[dict[str, Any]] = []
        self.cancel_event = threading.Event()
        self.worker_thread: threading.Thread | None = None
        self.close_after_run = False

        self._init_vars()
        self._build_ui()
        self._apply_config(self.config_data)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    def _init_vars(self) -> None:
        self.provider_var = tk.StringVar()
        self.base_url_var = tk.StringVar()
        self.model_var = tk.StringVar()
        self.api_key_var = tk.StringVar()
        self.remember_api_key_var = tk.BooleanVar(value=False)
        self.proxy_url_var = tk.StringVar()
        self.temperature_var = tk.StringVar()
        self.timeout_var = tk.StringVar()
        self.max_tokens_var = tk.StringVar()

        self.input_path_var = tk.StringVar()
        self.output_path_var = tk.StringVar()
        self.pdf_column_var = tk.StringVar()

        self.field_name_var = tk.StringVar()
        self.field_type_var = tk.StringVar(value="string")
        self.field_required_var = tk.BooleanVar(value=True)

        self.row_start_var = tk.StringVar()
        self.row_end_var = tk.StringVar()
        self.only_empty_var = tk.BooleanVar(value=False)
        self.autosave_var = tk.StringVar(value="3")
        self.delay_var = tk.StringVar(value="0")
        self.pdf_limit_var = tk.StringVar(value="10000")

    def _build_ui(self) -> None:
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.grid(row=0, column=0, sticky="nsew")

        self.model_tab = ttk.Frame(self.notebook, padding=12)
        self.excel_tab = ttk.Frame(self.notebook, padding=12)
        self.fields_tab = ttk.Frame(self.notebook, padding=12)
        self.run_tab = ttk.Frame(self.notebook, padding=12)

        self.notebook.add(self.model_tab, text="模型设置")
        self.notebook.add(self.excel_tab, text="Excel 输入")
        self.notebook.add(self.fields_tab, text="输出字段")
        self.notebook.add(self.run_tab, text="运行")

        self._build_model_tab()
        self._build_excel_tab()
        self._build_fields_tab()
        self._build_run_tab()

        footer = ttk.Frame(self, padding=(12, 6))
        footer.grid(row=1, column=0, sticky="ew")
        footer.columnconfigure(0, weight=1)
        self.config_label = ttk.Label(footer, text=f"配置文件: {config_path()}")
        self.config_label.grid(row=0, column=0, sticky="w")
        ttk.Button(footer, text="保存配置", command=self._save_config_clicked).grid(row=0, column=1, padx=4)
        ttk.Button(footer, text="退出", command=self._on_close).grid(row=0, column=2, padx=4)

    def _build_model_tab(self) -> None:
        self.model_tab.columnconfigure(1, weight=1)

        ttk.Label(self.model_tab, text="服务商").grid(row=0, column=0, sticky="w", pady=6)
        provider = ttk.Combobox(
            self.model_tab,
            textvariable=self.provider_var,
            values=list(PROVIDER_PRESETS.keys()),
            state="readonly",
        )
        provider.grid(row=0, column=1, sticky="ew", pady=6)
        provider.bind("<<ComboboxSelected>>", lambda _event: self._provider_changed())

        ttk.Label(self.model_tab, text="Base URL").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(self.model_tab, textvariable=self.base_url_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Label(
            self.model_tab,
            text="可填基础地址或完整 /chat/completions 地址；程序会自动拼接/识别。",
            foreground="#555555",
        ).grid(row=2, column=1, sticky="w")

        ttk.Label(self.model_tab, text="模型名").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(self.model_tab, textvariable=self.model_var).grid(row=3, column=1, sticky="ew", pady=6)

        ttk.Label(self.model_tab, text="API Key").grid(row=4, column=0, sticky="w", pady=6)
        ttk.Entry(self.model_tab, textvariable=self.api_key_var, show="*").grid(row=4, column=1, sticky="ew", pady=6)

        ttk.Checkbutton(
            self.model_tab,
            text="保存 API Key 到本机配置文件",
            variable=self.remember_api_key_var,
        ).grid(row=5, column=1, sticky="w", pady=2)

        ttk.Label(self.model_tab, text="代理地址").grid(row=6, column=0, sticky="w", pady=6)
        ttk.Entry(self.model_tab, textvariable=self.proxy_url_var).grid(row=6, column=1, sticky="ew", pady=6)
        ttk.Label(
            self.model_tab,
            text="示例：http://127.0.0.1:7897；也可直接填 127.0.0.1:7897。",
            foreground="#555555",
        ).grid(row=7, column=1, sticky="w")

        number_frame = ttk.Frame(self.model_tab)
        number_frame.grid(row=8, column=1, sticky="w", pady=6)
        ttk.Label(self.model_tab, text="参数").grid(row=8, column=0, sticky="w", pady=6)
        ttk.Label(number_frame, text="Temperature").grid(row=0, column=0, padx=(0, 4))
        ttk.Entry(number_frame, textvariable=self.temperature_var, width=8).grid(row=0, column=1, padx=(0, 12))
        ttk.Label(number_frame, text="Timeout").grid(row=0, column=2, padx=(0, 4))
        ttk.Entry(number_frame, textvariable=self.timeout_var, width=8).grid(row=0, column=3, padx=(0, 12))
        ttk.Label(number_frame, text="Max tokens").grid(row=0, column=4, padx=(0, 4))
        ttk.Entry(number_frame, textvariable=self.max_tokens_var, width=10).grid(row=0, column=5)

        ttk.Button(self.model_tab, text="测试连接", command=self._test_connection_clicked).grid(
            row=9, column=1, sticky="w", pady=12
        )

    def _build_excel_tab(self) -> None:
        self.excel_tab.columnconfigure(1, weight=1)
        self.excel_tab.rowconfigure(3, weight=1)

        ttk.Label(self.excel_tab, text="输入 Excel/CSV").grid(row=0, column=0, sticky="w", pady=6)
        ttk.Entry(self.excel_tab, textvariable=self.input_path_var).grid(row=0, column=1, sticky="ew", pady=6)
        ttk.Button(self.excel_tab, text="选择", command=self._browse_input).grid(row=0, column=2, padx=6)

        ttk.Label(self.excel_tab, text="输出 Excel").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(self.excel_tab, textvariable=self.output_path_var).grid(row=1, column=1, sticky="ew", pady=6)
        ttk.Button(self.excel_tab, text="选择", command=self._browse_output).grid(row=1, column=2, padx=6)

        ttk.Label(self.excel_tab, text="PDF 路径列").grid(row=2, column=0, sticky="w", pady=6)
        self.pdf_combo = ttk.Combobox(self.excel_tab, textvariable=self.pdf_column_var, values=[], state="readonly")
        self.pdf_combo.grid(row=2, column=1, sticky="ew", pady=6)
        ttk.Button(self.excel_tab, text="清空", command=lambda: self.pdf_column_var.set("")).grid(
            row=2, column=2, padx=6
        )

        field_frame = ttk.LabelFrame(self.excel_tab, text="选择要保留到输出结果的原始列", padding=8)
        field_frame.grid(row=3, column=0, columnspan=3, sticky="nsew", pady=(10, 0))
        field_frame.columnconfigure(0, weight=1)
        field_frame.rowconfigure(0, weight=1)
        self.input_columns_list = tk.Listbox(field_frame, selectmode=tk.EXTENDED, exportselection=False)
        self.input_columns_list.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(field_frame, orient="vertical", command=self.input_columns_list.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.input_columns_list.configure(yscrollcommand=scrollbar.set)

        buttons = ttk.Frame(field_frame)
        buttons.grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Button(buttons, text="保留全部", command=self._select_all_input_columns).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(buttons, text="全部不保留", command=lambda: self.input_columns_list.selection_clear(0, tk.END)).grid(
            row=0, column=1
        )

    def _build_fields_tab(self) -> None:
        self.fields_tab.columnconfigure(0, weight=1)
        self.fields_tab.columnconfigure(1, weight=1)
        self.fields_tab.rowconfigure(0, weight=1)

        left = ttk.Frame(self.fields_tab)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)

        columns = ("name", "type", "required", "description")
        self.fields_tree = ttk.Treeview(left, columns=columns, show="headings", height=16)
        self.fields_tree.heading("name", text="输出列名")
        self.fields_tree.heading("type", text="类型")
        self.fields_tree.heading("required", text="必填")
        self.fields_tree.heading("description", text="说明")
        self.fields_tree.column("name", width=150)
        self.fields_tree.column("type", width=80)
        self.fields_tree.column("required", width=60)
        self.fields_tree.column("description", width=360)
        self.fields_tree.grid(row=0, column=0, sticky="nsew")
        self.fields_tree.bind("<<TreeviewSelect>>", lambda _event: self._load_selected_field_into_form())

        tree_scroll = ttk.Scrollbar(left, orient="vertical", command=self.fields_tree.yview)
        tree_scroll.grid(row=0, column=1, sticky="ns")
        self.fields_tree.configure(yscrollcommand=tree_scroll.set)

        field_buttons = ttk.Frame(left)
        field_buttons.grid(row=1, column=0, sticky="w", pady=8)
        ttk.Button(field_buttons, text="新增", command=self._new_field).grid(row=0, column=0, padx=(0, 6))
        ttk.Button(field_buttons, text="删除", command=self._delete_field).grid(row=0, column=1, padx=(0, 6))

        right = ttk.LabelFrame(self.fields_tab, text="字段编辑", padding=8)
        right.grid(row=0, column=1, sticky="nsew")
        right.columnconfigure(1, weight=1)
        right.rowconfigure(3, weight=1)
        right.rowconfigure(5, weight=1)

        ttk.Label(right, text="输出列名").grid(row=0, column=0, sticky="w", pady=4)
        ttk.Entry(right, textvariable=self.field_name_var).grid(row=0, column=1, sticky="ew", pady=4)

        ttk.Label(right, text="类型").grid(row=1, column=0, sticky="w", pady=4)
        ttk.Combobox(
            right,
            textvariable=self.field_type_var,
            values=["string", "number", "boolean"],
            state="readonly",
            width=12,
        ).grid(row=1, column=1, sticky="w", pady=4)

        ttk.Checkbutton(right, text="必填字段", variable=self.field_required_var).grid(
            row=2, column=1, sticky="w", pady=4
        )

        ttk.Label(right, text="大白话说明").grid(row=3, column=0, sticky="nw", pady=4)
        self.field_description_text = ScrolledText(right, height=7, wrap="word")
        self.field_description_text.grid(row=3, column=1, sticky="nsew", pady=4)

        optimize_row = ttk.Frame(right)
        optimize_row.grid(row=4, column=1, sticky="w", pady=4)
        ttk.Button(optimize_row, text="AI 优化为稳定 Prompt", command=self._optimize_field_clicked).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(optimize_row, text="用大白话填入 Prompt", command=self._copy_plain_to_prompt).grid(row=0, column=1)

        ttk.Label(right, text="字段 Prompt").grid(row=5, column=0, sticky="nw", pady=4)
        self.field_prompt_text = ScrolledText(right, height=9, wrap="word")
        self.field_prompt_text.grid(row=5, column=1, sticky="nsew", pady=4)

        edit_actions = ttk.Frame(right)
        edit_actions.grid(row=6, column=1, sticky="e", pady=(8, 0))
        ttk.Button(edit_actions, text="上移", command=lambda: self._move_selected_field(-1)).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(edit_actions, text="下移", command=lambda: self._move_selected_field(1)).grid(
            row=0, column=1, padx=(0, 6)
        )
        ttk.Button(edit_actions, text="更新字段", command=self._update_field).grid(row=0, column=2)

        template_actions = ttk.Frame(right)
        template_actions.grid(row=7, column=1, sticky="e", pady=(8, 0))
        ttk.Button(template_actions, text="保存字段模板", command=self._save_field_template_clicked).grid(
            row=0, column=0, padx=(0, 6)
        )
        ttk.Button(template_actions, text="载入字段模板", command=self._load_field_template_clicked).grid(
            row=0, column=1
        )

    def _build_run_tab(self) -> None:
        self.run_tab.columnconfigure(0, weight=1)
        self.run_tab.rowconfigure(3, weight=1)

        task_frame = ttk.LabelFrame(self.run_tab, text="任务描述", padding=8)
        task_frame.grid(row=0, column=0, sticky="ew")
        task_frame.columnconfigure(0, weight=1)
        self.task_description_text = ScrolledText(task_frame, height=4, wrap="word")
        self.task_description_text.grid(row=0, column=0, sticky="ew")

        options = ttk.LabelFrame(self.run_tab, text="运行选项", padding=8)
        options.grid(row=1, column=0, sticky="ew", pady=10)
        for i in range(12):
            options.columnconfigure(i, weight=0)
        ttk.Label(options, text="起始行").grid(row=0, column=0, padx=(0, 4))
        ttk.Entry(options, textvariable=self.row_start_var, width=8).grid(row=0, column=1, padx=(0, 12))
        ttk.Label(options, text="结束行").grid(row=0, column=2, padx=(0, 4))
        ttk.Entry(options, textvariable=self.row_end_var, width=8).grid(row=0, column=3, padx=(0, 12))
        ttk.Label(options, text="自动保存间隔").grid(row=0, column=4, padx=(0, 4))
        ttk.Entry(options, textvariable=self.autosave_var, width=8).grid(row=0, column=5, padx=(0, 12))
        ttk.Label(options, text="请求间隔秒").grid(row=0, column=6, padx=(0, 4))
        ttk.Entry(options, textvariable=self.delay_var, width=8).grid(row=0, column=7, padx=(0, 12))
        ttk.Label(options, text="PDF 字符数").grid(row=0, column=8, padx=(0, 4))
        ttk.Entry(options, textvariable=self.pdf_limit_var, width=10).grid(row=0, column=9, padx=(0, 12))
        ttk.Checkbutton(options, text="只处理输出为空的行", variable=self.only_empty_var).grid(
            row=0, column=10, sticky="w"
        )

        actions = ttk.Frame(self.run_tab)
        actions.grid(row=2, column=0, sticky="ew", pady=(0, 8))
        self.preview_button = ttk.Button(actions, text="预览前 3 行", command=self._preview_clicked)
        self.preview_button.grid(row=0, column=0, padx=(0, 6))
        self.run_button = ttk.Button(actions, text="开始批处理", command=self._run_clicked)
        self.run_button.grid(row=0, column=1, padx=(0, 6))
        self.cancel_button = ttk.Button(actions, text="取消", command=self._cancel_clicked, state="disabled")
        self.cancel_button.grid(row=0, column=2, padx=(0, 6))
        ttk.Button(actions, text="打开输出文件夹", command=self._open_output_folder).grid(row=0, column=3)

        self.progress = ttk.Progressbar(self.run_tab, mode="determinate")
        self.progress.grid(row=3, column=0, sticky="ew", pady=(0, 8))

        log_frame = ttk.LabelFrame(self.run_tab, text="日志", padding=8)
        log_frame.grid(row=4, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_text = ScrolledText(log_frame, height=16, wrap="word")
        self.log_text.grid(row=0, column=0, sticky="nsew")

    def _apply_config(self, config: dict[str, Any]) -> None:
        self.provider_var.set(config.get("provider", "DeepSeek"))
        self.base_url_var.set(config.get("base_url", "https://api.deepseek.com/v1"))
        self.model_var.set(config.get("model", "deepseek-chat"))
        self.api_key_var.set(config.get("api_key", ""))
        self.remember_api_key_var.set(bool(config.get("remember_api_key", False)))
        self.proxy_url_var.set(config.get("proxy_url", ""))
        self.temperature_var.set(str(config.get("temperature", 0.2)))
        self.timeout_var.set(str(config.get("timeout_seconds", 180)))
        self.max_tokens_var.set(str(config.get("max_tokens", 3000)))
        self.input_path_var.set(config.get("input_path", ""))
        self.output_path_var.set(config.get("output_path", ""))
        self.pdf_column_var.set(config.get("pdf_path_column", "File Attachments"))
        self.row_start_var.set(str(config.get("row_start", "")))
        self.row_end_var.set(str(config.get("row_end", "")))
        self.only_empty_var.set(bool(config.get("only_empty_outputs", False)))
        self.autosave_var.set(str(config.get("autosave_every", 3)))
        self.delay_var.set(str(config.get("request_delay_seconds", 0)))
        self.pdf_limit_var.set(str(config.get("pdf_text_limit", 10000)))

        self.task_description_text.delete("1.0", tk.END)
        self.task_description_text.insert("1.0", config.get("task_description", ""))

        self.output_fields = list(config.get("output_fields", DEFAULT_OUTPUT_FIELDS))
        self._refresh_fields_tree()

        input_path = self.input_path_var.get().strip()
        if input_path and Path(input_path).exists():
            selected = config.get("retained_columns") or config.get("selected_input_columns") or DEFAULT_RETAINED_COLUMNS
            self._load_headers(input_path, selected=selected)

    def _collect_config(self) -> dict[str, Any]:
        return {
            "provider": self.provider_var.get().strip(),
            "base_url": self.base_url_var.get().strip(),
            "model": self.model_var.get().strip(),
            "api_key": self.api_key_var.get() if self.remember_api_key_var.get() else "",
            "remember_api_key": self.remember_api_key_var.get(),
            "proxy_url": self.proxy_url_var.get().strip(),
            "temperature": self._float_var(self.temperature_var, 0.2),
            "timeout_seconds": self._int_var(self.timeout_var, 180),
            "max_tokens": self._int_var(self.max_tokens_var, 3000),
            "input_path": self.input_path_var.get().strip(),
            "output_path": self.output_path_var.get().strip(),
            "retained_columns": self._retained_columns(),
            "selected_input_columns": self._retained_columns(),
            "pdf_path_column": self.pdf_column_var.get().strip(),
            "task_description": self.task_description_text.get("1.0", tk.END).strip(),
            "row_start": self.row_start_var.get().strip(),
            "row_end": self.row_end_var.get().strip(),
            "only_empty_outputs": self.only_empty_var.get(),
            "autosave_every": self._int_var(self.autosave_var, 3),
            "request_delay_seconds": self._float_var(self.delay_var, 0),
            "pdf_text_limit": self._int_var(self.pdf_limit_var, 10000),
            "output_fields": self.output_fields,
        }

    def _provider_changed(self) -> None:
        preset = PROVIDER_PRESETS.get(self.provider_var.get())
        if not preset:
            return
        self.base_url_var.set(preset["base_url"])
        self.model_var.set(preset["model"])

    def _save_config_clicked(self) -> None:
        try:
            path = save_config(self._collect_config())
        except Exception as exc:
            messagebox.showerror("保存失败", f"配置文件写入失败:\n{exc}")
            return
        messagebox.showinfo("已保存", f"配置已保存到:\n{path}")

    def _settings(self) -> ChatCompletionSettings:
        base_url = self.base_url_var.get().strip()
        model = self.model_var.get().strip()
        if not base_url:
            raise ValueError("Base URL 不能为空。")
        if not model:
            raise ValueError("模型名不能为空。")
        preset = PROVIDER_PRESETS.get(self.provider_var.get(), {})
        if preset.get("requires_api_key", self.provider_var.get() != "OpenAI-compatible") and not self.api_key_var.get().strip():
            raise ValueError("当前服务商需要填写 API Key。")
        return ChatCompletionSettings(
            api_key=self.api_key_var.get(),
            model=model,
            base_url=base_url,
            temperature=self._float_var(self.temperature_var, 0.2),
            timeout_seconds=self._int_var(self.timeout_var, 180),
            max_tokens=self._int_var(self.max_tokens_var, 3000),
            proxy_url=self.proxy_url_var.get().strip(),
        )

    def _test_connection_clicked(self) -> None:
        try:
            settings = self._settings()
        except Exception as exc:
            messagebox.showerror("无法测试", str(exc))
            return

        def work() -> None:
            try:
                from .core import call_chat_completion

                response = call_chat_completion(
                    [
                        {"role": "system", "content": "Return valid JSON only."},
                        {"role": "user", "content": 'Return {"ok": true} as JSON.'},
                    ],
                    settings,
                )
                parsed = parse_json_object(response, required_keys=["ok"])
                self._thread_log(f"Connection OK: {parsed}")
                self.after(0, lambda: messagebox.showinfo("连接成功", "模型接口可用。"))
            except Exception as exc:
                self._thread_log(f"Connection failed: {exc}")
                error = str(exc)
                self.after(0, lambda error=error: messagebox.showerror("连接失败", error))

        self._start_background(work)

    def _browse_input(self) -> None:
        path = filedialog.askopenfilename(
            title="选择输入 Excel/CSV",
            filetypes=[("Excel/CSV", "*.xlsx *.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        if not self.output_path_var.get().strip():
            self.output_path_var.set(str(default_output_path(path)))
        self._load_headers(path)

    def _browse_output(self) -> None:
        path = filedialog.asksaveasfilename(
            title="选择输出 Excel/CSV",
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")],
        )
        if path:
            self.output_path_var.set(path)

    def _load_headers(self, path: str, selected: list[str] | None = None) -> None:
        try:
            self.headers = read_excel_columns(path)
        except Exception as exc:
            messagebox.showerror("读取失败", str(exc))
            return
        self.input_columns_list.delete(0, tk.END)
        for header in self.headers:
            self.input_columns_list.insert(tk.END, header)
        selected_set = set(selected if selected is not None else DEFAULT_RETAINED_COLUMNS)
        for index, header in enumerate(self.headers):
            if header in selected_set:
                self.input_columns_list.selection_set(index)
        self.pdf_combo.configure(values=["", *self.headers])
        self._set_default_pdf_column()
        self._log(f"Loaded table headers: {len(self.headers)} columns.")

    def _select_all_input_columns(self) -> None:
        self.input_columns_list.selection_set(0, tk.END)

    def _retained_columns(self) -> list[str]:
        return [self.input_columns_list.get(index) for index in self.input_columns_list.curselection()]

    def _set_default_pdf_column(self) -> None:
        current = self.pdf_column_var.get().strip()
        if current in self.headers:
            return
        for header in self.headers:
            if header.strip().lower() == "file attachments":
                self.pdf_column_var.set(header)
                return
        self.pdf_column_var.set("")

    def _refresh_fields_tree(self, select_index: int | None = None) -> None:
        for item in self.fields_tree.get_children():
            self.fields_tree.delete(item)
        for index, field in enumerate(self.output_fields):
            description = field.get("optimized_prompt") or field.get("description", "")
            self.fields_tree.insert(
                "",
                tk.END,
                iid=str(index),
                values=(
                    field.get("name", ""),
                    field.get("field_type", "string"),
                    "是" if field.get("required", True) else "否",
                    str(description)[:120],
                ),
            )
        if select_index is not None and 0 <= select_index < len(self.output_fields):
            iid = str(select_index)
            self.fields_tree.selection_set(iid)
            self.fields_tree.focus(iid)
            self.fields_tree.see(iid)

    def _new_field(self) -> None:
        self.fields_tree.selection_remove(*self.fields_tree.selection())
        self.field_name_var.set("")
        self.field_type_var.set("string")
        self.field_required_var.set(True)
        self.field_description_text.delete("1.0", tk.END)
        self.field_prompt_text.delete("1.0", tk.END)

    def _current_field_index(self) -> int | None:
        selection = self.fields_tree.selection()
        if not selection:
            return None
        return int(selection[0])

    def _load_selected_field_into_form(self) -> None:
        index = self._current_field_index()
        if index is None or index >= len(self.output_fields):
            return
        field = self.output_fields[index]
        self.field_name_var.set(field.get("name", ""))
        self.field_type_var.set(field.get("field_type", "string"))
        self.field_required_var.set(bool(field.get("required", True)))
        self.field_description_text.delete("1.0", tk.END)
        self.field_description_text.insert("1.0", field.get("description", ""))
        self.field_prompt_text.delete("1.0", tk.END)
        self.field_prompt_text.insert("1.0", field.get("optimized_prompt", ""))

    def _field_from_form(self) -> dict[str, Any]:
        name = self.field_name_var.get().strip()
        if not name:
            raise ValueError("字段列名不能为空。")
        return {
            "name": name,
            "description": self.field_description_text.get("1.0", tk.END).strip(),
            "optimized_prompt": self.field_prompt_text.get("1.0", tk.END).strip(),
            "field_type": self.field_type_var.get().strip() or "string",
            "required": self.field_required_var.get(),
        }

    def _update_field(self) -> None:
        try:
            field = self._field_from_form()
            self._validate_field_name(field["name"], current_index=self._current_field_index())
        except Exception as exc:
            messagebox.showerror("字段无效", str(exc))
            return
        index = self._current_field_index()
        if index is None:
            self.output_fields.append(field)
            index = len(self.output_fields) - 1
        else:
            self.output_fields[index] = field
        self._refresh_fields_tree(select_index=index)

    def _delete_field(self) -> None:
        index = self._current_field_index()
        if index is None:
            return
        del self.output_fields[index]
        next_index = min(index, len(self.output_fields) - 1)
        self._refresh_fields_tree(select_index=next_index if next_index >= 0 else None)
        if next_index >= 0:
            self._load_selected_field_into_form()
        else:
            self._new_field()

    def _move_selected_field(self, delta: int) -> None:
        index = self._current_field_index()
        if index is None:
            return
        new_index = index + delta
        if new_index < 0 or new_index >= len(self.output_fields):
            return
        self.output_fields[index], self.output_fields[new_index] = (
            self.output_fields[new_index],
            self.output_fields[index],
        )
        self._refresh_fields_tree(select_index=new_index)
        self._load_selected_field_into_form()

    def _save_field_template_clicked(self) -> None:
        try:
            self._output_field_objects()
        except Exception as exc:
            messagebox.showerror("无法保存模板", str(exc))
            return
        template_dir = field_templates_dir()
        template_dir.mkdir(parents=True, exist_ok=True)
        first_name = self.output_fields[0].get("name", "paper2excel_template") if self.output_fields else "paper2excel_template"
        path = filedialog.asksaveasfilename(
            title="保存字段模板",
            initialdir=str(template_dir),
            initialfile=safe_template_filename(str(first_name)),
            defaultextension=".json",
            filetypes=[("Paper2Excel 模板", "*.json"), ("JSON", "*.json")],
        )
        if not path:
            return
        try:
            saved = save_field_template(
                path,
                task_description=self.task_description_text.get("1.0", tk.END).strip(),
                output_fields=self.output_fields,
            )
        except Exception as exc:
            messagebox.showerror("保存失败", f"模板写入失败:\n{exc}")
            return
        messagebox.showinfo("已保存", f"字段模板已保存到:\n{saved}")

    def _load_field_template_clicked(self) -> None:
        template_dir = field_templates_dir()
        template_dir.mkdir(parents=True, exist_ok=True)
        path = filedialog.askopenfilename(
            title="载入字段模板",
            initialdir=str(template_dir),
            filetypes=[("Paper2Excel 模板", "*.json"), ("JSON", "*.json"), ("All files", "*.*")],
        )
        if not path:
            return
        try:
            template = load_field_template(path)
        except Exception as exc:
            messagebox.showerror("载入失败", f"模板读取失败:\n{exc}")
            return
        self.output_fields = list(template["output_fields"])
        if template.get("task_description"):
            self.task_description_text.delete("1.0", tk.END)
            self.task_description_text.insert("1.0", template["task_description"])
        self._refresh_fields_tree(select_index=0 if self.output_fields else None)
        if self.output_fields:
            self._load_selected_field_into_form()
        messagebox.showinfo("已载入", f"字段模板已载入:\n{path}")

    def _validate_field_name(self, name: str, current_index: int | None = None) -> None:
        for index, field in enumerate(self.output_fields):
            if current_index is not None and index == current_index:
                continue
            if field.get("name") == name:
                raise ValueError(f"输出列名重复: {name}")

    def _copy_plain_to_prompt(self) -> None:
        text = self.field_description_text.get("1.0", tk.END).strip()
        self.field_prompt_text.delete("1.0", tk.END)
        self.field_prompt_text.insert("1.0", text)

    def _optimize_field_clicked(self) -> None:
        name = self.field_name_var.get().strip()
        description = self.field_description_text.get("1.0", tk.END).strip()
        if not name or not description:
            messagebox.showerror("缺少内容", "请先填写列名和大白话说明。")
            return
        try:
            settings = self._settings()
        except Exception as exc:
            messagebox.showerror("无法优化", str(exc))
            return

        def work() -> None:
            try:
                prompt = optimize_field_prompt(name, description, settings)
                self.after(0, lambda: self._set_field_prompt(prompt))
                self._thread_log(f"Optimized prompt for field: {name}")
            except Exception as exc:
                self._thread_log(f"Prompt optimization failed: {exc}")
                error = str(exc)
                self.after(0, lambda error=error: messagebox.showerror("优化失败", error))

        self._start_background(work)

    def _set_field_prompt(self, prompt: str) -> None:
        self.field_prompt_text.delete("1.0", tk.END)
        self.field_prompt_text.insert("1.0", prompt)

    def _output_field_objects(self) -> list[OutputField]:
        if not self.output_fields:
            raise ValueError("请至少添加一个输出字段。")
        fields = []
        names: set[str] = set()
        for field in self.output_fields:
            name = str(field.get("name", "")).strip()
            if not name:
                raise ValueError("存在空的输出字段列名。")
            if name in names:
                raise ValueError(f"输出字段重复: {name}")
            names.add(name)
            fields.append(
                OutputField(
                    name=name,
                    description=str(field.get("optimized_prompt") or field.get("description") or "").strip(),
                    field_type=str(field.get("field_type", "string")),
                    required=bool(field.get("required", True)),
                )
            )
        return fields

    def _column_mappings(self) -> list[ColumnMapping]:
        if not self.headers:
            raise ValueError("请先读取输入 Excel/CSV 的表头。")
        mappings = [
            ColumnMapping(source=column, alias=column, description=f"Original table column: {column}")
            for column in self.headers
        ]
        if self.pdf_column_var.get().strip():
            mappings.append(ColumnMapping(source="PDF Content", alias="PDF Content", description="Text extracted from PDF."))
        return mappings

    def _batch_options(self, preview: bool = False) -> BatchOptions:
        row_start = self._optional_int(self.row_start_var)
        row_end = self._optional_int(self.row_end_var)
        if preview:
            row_start = row_start or 1
            row_end = min(row_end or row_start + 2, row_start + 2)
        return BatchOptions(
            task_description=self.task_description_text.get("1.0", tk.END).strip()
            or "Analyze the Excel row and fill the requested output fields.",
            row_start=row_start,
            row_end=row_end,
            only_empty_outputs=self.only_empty_var.get() and not preview,
            autosave_every=self._int_var(self.autosave_var, 3),
            request_delay_seconds=0 if preview else self._float_var(self.delay_var, 0),
            pdf_path_column=self.pdf_column_var.get().strip(),
            pdf_text_limit=self._int_var(self.pdf_limit_var, 10000),
            retained_columns=self._retained_columns(),
        )

    def _validate_run_inputs(self, preview: bool = False) -> tuple[str, str, list[ColumnMapping], list[OutputField], BatchOptions]:
        input_path = self.input_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        if not input_path or not Path(input_path).exists():
            raise ValueError("请选择存在的输入 Excel/CSV 文件。")
        if not output_path:
            output_path = str(default_output_path(input_path))
            self.output_path_var.set(output_path)
        if preview:
            output = Path(output_path)
            output_path = str(output.with_name(f"{output.stem}_preview{output.suffix}"))
        return input_path, output_path, self._column_mappings(), self._output_field_objects(), self._batch_options(preview)

    def _preview_clicked(self) -> None:
        self._run_job(preview=True)

    def _run_clicked(self) -> None:
        self._run_job(preview=False)

    def _run_job(self, preview: bool) -> None:
        try:
            input_path, output_path, mappings, fields, options = self._validate_run_inputs(preview=preview)
            settings = self._settings()
        except Exception as exc:
            messagebox.showerror("无法运行", str(exc))
            return
        try:
            save_config(self._collect_config())
        except Exception as exc:
            self._log(f"Config save failed; continuing without saving settings: {exc}")
        self.cancel_event.clear()
        self.close_after_run = False
        self._set_running(True)
        self.progress.configure(value=0, maximum=100)
        self._log("Preview run started." if preview else "Batch run started.")

        def work() -> None:
            try:
                result = run_batch(
                    input_path=input_path,
                    output_path=output_path,
                    column_mappings=mappings,
                    output_fields=fields,
                    settings=settings,
                    options=options,
                    on_log=self._thread_log,
                    on_progress=self._thread_progress,
                    should_cancel=self.cancel_event.is_set,
                )
                message = (
                    f"Completed. success={result.success}, failed={result.failed}, "
                    f"cancelled={result.cancelled}. Output: {result.output_path}"
                )
                self._thread_log(message)
                if not self.close_after_run:
                    self.after(0, lambda: messagebox.showinfo("运行完成", message))
            except Exception as exc:
                self._thread_log(f"Run failed: {exc}")
                error = str(exc)
                if not self.close_after_run:
                    self.after(0, lambda error=error: messagebox.showerror("运行失败", error))
            finally:
                self.after(0, self._job_finished)

        self.worker_thread = threading.Thread(target=work, daemon=False)
        self.worker_thread.start()

    def _cancel_clicked(self) -> None:
        self.cancel_event.set()
        self._log("Cancel requested. The current row will finish first.")

    def _set_running(self, running: bool) -> None:
        state = "disabled" if running else "normal"
        self.preview_button.configure(state=state)
        self.run_button.configure(state=state)
        self.cancel_button.configure(state="normal" if running else "disabled")

    def _job_finished(self) -> None:
        self._set_running(False)
        if self.close_after_run:
            self.destroy()

    def _on_close(self) -> None:
        if self.worker_thread and self.worker_thread.is_alive():
            confirmed = messagebox.askyesno(
                "任务仍在运行",
                "当前批处理仍在运行。是否取消任务，并在当前行保存完成后退出？",
            )
            if not confirmed:
                return
            self.close_after_run = True
            self.cancel_event.set()
            self._log("Close requested. Cancelling after the current row is saved.")
            self.cancel_button.configure(state="disabled")
            return
        self.destroy()

    def _thread_progress(self, current: int, total: int, message: str) -> None:
        def update() -> None:
            self.progress.configure(maximum=max(total, 1), value=current)
            if message:
                self._log(message)

        self.after(0, update)

    def _thread_log(self, message: str) -> None:
        self.after(0, lambda: self._log(message))

    def _log(self, message: str) -> None:
        self.log_text.insert(tk.END, f"{message}\n")
        self.log_text.see(tk.END)

    def _start_background(self, work: Any) -> None:
        threading.Thread(target=work, daemon=True).start()

    def _set_window_icon(self) -> None:
        icon = resource_path("assets/Paper2Excel.ico")
        if not icon.exists():
            return
        try:
            self.iconbitmap(default=str(icon))
        except tk.TclError:
            pass

    def _open_output_folder(self) -> None:
        output = self.output_path_var.get().strip()
        folder = Path(output).parent if output else Path.cwd()
        if folder.exists():
            os.startfile(folder)

    @staticmethod
    def _float_var(var: tk.StringVar, default: float) -> float:
        try:
            return float(var.get())
        except ValueError:
            return default

    @staticmethod
    def _int_var(var: tk.StringVar, default: int) -> int:
        try:
            return int(float(var.get()))
        except ValueError:
            return default

    @staticmethod
    def _optional_int(var: tk.StringVar) -> int | None:
        value = var.get().strip()
        if not value:
            return None
        return int(float(value))


def main() -> None:
    app = Paper2ExcelApp()
    app.mainloop()
