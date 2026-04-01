from __future__ import annotations

import importlib
import os
import queue
import random
import subprocess
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox
from tkinter import ttk

import yaml

from .smtp_domains import (
    CONN_LABELS,
    CONN_LABEL_TO_INT,
    _flags_to_conn_type,
    _conn_type_to_flags,
    _flags_to_label,
    load_domains,
    parse_smtp_domains_file,
    save_smtp_domains_file,
)

from .campaign_queue import (
    CampaignQueueError,
    load_campaign_queue,
    run_campaign_queue,
    save_campaign_queue,
)
from .config import ConfigError, load_config
from .presets import CampaignPreset, PresetError, load_preset, save_preset
from .renderer import TemplateRenderer
from .service import CampaignError, render_preview, run_campaign, run_preflight
from .recipients import RecipientsError, load_recipients
from .stats import (
    StatsError,
    export_history_records,
    filter_history_records,
    load_history_records,
    summarize_history_records,
)
from .tinymce_editor import RichEditorError, RichTemplateEditorServer


class ModernEmailAppGUI:
    DEFAULT_CONFIG_PATH = "config/settings.yaml"

    def __init__(self, ctk: object, base_dir: Path, preset_path: Path | None = None) -> None:
        self.ctk = ctk
        self.base_dir = base_dir
        self.current_preset_path = preset_path
        self.queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None
        self.editor_window = None
        self.editor_text = None
        self.stats_window = None
        self.stats_text = None
        self.preview_window = None
        self.preview_html_widget = None
        self.preview_source_text = None
        self.rich_editor_server = RichTemplateEditorServer(base_dir)

        self.root = ctk.CTk()
        self.root.title("Email App – Modern SMTP Campaign Manager")
        self.root.geometry("1200x920")
        self.root.minsize(1000, 750)

        # Theme preference
        self.theme_var = ctk.StringVar(value="dark")

        self.config_var = ctk.StringVar(value=self.DEFAULT_CONFIG_PATH)
        self.recipients_var = ctk.StringVar(value="recipients.csv")
        self.templates_var = ctk.StringVar(value="templates")
        self.smtp_accounts_file_var = ctk.StringVar(value="")
        self.template_var = ctk.StringVar(value="")
        self.attachments_folder_var = ctk.StringVar(value="")
        self.proxy_file_var = ctk.StringVar(value="config/proxies.txt")
        self.subjects_file_var = ctk.StringVar(value="")
        self._subjects_list: list[str] = []
        self.subjects_mode_var = ctk.StringVar(value="random")  # "random" | "fixed"
        self.external_editor_path = ctk.StringVar(value="")
        self.proxy_enabled_var = ctk.BooleanVar(value=True)
        self.live_preview_var = ctk.BooleanVar(value=True)
        self.preview_recipient_var = ctk.StringVar(value="")
        self.mobile_preview_var = ctk.BooleanVar(value=False)
        self.replyto_mode_var = ctk.StringVar(value="random")
        self.replyto_count_var = ctk.StringVar(value="0")
        self.delay_var = ctk.StringVar(value="0")
        self.rate_limit_var = ctk.StringVar(value="")
        self.retry_attempts_var = ctk.StringVar(value="3")
        self.parallel_smtp_var = ctk.StringVar(value="1")
        self.parallel_enabled_var = ctk.BooleanVar(value=False)
        self.batch_interval_var = ctk.StringVar(value="0")
        self.retry_backoff_var = ctk.StringVar(value="5")
        self.dry_run_var = ctk.BooleanVar(value=True)
        self.status_var = ctk.StringVar(value="✓ Готово к запуску")
        self.editor_live_job = None
        # Отслеживание прогресса кампании
        self._stop_event: threading.Event = threading.Event()
        self._campaign_total: int = 0
        self._campaign_sent: int = 0
        self._campaign_failed: int = 0
        self._campaign_start_time: float | None = None
        self._recipients_count: int = 0
        self._campaign_failed_recipients: list[dict] = []  # [{email, reason}]

        self._build()
        self._setup_mousewheel_scrolling()
        self._on_theme_change(self.theme_var.get())
        self._load_last_session()
        if self.current_preset_path:
            self._load_preset(self.current_preset_path, silent=True)
        self._refresh_templates()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(150, self._poll_queue)

    def _build(self) -> None:
        main_container = self.ctk.CTkScrollableFrame(self.root, corner_radius=0)
        main_container.pack(fill="both", expand=True)
        main_container.grid_columnconfigure(0, weight=1)
        self.main_container = main_container

        # === ЗАГОЛОВОК С ПЕРЕКЛЮЧАТЕЛЕМ ТЕМЫ ===
        header = self.ctk.CTkFrame(main_container, fg_color=("gray90", "gray15"), corner_radius=12)
        header.pack(fill="x", padx=16, pady=(16, 8))
        
        header_row = self.ctk.CTkFrame(header, fg_color="transparent")
        header_row.pack(fill="x", padx=16, pady=12)
        header_row.grid_columnconfigure(0, weight=1)
        
        self.ctk.CTkLabel(
            header_row,
            text="📧 SMTP Campaign Manager",
            font=("", 20, "bold"),
            text_color=("gray10", "gray90"),
        ).grid(row=0, column=0, sticky="w")
        
        # Theme switcher
        theme_frame = self.ctk.CTkFrame(header_row, fg_color="transparent")
        theme_frame.grid(row=0, column=1, sticky="e", padx=(16, 0))
        self.ctk.CTkLabel(theme_frame, text="🌙 Тема:").pack(side="left", padx=(0, 6))
        theme_combo = self.ctk.CTkComboBox(
            theme_frame,
            values=["dark", "light"],
            variable=self.theme_var,
            width=80,
            command=self._on_theme_change,
        )
        theme_combo.pack(side="left")

        # === ПУТИ И КОНФИГ (Раздел 1) ===
        config_section = self.ctk.CTkFrame(main_container, fg_color=("gray95", "gray20"), corner_radius=12)
        config_section.pack(fill="x", padx=16, pady=8)
        self.ctk.CTkLabel(config_section, text="📋 Конфигурация", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(12, 8))

        config_info = self.ctk.CTkFrame(config_section, fg_color="transparent")
        config_info.pack(fill="x", padx=12, pady=(0, 8))
        self.ctk.CTkLabel(
            config_info,
            text=f"Используется фиксированный конфиг: {self.DEFAULT_CONFIG_PATH}",
            font=("", 10),
            text_color=("gray40", "gray60"),
        ).pack(anchor="w")

        self._add_path_row(config_section, 1, "Получатели", self.recipients_var, self._select_recipients)
        _rcp_info = self.ctk.CTkFrame(config_section, fg_color="transparent")
        _rcp_info.pack(fill="x", padx=14, pady=(0, 4))
        self.recipients_count_label = self.ctk.CTkLabel(_rcp_info, text="", font=("", 10), text_color=("gray40", "gray60"))
        self.recipients_count_label.pack(side="left")
        self.eta_estimate_label = self.ctk.CTkLabel(_rcp_info, text="", font=("", 10), text_color=("gray40", "gray60"))
        self.eta_estimate_label.pack(side="left", padx=(16, 0))
        self._add_path_row(config_section, 2, "Шаблоны", self.templates_var, self._select_templates)
        self._add_path_row(config_section, 3, "SMTP аккаунты", self.smtp_accounts_file_var, self._select_smtp_accounts_file)
        self._add_path_row(config_section, 4, "Темы писем", self.subjects_file_var, self._select_subjects_file)

        subjects_mode_row = self.ctk.CTkFrame(config_section, fg_color="transparent")
        subjects_mode_row.pack(fill="x", padx=14, pady=(0, 6))
        self.ctk.CTkLabel(subjects_mode_row, text="Режим темы:").grid(row=0, column=0, sticky="w")
        self.ctk.CTkRadioButton(subjects_mode_row, text="Случайная для каждого", variable=self.subjects_mode_var, value="random").grid(row=0, column=1, padx=(8, 4))
        self.ctk.CTkRadioButton(subjects_mode_row, text="Одна для всей кампании", variable=self.subjects_mode_var, value="campaign").grid(row=0, column=2, padx=4)

        # === REPLY-TO / PROXY (в блоке Конфигурация) ===
        replyto_section = self.ctk.CTkFrame(config_section, fg_color="transparent")
        replyto_section.pack(fill="x", padx=12, pady=(6, 8))
        self.ctk.CTkLabel(replyto_section, text="Reply-To (адрес для ответов):").grid(row=0, column=0, sticky="w")
        self.replyto_var = self.ctk.StringVar(value="")
        self.replyto_combo = self.ctk.CTkComboBox(replyto_section, variable=self.replyto_var, values=[], width=320)
        self.replyto_combo.grid(row=0, column=1, padx=(8, 8))
        self.ctk.CTkButton(replyto_section, text="Загрузить txt", command=self._load_replyto_txt).grid(row=0, column=2)
        self.ctk.CTkLabel(replyto_section, textvariable=self.replyto_count_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4,0))

        replyto_mode_section = self.ctk.CTkFrame(config_section, fg_color="transparent")
        replyto_mode_section.pack(fill="x", padx=12, pady=(0,8))
        self.ctk.CTkLabel(replyto_mode_section, text="Режим Reply-To:").grid(row=0, column=0, sticky="w")
        self.ctk.CTkRadioButton(replyto_mode_section, text="Случайный для каждого", variable=self.replyto_mode_var, value="random").grid(row=0, column=1, padx=4)
        self.ctk.CTkRadioButton(replyto_mode_section, text="Фиксированный выбранный", variable=self.replyto_mode_var, value="fixed").grid(row=0, column=2, padx=4)

        proxy_section = self.ctk.CTkFrame(config_section, fg_color="transparent")
        proxy_section.pack(fill="x", padx=12, pady=(0, 8))
        self.ctk.CTkCheckBox(proxy_section, text="Использовать прокси", variable=self.proxy_enabled_var, onvalue=True, offvalue=False).pack(anchor="w")
        proxy_file_row = self.ctk.CTkFrame(proxy_section, fg_color="transparent")
        proxy_file_row.pack(fill="x", pady=(6, 0))
        proxy_file_row.grid_columnconfigure(1, weight=1)
        self.ctk.CTkLabel(proxy_file_row, text="Файл прокси:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ctk.CTkEntry(proxy_file_row, textvariable=self.proxy_file_var).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.ctk.CTkButton(proxy_file_row, text="📁", width=36, command=self._select_proxy_file).grid(row=0, column=2)

        # Кнопка менеджера SMTP-доменов
        domains_btn_frame = self.ctk.CTkFrame(config_section, fg_color="transparent")
        domains_btn_frame.pack(fill="x", padx=12, pady=(0, 10))
        self.ctk.CTkButton(
            domains_btn_frame,
            text="🗂 Менеджер SMTP-доменов",
            command=self._open_smtp_domains_manager,
            width=220,
            height=30,
            fg_color=("gray70", "gray30"),
        ).pack(side="left")
        self.ctk.CTkButton(
            domains_btn_frame,
            text="🧩 Переменные {{...}}",
            command=self._open_content_variables_manager,
            width=210,
            height=30,
            fg_color=("gray70", "gray30"),
        ).pack(side="left", padx=(8, 0))
        self.ctk.CTkButton(
            domains_btn_frame,
            text="⚙️ Мастер настройки",
            command=self._open_settings_wizard,
            width=190,
            height=30,
            fg_color=("gray70", "gray30"),
        ).pack(side="left", padx=(8, 0))
        self._subjects_count_label = self.ctk.CTkLabel(domains_btn_frame, text="")
        self._subjects_count_label.pack(side="left", padx=(16, 0))

        # === ШАБЛОН И ОПЦИИ (Раздел 2) ===
        template_section = self.ctk.CTkFrame(main_container, fg_color=("gray95", "gray20"), corner_radius=12)
        template_section.pack(fill="x", padx=16, pady=8)
        self.ctk.CTkLabel(template_section, text="🎨 Шаблон & Опции", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        
        template_row = self.ctk.CTkFrame(template_section, fg_color="transparent")
        template_row.pack(fill="x", padx=12, pady=8)
        template_row.grid_columnconfigure(0, weight=1)
        self.ctk.CTkLabel(template_row, text="Выбор шаблона:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.template_combo = self.ctk.CTkComboBox(template_row, variable=self.template_var, values=[])
        self.template_combo.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.ctk.CTkButton(template_row, text="↻ Обновить", width=90, command=self._refresh_templates).grid(row=0, column=2)

        attachments_folder_row = self.ctk.CTkFrame(template_section, fg_color="transparent")
        attachments_folder_row.pack(fill="x", padx=12, pady=(0, 8))
        attachments_folder_row.grid_columnconfigure(0, weight=1)
        self.ctk.CTkLabel(attachments_folder_row, text="Папка вложений:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ctk.CTkEntry(attachments_folder_row, textvariable=self.attachments_folder_var, width=360).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.ctk.CTkButton(attachments_folder_row, text="📁", command=self._select_attachments_folder).grid(row=0, column=2)
        self.attachments_info_label = self.ctk.CTkLabel(attachments_folder_row, text="Файлов: 0")
        self.attachments_info_label.grid(row=1, column=0, columnspan=3, sticky="w", padx=(0, 8), pady=(4, 0))

        template_meta_row = self.ctk.CTkFrame(template_section, fg_color="transparent")
        template_meta_row.pack(fill="x", padx=12, pady=(0, 8))
        template_meta_row.grid_columnconfigure(0, weight=1)
        self.template_vars_label = self.ctk.CTkLabel(template_meta_row, text="Переменных в шаблоне: 0")
        self.template_vars_label.grid(row=0, column=0, sticky="w")
        self.ctk.CTkButton(template_meta_row, text="Проверить переменные", command=self._check_template_variables).grid(row=0, column=1, padx=(8,0))
        self.ctk.CTkButton(
            template_meta_row,
            text="❔ Подсказка {{...}}",
            command=self._show_template_variables_hint,
            fg_color=("gray70", "gray30"),
        ).grid(row=0, column=2, padx=(8, 0))

        # Сетка опций: delay, rate limit, retry
        options_frame = self.ctk.CTkFrame(template_section, fg_color="transparent")
        options_frame.pack(fill="x", padx=12, pady=8)
        for i in range(5):
            options_frame.grid_columnconfigure(i, weight=1)

        self.ctk.CTkLabel(options_frame, text="Задержка, сек").grid(row=0, column=0, sticky="w")
        self.ctk.CTkEntry(options_frame, textvariable=self.delay_var, width=100).grid(row=1, column=0, sticky="ew", padx=(0, 8))

        self.ctk.CTkLabel(options_frame, text="Rate limit (писем/мин)").grid(row=0, column=1, sticky="w")
        self.ctk.CTkEntry(options_frame, textvariable=self.rate_limit_var, width=100).grid(row=1, column=1, sticky="ew", padx=(0, 8))

        self.ctk.CTkLabel(options_frame, text="Retry попытки").grid(row=0, column=2, sticky="w")
        retry_row = self.ctk.CTkFrame(options_frame, fg_color="transparent")
        retry_row.grid(row=1, column=2, sticky="ew")
        retry_row.grid_columnconfigure(0, weight=1)
        self.ctk.CTkEntry(retry_row, textvariable=self.retry_attempts_var, width=50).grid(row=0, column=0, sticky="ew")
        self.ctk.CTkLabel(retry_row, text="откат (сек):").grid(row=0, column=1, padx=(8, 4), sticky="w")
        self.ctk.CTkEntry(retry_row, textvariable=self.retry_backoff_var, width=50).grid(row=0, column=2, sticky="ew")

        self.ctk.CTkLabel(options_frame, text="SMTP parallel").grid(row=0, column=3, sticky="w")
        self.ctk.CTkEntry(options_frame, textvariable=self.parallel_smtp_var, width=100).grid(row=1, column=3, sticky="ew", padx=(0, 8))

        self.ctk.CTkLabel(options_frame, text="Пауза batch (сек)").grid(row=0, column=4, sticky="w")
        self.ctk.CTkEntry(options_frame, textvariable=self.batch_interval_var, width=100).grid(row=1, column=4, sticky="ew", padx=(0, 8))

        self.ctk.CTkCheckBox(
            options_frame,
            text="Параллельная отправка",
            variable=self.parallel_enabled_var,
            onvalue=True,
            offvalue=False,
        ).grid(row=2, column=3, columnspan=2, sticky="w", pady=(6,0))

        # Чекбокс dry-run
        check_frame = self.ctk.CTkFrame(template_section, fg_color="transparent")
        check_frame.pack(fill="x", padx=12, pady=8)
        self.ctk.CTkCheckBox(
            check_frame,
            text="🔄 Dry-run (без реальной отправки)",
            variable=self.dry_run_var,
            onvalue=True,
            offvalue=False,
        ).pack(anchor="w")

        # === ДЕЙСТВИЯ (Раздел 3) ===
        actions_section = self.ctk.CTkFrame(main_container, fg_color=("gray95", "gray20"), corner_radius=12)
        actions_section.pack(fill="x", padx=16, pady=8)
        self.ctk.CTkLabel(actions_section, text="⚡ Действия", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(12, 8))

        buttons_frame = self.ctk.CTkFrame(actions_section, fg_color="transparent")
        buttons_frame.pack(fill="x", padx=12, pady=8)
        for i in range(6):
            buttons_frame.grid_columnconfigure(i, weight=1)

        button_specs = [
            ("▶️  Старт", self._start_send, 0),
            ("📊 Статистика", self._show_stats, 1),
            ("👁️ Предпросмотр", self._preview_email, 2),
            ("✏️ Редактор HTML", self._open_template_editor, 3),
            ("🎨 Визуальный", self._open_visual_template_editor, 4),
            ("💾 Пресет", self._save_preset_dialog, 5),
        ]
        for label, command, col in button_specs:
            self.ctk.CTkButton(
                buttons_frame,
                text=label,
                command=command,
                height=36,
                font=("", 11),
            ).grid(row=0, column=col, sticky="ew", padx=4)

        # Дополнительные кнопки на второй строке
        buttons_frame2 = self.ctk.CTkFrame(actions_section, fg_color="transparent")
        buttons_frame2.pack(fill="x", padx=12, pady=(0, 8))
        for i in range(6):
            buttons_frame2.grid_columnconfigure(i, weight=1)

        self.ctk.CTkButton(
            buttons_frame2,
            text="⏹ Стоп",
            command=self._stop_campaign,
            height=32,
            font=("", 10),
            fg_color=("#c0392b", "#7b241c"),
            hover_color=("#e74c3c", "#922b21"),
        ).grid(row=0, column=0, sticky="ew", padx=4)

        self.ctk.CTkButton(
            buttons_frame2,
            text="📦 Тест SMTP (все)",
            command=self._test_all_smtp_accounts,
            height=32,
            font=("", 10),
            fg_color=("gray70", "gray30"),
        ).grid(row=0, column=1, sticky="ew", padx=4)

        self.ctk.CTkButton(
            buttons_frame2,
            text="🧪 Тест прокси",
            command=self._test_proxies,
            height=32,
            font=("", 10),
            fg_color=("gray70", "gray30"),
        ).grid(row=0, column=2, sticky="ew", padx=4)

        button_specs2 = [
            ("📂 Очередь JSON", self._run_queue_dialog, 3),
            ("💾 Экспорт JSON/CSV", self._export_queue_dialog, 4),
            ("📂 Загрузить пресет", self._load_preset_dialog, 5),
        ]
        for label, command, col in button_specs2:
            self.ctk.CTkButton(
                buttons_frame2,
                text=label,
                command=command,
                height=32,
                font=("", 10),
                fg_color=("gray70", "gray30"),
            ).grid(row=0, column=col, sticky="ew", padx=4)

        # === ПРОГРЕСС (Раздел 4) ===
        progress_section = self.ctk.CTkFrame(main_container, fg_color=("gray88", "gray25"), corner_radius=12)
        progress_section.pack(fill="x", padx=16, pady=8)

        pb_row = self.ctk.CTkFrame(progress_section, fg_color="transparent")
        pb_row.pack(fill="x", padx=12, pady=(10, 4))
        pb_row.grid_columnconfigure(1, weight=1)
        self.ctk.CTkLabel(pb_row, text="Прогресс:").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.progress_bar = self.ctk.CTkProgressBar(pb_row, mode="determinate")
        self.progress_bar.set(0)
        self.progress_bar.grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.progress_pct_label = self.ctk.CTkLabel(pb_row, text="0%  (0/0)", width=100)
        self.progress_pct_label.grid(row=0, column=2, sticky="e")

        counters_row = self.ctk.CTkFrame(progress_section, fg_color="transparent")
        counters_row.pack(fill="x", padx=12, pady=(0, 4))
        self.sent_label = self.ctk.CTkLabel(counters_row, text="✅ Отправлено: 0")
        self.sent_label.pack(side="left", padx=(0, 16))
        self.failed_label = self.ctk.CTkLabel(counters_row, text="❌ Ошибок: 0")
        self.failed_label.pack(side="left", padx=(0, 16))
        self.remaining_label = self.ctk.CTkLabel(counters_row, text="⏳ Осталось: —")
        self.remaining_label.pack(side="left", padx=(0, 16))
        self.eta_label = self.ctk.CTkLabel(counters_row, text="⏱ ETA: —")
        self.eta_label.pack(side="left")

        self.errors_btn = self.ctk.CTkButton(
            counters_row,
            text="🔍 Список ошибок",
            width=150,
            state="disabled",
            fg_color=("gray75", "gray35"),
            hover_color=("gray65", "gray45"),
            command=self._show_failed_window,
        )
        self.errors_btn.pack(side="right")

        self.ctk.CTkLabel(
            progress_section,
            textvariable=self.status_var,
            anchor="w",
            font=("", 11),
        ).pack(fill="x", padx=12, pady=(0, 10))

        # === ЛОГ (Раздел 5) ===
        log_section = self.ctk.CTkFrame(main_container, fg_color=("gray95", "gray20"), corner_radius=12)
        log_section.pack(fill="both", expand=True, padx=16, pady=(8, 16))
        self.ctk.CTkLabel(log_section, text="📝 Лог отправки", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        
        log_buttons = self.ctk.CTkFrame(log_section, fg_color="transparent")
        log_buttons.pack(fill="x", padx=12, pady=(0, 8))
        self.ctk.CTkButton(log_buttons, text="🗑️  Очистить лог", width=120, command=self._clear_log).pack(side="left")
        self.ctk.CTkButton(log_buttons, text="📂 Открыть лог-файл", width=140, command=self._open_log_file, fg_color=("gray70", "gray30")).pack(side="left", padx=8)
        self.ctk.CTkButton(log_buttons, text="📁 Папка проекта", width=130, command=self._open_project_folder, fg_color=("gray70", "gray30")).pack(side="left")

        self.log_widget = self.ctk.CTkTextbox(log_section, wrap="word", font=("Courier", 10))
        self.log_widget.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    def _setup_mousewheel_scrolling(self) -> None:
        """Явно включает прокрутку колёсиком для основного CTkScrollableFrame."""
        canvas = getattr(getattr(self, "main_container", None), "_parent_canvas", None)
        if canvas is None:
            return

        def _is_inside_main_container(widget) -> bool:
            current = widget
            while current is not None:
                if current is self.main_container:
                    return True
                current = getattr(current, "master", None)
            return False

        def _on_mousewheel(event):
            widget = getattr(event, "widget", None)
            if widget is None or not _is_inside_main_container(widget):
                return
            delta = int(getattr(event, "delta", 0))
            if delta == 0:
                return
            # На macOS delta обычно маленький и событий много — ограничиваем шаг.
            magnitude = max(1, min(3, abs(delta) // 60 if abs(delta) > 1 else 1))
            step = -magnitude if delta > 0 else magnitude
            try:
                canvas.yview_scroll(step, "units")
            except Exception:
                return

        def _on_mousewheel_linux_up(event):
            widget = getattr(event, "widget", None)
            if widget is None or not _is_inside_main_container(widget):
                return
            try:
                canvas.yview_scroll(-1, "units")
            except Exception:
                return

        def _on_mousewheel_linux_down(event):
            widget = getattr(event, "widget", None)
            if widget is None or not _is_inside_main_container(widget):
                return
            try:
                canvas.yview_scroll(1, "units")
            except Exception:
                return

        # macOS / Windows
        self.root.bind_all("<MouseWheel>", _on_mousewheel, add="+")
        # Linux (на случай запуска вне macOS)
        self.root.bind_all("<Button-4>", _on_mousewheel_linux_up, add="+")
        self.root.bind_all("<Button-5>", _on_mousewheel_linux_down, add="+")

    def _add_path_row(self, parent: object, row: int, label: str, variable: object, command: object) -> None:
        row_frame = self.ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(fill="x", padx=12, pady=6)
        row_frame.grid_columnconfigure(0, minsize=100)
        row_frame.grid_columnconfigure(1, weight=1)
        row_frame.grid_columnconfigure(2, minsize=45)

        self.ctk.CTkLabel(row_frame, text=label + ":").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ctk.CTkEntry(row_frame, textvariable=variable).grid(row=0, column=1, sticky="ew", padx=(0, 8))
        self.ctk.CTkButton(row_frame, text="📁", command=command).grid(row=0, column=2)

    def _select_config(self) -> None:
        messagebox.showinfo(
            "Конфиг",
            f"В упрощённом режиме используется только {self.DEFAULT_CONFIG_PATH}",
        )

    def _open_settings_wizard(self) -> None:
        """Визуальный мастер заполнения settings.yaml без ручного редактирования файлов."""
        config_path = self.base_dir / self.config_var.get()
        try:
            raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        except Exception:
            raw = {}

        smtp_raw = dict(raw.get("smtp") or {})
        message_raw = dict(raw.get("message") or {})
        delivery_raw = dict(raw.get("delivery") or {})

        win = self.ctk.CTkToplevel(self.root)
        win.title("⚙️ Мастер настройки")
        win.geometry("900x760")
        win.grab_set()

        container = self.ctk.CTkScrollableFrame(win)
        container.pack(fill="both", expand=True, padx=10, pady=10)

        # SMTP
        smtp_section = self.ctk.CTkFrame(container, fg_color=("gray95", "gray20"), corner_radius=10)
        smtp_section.pack(fill="x", pady=(0, 10))
        self.ctk.CTkLabel(smtp_section, text="SMTP", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        smtp_host_var = self.ctk.StringVar(value=str(smtp_raw.get("host", "")))
        smtp_port_var = self.ctk.StringVar(value=str(smtp_raw.get("port", "587")))
        smtp_user_var = self.ctk.StringVar(value=str(smtp_raw.get("username", "")))
        smtp_pass_var = self.ctk.StringVar(value=str(smtp_raw.get("password", "")))
        smtp_from_email_var = self.ctk.StringVar(value=str(smtp_raw.get("from_email", "")))
        smtp_from_name_var = self.ctk.StringVar(value=str(smtp_raw.get("from_name", "")))
        smtp_use_tls_var = self.ctk.BooleanVar(value=bool(smtp_raw.get("use_tls", True)))
        smtp_use_ssl_var = self.ctk.BooleanVar(value=bool(smtp_raw.get("use_ssl", False)))
        smtp_timeout_var = self.ctk.StringVar(value=str(smtp_raw.get("timeout_seconds", "30")))
        smtp_accounts_file_var = self.ctk.StringVar(value=str(smtp_raw.get("accounts_file", self.smtp_accounts_file_var.get() or "")))
        smtp_proxy_file_var = self.ctk.StringVar(value=str(smtp_raw.get("proxy_file", self.proxy_file_var.get() or "config/proxies.txt")))
        smtp_mode_var = self.ctk.StringVar(value="bulk" if smtp_accounts_file_var.get().strip() else "single")

        def _row(parent: object, label: str, var: object, show: str | None = None) -> None:
            fr = self.ctk.CTkFrame(parent, fg_color="transparent")
            fr.pack(fill="x", padx=12, pady=4)
            fr.grid_columnconfigure(1, weight=1)
            self.ctk.CTkLabel(fr, text=label, width=180, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
            kwargs = {"textvariable": var}
            if show is not None:
                kwargs["show"] = show
            self.ctk.CTkEntry(fr, **kwargs).grid(row=0, column=1, sticky="ew")

        _row(smtp_section, "Хост", smtp_host_var)
        _row(smtp_section, "Порт", smtp_port_var)
        _row(smtp_section, "Логин", smtp_user_var)
        _row(smtp_section, "Пароль", smtp_pass_var, show="*")
        _row(smtp_section, "From Email", smtp_from_email_var)
        _row(smtp_section, "From Name", smtp_from_name_var)
        _row(smtp_section, "Timeout (сек)", smtp_timeout_var)

        mode_row = self.ctk.CTkFrame(smtp_section, fg_color="transparent")
        mode_row.pack(fill="x", padx=12, pady=(4, 4))
        self.ctk.CTkLabel(mode_row, text="Режим SMTP", width=180, anchor="w").pack(side="left")
        self.ctk.CTkRadioButton(mode_row, text="Один аккаунт", variable=smtp_mode_var, value="single").pack(side="left", padx=(0, 8))
        self.ctk.CTkRadioButton(mode_row, text="Пакет аккаунтов (TXT/CSV)", variable=smtp_mode_var, value="bulk").pack(side="left")

        self.ctk.CTkLabel(
            smtp_section,
            text="Подсказка: в режиме «Пакет аккаунтов» поля Хост/Логин/Пароль не обязательны — берутся из файла.",
            font=("", 10),
            text_color=("gray40", "gray60"),
        ).pack(anchor="w", padx=12, pady=(0, 6))

        flags_row = self.ctk.CTkFrame(smtp_section, fg_color="transparent")
        flags_row.pack(fill="x", padx=12, pady=(4, 8))
        self.ctk.CTkCheckBox(flags_row, text="STARTTLS", variable=smtp_use_tls_var, onvalue=True, offvalue=False).pack(side="left")
        self.ctk.CTkCheckBox(flags_row, text="SSL", variable=smtp_use_ssl_var, onvalue=True, offvalue=False).pack(side="left", padx=(12, 0))

        # accounts/proxy files
        smtp_accounts_info_var = self.ctk.StringVar(value="")

        def _load_accounts_with_wizard_defaults(accounts_path: Path) -> list[object]:
            from .config import _load_smtp_accounts, _load_smtp_accounts_txt
            from .smtp_domains import load_domains as _load_domains

            if accounts_path.suffix.lower() == ".txt":
                port_text = smtp_port_var.get().strip()
                timeout_text = smtp_timeout_var.get().strip()
                defaults = {
                    "host": smtp_host_var.get().strip() or None,
                    "port": int(port_text) if port_text else None,
                    "from_name": smtp_from_name_var.get().strip() or None,
                    "use_tls": bool(smtp_use_tls_var.get()),
                    "use_ssl": bool(smtp_use_ssl_var.get()),
                    "timeout_seconds": int(timeout_text or "30"),
                }
                return _load_smtp_accounts_txt(
                    accounts_path,
                    defaults,
                    domains_db=_load_domains(self.base_dir),
                    base_dir=self.base_dir,
                )
            return _load_smtp_accounts(accounts_path)

        def _count_non_comment_lines(accounts_path: Path) -> int:
            count = 0
            with accounts_path.open("r", encoding="utf-8-sig") as handle:
                for line in handle:
                    text = line.strip()
                    if text and not text.startswith("#"):
                        count += 1
            return count

        def _refresh_accounts_info() -> None:
            value = smtp_accounts_file_var.get().strip()
            if not value:
                smtp_accounts_info_var.set("")
                return
            accounts_path = (self.base_dir / value).resolve()
            if not accounts_path.exists():
                smtp_accounts_info_var.set("⚠️ Файл не найден")
                return
            if self._looks_like_smtp_domains_file(accounts_path):
                smtp_accounts_info_var.set("⚠️ Это smtp_domains.txt (файл доменов), выберите файл аккаунтов")
                return
            try:
                accounts = _load_accounts_with_wizard_defaults(accounts_path)
                smtp_accounts_info_var.set(f"✅ Найдено аккаунтов: {len(accounts)}")
            except Exception as error:  # noqa: BLE001
                try:
                    rough_count = _count_non_comment_lines(accounts_path)
                    if rough_count > 0:
                        smtp_accounts_info_var.set(
                            f"⚠️ Частичная проверка: строк={rough_count} (есть неразобранные записи)"
                        )
                    else:
                        smtp_accounts_info_var.set("⚠️ Файл пуст")
                except Exception:
                    smtp_accounts_info_var.set(f"⚠️ Ошибка файла: {error}")

        smtp_accounts_row = self.ctk.CTkFrame(smtp_section, fg_color="transparent")
        smtp_accounts_row.pack(fill="x", padx=12, pady=4)
        smtp_accounts_row.grid_columnconfigure(1, weight=1)
        self.ctk.CTkLabel(smtp_accounts_row, text="Файл SMTP аккаунтов", width=180, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ctk.CTkEntry(smtp_accounts_row, textvariable=smtp_accounts_file_var).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        def _pick_accounts_file() -> None:
            p = filedialog.askopenfilename(
                initialdir=self.base_dir,
                filetypes=[("TXT/CSV", "*.txt *.csv"), ("Все", "*.*")],
            )
            if p:
                picked_path = Path(p)
                if self._looks_like_smtp_domains_file(picked_path):
                    messagebox.showwarning(
                        "SMTP аккаунты",
                        "Вы выбрали файл SMTP-доменов (smtp_domains.txt).\n"
                        "Для поля «Файл SMTP аккаунтов» нужен файл с логинами/паролями:\n"
                        "- username:password\n"
                        "- или host|port|username|password|from_email|from_name",
                        parent=win,
                    )
                    return
                smtp_accounts_file_var.set(self._relative(Path(p)))
                _refresh_accounts_info()

        self.ctk.CTkButton(smtp_accounts_row, text="📁", width=36, command=_pick_accounts_file).grid(row=0, column=2)
        self.ctk.CTkLabel(
            smtp_accounts_row,
            textvariable=smtp_accounts_info_var,
            font=("", 10),
            text_color=("gray40", "gray60"),
        ).grid(row=1, column=1, sticky="w", pady=(4, 0))

        proxy_file_row = self.ctk.CTkFrame(smtp_section, fg_color="transparent")
        proxy_file_row.pack(fill="x", padx=12, pady=4)
        proxy_file_row.grid_columnconfigure(1, weight=1)
        self.ctk.CTkLabel(proxy_file_row, text="Файл прокси", width=180, anchor="w").grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.ctk.CTkEntry(proxy_file_row, textvariable=smtp_proxy_file_var).grid(row=0, column=1, sticky="ew", padx=(0, 8))

        def _pick_proxy_file() -> None:
            p = filedialog.askopenfilename(
                initialdir=self.base_dir,
                filetypes=[("TXT", "*.txt"), ("Все", "*.*")],
            )
            if p:
                smtp_proxy_file_var.set(self._relative(Path(p)))

        self.ctk.CTkButton(proxy_file_row, text="📁", width=36, command=_pick_proxy_file).grid(row=0, column=2)

        smtp_accounts_file_var.trace_add("write", lambda *_: _refresh_accounts_info())
        _refresh_accounts_info()

        # Message
        msg_section = self.ctk.CTkFrame(container, fg_color=("gray95", "gray20"), corner_radius=10)
        msg_section.pack(fill="x", pady=(0, 10))
        self.ctk.CTkLabel(msg_section, text="Письмо", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        msg_subject_var = self.ctk.StringVar(value=str(message_raw.get("subject", "")))
        msg_template_var = self.ctk.StringVar(value=str(message_raw.get("template", self.template_var.get() or "")))
        msg_reply_to_var = self.ctk.StringVar(value=str(message_raw.get("reply_to", self.replyto_var.get() or "")))
        _row(msg_section, "Тема по умолчанию", msg_subject_var)
        _row(msg_section, "Шаблон", msg_template_var)
        _row(msg_section, "Reply-To", msg_reply_to_var)

        # Delivery
        d_section = self.ctk.CTkFrame(container, fg_color=("gray95", "gray20"), corner_radius=10)
        d_section.pack(fill="x", pady=(0, 10))
        self.ctk.CTkLabel(d_section, text="Отправка", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(10, 6))

        d_delay_var = self.ctk.StringVar(value=str(delivery_raw.get("delay_seconds", self.delay_var.get() or "0")))
        d_rate_var = self.ctk.StringVar(value=str(delivery_raw.get("rate_limit_per_minute", self.rate_limit_var.get() or "")))
        d_retry_var = self.ctk.StringVar(value=str(delivery_raw.get("retry_attempts", self.retry_attempts_var.get() or "1")))
        d_backoff_var = self.ctk.StringVar(value=str(delivery_raw.get("retry_backoff_seconds", self.retry_backoff_var.get() or "5")))
        d_parallel_var = self.ctk.StringVar(value=str(delivery_raw.get("parallel_smtp_accounts", self.parallel_smtp_var.get() or "1")))
        d_batch_var = self.ctk.StringVar(value=str(delivery_raw.get("batch_interval_seconds", self.batch_interval_var.get() or "0")))
        d_log_var = self.ctk.StringVar(value=str(delivery_raw.get("log_file", "logs/email_app.log")))
        d_hist_csv_var = self.ctk.StringVar(value=str(delivery_raw.get("history_csv", "history/email_history.csv")))
        d_hist_jsonl_var = self.ctk.StringVar(value=str(delivery_raw.get("history_jsonl", "history/email_history.jsonl")))

        _row(d_section, "Delay (сек)", d_delay_var)
        _row(d_section, "Rate limit / мин", d_rate_var)
        _row(d_section, "Retry попытки", d_retry_var)
        _row(d_section, "Retry backoff (сек)", d_backoff_var)
        _row(d_section, "Parallel SMTP", d_parallel_var)
        _row(d_section, "Batch interval", d_batch_var)
        _row(d_section, "Лог файл", d_log_var)
        _row(d_section, "History CSV", d_hist_csv_var)
        _row(d_section, "History JSONL", d_hist_jsonl_var)

        btns = self.ctk.CTkFrame(container, fg_color="transparent")
        btns.pack(fill="x", pady=(2, 8))

        def _save_settings() -> None:
            try:
                accounts_file_text = smtp_accounts_file_var.get().strip()
                proxy_file_text = smtp_proxy_file_var.get().strip()
                smtp_mode = smtp_mode_var.get().strip() or "single"

                smtp_payload = dict(raw.get("smtp") or {})

                # Общие поля (для single и bulk)
                timeout_text = smtp_timeout_var.get().strip()
                smtp_payload["timeout_seconds"] = int(timeout_text or "30")
                smtp_payload["use_tls"] = bool(smtp_use_tls_var.get())
                smtp_payload["use_ssl"] = bool(smtp_use_ssl_var.get())

                # Прокси-файл
                if proxy_file_text:
                    smtp_payload["proxy_file"] = self._portable_path_value(proxy_file_text)
                else:
                    smtp_payload.pop("proxy_file", None)

                if smtp_mode == "bulk":
                    if not accounts_file_text:
                        raise ValueError("В режиме «Пакет аккаунтов» укажите файл SMTP-аккаунтов")
                    accounts_path = (self.base_dir / accounts_file_text).resolve()
                    if not accounts_path.exists():
                        raise ValueError(f"Файл SMTP-аккаунтов не найден: {accounts_file_text}")
                    rough_count = 0
                    try:
                        _accounts = _load_accounts_with_wizard_defaults(accounts_path)
                        if not _accounts:
                            raise ValueError("Файл SMTP-аккаунтов пуст")
                    except Exception as error:  # noqa: BLE001
                        # Не блокируем сохранение мастера полностью, если файл неполный:
                        # допускаем сохранение при наличии строк, чтобы пользователь мог
                        # донастроить домены/host позже из GUI.
                        rough_count = _count_non_comment_lines(accounts_path)
                        if rough_count <= 0:
                            raise ValueError(f"Ошибка SMTP-файла: {error}") from error
                        self._append_log(
                            "⚠️ SMTP-файл сохранён с предупреждением: есть строки, которые пока не удалось разобрать. "
                            f"Строк в файле: {rough_count}. Подробно: {error}"
                        )

                    smtp_payload["accounts_file"] = self._portable_path_value(accounts_file_text)

                    # Не заставляем пользователя заполнять single-поля,
                    # но если он их ввёл — сохраняем как fallback.
                    if smtp_host_var.get().strip():
                        smtp_payload["host"] = smtp_host_var.get().strip()
                    if smtp_port_var.get().strip():
                        smtp_payload["port"] = int(smtp_port_var.get().strip())
                    if smtp_user_var.get().strip():
                        smtp_payload["username"] = smtp_user_var.get().strip()
                    if smtp_pass_var.get().strip():
                        smtp_payload["password"] = smtp_pass_var.get().strip()
                    if smtp_from_email_var.get().strip():
                        smtp_payload["from_email"] = smtp_from_email_var.get().strip()
                    if smtp_from_name_var.get().strip():
                        smtp_payload["from_name"] = smtp_from_name_var.get().strip()
                else:
                    # Single-account режим
                    smtp_payload.pop("accounts_file", None)
                    smtp_payload["host"] = smtp_host_var.get().strip()
                    smtp_payload["port"] = int(smtp_port_var.get().strip() or "0")
                    smtp_payload["username"] = smtp_user_var.get().strip()
                    smtp_payload["password"] = smtp_pass_var.get().strip()
                    smtp_payload["from_email"] = smtp_from_email_var.get().strip()
                    smtp_payload["from_name"] = smtp_from_name_var.get().strip()

                    required = ["host", "port", "username", "password", "from_email", "from_name"]
                    missing = [key for key in required if not smtp_payload.get(key)]
                    if missing:
                        raise ValueError("Заполните SMTP поля: " + ", ".join(missing))

                message_payload = {
                    "subject": msg_subject_var.get().strip() or "Без темы",
                    "template": msg_template_var.get().strip(),
                    "reply_to": msg_reply_to_var.get().strip() or None,
                    "attachments": list((raw.get("message") or {}).get("attachments") or []),
                    "inline_images": dict((raw.get("message") or {}).get("inline_images") or {}),
                }

                delivery_payload = {
                    "delay_seconds": float(d_delay_var.get().strip() or "0"),
                    "log_file": d_log_var.get().strip() or "logs/email_app.log",
                    "history_csv": d_hist_csv_var.get().strip() or "history/email_history.csv",
                    "history_jsonl": d_hist_jsonl_var.get().strip() or "history/email_history.jsonl",
                    "parallel_smtp_accounts": int(d_parallel_var.get().strip() or "1"),
                    "batch_interval_seconds": float(d_batch_var.get().strip() or "0"),
                    "retry_attempts": int(d_retry_var.get().strip() or "1"),
                    "retry_backoff_seconds": float(d_backoff_var.get().strip() or "5"),
                }
                if d_rate_var.get().strip():
                    delivery_payload["rate_limit_per_minute"] = int(d_rate_var.get().strip())

                raw["smtp"] = smtp_payload
                raw["message"] = message_payload
                raw["delivery"] = {**(raw.get("delivery") or {}), **delivery_payload}

                config_path.parent.mkdir(parents=True, exist_ok=True)
                config_path.write_text(yaml.safe_dump(raw, allow_unicode=True, sort_keys=False), encoding="utf-8")

                # синхронизация текущих виджетов
                self.smtp_accounts_file_var.set(accounts_file_text if smtp_mode == "bulk" else "")
                self.proxy_file_var.set(proxy_file_text or self.proxy_file_var.get())
                self.template_var.set(msg_template_var.get().strip())
                self.replyto_var.set(msg_reply_to_var.get().strip())
                self.delay_var.set(str(delivery_payload["delay_seconds"]))
                self.rate_limit_var.set(str(delivery_payload.get("rate_limit_per_minute", "")))
                self.retry_attempts_var.set(str(delivery_payload["retry_attempts"]))
                self.retry_backoff_var.set(str(delivery_payload["retry_backoff_seconds"]))
                self.parallel_smtp_var.set(str(delivery_payload["parallel_smtp_accounts"]))
                self.batch_interval_var.set(str(delivery_payload["batch_interval_seconds"]))

                self._append_log(f"⚙️ Настройки сохранены: {self.config_var.get()}")
                self.status_var.set("✓ Настройки сохранены")
                self._update_eta_estimate()
                self._refresh_templates()
                win.destroy()
            except Exception as error:  # noqa: BLE001
                messagebox.showerror("Мастер настройки", f"Не удалось сохранить: {error}", parent=win)

        self.ctk.CTkButton(btns, text="💾 Сохранить настройки", command=_save_settings).pack(side="left")
        self.ctk.CTkButton(btns, text="✕ Закрыть", command=win.destroy, fg_color=("gray70", "gray30")).pack(side="right")

    def _select_smtp_accounts_file(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("SMTP accounts", "*.txt *.csv"), ("TXT", "*.txt"), ("CSV", "*.csv"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        if self._looks_like_smtp_domains_file(Path(path)):
            messagebox.showwarning(
                "SMTP аккаунты",
                "Вы выбрали smtp_domains.txt (база доменов).\n"
                "Выберите файл SMTP-аккаунтов с логинами/паролями.",
            )
            return

        relative_path = self._relative(Path(path))
        self.smtp_accounts_file_var.set(relative_path)
        self._apply_smtp_accounts_file_to_config(relative_path)

    def _apply_smtp_accounts_file_to_config(self, accounts_file: str) -> None:
        config_path = self.base_dir / self.config_var.get()
        try:
            raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
            smtp = raw.get("smtp") or {}
            smtp["accounts_file"] = self._portable_path_value(accounts_file)
            raw["smtp"] = smtp
            config_path.write_text(yaml.safe_dump(raw, allow_unicode=True, sort_keys=False), encoding="utf-8")
            try:
                from .config import _load_smtp_accounts
                _accs = _load_smtp_accounts(self.base_dir / accounts_file)
                self._append_log(f"SMTP accounts file подключён: {accounts_file} ({len(_accs)} аккаунтов)")
                self.status_var.set(f"✓ SMTP-аккаунты подключены: {len(_accs)}")
            except Exception as error:  # noqa: BLE001
                self._append_log(f"⚠️ SMTP файл подключён, но есть проблемы валидации: {error}")
                self.status_var.set("⚠️ Проверьте SMTP файл")
                messagebox.showwarning(
                    "SMTP аккаунты",
                    "Файл подключён, но есть ошибки формата/данных.\n"
                    "Для реальной проверки логин/пароль используйте кнопку «📦 Тест SMTP (все)».\n\n"
                    f"Детали: {error}",
                )
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("SMTP аккаунты", f"Не удалось обновить конфиг: {error}")

    def _show_template_variables_hint(self) -> None:
        recipients_file = self.recipients_var.get().strip() or "recipients.csv"
        smtp_accounts_file = self.smtp_accounts_file_var.get().strip() or "(не выбран)"
        messagebox.showinfo(
            "Подсказка {{...}}",
            "Как использовать переменные в шаблоне:\n\n"
            "1) Переменные из content:\n"
            "   {{ приветствие }}\n"
            "   {{ кнопка }}\n\n"
            "2) Случайное значение из списка:\n"
            "   {{ приветствие | random }}\n\n"
            "3) Данные получателя:\n"
            "   {{ recipient.email }}\n"
            "   {{ recipient.name }}\n\n"
            "4) Поля из recipients.csv/txt:\n"
            "   {{ recipient.phone }}\n"
            "   {{ recipient.city }}\n\n"
            "5) Данные SMTP аккаунта (из выбранного файла SMTP аккаунтов):\n"
            "   {{ smtp.username }}\n"
            "   {{ smtp.from_email }}\n"
            "   {{ smtp.from_name }}\n"
            "   {{ smtp.host }} / {{ smtp.port }}\n\n"
            "6) Текст письма из поля «Тексты писем»:\n"
            "   {{ body_text }}  (или {{ message_text }})\n\n"
            "Важно: все {{ recipient.* }} берутся из файла получателей, выбранного в поле «Получатели»:\n"
            f"{recipients_file}\n\n"
            "Источник {{ smtp.* }}: поле «SMTP аккаунты»:\n"
            f"{smtp_accounts_file}",
        )

    def _select_recipients(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("CSV/TXT", "*.csv *.txt"), ("CSV", "*.csv"), ("TXT", "*.txt"), ("Все файлы", "*.*")],
        )
        if path:
            self.recipients_var.set(self._relative(Path(path)))
            self._load_recipients_for_preview()
            self._update_recipients_count(Path(path))

    def _select_templates(self) -> None:
        path = filedialog.askdirectory(initialdir=self.base_dir)
        if path:
            self.templates_var.set(self._relative(Path(path)))
            self._refresh_templates()

    def _select_attachments_folder(self) -> None:
        path = filedialog.askdirectory(initialdir=self.base_dir)
        if path:
            self.attachments_folder_var.set(self._relative(Path(path)))
            self._update_random_attachment_count()

    def _select_proxy_file(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("Text files", "*.txt"), ("Все файлы", "*.*")],
        )
        if path:
            relative_path = self._relative(Path(path))
            self.proxy_file_var.set(relative_path)
            self._apply_proxy_file_to_config(relative_path)

    def _apply_proxy_file_to_config(self, proxy_file: str) -> None:
        config_path = self.base_dir / self.config_var.get()
        try:
            raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
            smtp = raw.get("smtp") or {}
            smtp["proxy_file"] = self._portable_path_value(proxy_file)
            raw["smtp"] = smtp
            config_path.write_text(yaml.safe_dump(raw, allow_unicode=True, sort_keys=False), encoding="utf-8")
            self._append_log(f"Прокси-файл подключён: {proxy_file}")
            self.status_var.set("✓ Прокси-файл подключён")
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Прокси", f"Не удалось обновить конфиг: {error}")

    # ------------------------------------------------------------------ #
    # Темы писем (subjects TXT)                                           #
    # ------------------------------------------------------------------ #

    def _select_subjects_file(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("TXT", "*.txt"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        relative_path = self._relative(Path(path))
        self.subjects_file_var.set(relative_path)
        self._load_subjects_file(Path(path))

    def _load_subjects_file(self, path: Path) -> None:
        """Загружает список тем из TXT-файла (одна тема — одна строка)."""
        subjects: list[str] = []
        try:
            with path.open("r", encoding="utf-8") as handle:
                for line in handle:
                    s = line.strip()
                    if s and not s.startswith("#"):
                        subjects.append(s)
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Темы", f"Не удалось загрузить файл тем: {error}")
            return
        self._subjects_list = subjects
        self._subjects_count_label.configure(text=f"Тем загружено: {len(subjects)}")
        self._append_log(f"Загружено тем: {len(subjects)} из {self.subjects_file_var.get()}")

    def _pick_random_subject(self) -> str | None:
        """Случайно выбирает тему из загруженного списка (или None если список пуст)."""
        if self._subjects_list:
            import random as _random
            return _random.choice(self._subjects_list)
        return None

    # ------------------------------------------------------------------ #
    # SMTP Domains Manager (редактор базы доменов)                        #
    # ------------------------------------------------------------------ #

    def _open_smtp_domains_manager(self) -> None:
        """Открывает диалог редактора базы SMTP-доменов."""
        import tkinter as tk

        domains_file = self.base_dir / "config" / "smtp_domains.txt"
        domains: dict = load_domains(self.base_dir)

        win = self.ctk.CTkToplevel(self.root)
        win.title("🗂 SMTP Домены — менеджер")
        win.geometry("780x540")
        win.grab_set()

        # --- Заголовок ---
        self.ctk.CTkLabel(
            win,
            text="🗂 База SMTP-доменов (автоопределение по email)",
            font=("", 14, "bold"),
        ).pack(anchor="w", padx=12, pady=(12, 4))
        self.ctk.CTkLabel(
            win,
            text=f"Файл: {domains_file}",
            font=("", 10),
            text_color="gray",
        ).pack(anchor="w", padx=12, pady=(0, 8))

        # --- Treeview ---
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=4)

        columns = ("domain", "host", "port", "type")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", height=16)
        tree.heading("domain", text="Домен")
        tree.heading("host", text="SMTP Хост")
        tree.heading("port", text="Порт")
        tree.heading("type", text="Тип соединения")
        tree.column("domain", width=160, anchor="w")
        tree.column("host", width=220, anchor="w")
        tree.column("port", width=60, anchor="center")
        tree.column("type", width=120, anchor="center")

        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscroll=scrollbar.set)
        tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        def refresh_tree() -> None:
            tree.delete(*tree.get_children())
            for domain, s in sorted(domains.items()):
                conn_label = _flags_to_label(s.get("use_tls", False), s.get("use_ssl", False))
                tree.insert("", "end", values=(f"@{domain}", s["host"], s["port"], conn_label))

        refresh_tree()

        # --- Кнопки ---
        btn_frame = tk.Frame(win)
        btn_frame.pack(fill="x", padx=12, pady=8)

        def add_domain() -> None:
            self._open_domain_edit_dialog(win, domains, None, refresh_tree)

        def edit_domain() -> None:
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("SMTP Домены", "Выберите строку для редактирования")
                return
            item = tree.item(sel[0])
            domain_key = str(item["values"][0]).lstrip("@")
            self._open_domain_edit_dialog(win, domains, domain_key, refresh_tree)

        def delete_domain() -> None:
            sel = tree.selection()
            if not sel:
                return
            item = tree.item(sel[0])
            domain_key = str(item["values"][0]).lstrip("@")
            if messagebox.askyesno("Удалить", f"Удалить домен @{domain_key}?", parent=win):
                domains.pop(domain_key, None)
                refresh_tree()

        def save_domains() -> None:
            try:
                save_smtp_domains_file(domains_file, domains)
                self._append_log(f"SMTP домены сохранены: {domains_file}")
                messagebox.showinfo("Сохранено", f"Файл обновлён:\n{domains_file}", parent=win)
            except Exception as error:  # noqa: BLE001
                messagebox.showerror("Ошибка", f"Не удалось сохранить: {error}", parent=win)

        tk.Button(btn_frame, text="➕  Добавить", command=add_domain, padx=8, pady=4).pack(side="left", padx=4)
        tk.Button(btn_frame, text="✏️  Изменить", command=edit_domain, padx=8, pady=4).pack(side="left", padx=4)
        tk.Button(btn_frame, text="🗑  Удалить",  command=delete_domain, padx=8, pady=4).pack(side="left", padx=4)
        tk.Button(btn_frame, text="💾  Сохранить", command=save_domains, padx=8, pady=4, bg="#2a7d2a", fg="white").pack(side="left", padx=4)
        tk.Button(btn_frame, text="✕  Закрыть",  command=win.destroy, padx=8, pady=4).pack(side="right", padx=4)

    def _open_content_variables_manager(self) -> None:
        """Менеджер переменных из settings.yaml/content для {{ ... }} в шаблоне."""
        import tkinter as tk

        config_path = self.base_dir / self.config_var.get()
        try:
            raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        except Exception:
            raw = {}

        content_data: dict = dict(raw.get("content") or {})
        recipients_file = self.recipients_var.get().strip() or "recipients.csv"
        smtp_accounts_file = self.smtp_accounts_file_var.get().strip() or "(не выбран)"

        win = self.ctk.CTkToplevel(self.root)
        win.title("🧩 Переменные письма ({{ ... }})")
        win.geometry("960x620")
        win.grab_set()

        header = self.ctk.CTkFrame(win, fg_color="transparent")
        header.pack(fill="x", padx=12, pady=(10, 6))
        self.ctk.CTkLabel(
            header,
            text=(
                "Подсказка: {{ имя_переменной }}; для списков — {{ имя_переменной | random }}. "
                f"{{{{ recipient.* }}}} берётся из файла: {recipients_file}"
            ),
            font=("", 10),
            text_color=("gray40", "gray60"),
        ).pack(anchor="w")
        self.ctk.CTkLabel(
            header,
            text=(
                "SMTP-переменные в шаблоне: {{ smtp.username }}, {{ smtp.from_email }}, "
                "{{ smtp.from_name }}, {{ smtp.host }}, {{ smtp.port }}. "
                f"Источник: {smtp_accounts_file}"
            ),
            font=("", 10),
            text_color=("gray40", "gray60"),
        ).pack(anchor="w", pady=(2, 0))

        try:
            _accs = self._resolve_all_smtp_accounts_for_test()
            if _accs:
                _a = _accs[0]
                self.ctk.CTkLabel(
                    header,
                    text=(
                        "Пример значений smtp.* (1-й аккаунт): "
                        f"smtp.username={_a.username}, smtp.from_email={_a.from_email}, "
                        f"smtp.host={_a.host}, smtp.port={_a.port}"
                    ),
                    font=("", 10),
                    text_color=("gray40", "gray60"),
                ).pack(anchor="w", pady=(2, 0))
        except Exception:
            pass

        body = self.ctk.CTkFrame(win)
        body.pack(fill="both", expand=True, padx=12, pady=8)
        body.grid_columnconfigure(0, weight=1)
        body.grid_columnconfigure(1, weight=1)
        body.grid_rowconfigure(0, weight=1)

        left = self.ctk.CTkFrame(body)
        right = self.ctk.CTkFrame(body)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
        right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

        # Левая часть: список переменных
        tree = ttk.Treeview(left, columns=("name", "type", "preview"), show="headings", height=18)
        tree.heading("name", text="Переменная")
        tree.heading("type", text="Тип")
        tree.heading("preview", text="Значение")
        tree.column("name", width=180, anchor="w")
        tree.column("type", width=90, anchor="center")
        tree.column("preview", width=320, anchor="w")
        tree.pack(fill="both", expand=True, padx=8, pady=8)

        def _preview_value(value: object) -> tuple[str, str]:
            if isinstance(value, list):
                text = " | ".join(str(x) for x in value[:3])
                if len(value) > 3:
                    text += " ..."
                return "list", text
            return "string", str(value)

        def _refresh_tree() -> None:
            tree.delete(*tree.get_children())
            for key, value in sorted(content_data.items()):
                vtype, preview = _preview_value(value)
                tree.insert("", "end", values=(str(key), vtype, preview))

        _refresh_tree()

        # Правая часть: редактор переменной
        self.ctk.CTkLabel(right, text="Имя переменной", anchor="w").pack(fill="x", padx=10, pady=(10, 4))
        key_var = self.ctk.StringVar(value="")
        key_entry = self.ctk.CTkEntry(right, textvariable=key_var)
        key_entry.pack(fill="x", padx=10)

        self.ctk.CTkLabel(
            right,
            text="Значение (1 строка = string, несколько строк = list/random)",
            anchor="w",
        ).pack(fill="x", padx=10, pady=(10, 4))
        value_box = self.ctk.CTkTextbox(right, wrap="word", height=260)
        value_box.pack(fill="both", expand=True, padx=10, pady=(0, 10))

        def _set_editor(key: str, value: object) -> None:
            key_var.set(key)
            value_box.delete("1.0", "end")
            if isinstance(value, list):
                value_box.insert("1.0", "\n".join(str(x) for x in value))
            else:
                value_box.insert("1.0", str(value))

        def _on_select(_event=None) -> None:
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values")
            if not vals:
                return
            key = str(vals[0])
            if key in content_data:
                _set_editor(key, content_data[key])

        tree.bind("<<TreeviewSelect>>", _on_select)

        def _validate_key(name: str) -> str:
            key = name.strip()
            if not key:
                raise ValueError("Имя переменной не может быть пустым")
            if any(ch in key for ch in " {}"):
                raise ValueError("Имя переменной не должно содержать пробелы, '{' или '}'")
            return key

        def _read_editor_value() -> object:
            text = value_box.get("1.0", "end").strip()
            lines = [line.strip() for line in text.splitlines() if line.strip()]
            if len(lines) <= 1:
                return lines[0] if lines else ""
            return lines

        def _add_or_update() -> None:
            try:
                key = _validate_key(key_var.get())
            except ValueError as error:
                messagebox.showerror("Переменные", str(error), parent=win)
                return
            content_data[key] = _read_editor_value()
            _refresh_tree()

        def _delete_var() -> None:
            sel = tree.selection()
            if not sel:
                return
            vals = tree.item(sel[0], "values")
            if not vals:
                return
            key = str(vals[0])
            if key in content_data and messagebox.askyesno("Удалить", f"Удалить переменную '{key}'?", parent=win):
                content_data.pop(key, None)
                _refresh_tree()
                key_var.set("")
                value_box.delete("1.0", "end")

        def _new_var() -> None:
            key_var.set("")
            value_box.delete("1.0", "end")
            key_entry.focus_set()

        buttons = self.ctk.CTkFrame(right, fg_color="transparent")
        buttons.pack(fill="x", padx=10, pady=(0, 10))
        self.ctk.CTkButton(buttons, text="➕ Новый", command=_new_var, width=90).pack(side="left")
        self.ctk.CTkButton(buttons, text="💾 Добавить/Обновить", command=_add_or_update).pack(side="left", padx=(8, 0))
        self.ctk.CTkButton(buttons, text="🗑 Удалить", command=_delete_var, fg_color=("#a63d3d", "#7a2a2a")).pack(side="left", padx=(8, 0))

        footer = self.ctk.CTkFrame(win, fg_color="transparent")
        footer.pack(fill="x", padx=12, pady=(0, 10))

        def _save_all() -> None:
            try:
                raw_local = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
            except Exception:
                raw_local = {}
            raw_local["content"] = content_data
            try:
                config_path.write_text(yaml.safe_dump(raw_local, allow_unicode=True, sort_keys=False), encoding="utf-8")
                self._append_log("🧩 Переменные content сохранены")
                self.status_var.set("✓ Переменные сохранены")
                messagebox.showinfo(
                    "Сохранено",
                    "Переменные сохранены.\n\nПримеры использования в шаблоне:\n"
                    "- {{ имя_переменной }}\n"
                    "- {{ имя_переменной | random }}\n"
                    "- {{ recipient.email }}\n"
                    "- {{ recipient.name }}\n"
                    "- {{ body_text }} (текст письма из TXT)\n"
                    "- {{ smtp.username }}\n"
                    "- {{ smtp.from_email }}\n\n"
                    "Источник для {{ recipient.* }}: файл из поля «Получатели».\n"
                    f"Текущий файл: {recipients_file}\n\n"
                    "Источник для {{ smtp.* }}: файл из поля «SMTP аккаунты».\n"
                    f"Текущий файл: {smtp_accounts_file}",
                    parent=win,
                )
                win.destroy()
            except Exception as error:  # noqa: BLE001
                messagebox.showerror("Переменные", f"Не удалось сохранить: {error}", parent=win)

        self.ctk.CTkButton(footer, text="💾 Сохранить переменные", command=_save_all).pack(side="left")
        self.ctk.CTkButton(footer, text="❔ Подсказка {{...}}", command=lambda: messagebox.showinfo(
            "Подсказка",
            "Как использовать переменные в HTML-шаблоне:\n\n"
            "1) Обычное значение:\n"
            "   {{ приветствие }}\n\n"
            "2) Случайное значение из списка:\n"
            "   {{ приветствие | random }}\n\n"
            "3) Данные получателя:\n"
            "   {{ recipient.email }}\n"
            "   {{ recipient.name }}\n\n"
            "4) Данные SMTP аккаунта:\n"
            "   {{ smtp.username }}\n"
            "   {{ smtp.from_email }}\n"
            "   {{ smtp.host }}\n\n"
            "5) Текст письма из TXT:\n"
            "   {{ body_text }} (или {{ message_text }})\n\n"
            "6) Цепочка с макросами из CSV поддерживается автоматически.\n\n"
            "Важно: {{ recipient.email }}, {{ recipient.name }} и другие {{ recipient.* }} "
            "читаются из файла, указанного в поле «Получатели».\n"
            f"Текущий файл: {recipients_file}\n\n"
            "{{ smtp.* }} читаются из активного SMTP аккаунта (из поля «SMTP аккаунты»).\n"
            f"Текущий файл: {smtp_accounts_file}",
            parent=win,
        ), fg_color=("gray70", "gray30")).pack(side="left", padx=(8, 0))
        self.ctk.CTkButton(footer, text="✕ Закрыть", command=win.destroy, fg_color=("gray70", "gray30")).pack(side="right")

    def _open_domain_edit_dialog(
        self,
        parent: object,
        domains: dict,
        domain_key: str | None,
        on_save: object,
    ) -> None:
        """Диалог добавления / редактирования записи домена."""
        import tkinter as tk

        existing = domains.get(domain_key, {}) if domain_key else {}
        use_ssl = existing.get("use_ssl", False)
        use_tls = existing.get("use_tls", False)
        conn_default = _flags_to_label(use_tls, use_ssl)

        dialog = tk.Toplevel(parent)
        dialog.title("Добавить домен" if domain_key is None else f"Изменить @{domain_key}")
        dialog.geometry("430x290")
        dialog.grab_set()
        dialog.resizable(False, False)

        pad = {"padx": 14, "pady": 8}

        tk.Label(dialog, text="Домен (без @):").grid(row=0, column=0, sticky="w", **pad)
        domain_entry = tk.Entry(dialog, width=34)
        domain_entry.insert(0, domain_key or "")
        domain_entry.grid(row=0, column=1, **pad)

        tk.Label(dialog, text="SMTP Хост:").grid(row=1, column=0, sticky="w", **pad)
        host_entry = tk.Entry(dialog, width=34)
        host_entry.insert(0, existing.get("host", ""))
        host_entry.grid(row=1, column=1, **pad)

        tk.Label(dialog, text="Порт:").grid(row=2, column=0, sticky="w", **pad)
        port_entry = tk.Entry(dialog, width=34)
        port_entry.insert(0, str(existing.get("port", "465")))
        port_entry.grid(row=2, column=1, **pad)

        tk.Label(dialog, text="Тип соединения:").grid(row=3, column=0, sticky="w", **pad)
        conn_var = tk.StringVar(value=conn_default)
        conn_combo = ttk.Combobox(
            dialog,
            textvariable=conn_var,
            values=["SSL/TLS", "STARTTLS", "plain"],
            state="readonly",
            width=22,
        )
        conn_combo.grid(row=3, column=1, sticky="w", **pad)

        def _save() -> None:
            domain = domain_entry.get().strip().lstrip("@").lower()
            host = host_entry.get().strip()
            try:
                port = int(port_entry.get().strip())
            except ValueError:
                messagebox.showerror("Ошибка", "Порт должен быть числом", parent=dialog)
                return
            if not domain or not host:
                messagebox.showerror("Ошибка", "Заполните домен и хост", parent=dialog)
                return
            conn = conn_var.get()
            new_use_ssl = conn == "SSL/TLS"
            new_use_tls = conn == "STARTTLS"
            # Если домен переименован — удалить старый ключ
            if domain_key and domain_key != domain:
                domains.pop(domain_key, None)
            domains[domain] = {"host": host, "port": port, "use_tls": new_use_tls, "use_ssl": new_use_ssl}
            on_save()
            dialog.destroy()

        tk.Button(dialog, text="💾  Сохранить", command=_save, padx=10, pady=6, bg="#2a7d2a", fg="white").grid(
            row=4, column=0, columnspan=2, pady=16,
        )

    def _update_random_attachment_count(self) -> None:
        folder = self.attachments_folder_var.get().strip()
        if not folder:
            count = 0
        else:
            folder_path = (self.base_dir / folder).resolve()
            if folder_path.exists() and folder_path.is_dir():
                allowed_ext = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg"}
                count = len([p for p in folder_path.iterdir() if p.is_file() and p.suffix.lower() in allowed_ext])
            else:
                count = 0
        self.attachments_info_label.configure(text=f"Файлов: {count}")

    def _refresh_templates(self) -> None:
        template_dir = self.base_dir / self.templates_var.get()
        renderer = TemplateRenderer(template_dir)
        templates = renderer.list_templates()
        if templates:
            self.template_combo.configure(values=templates)
            if self.template_var.get() not in templates:
                self.template_var.set(templates[0])
        else:
            self.template_combo.configure(values=[""])
            self.template_var.set("")

    def _pick_random_replyto(self) -> str:
        emails = [e.strip() for e in self.replyto_combo.cget("values") if str(e).strip()]
        selected = self.replyto_var.get().strip()

        if self.replyto_mode_var.get() == "fixed":
            if selected:
                return selected
            if emails:
                self.replyto_var.set(emails[0])
                return emails[0]

        if emails:
            chosen = random.choice(emails)
            self.replyto_var.set(chosen)
            return chosen

        try:
            config = load_config(self.base_dir / self.config_var.get())
            if config.message.reply_to:
                return config.message.reply_to
        except ConfigError:
            pass

        return ""

    def _start_send(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Email App Modern", "Отправка уже выполняется")
            return

        try:
            delay = float(self.delay_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Email App Modern", "Задержка должна быть числом")
            return

        self._update_random_attachment_count()

        try:
            preflight = run_preflight(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                template_override=self.template_var.get() or None,
                body_text_override=None,
            )
        except CampaignError as error:
            messagebox.showerror("Email App Modern", str(error))
            return

        if preflight.errors:
            messagebox.showerror("Email App Modern", "\n".join(preflight.errors))
            return

        if preflight.warnings:
            proceed = messagebox.askyesno(
                "Email App Modern",
                "⚠️ Найдены предупреждения preflight:\n\n"
                + "\n".join(f"- {item}" for item in preflight.warnings)
                + "\n\nПродолжить?",
            )
            if not proceed:
                self.status_var.set("⊘ Отправка отменена")
                return

        if not self.dry_run_var.get():
            proceed = messagebox.askyesno("Email App Modern", "🔥 Запустить РЕАЛЬНУЮ отправку?")
            if not proceed:
                self.status_var.set("⊘ Отправка отменена")
                return

        # Рандомный выбор Reply-To
        reply_to_random = self._pick_random_replyto()
        if not reply_to_random:
            messagebox.showerror("Reply-To", "Не выбран ни один email для Reply-To!")
            return
        self.replyto_var.set(reply_to_random)

        mode_label = "фиксированный" if self.replyto_mode_var.get() == "fixed" else "случайный"
        # Сброс состояния прогресса
        self._stop_event.clear()
        self._campaign_sent = 0
        self._campaign_failed = 0
        self._campaign_total = 0
        self._campaign_start_time = None
        self._campaign_failed_recipients = []
        self.errors_btn.configure(state="disabled", fg_color=("gray75", "gray35"))
        self.progress_bar.set(0)
        self.progress_pct_label.configure(text="0%  (0/0)")
        self.sent_label.configure(text="✅ Отправлено: 0")
        self.failed_label.configure(text="❌ Ошибок: 0")
        self.remaining_label.configure(text="⏳ Осталось: —")
        self.eta_label.configure(text="⏱ ETA: —")
        self.status_var.set("⏳ Запуск...")
        self._append_log(f"▶ Старт задачи (Reply-To: {reply_to_random}, режим: {mode_label})")
        self.worker = threading.Thread(target=self._run_worker, args=(delay,), daemon=True)
        self.worker.start()
    def _run_queue_dialog(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Email App Modern", "Задача уже выполняется")
            return

        path = filedialog.askopenfilename(
            initialdir=self.base_dir / "presets",
            filetypes=[("JSON/CSV", "*.json *.csv"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        self.status_var.set("Запуск очереди...")
        self._append_log(f"Старт очереди: {path}")
        self.worker = threading.Thread(target=self._run_queue_worker, args=(Path(path),), daemon=True)
        self.worker.start()

    def _export_queue_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.base_dir / "presets",
            defaultextension=".json",
            filetypes=[("JSON", "*.json"), ("CSV", "*.csv"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        exported_path = save_campaign_queue(
            path,
            [
                CampaignPreset(
                    config=self._portable_path_value(self.config_var.get()),
                    recipients=self._portable_path_value(self.recipients_var.get()),
                    templates=self._portable_path_value(self.templates_var.get()),
                    template=self.template_var.get() or None,
                    delay_seconds=float(self.delay_var.get().strip() or "0"),
                    dry_run=self.dry_run_var.get(),
                    attachments_folder=self._portable_path_value(self.attachments_folder_var.get()) or None,
                    use_proxy=self.proxy_enabled_var.get(),
                    proxy_file=self._portable_path_value(self.proxy_file_var.get()) or None,
                    rate_limit_per_minute=(int(self.rate_limit_var.get().strip()) if self.rate_limit_var.get().strip() else None),
                    retry_attempts=(int(self.retry_attempts_var.get().strip()) if self.retry_attempts_var.get().strip() else None),
                    retry_backoff_seconds=(float(self.retry_backoff_var.get().strip()) if self.retry_backoff_var.get().strip() else None),
                    parallel_smtp_enabled=self.parallel_enabled_var.get(),
                    parallel_smtp_accounts=(int(self.parallel_smtp_var.get().strip()) if self.parallel_smtp_var.get().strip() else None),
                    batch_interval_seconds=(float(self.batch_interval_var.get().strip()) if self.batch_interval_var.get().strip() else None),
                    reply_to=self.replyto_var.get().strip() or None,
                    reply_to_mode=self.replyto_mode_var.get().strip() or None,
                )
            ],
        )
        self._append_log(f"Очередь экспортирована: {exported_path}")
        self.status_var.set("Очередь экспортирована")

    def _run_worker(self, delay: float) -> None:
        try:
            # Load config to get and apply rate_limit and retry settings
            config = load_config(self.base_dir / self.config_var.get())
            
            # Override rate limit if set in GUI
            if self.rate_limit_var.get().strip():
                try:
                    config.delivery.rate_limit_per_minute = int(self.rate_limit_var.get().strip())
                except ValueError:
                    pass
            
            # Override retry settings if set in GUI
            try:
                config.delivery.retry_attempts = int(self.retry_attempts_var.get().strip() or "1")
                config.delivery.retry_backoff_seconds = float(self.retry_backoff_var.get().strip() or "5")
            except ValueError:
                pass

            # Для прокси минимум 3 попытки (иначе слишком часто attempts=1 и мгновенный фейл)
            if self.proxy_enabled_var.get() and config.delivery.retry_attempts < 3:
                config.delivery.retry_attempts = 3
                self.queue.put(("log", "ℹ️ Retry увеличен до 3 для стабильной отправки через прокси"))

            # Override concurrency settings if set in GUI
            try:
                config.delivery.parallel_smtp_enabled = self.parallel_enabled_var.get()
                config.delivery.parallel_smtp_accounts = max(1, int(self.parallel_smtp_var.get().strip() or "1"))
                config.delivery.batch_interval_seconds = float(self.batch_interval_var.get().strip() or "0")
            except ValueError:
                pass
            
            # Передать reply_to в message
            config.message.reply_to = self.replyto_var.get()

            summary = run_campaign(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                dry_run=self.dry_run_var.get(),
                template_override=self.template_var.get() or None,
                subject_override=None,
                subject_mode=("random_recipient" if self.subjects_mode_var.get() == "random" else "random_campaign") if self._subjects_list else "fixed",
                subject_variants=list(self._subjects_list),
                body_text_override=None,
                body_text_mode="fixed",
                body_text_variants=[],
                random_attachments_folder_override=self.attachments_folder_var.get() or None,
                use_proxy=self.proxy_enabled_var.get(),
                proxy_file_override=self.proxy_file_var.get().strip() or None,
                delay_override=delay,
                rate_limit_per_minute=config.delivery.rate_limit_per_minute,
                retry_attempts=config.delivery.retry_attempts,
                retry_backoff_seconds=config.delivery.retry_backoff_seconds,
                parallel_smtp_enabled=self.parallel_enabled_var.get(),
                parallel_smtp_accounts=config.delivery.parallel_smtp_accounts,
                batch_interval_seconds=config.delivery.batch_interval_seconds,
                reply_to_override=self.replyto_var.get().strip() or None,
                reply_to_mode_override=self.replyto_mode_var.get().strip() or None,
                stop_event=self._stop_event,
                progress_callback=self._make_progress_callback(),
            )
            self.queue.put(("log", f"История CSV: {summary.history_csv}"))
            self.queue.put(("log", f"История JSONL: {summary.history_jsonl}"))
            self.queue.put(
                (
                    "done",
                    f"✓ Готово. Обработано: {summary.processed}/{summary.total}, успешно: {summary.successful}, ошибок: {summary.failed}",
                )
            )
        except CampaignError as error:
            self.queue.put(("error", str(error)))
        except Exception as error:  # noqa: BLE001
            self.queue.put(("error", f"Неожиданная ошибка: {error}"))

    def _run_queue_worker(self, queue_path: Path) -> None:
        try:
            summary = run_campaign_queue(
                base_dir=self.base_dir,
                campaigns=load_campaign_queue(queue_path),
                progress_callback=lambda message: self.queue.put(("log", message)),
            )
            self.queue.put(
                (
                    "done",
                    f"Очередь завершена. Кампаний: {summary.campaigns_completed}/{summary.campaigns_total}, обработано: {summary.total_processed}, успешно: {summary.total_successful}, ошибок: {summary.total_failed}",
                )
            )
        except (CampaignQueueError, CampaignError) as error:
            self.queue.put(("error", str(error)))
        except Exception as error:  # noqa: BLE001
            self.queue.put(("error", f"Неожиданная ошибка: {error}"))

    def _refresh_stats_window(self) -> None:
        records = filter_history_records(
            self.stats_all_records,
            status=self.stats_status_var.get(),
            template_query=self.stats_template_var.get(),
            smtp_query=self.stats_smtp_var.get(),
        )
        stats = summarize_history_records(records)
        lines = [
            f"Всего записей: {stats.total}",
            f"Успешно отправлено: {stats.sent}",
            f"Dry-run: {stats.dry_run}",
            f"Ошибок: {stats.errors}",
            f"Уникальных получателей: {stats.unique_recipients}",
            "",
            "Топ шаблонов:",
            *[f"- {name}: {count}" for name, count in stats.top_templates],
            "",
            "Топ SMTP-аккаунтов:",
            *[f"- {name}: {count}" for name, count in stats.top_smtp_accounts],
        ]
        self.stats_text.delete("1.0", "end")
        self.stats_text.insert("1.0", "\n".join(lines))
        row_lines = [
            f"{record.timestamp} | {record.status} | {record.recipient} | {record.template} | {record.smtp_account} | {record.error}"
            for record in records
        ]
        self.stats_records_text.delete("1.0", "end")
        self.stats_records_text.insert("1.0", "\n".join(row_lines) or "Нет записей")
        self.stats_filtered_records = records

    def _export_stats(self, extension: str) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.base_dir / "history",
            defaultextension=extension,
            filetypes=[("CSV", "*.csv"), ("JSON", "*.json"), ("Все файлы", "*.*")],
        )
        if not path:
            return
        export_path = export_history_records(path, self.stats_filtered_records)
        self._append_log(f"Статистика экспортирована: {export_path}")
        self.status_var.set("Статистика экспортирована")

    def _show_preview_window(self, html_content: str, preview_path: Path, mobile: bool = False) -> None:
        import webbrowser
        if self.preview_window is None or not self.preview_window.winfo_exists():
            self.preview_window = self.ctk.CTkToplevel(self.root)
            self.preview_window.title("HTML Preview")
            self.preview_window.geometry("420x800" if mobile else "1180x800")

            toolbar = self.ctk.CTkFrame(self.preview_window)
            toolbar.pack(fill="x", padx=12, pady=12)
            self.ctk.CTkButton(
                toolbar,
                text="Открыть предпросмотр в браузере",
                command=lambda: webbrowser.open(preview_path.as_uri()),
            ).pack(side="left")
            self.ctk.CTkButton(
                toolbar,
                text="Показать CSS",
                command=lambda: self._show_css_window(html_content),
            ).pack(side="left", padx=(6, 0))
            self.ctk.CTkLabel(toolbar, text=str(preview_path)).pack(side="left", padx=12)

            content = self.ctk.CTkFrame(self.preview_window)
            content.pack(fill="both", expand=True, padx=12, pady=(0, 12))
            content.grid_columnconfigure(0, weight=2)
            content.grid_columnconfigure(1, weight=1)
            content.grid_rowconfigure(0, weight=1)

            left = self.ctk.CTkFrame(content)
            right = self.ctk.CTkFrame(content)
            left.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
            right.grid(row=0, column=1, sticky="nsew", padx=(6, 0))

            try:
                html_module = importlib.import_module("tkhtmlview")
                self.preview_html_widget = html_module.HTMLScrolledText(left, html="")
                self.preview_html_widget.pack(fill="both", expand=True)
            except ImportError:
                self.preview_html_widget = self.ctk.CTkTextbox(left, wrap="word")
                self.preview_html_widget.pack(fill="both", expand=True)

            self.preview_source_text = self.ctk.CTkTextbox(right, wrap="none")
            self.preview_source_text.pack(fill="both", expand=True)

        if hasattr(self.preview_html_widget, "set_html"):
            self.preview_html_widget.set_html(html_content)
        else:
            self.preview_html_widget.delete("1.0", "end")
            self.preview_html_widget.insert("1.0", html_content)
        self.preview_source_text.delete("1.0", "end")
        self.preview_source_text.insert("1.0", html_content)

    def _show_css_window(self, html_content: str) -> None:
        import re
        styles = re.findall(r"<style[^>]*>(.*?)</style>", html_content, flags=re.DOTALL | re.IGNORECASE)
        if not styles:
            messagebox.showinfo("CSS", "В шаблоне не найдено блоков <style>")
            return

        css_window = self.ctk.CTkToplevel(self.root)
        css_window.title("CSS из шаблона")
        css_window.geometry("800x500")
        css_box = self.ctk.CTkTextbox(css_window, wrap="none")
        css_box.pack(fill="both", expand=True, padx=12, pady=12)
        css_box.insert("1.0", "\n\n---\n\n".join(styles))

    def _preview_email(self) -> None:
        import webbrowser
        try:
            reply_to_random = self._pick_random_replyto()
            if reply_to_random:
                self.replyto_var.set(reply_to_random)
            summary = render_preview(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                template_override=self.template_var.get() or None,
                body_text_override=self._pick_message_text(),
                recipient_email=self.preview_recipient_var.get() or None,
                open_in_browser=True,
            )
            message = (
                f"Предпросмотр сохранён: {summary.preview_path}\n"
                f"Шаблон: {summary.template_name}\n"
                f"Получатель: {summary.recipient_email}\n"
                f"Reply-To: {reply_to_random or '(из конфига)'}"
            )
            self._append_log(message)
            self.status_var.set("Предпросмотр открыт в браузере")
        except CampaignError as error:
            # Логируем список шаблонов и выбранный шаблон для диагностики
            template_dir = self.base_dir / self.templates_var.get()
            renderer = TemplateRenderer(template_dir)
            templates = renderer.list_templates()
            selected_template = self.template_var.get()
            log_message = (
                f"Ошибка предпросмотра: {error}\n"
                f"Доступные шаблоны: {templates}\n"
                f"Выбранный шаблон: '{selected_template}'\n"
                f"Папка шаблонов: {template_dir}"
            )
            self._append_log(log_message)
            messagebox.showerror("Email App Modern", log_message)

    def _show_stats(self) -> None:
        try:
            history_csv_path = self._current_history_csv_path()
            records = load_history_records(history_csv_path)
        except (StatsError, CampaignError) as error:
            self._append_log(f"Ошибка статистики: {error}")
            messagebox.showerror("Email App Modern", str(error))
            return

        if self.stats_window is None or not self.stats_window.winfo_exists():
            self.stats_window = self.ctk.CTkToplevel(self.root)
            self.stats_window.title("Статистика отправок")
            self.stats_window.geometry("960x680")

            filters = self.ctk.CTkFrame(self.stats_window)
            filters.pack(fill="x", padx=12, pady=12)
            self.stats_status_var = self.ctk.StringVar(value="all")
            self.stats_template_var = self.ctk.StringVar(value="")
            self.stats_smtp_var = self.ctk.StringVar(value="")
            self.ctk.CTkComboBox(filters, variable=self.stats_status_var, values=["all", "sent", "dry-run", "error"]).pack(side="left", padx=4)
            self.ctk.CTkEntry(filters, textvariable=self.stats_template_var, width=180, placeholder_text="Фильтр шаблона").pack(side="left", padx=4)
            self.ctk.CTkEntry(filters, textvariable=self.stats_smtp_var, width=180, placeholder_text="Фильтр SMTP").pack(side="left", padx=4)
            self.ctk.CTkButton(filters, text="Применить", command=self._refresh_stats_window).pack(side="left", padx=4)
            self.ctk.CTkButton(filters, text="Экспорт CSV", command=lambda: self._export_stats(".csv")).pack(side="left", padx=4)
            self.ctk.CTkButton(filters, text="Экспорт JSON", command=lambda: self._export_stats(".json")).pack(side="left", padx=4)

            self.stats_text = self.ctk.CTkTextbox(self.stats_window, wrap="word", height=180)
            self.stats_text.pack(fill="x", padx=12, pady=(0, 12))
            self.stats_records_text = self.ctk.CTkTextbox(self.stats_window, wrap="none")
            self.stats_records_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        else:
            self.stats_window.lift()
        self.stats_all_records = records
        self._refresh_stats_window()
        self._append_log(f"Статистика загружена: {history_csv_path}")
        self.status_var.set("Статистика загружена")

    def _current_template_path(self) -> Path:
        template_name = self.template_var.get().strip()
        if not template_name:
            raise CampaignError("Не выбран HTML-шаблон")
        return self.base_dir / self.templates_var.get() / template_name

    def _current_history_csv_path(self) -> Path:
        try:
            config = load_config(self.base_dir / self.config_var.get())
        except ConfigError as error:
            raise CampaignError(str(error)) from error
        return self.base_dir / config.delivery.history_csv

    def _open_template_editor(self) -> None:
        try:
            template_path = self._current_template_path()
        except CampaignError as error:
            messagebox.showerror("Email App Modern", str(error))
            return

        if self.editor_window is None or not self.editor_window.winfo_exists():
            self.editor_window = self.ctk.CTkToplevel(self.root)
            self.editor_window.title("Редактор HTML-шаблона")
            self.editor_window.geometry("1000x760")

            toolbar = self.ctk.CTkFrame(self.editor_window)
            toolbar.pack(fill="x", padx=12, pady=12)
            self.ctk.CTkButton(toolbar, text="Перезагрузить", command=self._load_template_into_editor).pack(side="left")
            self.ctk.CTkButton(toolbar, text="Сохранить", command=self._save_template_from_editor).pack(side="left", padx=8)
            self.ctk.CTkButton(toolbar, text="Сохранить и показать", command=self._save_and_preview_from_editor).pack(side="left", padx=8)
            self.ctk.CTkButton(toolbar, text="Выбрать редактор", command=self._select_external_editor).pack(side="left", padx=8)
            self.ctk.CTkButton(toolbar, text="Открыть в редакторе ОС", command=self._open_template_in_external).pack(side="left", padx=8)

            self.editor_text = self.ctk.CTkTextbox(self.editor_window, wrap="none")
            self.editor_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
            self.editor_text.bind("<KeyRelease>", self._on_editor_key_release)
            self.editor_window.bind_all("<Control-s>", lambda event: self._save_and_preview_from_editor())
        else:
            self.editor_window.lift()

        self._load_template_into_editor()
        self._append_log(f"Шаблон открыт в редакторе: {template_path}")

    def _select_external_editor(self) -> None:
        editor_path = filedialog.askopenfilename(
            title="Выберите исполняемый файл внешнего редактора",
            filetypes=[("Executable", "*.exe *.app *")],
        )
        if not editor_path:
            return
        self.external_editor_path.set(editor_path)
        self._append_log(f"Выбран внешний редактор: {editor_path}")
        messagebox.showinfo("Email App Modern", f"Выбран внешний редактор: {editor_path}")

    def _open_template_in_external(self) -> None:
        try:
            template_path = self._current_template_path()
        except CampaignError as error:
            messagebox.showerror("Email App Modern", str(error))
            return

        editor_path = self.external_editor_path.get().strip()
        try:
            if editor_path:
                subprocess.Popen([editor_path, str(template_path)])
            else:
                if sys.platform.startswith("darwin"):
                    subprocess.run(["open", str(template_path)])
                elif sys.platform.startswith("win"):
                    os.startfile(str(template_path))
                else:
                    subprocess.run(["xdg-open", str(template_path)])
            self._append_log(f"Шаблон открыт во внешнем редакторе: {template_path}")
        except Exception as error:
            messagebox.showerror("Email App Modern", f"Не удалось открыть файл в системе: {error}")
            self._append_log(f"Ошибка внешнего открытия шаблона: {error}")

    def _open_visual_template_editor(self) -> None:
        try:
            template_path = self._current_template_path()
            editor_url = self.rich_editor_server.open_template(template_path)
        except CampaignError as error:
            messagebox.showerror("Email App Modern", str(error))
            return
        except RichEditorError as error:
            messagebox.showerror("Email App Modern", str(error))
            return
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Email App Modern", f"Не удалось открыть визуальный редактор: {error}")
            return

        self._append_log(f"Визуальный редактор открыт: {template_path} | {editor_url}")
        self.status_var.set("Визуальный редактор открыт")
        messagebox.showinfo(
            "Email App Modern",
            "Открыт визуальный редактор в браузере. После сохранения шаблон сразу обновится в проекте.",
        )

    def _load_template_into_editor(self) -> None:
        if self.editor_text is None:
            return
        template_path = self._current_template_path()
        if not template_path.exists():
            messagebox.showerror("Email App Modern", f"Шаблон не найден: {template_path}")
            return
        self.editor_text.delete("1.0", "end")
        self.editor_text.insert("1.0", template_path.read_text(encoding="utf-8"))
        self.status_var.set("Шаблон открыт")
        self._check_template_variables()
        self._load_recipients_for_preview()

    def _save_template_from_editor(self, suppress_message: bool = False) -> None:
        if self.editor_text is None:
            return
        template_path = self._current_template_path()
        template_path.parent.mkdir(parents=True, exist_ok=True)
        template_path.write_text(self.editor_text.get("1.0", "end").rstrip() + "\n", encoding="utf-8")
        self._append_log(f"Шаблон сохранён: {template_path}")
        self.status_var.set("Шаблон сохранён")
        if not suppress_message:
            messagebox.showinfo("Email App Modern", f"Шаблон сохранён: {template_path}")

    def _save_preset_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.base_dir / "presets",
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml *.yml"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        try:
            delay_seconds = float(self.delay_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Email App Modern", "Задержка должна быть числом")
            return

        preset = CampaignPreset(
            config=self._portable_path_value(self.config_var.get()),
            recipients=self._portable_path_value(self.recipients_var.get()),
            templates=self._portable_path_value(self.templates_var.get()),
            template=self.template_var.get() or None,
            delay_seconds=delay_seconds,
            dry_run=self.dry_run_var.get(),
            attachments_folder=self._portable_path_value(self.attachments_folder_var.get()) or None,
            use_proxy=self.proxy_enabled_var.get(),
            proxy_file=self._portable_path_value(self.proxy_file_var.get()) or None,
            rate_limit_per_minute=(int(self.rate_limit_var.get().strip()) if self.rate_limit_var.get().strip() else None),
            retry_attempts=(int(self.retry_attempts_var.get().strip()) if self.retry_attempts_var.get().strip() else None),
            retry_backoff_seconds=(float(self.retry_backoff_var.get().strip()) if self.retry_backoff_var.get().strip() else None),
            parallel_smtp_enabled=self.parallel_enabled_var.get(),
            parallel_smtp_accounts=(int(self.parallel_smtp_var.get().strip()) if self.parallel_smtp_var.get().strip() else None),
            batch_interval_seconds=(float(self.batch_interval_var.get().strip()) if self.batch_interval_var.get().strip() else None),
            reply_to=self.replyto_var.get().strip() or None,
            reply_to_mode=self.replyto_mode_var.get().strip() or None,
        )
        saved_path = save_preset(path, preset)
        self.current_preset_path = saved_path
        self._append_log(f"Пресет сохранён: {saved_path}")
        self.status_var.set("Пресет сохранён")

    def _load_preset_dialog(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir / "presets",
            filetypes=[("YAML", "*.yaml *.yml"), ("Все файлы", "*.*")],
        )
        if path:
            self._load_preset(Path(path))

    def _load_preset(self, path: Path, silent: bool = False) -> None:
        try:
            preset = load_preset(path)
        except PresetError as error:
            self._append_log(f"Ошибка пресета: {error}")
            if not silent:
                messagebox.showerror("Email App Modern", str(error))
            return

        self.current_preset_path = path
        self.config_var.set(self.DEFAULT_CONFIG_PATH)
        self.recipients_var.set(self._portable_path_value(preset.recipients))
        self.templates_var.set(self._portable_path_value(preset.templates))
        self.delay_var.set("" if preset.delay_seconds is None else str(preset.delay_seconds))
        self.dry_run_var.set(preset.dry_run)
        self.attachments_folder_var.set(self._portable_path_value(preset.attachments_folder or ""))
        self.proxy_enabled_var.set(preset.use_proxy)
        self.proxy_file_var.set(self._portable_path_value(preset.proxy_file or "config/proxies.txt"))
        self.rate_limit_var.set("" if preset.rate_limit_per_minute is None else str(preset.rate_limit_per_minute))
        self.retry_attempts_var.set("" if preset.retry_attempts is None else str(preset.retry_attempts))
        self.retry_backoff_var.set("" if preset.retry_backoff_seconds is None else str(preset.retry_backoff_seconds))
        self.parallel_enabled_var.set(bool(preset.parallel_smtp_enabled) if preset.parallel_smtp_enabled is not None else False)
        self.parallel_smtp_var.set("" if preset.parallel_smtp_accounts is None else str(preset.parallel_smtp_accounts))
        self.batch_interval_var.set("" if preset.batch_interval_seconds is None else str(preset.batch_interval_seconds))
        self.replyto_var.set(preset.reply_to or "")
        if preset.reply_to_mode in {"random", "fixed"}:
            self.replyto_mode_var.set(preset.reply_to_mode)
        self._refresh_templates()
        if preset.template:
            self.template_var.set(preset.template)
        self._append_log(f"Пресет загружен: {path} (конфиг зафиксирован: {self.DEFAULT_CONFIG_PATH})")
        self.status_var.set("Пресет загружен")
        if not silent:
            messagebox.showinfo("Email App Modern", f"Пресет загружен: {path}")

    def _poll_queue(self) -> None:
        while True:
            try:
                event, payload = self.queue.get_nowait()
            except queue.Empty:
                break

            if event == "log":
                self._append_log(payload)
                self.status_var.set(payload[:120])
            elif event == "progress":
                current, total, sent, failed = payload
                self._update_progress_ui(current, total, sent, failed)
            elif event == "done":
                self._append_log(payload)
                self.status_var.set(payload)
                self.progress_bar.set(1.0)
                self.progress_pct_label.configure(text="100%")
                if self._campaign_failed_recipients:
                    self.errors_btn.configure(state="normal", fg_color=("#c0392b", "#922b21"))
                messagebox.showinfo("✓ Результат", payload)
            elif event == "error":
                self._append_log(f"✗ Ошибка: {payload}")
                self.status_var.set("✗ Ошибка")
                messagebox.showerror("✗ Ошибка", payload)

        self.root.after(150, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_widget.insert("end", f"{message}\n")
        self.log_widget.see("end")

    def _clear_log(self) -> None:
        self.log_widget.delete("1.0", "end")

    def _relative(self, path: Path) -> str:
        try:
            return str(path.relative_to(self.base_dir))
        except ValueError:
            return str(path)

    def _portable_path_value(self, value: str) -> str:
        text = str(value).strip()
        if not text:
            return text
        candidate = Path(text)
        if not candidate.is_absolute():
            return text
        try:
            return str(candidate.resolve().relative_to(self.base_dir.resolve()))
        except ValueError:
            return text

    def _looks_like_smtp_domains_file(self, path: Path) -> bool:
        """Эвристика: отличаем файл smtp_domains.txt от файла SMTP-аккаунтов."""
        try:
            with path.open("r", encoding="utf-8-sig") as handle:
                for line in handle:
                    text = line.strip()
                    if not text or text.startswith("#"):
                        continue
                    # Типичный формат доменов: @domain;smtp.host:port:1:type
                    if text.startswith("@") and ";" in text and "|" not in text:
                        return True
                    # Первая полезная строка не похожа на domains-format
                    return False
        except Exception:
            return False
        return False

    def _on_theme_change(self, selected_theme: str | None = None) -> None:
        """Handle theme change from combo callback and apply safely."""
        theme = (selected_theme or self.theme_var.get() or "dark").strip().lower()
        if theme not in {"dark", "light", "system"}:
            theme = "dark"
        self.theme_var.set(theme)
        self.ctk.set_appearance_mode(theme)

    def _load_replyto_txt(self):
        file_path = filedialog.askopenfilename(
            title="Выберите txt-файл с email-адресами",
            filetypes=[("Text files", "*.txt")],
        )
        if not file_path:
            return
        emails = []
        with open(file_path, encoding="utf-8") as f:
            for line in f:
                email = line.strip()
                if email and "@" in email:
                    emails.append(email)
        if not emails:
            messagebox.showerror("Reply-To", "В файле нет ни одного email!")
            return
        self.replyto_combo.configure(values=emails)
        self.replyto_var.set(emails[0])
        self.replyto_count_var.set(f"Reply-To адресов: {len(emails)}")

    def _pick_random_replyto(self) -> str:
        current_values = list(self.replyto_combo.cget("values"))
        current_values = [e for e in current_values if e.strip()]

        if self.replyto_mode_var.get() == "fixed":
            selected = self.replyto_var.get().strip()
            if selected and selected in current_values:
                return selected
            if current_values:
                return current_values[0]

        if current_values:
            return random.choice(current_values)

        # fallback to config reply_to
        try:
            config = load_config(self.base_dir / self.config_var.get())
            if config.message.reply_to:
                return config.message.reply_to
        except ConfigError:
            pass

        return ""

    def _load_recipients_for_preview(self) -> None:
        try:
            recipients = load_recipients(self.base_dir / self.recipients_var.get())
        except RecipientsError:
            self.preview_recipient_combo.configure(values=[])
            self.preview_recipient_var.set("")
            return
        values = [recipient.email for recipient in recipients]
        self.preview_recipient_combo.configure(values=values)
        if values:
            self.preview_recipient_var.set(values[0])

    def _check_template_variables(self) -> None:
        try:
            template_dir = self.base_dir / self.templates_var.get()
            renderer = TemplateRenderer(template_dir)
            template_name = self.template_var.get().strip()
            if not template_name:
                messagebox.showinfo("Переменные шаблона", "Шаблон не выбран")
                return
            vars_set = renderer.extract_template_variables(template_name)
            self.template_vars_label.configure(text=f"Переменных в шаблоне: {len(vars_set)}")
            if vars_set:
                messagebox.showinfo("Переменные шаблона", "Найденные переменные:\n" + "\n".join(sorted(vars_set)))
            else:
                messagebox.showinfo("Переменные шаблона", "Переменные не найдены")
        except Exception as err:
            messagebox.showerror("Переменные шаблона", f"Ошибка при анализе шаблона: {err}")

    def _on_editor_key_release(self, event=None) -> None:
        if self.editor_live_job:
            self.editor_window.after_cancel(self.editor_live_job)
        if self.live_preview_var.get():
            self.editor_live_job = self.editor_window.after(700, self._save_and_preview_from_editor)

    def _save_and_preview_from_editor(self) -> None:
        self._save_template_from_editor(suppress_message=True)
        self._check_template_variables()
        self._preview_email()

    # ------------------------------------------------------------------ #
    # Прогресс и счётчики                                                 #
    # ------------------------------------------------------------------ #

    def _make_progress_callback(self):
        """Возвращает progress_callback с разбором структуры сообщений."""
        import re as _re
        import time as _time

        try:
            _recs = load_recipients(self.base_dir / self.recipients_var.get())
            self._campaign_total = len(_recs)
        except RecipientsError:
            self._campaign_total = 0
        self._campaign_start_time = _time.monotonic()

        def _cb(message: str) -> None:
            self.queue.put(("log", message))
            _m = _re.search(r'\[(?:OK|ERROR|DRY-RUN|SKIP)\] (\d+)/(\d+)', message)
            if _m:
                _current = int(_m.group(1))
                _total = int(_m.group(2))
                if "[OK]" in message or "[DRY-RUN]" in message:
                    self._campaign_sent += 1
                elif "[ERROR]" in message:
                    self._campaign_failed += 1
                    # Извлекаем email и причину ошибки из формата: [ERROR] N/M email: REASON (subject=...)
                    _base = message.split(" (subject=", 1)[0]
                    _m_err = _re.search(r'\[ERROR\] \d+/\d+ (\S+): (.+)', _base)
                    if _m_err:
                        self._campaign_failed_recipients.append({
                            "email": _m_err.group(1),
                            "reason": _m_err.group(2).strip(),
                        })
                self.queue.put(("progress", (_current, _total, self._campaign_sent, self._campaign_failed)))

        return _cb

    def _update_progress_ui(self, current: int, total: int, sent: int, failed: int) -> None:
        import time as _time
        if total > 0:
            frac = min(current / total, 1.0)
            self.progress_bar.set(frac)
            self.progress_pct_label.configure(text=f"{int(frac * 100)}%  ({current}/{total})")
        self.sent_label.configure(text=f"✅ Отправлено: {sent}")
        self.failed_label.configure(text=f"❌ Ошибок: {failed}")
        remaining = max(0, total - current)
        self.remaining_label.configure(text=f"⏳ Осталось: {remaining}")
        if self._campaign_start_time is not None and current > 0:
            elapsed = _time.monotonic() - self._campaign_start_time
            rate = current / elapsed
            if rate > 0 and remaining > 0:
                eta_secs = int(remaining / rate)
                m, s = divmod(eta_secs, 60)
                self.eta_label.configure(text=f"⏱ ETA: ~{m}м {s:02d}с")
            else:
                self.eta_label.configure(text="⏱ ETA: —")

    def _show_failed_window(self) -> None:
        """Показывает окно со списком адресов, на которые не удалось отправить письма."""
        failed = self._campaign_failed_recipients
        if not failed:
            messagebox.showinfo("Ошибки отправки", "Список ошибок пуст.")
            return

        win = self.ctk.CTkToplevel(self.root)
        win.title(f"❌ Ошибки отправки ({len(failed)})")
        win.geometry("700x480")
        win.grab_set()

        self.ctk.CTkLabel(
            win,
            text=f"Не удалось отправить: {len(failed)} адрес(а/ов)",
            font=("", 13, "bold"),
            anchor="w",
        ).pack(anchor="w", padx=16, pady=(14, 4))

        # Текстовый виджет со списком
        text = self.ctk.CTkTextbox(win, wrap="word", font=("Courier", 11))
        text.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        lines = []
        for i, item in enumerate(failed, start=1):
            lines.append(f"{i:>3}. {item['email']}")
            lines.append(f"       └ {item['reason']}")
            lines.append("")

        text.insert("1.0", "\n".join(lines).rstrip())
        text.configure(state="disabled")

        def _copy_emails() -> None:
            emails = "\n".join(item["email"] for item in failed)
            win.clipboard_clear()
            win.clipboard_append(emails)
            messagebox.showinfo("Скопировано", f"Скопировано {len(failed)} адрес(а/ов) в буфер обмена.")

        def _copy_all() -> None:
            content = "\n".join(
                f"{item['email']}\t{item['reason']}" for item in failed
            )
            win.clipboard_clear()
            win.clipboard_append(content)
            messagebox.showinfo("Скопировано", "Скопированы адреса и причины ошибок.")

        btn_row = self.ctk.CTkFrame(win, fg_color="transparent")
        btn_row.pack(fill="x", padx=16, pady=(0, 14))
        self.ctk.CTkButton(btn_row, text="📋 Скопировать адреса", width=180, command=_copy_emails).pack(side="left", padx=(0, 8))
        self.ctk.CTkButton(btn_row, text="📋 Скопировать всё", width=160, fg_color=("gray70", "gray35"), command=_copy_all).pack(side="left")
        self.ctk.CTkButton(btn_row, text="✕ Закрыть", width=100, fg_color=("gray70", "gray35"), command=win.destroy).pack(side="right")

    def _update_recipients_count(self, path: Path) -> None:
        """Пересчитывает получателей и обновляет метку + ETA."""
        try:
            recs = load_recipients(path)
            count = len(recs)
            self._recipients_count = count
            self.recipients_count_label.configure(text=f"👥 Получателей: {count}")
        except RecipientsError:
            self.recipients_count_label.configure(text="⚠️ Файл не распознан")
            self._recipients_count = 0
        self._update_eta_estimate()

    def _update_eta_estimate(self) -> None:
        """Обновляет оценочное время завершения рядом с числом получателей."""
        try:
            delay = float(self.delay_var.get().strip() or "0")
        except ValueError:
            delay = 0.0
        count = self._recipients_count
        if count and delay:
            total_secs = count * delay
            m, s = divmod(int(total_secs), 60)
            h, m2 = divmod(m, 60)
            if h:
                eta = f"⏱ ~{h}ч {m2}м"
            else:
                eta = f"⏱ ~{m}м {s}с"
        elif count:
            eta = f"⏱ ~{count} писем без задержки"
        else:
            eta = ""
        self.eta_estimate_label.configure(text=eta)

    # ------------------------------------------------------------------ #
    # Stop / Test SMTP                                                     #
    # ------------------------------------------------------------------ #

    def _stop_campaign(self) -> None:
        """Запрашивает остановку текущей рассылки."""
        if self.worker and self.worker.is_alive():
            self._stop_event.set()
            self.status_var.set("⏹ Остановка...")
            self._append_log("⏹ Запрошена остановка рассылки")
        else:
            messagebox.showinfo("Стоп", "Рассылка не запущена")

    def _resolve_all_smtp_accounts_for_test(self) -> list[object]:
        """Возвращает список SMTP-аккаунтов для тестов (из accounts_file или smtp секции)."""
        from .config import _build_smtp_settings, _load_smtp_accounts_txt, _load_smtp_accounts
        from .smtp_domains import load_domains as _ld, get_smtp_defaults_for_email

        config_path = self.base_dir / self.config_var.get()
        raw = yaml.safe_load(config_path.read_text(encoding="utf-8")) or {}
        smtp_raw: dict = dict(raw.get("smtp") or {})

        acf_value = smtp_raw.get("accounts_file")
        if acf_value:
            acf_path = self.base_dir / str(acf_value)
            if acf_path.exists() and acf_path.suffix.lower() in (".txt", ".csv"):
                domains_db = _ld(self.base_dir)
                smtp_defaults = {
                    "host": smtp_raw.get("host"),
                    "port": smtp_raw.get("port"),
                    "from_name": smtp_raw.get("from_name"),
                    "use_tls": smtp_raw.get("use_tls"),
                    "use_ssl": smtp_raw.get("use_ssl"),
                    "timeout_seconds": smtp_raw.get("timeout_seconds"),
                }
                if acf_path.suffix.lower() == ".txt":
                    return _load_smtp_accounts_txt(acf_path, smtp_defaults, domains_db=domains_db, base_dir=self.base_dir)
                return _load_smtp_accounts(acf_path)

        mapping = {k: v for k, v in smtp_raw.items() if k not in ("accounts_file", "proxy_file")}
        if not mapping.get("host"):
            username = str(mapping.get("username", ""))
            if username:
                guessed = get_smtp_defaults_for_email(username, domains=_ld(self.base_dir))
                mapping.setdefault("host", guessed.get("host"))
                mapping.setdefault("port", guessed.get("port"))
                if mapping.get("use_tls") is None:
                    mapping["use_tls"] = guessed.get("use_tls", True)
                if mapping.get("use_ssl") is None:
                    mapping["use_ssl"] = guessed.get("use_ssl", False)
        return [_build_smtp_settings(mapping)]

    def _test_all_smtp_accounts(self) -> None:
        """Тестирует авторизацию всех SMTP-аккаунтов из пакета."""
        import smtplib
        import socket as _socket

        def _humanize_error_ru(error: Exception | str) -> str:
            raw = str(error)
            low = raw.lower()
            if "authentication" in low or "535" in low or "username and password not accepted" in low:
                return "Ошибка авторизации SMTP (неверный логин/пароль или app-password)"
            if "timed out" in low or "timeout" in low:
                return "Таймаут соединения"
            if "connection refused" in low:
                return "Соединение отклонено SMTP-сервером"
            if "name or service not known" in low or "nodename nor servname provided" in low:
                return "Ошибка DNS/имени SMTP-сервера"
            if "ssl" in low or "certificate" in low:
                return "Ошибка SSL/TLS"
            return raw

        try:
            accounts = self._resolve_all_smtp_accounts_for_test()
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Тест SMTP (все)", f"Не удалось загрузить аккаунты: {error}")
            return

        if not accounts:
            messagebox.showinfo("Тест SMTP (все)", "SMTP аккаунты не найдены")
            return

        win = self.ctk.CTkToplevel(self.root)
        win.title("📦 Тест SMTP всех аккаунтов")
        win.geometry("760x520")
        win.grab_set()

        self.ctk.CTkLabel(win, text=f"Аккаунтов для теста: {len(accounts)}", font=("", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
        result_text = self.ctk.CTkTextbox(win, wrap="word")
        result_text.pack(fill="both", expand=True, padx=12, pady=8)

        def append(line: str) -> None:
            result_text.insert("end", line + "\n")
            result_text.see("end")

        def run_all() -> None:
            ok = 0
            bad = 0
            result_text.delete("1.0", "end")
            for idx, smtp in enumerate(accounts, start=1):
                conn_type = "SSL" if smtp.use_ssl else "STARTTLS" if smtp.use_tls else "plain"
                append(f"[{idx}/{len(accounts)}] {smtp.username} -> {smtp.host}:{smtp.port} ({conn_type})")
                win.update()
                try:
                    if smtp.use_ssl:
                        conn = smtplib.SMTP_SSL(smtp.host, smtp.port, timeout=smtp.timeout_seconds)
                    else:
                        conn = smtplib.SMTP(smtp.host, smtp.port, timeout=smtp.timeout_seconds)
                        if smtp.use_tls:
                            conn.ehlo()
                            conn.starttls()
                            conn.ehlo()
                    conn.login(smtp.username, smtp.password)
                    conn.quit()
                    ok += 1
                    append("  ✅ OK")
                except smtplib.SMTPAuthenticationError as exc:
                    bad += 1
                    append(f"  ❌ AUTH: {_humanize_error_ru(exc)}")
                except _socket.timeout:
                    bad += 1
                    append("  ❌ TIMEOUT")
                except Exception as exc:  # noqa: BLE001
                    bad += 1
                    append(f"  ❌ ERROR: {_humanize_error_ru(exc)}")

            append("")
            append(f"Итог: ✅ {ok} | ❌ {bad}")

        btns = self.ctk.CTkFrame(win, fg_color="transparent")
        btns.pack(fill="x", padx=12, pady=(0, 10))
        self.ctk.CTkButton(btns, text="▶ Запустить тест", command=run_all).pack(side="left")
        self.ctk.CTkButton(btns, text="✕ Закрыть", command=win.destroy).pack(side="right")

    def _test_proxies(self) -> None:
        """Тест прокси: доступность proxy host:port и туннель до SMTP через прокси."""
        import socket
        try:
            import socks  # type: ignore
        except Exception:
            socks = None
        from .proxy_utils import load_proxies

        proxy_file = (self.base_dir / self.proxy_file_var.get().strip()).resolve()
        if not proxy_file.exists():
            messagebox.showerror("Тест прокси", f"Файл не найден: {proxy_file}")
            return

        try:
            proxies = load_proxies(proxy_file)
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Тест прокси", f"Ошибка чтения прокси: {error}")
            return

        if not proxies:
            messagebox.showinfo("Тест прокси", "В файле нет прокси")
            return

        smtp_target_host = None
        smtp_target_port = None
        smtp_probe_targets: list[tuple[str, int, str]] = [
            ("smtp.gmail.com", 465, "Gmail SSL"),
            ("smtp.gmail.com", 587, "Gmail STARTTLS"),
        ]
        try:
            accounts = self._resolve_all_smtp_accounts_for_test()
            if accounts:
                smtp_target_host = accounts[0].host
                smtp_target_port = int(accounts[0].port)
                known = {(host, port) for host, port, _label in smtp_probe_targets}
                for account in accounts:
                    h = str(account.host).strip()
                    p = int(account.port)
                    if h and p and (h, p) not in known:
                        smtp_probe_targets.append((h, p, "Из SMTP аккаунта"))
                        known.add((h, p))
        except Exception:
            smtp_target_host = None
            smtp_target_port = None

        win = self.ctk.CTkToplevel(self.root)
        win.title("🧪 Тест прокси")
        win.geometry("760x520")
        win.grab_set()

        self.ctk.CTkLabel(win, text=f"Прокси для теста: {len(proxies)}", font=("", 12, "bold")).pack(anchor="w", padx=12, pady=(12, 6))
        result_text = self.ctk.CTkTextbox(win, wrap="word")
        result_text.pack(fill="both", expand=True, padx=12, pady=8)

        def append(line: str) -> None:
            result_text.insert("end", line + "\n")
            result_text.see("end")

        def _humanize_error_ru(error: Exception | str) -> str:
            raw = str(error)
            low = raw.lower()
            if "network unreachable" in low:
                return "Сеть недоступна для SMTP через этот прокси"
            if "timed out" in low or "timeout" in low:
                return "Таймаут соединения"
            if "connection refused" in low:
                return "Соединение отклонено"
            if "name or service not known" in low or "nodename nor servname provided" in low:
                return "Ошибка DNS/имени хоста"
            if "authentication" in low and "proxy" in low:
                return "Ошибка авторизации прокси"
            return raw

        def run_test() -> None:
            ok = 0
            bad = 0
            smtp_ready = 0
            result_text.delete("1.0", "end")
            if smtp_target_host and smtp_target_port:
                append(f"Цель для туннеля: {smtp_target_host}:{smtp_target_port}")
            else:
                append("Цель для туннеля: не определена (нет SMTP в конфиге)")
            append("Проверка SMTP-портов через прокси: 465/587 + цели из SMTP-аккаунтов")
            append("")

            def _connect_via_proxy(proxy: dict, target_host: str, target_port: int, timeout: int = 7) -> tuple[bool, str]:
                p_host = str(proxy.get("proxy_host") or "")
                p_port = int(proxy.get("proxy_port") or 0)
                p_type = str(proxy.get("proxy_type") or "").lower()
                p_user = proxy.get("proxy_user")
                p_pass = proxy.get("proxy_pass")

                if socks is None:
                    return False, "PySocks не установлен"

                proxy_map = {
                    "socks5": socks.SOCKS5,
                    "socks4": socks.SOCKS4,
                    "http": socks.HTTP,
                    "https": socks.HTTP,
                }
                pconst = proxy_map.get(p_type)
                if pconst is None:
                    return False, f"неизвестный тип прокси: {p_type}"

                s = socks.socksocket()
                s.settimeout(timeout)
                try:
                    s.set_proxy(
                        proxy_type=pconst,
                        addr=p_host,
                        port=p_port,
                        rdns=True,
                        username=str(p_user) if p_user else None,
                        password=str(p_pass) if p_pass else None,
                    )
                    s.connect((target_host, target_port))
                    return True, "ok"
                except Exception as exc:  # noqa: BLE001
                    return False, _humanize_error_ru(exc)
                finally:
                    try:
                        s.close()
                    except Exception:
                        pass

            for idx, proxy in enumerate(proxies, start=1):
                host = proxy.get("proxy_host")
                port = int(proxy.get("proxy_port") or 0)
                ptype = proxy.get("proxy_type")
                append(f"[{idx}/{len(proxies)}] {ptype}:{host}:{port}")
                win.update()
                try:
                    with socket.create_connection((str(host), int(port)), timeout=5):
                        pass
                    append("  ✅ proxy reachable")
                except Exception as exc:  # noqa: BLE001
                    bad += 1
                    append(f"  ❌ proxy unreachable: {_humanize_error_ru(exc)}")
                    continue

                if not (smtp_target_host and smtp_target_port):
                    ok += 1
                    append("  ℹ️ туннель SMTP пропущен")
                    continue

                tunnel_ok, tunnel_msg = _connect_via_proxy(proxy, str(smtp_target_host), int(smtp_target_port))
                if tunnel_ok:
                    ok += 1
                    append(f"  ✅ SMTP tunnel OK -> {smtp_target_host}:{smtp_target_port}")
                else:
                    bad += 1
                    append(f"  ❌ SMTP tunnel failed: {tunnel_msg}")

                probe_465_ok = False
                probe_587_ok = False
                for target_host, target_port, label in smtp_probe_targets:
                    probe_ok, probe_msg = _connect_via_proxy(proxy, target_host, target_port, timeout=8)
                    if probe_ok:
                        append(f"  ✅ {label}: {target_host}:{target_port}")
                        if target_port == 465:
                            probe_465_ok = True
                        if target_port == 587:
                            probe_587_ok = True
                    else:
                        append(f"  ❌ {label}: {target_host}:{target_port} -> {probe_msg}")

                if probe_465_ok or probe_587_ok:
                    smtp_ready += 1
                    ports_ok = []
                    if probe_465_ok:
                        ports_ok.append("465")
                    if probe_587_ok:
                        ports_ok.append("587")
                    append(f"  🟢 SMTP-ready: да (порты: {', '.join(ports_ok)})")
                else:
                    append("  🔴 SMTP-ready: нет (465/587 недоступны через прокси)")

            append("")
            append(f"Итог: ✅ {ok} | ❌ {bad} | SMTP-ready: {smtp_ready}/{len(proxies)}")

        btns = self.ctk.CTkFrame(win, fg_color="transparent")
        btns.pack(fill="x", padx=12, pady=(0, 10))
        self.ctk.CTkButton(btns, text="▶ Запустить тест", command=run_test).pack(side="left")
        self.ctk.CTkButton(btns, text="✕ Закрыть", command=win.destroy).pack(side="right")

    # ------------------------------------------------------------------ #
    # Открытие файлов / папки                                             #
    # ------------------------------------------------------------------ #

    def _open_log_file(self) -> None:
        try:
            config = load_config(self.base_dir / self.config_var.get())
            log_path = (self.base_dir / config.delivery.log_file).resolve()
        except ConfigError:
            log_path = (self.base_dir / "logs" / "email_app.log").resolve()

        if not log_path.exists():
            messagebox.showinfo("Лог", f"Файл лога не найден:\n{log_path}")
            return
        try:
            if sys.platform.startswith("darwin"):
                subprocess.run(["open", str(log_path)])
            elif sys.platform.startswith("win"):
                os.startfile(str(log_path))  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(log_path)])
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Лог", f"Не удалось открыть файл: {error}")

    def _open_project_folder(self) -> None:
        try:
            if sys.platform.startswith("darwin"):
                subprocess.run(["open", str(self.base_dir)])
            elif sys.platform.startswith("win"):
                os.startfile(str(self.base_dir))  # type: ignore[attr-defined]
            else:
                subprocess.run(["xdg-open", str(self.base_dir)])
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Папка", f"Не удалось открыть папку: {error}")

    # ------------------------------------------------------------------ #
    # Сохранение / загрузка последней сессии                              #
    # ------------------------------------------------------------------ #

    _SESSION_FILE = "config/.last_session.json"

    def _save_last_session(self) -> None:
        import json
        data: dict = {
            "recipients": self.recipients_var.get(),
            "templates": self.templates_var.get(),
            "smtp_accounts_file": self.smtp_accounts_file_var.get(),
            "subjects_file": self.subjects_file_var.get(),
            "subjects_mode": self.subjects_mode_var.get(),
            "template": self.template_var.get(),
            "proxy_file": self.proxy_file_var.get(),
            "proxy_enabled": self.proxy_enabled_var.get(),
            "delay": self.delay_var.get(),
            "rate_limit": self.rate_limit_var.get(),
            "retry_attempts": self.retry_attempts_var.get(),
            "retry_backoff": self.retry_backoff_var.get(),
            "parallel_smtp": self.parallel_smtp_var.get(),
            "parallel_enabled": self.parallel_enabled_var.get(),
            "batch_interval": self.batch_interval_var.get(),
            "dry_run": self.dry_run_var.get(),
            "replyto": self.replyto_var.get(),
            "replyto_mode": self.replyto_mode_var.get(),
            "theme": self.theme_var.get(),
            "attachments_folder": self.attachments_folder_var.get(),
        }
        session_path = self.base_dir / self._SESSION_FILE
        try:
            session_path.parent.mkdir(parents=True, exist_ok=True)
            session_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:  # noqa: BLE001
            pass

    def _load_last_session(self) -> None:
        import json
        session_path = self.base_dir / self._SESSION_FILE
        if not session_path.exists():
            return
        try:
            data: dict = json.loads(session_path.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            return

        str_pairs = [
            ("recipients", self.recipients_var),
            ("templates", self.templates_var),
            ("smtp_accounts_file", self.smtp_accounts_file_var),
            ("subjects_file", self.subjects_file_var),
            ("subjects_mode", self.subjects_mode_var),
            ("template", self.template_var),
            ("proxy_file", self.proxy_file_var),
            ("delay", self.delay_var),
            ("rate_limit", self.rate_limit_var),
            ("retry_attempts", self.retry_attempts_var),
            ("retry_backoff", self.retry_backoff_var),
            ("parallel_smtp", self.parallel_smtp_var),
            ("batch_interval", self.batch_interval_var),
            ("replyto", self.replyto_var),
            ("replyto_mode", self.replyto_mode_var),
            ("theme", self.theme_var),
            ("attachments_folder", self.attachments_folder_var),
        ]
        for key, var in str_pairs:
            if key in data and data[key] is not None:
                var.set(str(data[key]))

        # Конфиг в упрощённом режиме всегда фиксированный
        self.config_var.set(self.DEFAULT_CONFIG_PATH)

        bool_pairs = [
            ("proxy_enabled", self.proxy_enabled_var),
            ("parallel_enabled", self.parallel_enabled_var),
            ("dry_run", self.dry_run_var),
        ]
        for key, var in bool_pairs:
            if key in data:
                var.set(bool(data[key]))

        if "theme" in data:
            self._on_theme_change(data["theme"])

        # Пересчитать получателей из сохранённого пути
        rcp_path = self.base_dir / self.recipients_var.get()
        if rcp_path.exists():
            self._update_recipients_count(rcp_path)

        # Перезагрузить темы писем из сохранённого пути
        subjects = data.get("subjects_file", "")
        if subjects:
            subjects_path = self.base_dir / subjects
            if subjects_path.exists():
                self._load_subjects_file(subjects_path)

    def _on_close(self) -> None:
        self._save_last_session()
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()


def launch_modern_gui(base_dir: Path, preset_path: Path | None = None) -> None:
    try:
        ctk = importlib.import_module("customtkinter")
    except ImportError:
        from .gui import launch_gui

        messagebox.showwarning(
            "Email App Modern",
            "Пакет customtkinter не установлен. Открывается обычный GUI.",
        )
        launch_gui(base_dir=base_dir, preset_path=preset_path)
        return

    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    app = ModernEmailAppGUI(ctk=ctk, base_dir=base_dir, preset_path=preset_path)
    app.run()
