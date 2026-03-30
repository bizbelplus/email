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


class ModernEmailAppGUI:
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

        self.root = ctk.CTk()
        self.root.title("Email App – Modern SMTP Campaign Manager")
        self.root.geometry("1200x920")
        self.root.minsize(1000, 750)

        # Theme preference
        self.theme_var = ctk.StringVar(value="dark")

        self.config_var = ctk.StringVar(value="config/settings.yaml")
        self.recipients_var = ctk.StringVar(value="recipients.csv")
        self.templates_var = ctk.StringVar(value="templates")
        self.template_var = ctk.StringVar(value="")
        self.attachments_folder_var = ctk.StringVar(value="")
        self.external_editor_path = ctk.StringVar(value="")
        self.proxy_enabled_var = ctk.BooleanVar(value=True)
        self.live_preview_var = ctk.BooleanVar(value=True)
        self.preview_recipient_var = ctk.StringVar(value="")
        self.mobile_preview_var = ctk.BooleanVar(value=False)
        self.replyto_mode_var = ctk.StringVar(value="random")
        self.replyto_count_var = ctk.StringVar(value="0")
        self.delay_var = ctk.StringVar(value="0")
        self.rate_limit_var = ctk.StringVar(value="")
        self.retry_attempts_var = ctk.StringVar(value="1")
        self.parallel_smtp_var = ctk.StringVar(value="1")
        self.parallel_enabled_var = ctk.BooleanVar(value=False)
        self.batch_interval_var = ctk.StringVar(value="0")
        self.retry_backoff_var = ctk.StringVar(value="5")
        self.dry_run_var = ctk.BooleanVar(value=True)
        self.status_var = ctk.StringVar(value="✓ Готово к запуску")
        self.editor_live_job = None

        self._build()
        if self.current_preset_path:
            self._load_preset(self.current_preset_path, silent=True)
        self._refresh_templates()
        self.root.after(150, self._poll_queue)

    def _build(self) -> None:
        main_container = self.ctk.CTkScrollableFrame(self.root, corner_radius=0)
        main_container.pack(fill="both", expand=True)
        main_container.grid_columnconfigure(0, weight=1)

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

        self._add_path_row(config_section, 0, "Конфиг", self.config_var, self._select_config)
        self._add_path_row(config_section, 1, "Получатели", self.recipients_var, self._select_recipients)
        self._add_path_row(config_section, 2, "Шаблоны", self.templates_var, self._select_templates)

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

        # === REPLY-TO (новый блок) ===
        replyto_section = self.ctk.CTkFrame(template_section, fg_color="transparent")
        replyto_section.pack(fill="x", padx=12, pady=8)
        self.ctk.CTkLabel(replyto_section, text="Reply-To (адрес для ответов):").grid(row=0, column=0, sticky="w")
        self.replyto_var = self.ctk.StringVar(value="")
        self.replyto_combo = self.ctk.CTkComboBox(replyto_section, variable=self.replyto_var, values=[], width=320)
        self.replyto_combo.grid(row=0, column=1, padx=(8, 8))
        self.ctk.CTkButton(replyto_section, text="Загрузить txt", command=self._load_replyto_txt).grid(row=0, column=2)
        self.ctk.CTkLabel(replyto_section, textvariable=self.replyto_count_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=(4,0))

        replyto_mode_section = self.ctk.CTkFrame(template_section, fg_color="transparent")
        replyto_mode_section.pack(fill="x", padx=12, pady=(0,8))
        self.ctk.CTkLabel(replyto_mode_section, text="Режим Reply-To:").grid(row=0, column=0, sticky="w")
        self.ctk.CTkRadioButton(replyto_mode_section, text="Случайный для каждого", variable=self.replyto_mode_var, value="random").grid(row=0, column=1, padx=4)
        self.ctk.CTkRadioButton(replyto_mode_section, text="Фиксированный выбранный", variable=self.replyto_mode_var, value="fixed").grid(row=0, column=2, padx=4)

        proxy_section = self.ctk.CTkFrame(template_section, fg_color="transparent")
        proxy_section.pack(fill="x", padx=12, pady=4)
        self.ctk.CTkCheckBox(proxy_section, text="Использовать прокси", variable=self.proxy_enabled_var, onvalue=True, offvalue=False).pack(anchor="w")

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
            ("✏️ Редактор", self._open_template_editor, 3),
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
        for i in range(3):
            buttons_frame2.grid_columnconfigure(i, weight=1)

        button_specs2 = [
            ("📂 Очередь JSON", self._run_queue_dialog, 0),
            ("💾 Экспорт JSON/CSV", self._export_queue_dialog, 1),
            ("📂 Загрузить пресет", self._load_preset_dialog, 2),
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

        # === СТАТУС (Раздел 4) ===
        status_section = self.ctk.CTkFrame(main_container, fg_color=("gray88", "gray25"), corner_radius=12)
        status_section.pack(fill="x", padx=16, pady=8)
        self.ctk.CTkLabel(
            status_section,
            textvariable=self.status_var,
            anchor="w",
            font=("", 11),
        ).pack(fill="x", padx=12, pady=12)

        # === ЛОГ (Раздел 5) ===
        log_section = self.ctk.CTkFrame(main_container, fg_color=("gray95", "gray20"), corner_radius=12)
        log_section.pack(fill="both", expand=True, padx=16, pady=(8, 16))
        self.ctk.CTkLabel(log_section, text="📝 Лог отправки", font=("", 14, "bold")).pack(anchor="w", padx=12, pady=(12, 8))
        
        log_buttons = self.ctk.CTkFrame(log_section, fg_color="transparent")
        log_buttons.pack(fill="x", padx=12, pady=(0, 8))
        self.ctk.CTkButton(log_buttons, text="🗑️  Очистить лог", width=120, command=self._clear_log).pack(side="left")

        self.log_widget = self.ctk.CTkTextbox(log_section, wrap="word", font=("Courier", 10))
        self.log_widget.pack(fill="both", expand=True, padx=12, pady=(0, 12))

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
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("YAML", "*.yaml *.yml"), ("Все файлы", "*.*")],
        )
        if path:
            self.config_var.set(self._relative(Path(path)))

    def _select_recipients(self) -> None:
        path = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("CSV", "*.csv"), ("Все файлы", "*.*")],
        )
        if path:
            self.recipients_var.set(self._relative(Path(path)))
            self._load_recipients_for_preview()

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
        # Получить список email из комбобокса (или файла)
        emails = self.replyto_combo.cget("values")
        import random
        emails = [e for e in emails if e.strip()]

        if emails:
            return random.choice(emails)

        # fallback: взять reply_to из конфига, если он задан
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

        self.status_var.set("⏳ Запуск...")
        self._append_log(f"▶ Старт задачи (Reply-To: {reply_to_random})")
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
                    config=self.config_var.get(),
                    recipients=self.recipients_var.get(),
                    templates=self.templates_var.get(),
                    template=self.template_var.get() or None,
                    delay_seconds=float(self.delay_var.get().strip() or "0"),
                    dry_run=self.dry_run_var.get(),
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
                random_attachments_folder_override=self.attachments_folder_var.get() or None,
                use_proxy=self.proxy_enabled_var.get(),
                delay_override=delay,
                rate_limit_per_minute=config.delivery.rate_limit_per_minute,
                retry_attempts=config.delivery.retry_attempts,
                retry_backoff_seconds=config.delivery.retry_backoff_seconds,
                parallel_smtp_enabled=self.parallel_enabled_var.get(),
                parallel_smtp_accounts=config.delivery.parallel_smtp_accounts,
                batch_interval_seconds=config.delivery.batch_interval_seconds,
                progress_callback=lambda message: self.queue.put(("log", message)),
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
                templates_path=self.templates_var.get(),
                template_override=self.template_var.get() or None,
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
            subprocess.Popen(
                [
                    sys.executable,
                    "-m",
                    "email_app.desktop_rich_editor",
                    "--template",
                    str(template_path),
                ],
                cwd=str(self.base_dir),
            )
        except CampaignError as error:
            messagebox.showerror("Email App Modern", str(error))
            return
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Email App Modern", f"Не удалось открыть desktop-редактор: {error}")
            return

        self._append_log(f"Desktop-редактор открыт: {template_path}")
        self.status_var.set("Desktop-редактор открыт")
        messagebox.showinfo(
            "Email App Modern",
            "Открыт desktop-редактор письма. После сохранения шаблон сразу обновится в проекте.",
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
            config=self.config_var.get(),
            recipients=self.recipients_var.get(),
            templates=self.templates_var.get(),
            template=self.template_var.get() or None,
            delay_seconds=delay_seconds,
            dry_run=self.dry_run_var.get(),
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
        self.config_var.set(preset.config)
        self.recipients_var.set(preset.recipients)
        self.templates_var.set(preset.templates)
        self.delay_var.set("" if preset.delay_seconds is None else str(preset.delay_seconds))
        self.dry_run_var.set(preset.dry_run)
        self._refresh_templates()
        if preset.template:
            self.template_var.set(preset.template)
        self._append_log(f"Пресет загружен: {path}")
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
                self.status_var.set(payload)
            elif event == "done":
                self._append_log(payload)
                self.status_var.set(payload)
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

    def _on_theme_change(self) -> None:
        """Handle theme change."""
        theme = self.theme_var.get()
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
