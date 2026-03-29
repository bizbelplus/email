from __future__ import annotations

import importlib
import queue
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
from .service import CampaignError, render_preview, run_campaign
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
        self.root.title("Email App Modern")
        self.root.geometry("1100x820")
        self.root.minsize(900, 680)

        self.config_var = ctk.StringVar(value="config/settings.yaml")
        self.recipients_var = ctk.StringVar(value="recipients.csv")
        self.templates_var = ctk.StringVar(value="templates")
        self.template_var = ctk.StringVar(value="")
        self.delay_var = ctk.StringVar(value="0")
        self.dry_run_var = ctk.BooleanVar(value=True)
        self.status_var = ctk.StringVar(value="Готово к запуску")

        self._build()
        if self.current_preset_path:
            self._load_preset(self.current_preset_path, silent=True)
        self._refresh_templates()
        self.root.after(150, self._poll_queue)

    def _build(self) -> None:
        container = self.ctk.CTkFrame(self.root, corner_radius=16)
        container.pack(fill="both", expand=True, padx=16, pady=16)
        container.grid_columnconfigure(1, weight=1)
        container.grid_rowconfigure(8, weight=1)

        self._add_path_row(container, 0, "Конфиг", self.config_var, self._select_config)
        self._add_path_row(container, 1, "Получатели", self.recipients_var, self._select_recipients)
        self._add_path_row(container, 2, "Шаблоны", self.templates_var, self._select_templates)

        self.ctk.CTkLabel(container, text="HTML-шаблон").grid(row=3, column=0, sticky="w", padx=12, pady=8)
        template_row = self.ctk.CTkFrame(container, fg_color="transparent")
        template_row.grid(row=3, column=1, columnspan=2, sticky="ew", padx=12, pady=8)
        template_row.grid_columnconfigure(0, weight=1)
        self.template_combo = self.ctk.CTkComboBox(template_row, variable=self.template_var, values=[])
        self.template_combo.grid(row=0, column=0, sticky="ew")
        self.ctk.CTkButton(template_row, text="Обновить", width=120, command=self._refresh_templates).grid(
            row=0,
            column=1,
            padx=(8, 0),
        )

        self.ctk.CTkLabel(container, text="Задержка, сек").grid(row=4, column=0, sticky="w", padx=12, pady=8)
        self.ctk.CTkEntry(container, textvariable=self.delay_var).grid(row=4, column=1, columnspan=2, sticky="ew", padx=12, pady=8)

        self.ctk.CTkCheckBox(
            container,
            text="Dry-run без реальной отправки",
            variable=self.dry_run_var,
            onvalue=True,
            offvalue=False,
        ).grid(row=5, column=0, columnspan=3, sticky="w", padx=12, pady=(8, 12))

        buttons = self.ctk.CTkFrame(container, fg_color="transparent")
        buttons.grid(row=6, column=0, columnspan=3, sticky="ew", padx=12, pady=8)
        for index in range(5):
            buttons.grid_columnconfigure(index, weight=1)
        button_specs = [
            ("Старт", self._start_send),
            ("Очередь JSON", self._run_queue_dialog),
            ("Экспорт JSON/CSV", self._export_queue_dialog),
            ("Предпросмотр", self._preview_email),
            ("Статистика", self._show_stats),
            ("Визуальный редактор", self._open_visual_template_editor),
            ("Редактор шаблона", self._open_template_editor),
            ("Сохранить пресет", self._save_preset_dialog),
            ("Загрузить пресет", self._load_preset_dialog),
        ]
        for index, (label, command) in enumerate(button_specs):
            row = index // 5
            column = index % 5
            self.ctk.CTkButton(buttons, text=label, command=command).grid(
                row=row,
                column=column,
                padx=4,
                pady=4,
                sticky="ew",
            )

        self.ctk.CTkLabel(container, textvariable=self.status_var, anchor="w").grid(
            row=7,
            column=0,
            columnspan=3,
            sticky="ew",
            padx=12,
            pady=8,
        )

        self.log_widget = self.ctk.CTkTextbox(container, wrap="word")
        self.log_widget.grid(row=8, column=0, columnspan=3, sticky="nsew", padx=12, pady=(0, 12))

    def _add_path_row(self, parent: object, row: int, label: str, variable: object, command: object) -> None:
        self.ctk.CTkLabel(parent, text=label).grid(row=row, column=0, sticky="w", padx=12, pady=8)
        self.ctk.CTkEntry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=12, pady=8)
        self.ctk.CTkButton(parent, text="Выбрать", width=120, command=command).grid(
            row=row,
            column=2,
            padx=12,
            pady=8,
        )

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

    def _select_templates(self) -> None:
        path = filedialog.askdirectory(initialdir=self.base_dir)
        if path:
            self.templates_var.set(self._relative(Path(path)))
            self._refresh_templates()

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

    def _start_send(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Email App Modern", "Отправка уже выполняется")
            return

        try:
            delay = float(self.delay_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Email App Modern", "Задержка должна быть числом")
            return

        self.status_var.set("Запуск...")
        self._append_log("Старт задачи")
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
            summary = run_campaign(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                dry_run=self.dry_run_var.get(),
                template_override=self.template_var.get() or None,
                delay_override=delay,
                progress_callback=lambda message: self.queue.put(("log", message)),
            )
            self.queue.put(("log", f"История CSV: {summary.history_csv}"))
            self.queue.put(("log", f"История JSONL: {summary.history_jsonl}"))
            self.queue.put(
                (
                    "done",
                    f"Готово. Обработано: {summary.processed}/{summary.total}, успешно: {summary.successful}, ошибок: {summary.failed}",
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

    def _show_preview_window(self, html_content: str, preview_path: Path) -> None:
        if self.preview_window is None or not self.preview_window.winfo_exists():
            self.preview_window = self.ctk.CTkToplevel(self.root)
            self.preview_window.title("HTML Preview")
            self.preview_window.geometry("1180x800")

            toolbar = self.ctk.CTkFrame(self.preview_window)
            toolbar.pack(fill="x", padx=12, pady=12)
            self.ctk.CTkButton(
                toolbar,
                text="Открыть в браузере",
                command=lambda: render_preview(
                    base_dir=self.base_dir,
                    config_path=self.base_dir / self.config_var.get(),
                    recipients_path=self.base_dir / self.recipients_var.get(),
                    templates_path=self.base_dir / self.templates_var.get(),
                    template_override=self.template_var.get() or None,
                    open_in_browser=True,
                ),
            ).pack(side="left")
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

    def _preview_email(self) -> None:
        try:
            summary = render_preview(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                template_override=self.template_var.get() or None,
                open_in_browser=False,
            )
            html_content = summary.preview_path.read_text(encoding="utf-8")
            message = (
                f"Предпросмотр сохранён: {summary.preview_path}\n"
                f"Шаблон: {summary.template_name}\n"
                f"Получатель: {summary.recipient_email}"
            )
            self._show_preview_window(html_content=html_content, preview_path=summary.preview_path)
            self._append_log(message)
            self.status_var.set("Предпросмотр открыт в GUI")
        except CampaignError as error:
            self._append_log(f"Ошибка предпросмотра: {error}")
            messagebox.showerror("Email App Modern", str(error))

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

            self.editor_text = self.ctk.CTkTextbox(self.editor_window, wrap="none")
            self.editor_text.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        else:
            self.editor_window.lift()

        self._load_template_into_editor()
        self._append_log(f"Шаблон открыт в редакторе: {template_path}")

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

    def _save_template_from_editor(self) -> None:
        if self.editor_text is None:
            return
        template_path = self._current_template_path()
        template_path.parent.mkdir(parents=True, exist_ok=True)
        template_path.write_text(self.editor_text.get("1.0", "end").rstrip() + "\n", encoding="utf-8")
        self._append_log(f"Шаблон сохранён: {template_path}")
        self.status_var.set("Шаблон сохранён")
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
                messagebox.showinfo("Email App Modern", payload)
            elif event == "error":
                self._append_log(f"Ошибка: {payload}")
                self.status_var.set("Ошибка")
                messagebox.showerror("Email App Modern", payload)

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
