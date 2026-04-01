from __future__ import annotations

import importlib
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import Callable

from .campaign_queue import (
    CampaignQueueError,
    load_campaign_queue,
    run_campaign_queue,
    save_campaign_queue,
)
from .config import ConfigError, load_config
from .presets import CampaignPreset, PresetError, load_preset, save_preset
from .service import CampaignError, render_preview, run_campaign, run_preflight
from .renderer import TemplateRenderer
from .stats import (
    StatsError,
    export_history_records,
    filter_history_records,
    load_history_records,
    summarize_history_records,
)
from .tinymce_editor import RichEditorError, RichTemplateEditorServer


class EmailAppGUI:
    def __init__(self, root: tk.Tk, base_dir: Path, preset_path: Path | None = None) -> None:
        self.root = root
        self.base_dir = base_dir
        self.current_preset_path = preset_path
        self.queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self.worker: threading.Thread | None = None

        self.root.title("Email App")
        self.root.geometry("900x700")
        self.root.minsize(760, 560)

        self.config_var = tk.StringVar(value="config/settings.yaml")
        self.recipients_var = tk.StringVar(value="recipients.csv")
        self.templates_var = tk.StringVar(value="templates")
        self.template_var = tk.StringVar(value="")
        self.delay_var = tk.StringVar(value="0")
        self.dry_run_var = tk.BooleanVar(value=True)
        self.status_var = tk.StringVar(value="Готово к запуску")
        self.editor_window: tk.Toplevel | None = None
        self.editor_text: tk.Text | None = None
        self.stats_window: tk.Toplevel | None = None
        self.preview_window: tk.Toplevel | None = None
        self.preview_source_text: tk.Text | None = None
        self.preview_html_widget = None
        self.rich_editor_server = RichTemplateEditorServer(base_dir)

        self._build()
        if self.current_preset_path:
            self._load_preset(self.current_preset_path, silent=True)
        self._refresh_templates()
        self.root.after(150, self._poll_queue)

    def _build(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(7, weight=1)

        self._add_path_row(container, 0, "Конфиг", self.config_var, self._select_config)
        self._add_path_row(container, 1, "Получатели", self.recipients_var, self._select_recipients)
        self._add_path_row(container, 2, "Шаблоны", self.templates_var, self._select_templates)

        ttk.Label(container, text="HTML-шаблон").grid(row=3, column=0, sticky="w", pady=(8, 0))
        template_row = ttk.Frame(container)
        template_row.grid(row=3, column=1, columnspan=2, sticky="ew", pady=(8, 0))
        template_row.columnconfigure(0, weight=1)
        self.template_combo = ttk.Combobox(
            template_row,
            textvariable=self.template_var,
            state="readonly",
        )
        self.template_combo.grid(row=0, column=0, sticky="ew")
        ttk.Button(template_row, text="Обновить", command=self._refresh_templates).grid(
            row=0,
            column=1,
            padx=(8, 0),
        )

        ttk.Label(container, text="Задержка, сек").grid(row=4, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(container, textvariable=self.delay_var).grid(row=4, column=1, sticky="ew", pady=(8, 0))

        ttk.Checkbutton(
            container,
            text="Dry-run без реальной отправки",
            variable=self.dry_run_var,
        ).grid(row=5, column=0, columnspan=3, sticky="w", pady=(12, 0))

        buttons = ttk.Frame(container)
        buttons.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(16, 8))
        for index in range(5):
            buttons.columnconfigure(index, weight=1)

        button_specs = [
            ("Старт", self._start_send),
            ("Очередь JSON", self._run_queue_dialog),
            ("Экспорт JSON/CSV", self._export_queue_dialog),
            ("Предпросмотр", self._preview_email),
            ("Статистика", self._show_stats),
            ("Визуальный редактор", self._open_visual_template_editor),
            ("Редактор HTML", self._open_template_editor),
            ("Сохранить пресет", self._save_preset_dialog),
            ("Загрузить пресет", self._load_preset_dialog),
            ("Очистить лог", self._clear_log),
        ]
        for index, (label, command) in enumerate(button_specs):
            row = index // 5
            column = index % 5
            ttk.Button(buttons, text=label, command=command).grid(
                row=row,
                column=column,
                sticky="ew",
                padx=4,
                pady=4,
            )

        ttk.Label(container, textvariable=self.status_var).grid(
            row=7,
            column=0,
            columnspan=3,
            sticky="w",
            pady=(0, 8),
        )

        self.log_widget = tk.Text(container, wrap="word", height=18)
        self.log_widget.grid(row=8, column=0, columnspan=3, sticky="nsew")
        container.rowconfigure(8, weight=1)

        scrollbar = ttk.Scrollbar(container, orient="vertical", command=self.log_widget.yview)
        scrollbar.grid(row=8, column=3, sticky="ns")
        self.log_widget.configure(yscrollcommand=scrollbar.set)

    def _add_path_row(
        self,
        parent: ttk.Frame,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Callable[[], None],
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 8))
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=(0, 8))
        ttk.Button(parent, text="Выбрать", command=command).grid(row=row, column=2, padx=(8, 0), pady=(0, 8))

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
        self.template_combo["values"] = templates
        if templates and self.template_var.get() not in templates:
            self.template_var.set(templates[0])
        elif not templates:
            self.template_var.set("")

    def _start_send(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Email App", "Отправка уже выполняется")
            return

        try:
            delay = float(self.delay_var.get().strip() or "0")
        except ValueError:
            messagebox.showerror("Email App", "Задержка должна быть числом")
            return

        try:
            preflight = run_preflight(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                template_override=self.template_var.get() or None,
            )
        except CampaignError as error:
            messagebox.showerror("Email App", str(error))
            return

        if preflight.errors:
            messagebox.showerror("Email App", "\n".join(preflight.errors))
            return

        if preflight.warnings:
            proceed = messagebox.askyesno(
                "Email App",
                "Найдены предупреждения preflight:\n\n"
                + "\n".join(f"- {item}" for item in preflight.warnings)
                + "\n\nПродолжить?",
            )
            if not proceed:
                self.status_var.set("Отправка отменена")
                return

        if not self.dry_run_var.get():
            proceed = messagebox.askyesno("Email App", "Запустить боевую отправку?")
            if not proceed:
                self.status_var.set("Отправка отменена")
                return

        self.status_var.set("Запуск...")
        self._append_log("Старт задачи")
        self.worker = threading.Thread(
            target=self._run_worker,
            args=(delay,),
            daemon=True,
        )
        self.worker.start()

    def _run_queue_dialog(self) -> None:
        if self.worker and self.worker.is_alive():
            messagebox.showinfo("Email App", "Задача уже выполняется")
            return

        path = filedialog.askopenfilename(
            initialdir=self.base_dir / "presets",
            filetypes=[("JSON/CSV", "*.json *.csv"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        self.status_var.set("Запуск очереди...")
        self._append_log(f"Старт очереди: {path}")
        self.worker = threading.Thread(
            target=self._run_queue_worker,
            args=(Path(path),),
            daemon=True,
        )
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
                )
            ],
        )
        self._append_log(f"Очередь экспортирована: {exported_path}")
        self.status_var.set("Очередь экспортирована")

    def _preview_email(self) -> None:
        try:
            summary = render_preview(
                base_dir=self.base_dir,
                config_path=self.base_dir / self.config_var.get(),
                recipients_path=self.base_dir / self.recipients_var.get(),
                templates_path=self.base_dir / self.templates_var.get(),
                template_override=self.template_var.get() or None,
                open_in_browser=True,
            )
            message = (
                f"Предпросмотр сохранён: {summary.preview_path}\n"
                f"Шаблон: {summary.template_name}\n"
                f"Получатель: {summary.recipient_email}"
            )
            self._append_log(message)
            self.status_var.set("Предпросмотр открыт в браузере")
        except CampaignError as error:
            self._append_log(f"Ошибка предпросмотра: {error}")
            messagebox.showerror("Email App", str(error))

    def _show_stats(self) -> None:
        try:
            history_csv_path = self._current_history_csv_path()
            records = load_history_records(history_csv_path)
        except (StatsError, CampaignError) as error:
            self._append_log(f"Ошибка статистики: {error}")
            messagebox.showerror("Email App", str(error))
            return

        if self.stats_window is None or not self.stats_window.winfo_exists():
            self.stats_window = tk.Toplevel(self.root)
            self.stats_window.title("Статистика отправок")
            self.stats_window.geometry("920x620")

            filters = ttk.Frame(self.stats_window, padding=8)
            filters.pack(fill=tk.X)
            ttk.Label(filters, text="Статус").grid(row=0, column=0, sticky="w")
            self.stats_status_var = tk.StringVar(value="all")
            ttk.Combobox(
                filters,
                textvariable=self.stats_status_var,
                values=["all", "sent", "dry-run", "error"],
                state="readonly",
                width=12,
            ).grid(row=0, column=1, padx=(6, 12))
            ttk.Label(filters, text="Шаблон").grid(row=0, column=2, sticky="w")
            self.stats_template_var = tk.StringVar(value="")
            ttk.Entry(filters, textvariable=self.stats_template_var, width=22).grid(row=0, column=3, padx=(6, 12))
            ttk.Label(filters, text="SMTP").grid(row=0, column=4, sticky="w")
            self.stats_smtp_var = tk.StringVar(value="")
            ttk.Entry(filters, textvariable=self.stats_smtp_var, width=22).grid(row=0, column=5, padx=(6, 12))
            ttk.Button(filters, text="Применить", command=self._refresh_stats_window).grid(row=0, column=6)
            ttk.Button(filters, text="Экспорт CSV", command=lambda: self._export_stats(".csv")).grid(row=0, column=7, padx=(8, 0))
            ttk.Button(filters, text="Экспорт JSON", command=lambda: self._export_stats(".json")).grid(row=0, column=8, padx=(8, 0))

            summary_text = tk.Text(self.stats_window, wrap="word", height=10)
            summary_text.pack(fill=tk.X, padx=8, pady=(0, 8))
            self.stats_text = summary_text

            records_text = tk.Text(self.stats_window, wrap="none")
            records_text.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 8))
            self.stats_records_text = records_text
        else:
            self.stats_window.lift()
        self.stats_all_records = records
        self._refresh_stats_window()
        self._append_log(f"Статистика загружена: {history_csv_path}")
        self.status_var.set("Статистика загружена")

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
        self.stats_text.delete("1.0", tk.END)
        self.stats_text.insert("1.0", "\n".join(lines))
        row_lines = [
            f"{record.timestamp} | {record.status} | {record.recipient} | {record.template} | {record.smtp_account} | {record.error}"
            for record in records
        ]
        self.stats_records_text.delete("1.0", tk.END)
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

    def _save_preset_dialog(self) -> None:
        path = filedialog.asksaveasfilename(
            initialdir=self.base_dir / "presets",
            defaultextension=".yaml",
            filetypes=[("YAML", "*.yaml *.yml"), ("Все файлы", "*.*")],
        )
        if not path:
            return

        preset = CampaignPreset(
            config=self._portable_path_value(self.config_var.get()),
            recipients=self._portable_path_value(self.recipients_var.get()),
            templates=self._portable_path_value(self.templates_var.get()),
            template=self.template_var.get() or None,
            delay_seconds=float(self.delay_var.get().strip() or "0"),
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
                messagebox.showerror("Email App", str(error))
            return

        self.current_preset_path = path
        self.config_var.set(self._portable_path_value(preset.config))
        self.recipients_var.set(self._portable_path_value(preset.recipients))
        self.templates_var.set(self._portable_path_value(preset.templates))
        self.delay_var.set("" if preset.delay_seconds is None else str(preset.delay_seconds))
        self.dry_run_var.set(preset.dry_run)
        self._refresh_templates()
        if preset.template:
            self.template_var.set(preset.template)
        self._append_log(f"Пресет загружен: {path}")
        self.status_var.set("Пресет загружен")
        if not silent:
            messagebox.showinfo("Email App", f"Пресет загружен: {path}")

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
            self.queue.put(
                (
                    "done",
                    f"Готово. Обработано: {summary.processed}/{summary.total}, успешно: {summary.successful}, ошибок: {summary.failed}",
                )
            )
            self.queue.put(("log", f"История CSV: {summary.history_csv}"))
            self.queue.put(("log", f"История JSONL: {summary.history_jsonl}"))
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

    def _show_preview_window(self, html_content: str, preview_path: Path) -> None:
        if self.preview_window is None or not self.preview_window.winfo_exists():
            self.preview_window = tk.Toplevel(self.root)
            self.preview_window.title("HTML Preview")
            self.preview_window.geometry("1100x760")

            toolbar = ttk.Frame(self.preview_window, padding=8)
            toolbar.pack(fill=tk.X)
            ttk.Button(
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
            ).pack(side=tk.LEFT)
            ttk.Label(toolbar, text=str(preview_path)).pack(side=tk.LEFT, padx=(12, 0))

            panes = ttk.Panedwindow(self.preview_window, orient=tk.HORIZONTAL)
            panes.pack(fill=tk.BOTH, expand=True)

            left = ttk.Frame(panes)
            right = ttk.Frame(panes)
            panes.add(left, weight=2)
            panes.add(right, weight=1)

            try:
                html_module = importlib.import_module("tkhtmlview")
                self.preview_html_widget = html_module.HTMLScrolledText(left, html="")
                self.preview_html_widget.pack(fill=tk.BOTH, expand=True)
            except ImportError:
                self.preview_html_widget = tk.Text(left, wrap="word")
                self.preview_html_widget.pack(fill=tk.BOTH, expand=True)

            self.preview_source_text = tk.Text(right, wrap="none")
            self.preview_source_text.pack(fill=tk.BOTH, expand=True)

        if hasattr(self.preview_html_widget, "set_html"):
            self.preview_html_widget.set_html(html_content)
        else:
            self.preview_html_widget.delete("1.0", tk.END)
            self.preview_html_widget.insert("1.0", html_content)
        self.preview_source_text.delete("1.0", tk.END)
        self.preview_source_text.insert("1.0", html_content)

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
            messagebox.showerror("Email App", str(error))
            return

        if self.editor_window is not None and self.editor_window.winfo_exists():
            self.editor_window.lift()
        else:
            self.editor_window = tk.Toplevel(self.root)
            self.editor_window.title("Редактор HTML-шаблона")
            self.editor_window.geometry("900x700")

            toolbar = ttk.Frame(self.editor_window, padding=8)
            toolbar.pack(fill=tk.X)
            ttk.Button(toolbar, text="Перезагрузить", command=self._load_template_into_editor).pack(side=tk.LEFT)
            ttk.Button(toolbar, text="Сохранить", command=self._save_template_from_editor).pack(
                side=tk.LEFT,
                padx=(8, 0),
            )

            self.editor_text = tk.Text(self.editor_window, wrap="none", undo=True)
            self.editor_text.pack(fill=tk.BOTH, expand=True)

            scrollbar_y = ttk.Scrollbar(self.editor_window, orient="vertical", command=self.editor_text.yview)
            scrollbar_y.pack(side=tk.RIGHT, fill=tk.Y)
            self.editor_text.configure(yscrollcommand=scrollbar_y.set)

            scrollbar_x = ttk.Scrollbar(self.editor_window, orient="horizontal", command=self.editor_text.xview)
            scrollbar_x.pack(side=tk.BOTTOM, fill=tk.X)
            self.editor_text.configure(xscrollcommand=scrollbar_x.set)

            self.editor_window.bind("<Control-s>", lambda _event: self._save_template_from_editor())

        self._load_template_into_editor()

    def _open_visual_template_editor(self) -> None:
        try:
            template_path = self._current_template_path()
            editor_url = self.rich_editor_server.open_template(template_path)
        except CampaignError as error:
            messagebox.showerror("Email App", str(error))
            return
        except RichEditorError as error:
            messagebox.showerror("Email App", str(error))
            return
        except Exception as error:  # noqa: BLE001
            messagebox.showerror("Email App", f"Не удалось открыть визуальный редактор: {error}")
            return

        self._append_log(f"Визуальный редактор открыт: {template_path} | {editor_url}")
        self.status_var.set("Визуальный редактор открыт")
        messagebox.showinfo("Email App", "Открыт визуальный редактор в браузере. После сохранения шаблон сразу обновится в проекте.")

    def _load_template_into_editor(self) -> None:
        if self.editor_text is None:
            return
        template_path = self._current_template_path()
        if not template_path.exists():
            messagebox.showerror("Email App", f"Шаблон не найден: {template_path}")
            return
        self.editor_text.delete("1.0", tk.END)
        self.editor_text.insert("1.0", template_path.read_text(encoding="utf-8"))
        self._append_log(f"Шаблон открыт в редакторе: {template_path}")
        self.status_var.set("Шаблон открыт")

    def _save_template_from_editor(self) -> None:
        if self.editor_text is None:
            return
        template_path = self._current_template_path()
        template_path.parent.mkdir(parents=True, exist_ok=True)
        template_path.write_text(self.editor_text.get("1.0", tk.END).rstrip() + "\n", encoding="utf-8")
        self._append_log(f"Шаблон сохранён: {template_path}")
        self.status_var.set("Шаблон сохранён")
        messagebox.showinfo("Email App", f"Шаблон сохранён: {template_path}")

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
                messagebox.showinfo("Email App", payload)
            elif event == "error":
                self._append_log(f"Ошибка: {payload}")
                self.status_var.set("Ошибка")
                messagebox.showerror("Email App", payload)

        self.root.after(150, self._poll_queue)

    def _append_log(self, message: str) -> None:
        self.log_widget.insert(tk.END, f"{message}\n")
        self.log_widget.see(tk.END)

    def _clear_log(self) -> None:
        self.log_widget.delete("1.0", tk.END)

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


def launch_gui(base_dir: Path, preset_path: Path | None = None) -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "clam" in style.theme_names():
        style.theme_use("clam")
    EmailAppGUI(root=root, base_dir=base_dir, preset_path=preset_path)
    root.mainloop()
