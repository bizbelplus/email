from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

from .campaign_queue import CampaignQueueError, load_campaign_queue, run_campaign_queue, save_campaign_queue
from .gui import launch_gui
from .modern_gui import launch_modern_gui
from .presets import CampaignPreset, PresetError, load_preset, save_preset
from .service import CampaignError, render_preview, run_campaign, run_preflight
from .stats import StatsError, filter_history_records, load_history_records, summarize_history_records


def _get_base_dir() -> Path:
    """Return stable project/app directory across script and frozen exe runs."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent

    def is_project_root(path: Path) -> bool:
        return (path / "config" / "settings.yaml").exists() or (path / "config" / "settings.example.yaml").exists()

    module_root = Path(__file__).resolve().parent.parent
    if is_project_root(module_root):
        return module_root

    cwd = Path.cwd().resolve()
    if is_project_root(cwd):
        return cwd

    for parent in cwd.parents:
        if is_project_root(parent):
            return parent

    return cwd


def _to_portable_path(value: str, base_dir: Path) -> str:
    text = str(value).strip()
    if not text:
        return text
    candidate = Path(text)
    if not candidate.is_absolute():
        return text
    try:
        return str(candidate.resolve().relative_to(base_dir.resolve()))
    except ValueError:
        return text


def _ensure_default_config(base_dir: Path) -> None:
    config_dir = base_dir / "config"
    settings_path = config_dir / "settings.yaml"
    example_path = config_dir / "settings.example.yaml"
    if settings_path.exists() or not example_path.exists():
        return
    config_dir.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(example_path, settings_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="SMTP-приложение для отправки HTML-писем по списку получателей"
    )
    parser.add_argument(
        "--config",
        default="config/settings.yaml",
        help="Путь до YAML-конфига",
    )
    parser.add_argument(
        "--recipients",
        default="recipients.csv",
        help="Путь до CSV (столбец email) или TXT (по одному email на строку)",
    )
    parser.add_argument(
        "--templates",
        default="templates",
        help="Папка с HTML-шаблонами",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Только отрендерить письма без реальной отправки",
    )
    parser.add_argument(
        "--template",
        default=None,
        help="Переопределить имя шаблона (HTML/TXT) из папки templates",
    )
    parser.add_argument(
        "--delay-seconds",
        type=float,
        default=None,
        help="Переопределить задержку между письмами",
    )
    parser.add_argument(
        "--gui",
        action="store_true",
        help="Запустить desktop-интерфейс на tkinter",
    )
    parser.add_argument(
        "--modern-gui",
        action="store_true",
        help="Запустить modern-интерфейс на customtkinter",
    )
    parser.add_argument(
        "--preview",
        action="store_true",
        help="Сгенерировать HTML-предпросмотр по первому получателю и открыть в браузере",
    )
    parser.add_argument(
        "--preflight",
        action="store_true",
        help="Проверить кампанию перед запуском и вывести чеклист",
    )
    parser.add_argument(
        "--yes",
        action="store_true",
        help="Запускать боевую отправку без интерактивного подтверждения",
    )
    parser.add_argument(
        "--preset",
        default=None,
        help="Загрузить параметры запуска из YAML-пресета",
    )
    parser.add_argument(
        "--save-preset",
        default=None,
        help="Сохранить текущие параметры запуска в YAML-пресет и выйти",
    )
    parser.add_argument(
        "--queue-file",
        default=None,
        help="Запустить очередь кампаний из JSON-файла",
    )
    parser.add_argument(
        "--export-queue",
        default=None,
        help="Экспортировать текущую кампанию в JSON/CSV-файл очереди и выйти",
    )
    parser.add_argument(
        "--show-stats",
        action="store_true",
        help="Показать статистику по history CSV и выйти",
    )
    parser.add_argument(
        "--history-csv",
        default="history/email_history.csv",
        help="Путь до CSV-файла истории для статистики",
    )
    parser.add_argument(
        "--status-filter",
        default=None,
        help="Фильтр статистики по статусу: sent, dry-run, error",
    )
    parser.add_argument(
        "--template-filter",
        default=None,
        help="Фильтр статистики по части имени шаблона",
    )
    parser.add_argument(
        "--smtp-filter",
        default=None,
        help="Фильтр статистики по части SMTP-аккаунта",
    )
    parser.add_argument(
        "--export-stats",
        default=None,
        help="Экспортировать отфильтрованные записи статистики в CSV/JSON и выйти",
    )
    return parser


def main() -> int:
    args = build_parser().parse_args()
    base_dir = _get_base_dir()
    _ensure_default_config(base_dir)

    try:
        if args.preset:
            preset = load_preset(base_dir / args.preset)
            args.config = preset.config
            args.recipients = preset.recipients
            args.templates = preset.templates
            args.template = preset.template or args.template
            args.delay_seconds = preset.delay_seconds if args.delay_seconds is None else args.delay_seconds
            if not args.dry_run:
                args.dry_run = preset.dry_run

        if args.save_preset:
            saved_path = save_preset(
                base_dir / args.save_preset,
                CampaignPreset(
                    config=_to_portable_path(args.config, base_dir),
                    recipients=_to_portable_path(args.recipients, base_dir),
                    templates=_to_portable_path(args.templates, base_dir),
                    template=args.template,
                    delay_seconds=args.delay_seconds,
                    dry_run=args.dry_run,
                ),
            )
            print(f"Пресет сохранён: {saved_path}")
            return 0

        if args.export_queue:
            exported_path = save_campaign_queue(
                base_dir / args.export_queue,
                [
                    CampaignPreset(
                        config=_to_portable_path(args.config, base_dir),
                        recipients=_to_portable_path(args.recipients, base_dir),
                        templates=_to_portable_path(args.templates, base_dir),
                        template=args.template,
                        delay_seconds=args.delay_seconds,
                        dry_run=args.dry_run,
                    )
                ],
            )
            print(f"Очередь экспортирована: {exported_path}")
            return 0

        # При запуске собранного .exe без аргументов запускаем modern GUI по умолчанию
        _is_frozen = getattr(sys, "frozen", False)
        _no_cli_flags = not any([
            args.preview, args.preflight, args.yes, args.save_preset,
            args.export_queue, args.queue_file, args.show_stats, args.export_stats,
        ])
        if args.modern_gui or (_is_frozen and _no_cli_flags and not args.gui):
            launch_modern_gui(base_dir=base_dir, preset_path=(base_dir / args.preset) if args.preset else None)
            return 0

        if args.gui:
            launch_gui(base_dir=base_dir, preset_path=(base_dir / args.preset) if args.preset else None)
            return 0

        if args.show_stats:
            records = filter_history_records(
                load_history_records(base_dir / args.history_csv),
                status=args.status_filter,
                template_query=args.template_filter,
                smtp_query=args.smtp_filter,
            )
            stats = summarize_history_records(records)
            print(
                f"Статистика: total={stats.total}, sent={stats.sent}, dry_run={stats.dry_run}, "
                f"errors={stats.errors}, unique_recipients={stats.unique_recipients}"
            )
            if stats.top_templates:
                print(f"Топ шаблонов: {stats.top_templates}")
            if stats.top_smtp_accounts:
                print(f"Топ SMTP: {stats.top_smtp_accounts}")
            if args.export_stats:
                from .stats import export_history_records

                exported_path = export_history_records(base_dir / args.export_stats, records)
                print(f"Экспорт статистики: {exported_path}")
            return 0

        if args.queue_file:
            queue_summary = run_campaign_queue(
                base_dir=base_dir,
                campaigns=load_campaign_queue(base_dir / args.queue_file),
                progress_callback=print,
            )
            print(
                f"Очередь завершена. Кампаний: {queue_summary.campaigns_completed}/"
                f"{queue_summary.campaigns_total}, обработано: {queue_summary.total_processed}, "
                f"успешно: {queue_summary.total_successful}, ошибок: {queue_summary.total_failed}"
            )
            return 0 if queue_summary.total_failed == 0 else 1

        if args.preview:
            preview = render_preview(
                base_dir=base_dir,
                config_path=base_dir / args.config,
                recipients_path=base_dir / args.recipients,
                templates_path=base_dir / args.templates,
                template_override=args.template,
                open_in_browser=True,
            )
            print(
                f"Предпросмотр сохранён: {preview.preview_path} | "
                f"шаблон: {preview.template_name} | получатель: {preview.recipient_email}"
            )
            return 0

        preflight = run_preflight(
            base_dir=base_dir,
            config_path=base_dir / args.config,
            recipients_path=base_dir / args.recipients,
            templates_path=base_dir / args.templates,
            template_override=args.template,
        )
        print("Чеклист:")
        for item in preflight.checks:
            print(f"  - OK: {item}")
        for item in preflight.warnings:
            print(f"  - WARN: {item}")
        for item in preflight.errors:
            print(f"  - ERROR: {item}")

        if preflight.errors:
            return 1

        if args.preflight:
            return 0

        if not args.dry_run and not args.yes:
            answer = input("Продолжить боевую отправку? Введите 'yes': ").strip().lower()
            if answer != "yes":
                print("Отправка отменена")
                return 1

        summary = run_campaign(
            base_dir=base_dir,
            config_path=base_dir / args.config,
            recipients_path=base_dir / args.recipients,
            templates_path=base_dir / args.templates,
            dry_run=args.dry_run,
            template_override=args.template,
            delay_override=args.delay_seconds,
            progress_callback=print,
        )
    except (CampaignError, PresetError, CampaignQueueError, StatsError) as error:
        print(f"Ошибка конфигурации: {error}")
        return 1

    print(
        f"Готово. Обработано: {summary.processed}/{summary.total}, "
        f"успешно: {summary.successful}, ошибок: {summary.failed}"
    )
    print(f"История CSV: {summary.history_csv}")
    print(f"История JSONL: {summary.history_jsonl}")
    return 0 if summary.failed == 0 else 1
