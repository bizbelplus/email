from __future__ import annotations

import base64
import csv
import json
import logging
import time
import webbrowser
from dataclasses import dataclass
import mimetypes
from datetime import datetime, timezone
from pathlib import Path
from typing import Callable

from .config import ConfigError, load_config
from .recipients import RecipientsError, load_recipients
from .renderer import TemplateRenderer
from .smtp_client import SMTPMailer


class CampaignError(RuntimeError):
    """Raised when campaign data is invalid."""


@dataclass(slots=True)
class CampaignSummary:
    total: int
    processed: int
    successful: int
    failed: int
    log_file: Path
    history_csv: Path
    history_jsonl: Path


@dataclass(slots=True)
class PreviewSummary:
    preview_path: Path
    template_name: str
    recipient_email: str


def _create_logger(base_dir: Path, log_file: str) -> tuple[logging.Logger, Path]:
    log_path = (base_dir / log_file).resolve()
    log_path.parent.mkdir(parents=True, exist_ok=True)

    logger_name = f"email_app.{log_path}"
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    logger.propagate = False

    if not logger.handlers:
        handler = logging.FileHandler(log_path, encoding="utf-8")
        formatter = logging.Formatter("%(asctime)s | %(levelname)s | %(message)s")
        handler.setFormatter(formatter)
        logger.addHandler(handler)

    return logger, log_path


def _resolve_output_path(base_dir: Path, relative_path: str) -> Path:
    output_path = (base_dir / relative_path).resolve()
    output_path.parent.mkdir(parents=True, exist_ok=True)
    return output_path


def _append_history(
    *,
    csv_path: Path,
    jsonl_path: Path,
    record: dict[str, str],
) -> None:
    headers = [
        "timestamp",
        "recipient",
        "status",
        "template",
        "smtp_account",
        "dry_run",
        "error",
    ]

    csv_exists = csv_path.exists()
    with csv_path.open("a", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=headers)
        if not csv_exists:
            writer.writeheader()
        writer.writerow({key: record.get(key, "") for key in headers})

    with jsonl_path.open("a", encoding="utf-8") as handle:
        handle.write(json.dumps(record, ensure_ascii=False) + "\n")


def _resolve_attachment_paths(base_dir: Path, attachments: list[str]) -> list[Path]:
    resolved: list[Path] = []
    for attachment in attachments:
        candidate = (base_dir / attachment).resolve()
        if not candidate.exists() or not candidate.is_file():
            raise CampaignError(f"Не найдено вложение: {candidate}")
        resolved.append(candidate)
    return resolved


def _resolve_inline_image_paths(base_dir: Path, inline_images: dict[str, str]) -> dict[str, Path]:
    resolved: dict[str, Path] = {}
    for cid, image_path in inline_images.items():
        candidate = (base_dir / image_path).resolve()
        if not candidate.exists() or not candidate.is_file():
            raise CampaignError(f"Не найдено inline-изображение '{cid}': {candidate}")
        resolved[cid] = candidate
    return resolved


def _build_inline_context(inline_image_paths: dict[str, Path]) -> dict[str, dict[str, str]]:
    return {
        cid: {
            "cid": cid,
            "path": str(path),
        }
        for cid, path in inline_image_paths.items()
    }


def _inject_preview_inline_images(html_body: str, inline_image_paths: dict[str, Path]) -> str:
    preview_html = html_body
    for cid, path in inline_image_paths.items():
        mime_type, _ = mimetypes.guess_type(path.name)
        mime_type = mime_type or "application/octet-stream"
        encoded = base64.b64encode(path.read_bytes()).decode("ascii")
        preview_html = preview_html.replace(f"cid:{cid}", f"data:{mime_type};base64,{encoded}")
    return preview_html


def render_preview(
    *,
    base_dir: Path,
    config_path: Path,
    recipients_path: Path,
    templates_path: Path,
    template_override: str | None = None,
    preview_path: Path | None = None,
    open_in_browser: bool = False,
) -> PreviewSummary:
    try:
        config = load_config(config_path)
        recipients = load_recipients(recipients_path)
    except (ConfigError, RecipientsError) as error:
        raise CampaignError(str(error)) from error

    renderer = TemplateRenderer(templates_path)
    available_templates = renderer.list_templates()
    template_name = template_override or config.message.template
    if template_name not in available_templates:
        raise CampaignError(f"Шаблон не найден: {template_name}")

    inline_image_paths = _resolve_inline_image_paths(base_dir, config.message.inline_images)
    recipient = recipients[0]
    html_body = renderer.render(
        template_name=template_name,
        recipient=recipient,
        context={
            **config.content,
            "subject": config.message.subject,
            "inline_images": _build_inline_context(inline_image_paths),
        },
    )
    html_body = _inject_preview_inline_images(html_body, inline_image_paths)
    output_path = preview_path or (base_dir / "preview" / "email_preview.html")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(html_body, encoding="utf-8")
    if open_in_browser:
        webbrowser.open(output_path.as_uri())
    return PreviewSummary(
        preview_path=output_path,
        template_name=template_name,
        recipient_email=recipient.email,
    )


def run_campaign(
    *,
    base_dir: Path,
    config_path: Path,
    recipients_path: Path,
    templates_path: Path,
    dry_run: bool = False,
    template_override: str | None = None,
    delay_override: float | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> CampaignSummary:
    try:
        config = load_config(config_path)
        recipients = load_recipients(recipients_path)
    except (ConfigError, RecipientsError) as error:
        raise CampaignError(str(error)) from error

    renderer = TemplateRenderer(templates_path)
    available_templates = renderer.list_templates()
    template_name = template_override or config.message.template
    if template_name not in available_templates:
        raise CampaignError(f"Шаблон не найден: {template_name}")

    attachment_paths = _resolve_attachment_paths(base_dir, config.message.attachments)
    inline_image_paths = _resolve_inline_image_paths(base_dir, config.message.inline_images)
    logger, log_path = _create_logger(base_dir, config.delivery.log_file)
    history_csv_path = _resolve_output_path(base_dir, config.delivery.history_csv)
    history_jsonl_path = _resolve_output_path(base_dir, config.delivery.history_jsonl)
    mailers = [SMTPMailer(settings) for settings in config.smtp_accounts]
    delay_seconds = config.delivery.delay_seconds if delay_override is None else max(delay_override, 0.0)

    successful = 0
    failed = 0
    processed = 0

    def emit(message: str) -> None:
        if progress_callback is not None:
            progress_callback(message)

    logger.info(
        "Старт кампании: dry_run=%s, recipients=%s, template=%s, attachments=%s, inline_images=%s, smtp_accounts=%s",
        dry_run,
        len(recipients),
        template_name,
        len(attachment_paths),
        len(inline_image_paths),
        len(mailers),
    )
    emit(f"Лог: {log_path}")
    emit(f"История CSV: {history_csv_path}")
    emit(f"История JSONL: {history_jsonl_path}")

    for index, recipient in enumerate(recipients, start=1):
        mailer = mailers[(index - 1) % len(mailers)]
        smtp_account = mailer.settings.from_email
        html_body = renderer.render(
            template_name=template_name,
            recipient=recipient,
            context={
                **config.content,
                "subject": config.message.subject,
                "inline_images": _build_inline_context(inline_image_paths),
            },
        )

        try:
            if dry_run:
                message = f"[DRY-RUN] {index}/{len(recipients)} подготовлено для {recipient.email}"
                successful += 1
                history_status = "dry-run"
                history_error = ""
            else:
                mailer.send(
                    recipient=recipient,
                    message_settings=config.message,
                    html_body=html_body,
                    attachment_paths=attachment_paths,
                    inline_image_paths=inline_image_paths,
                )
                message = (
                    f"[OK] {index}/{len(recipients)} отправлено: {recipient.email} "
                    f"через {smtp_account}"
                )
                successful += 1
                history_status = "sent"
                history_error = ""
            logger.info(message)
            emit(message)
        except Exception as error:  # noqa: BLE001
            failed += 1
            message = f"[ERROR] {index}/{len(recipients)} {recipient.email}: {error}"
            logger.exception(message)
            emit(message)
            history_status = "error"
            history_error = str(error)
        finally:
            processed += 1
            _append_history(
                csv_path=history_csv_path,
                jsonl_path=history_jsonl_path,
                record={
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "recipient": recipient.email,
                    "status": history_status,
                    "template": template_name,
                    "smtp_account": smtp_account,
                    "dry_run": str(dry_run).lower(),
                    "error": history_error,
                },
            )

        if delay_seconds > 0 and index < len(recipients):
            emit(f"Пауза {delay_seconds:.2f} сек.")
            time.sleep(delay_seconds)

    logger.info(
        "Завершено: processed=%s, successful=%s, failed=%s",
        processed,
        successful,
        failed,
    )
    return CampaignSummary(
        total=len(recipients),
        processed=processed,
        successful=successful,
        failed=failed,
        log_file=log_path,
        history_csv=history_csv_path,
        history_jsonl=history_jsonl_path,
    )
