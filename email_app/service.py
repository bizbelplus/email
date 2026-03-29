from __future__ import annotations

import base64
import csv
import json
import logging
import re
import time
import webbrowser
from dataclasses import dataclass
import mimetypes
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

from .config import ConfigError, load_config
from .recipients import RecipientsError, load_recipients
from .renderer import TemplateRenderer
from .smtp_client import SMTPMailer
from .validators import validate_csv_recipients


def _send_with_retry(
    mailer,
    recipient,
    message_settings,
    html_body,
    attachment_paths,
    inline_image_paths,
    retry_attempts: int = 1,
    retry_backoff_seconds: float = 5.0,
) -> None:
    """Send email with automatic retry on temporary failures."""
    last_error = None
    for attempt in range(retry_attempts):
        try:
            mailer.send(
                recipient=recipient,
                message_settings=message_settings,
                html_body=html_body,
                attachment_paths=attachment_paths,
                inline_image_paths=inline_image_paths,
            )
            return
        except Exception as error:
            last_error = error
            if attempt < retry_attempts - 1:
                time.sleep(retry_backoff_seconds)
    if last_error:
        raise last_error


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


@dataclass(slots=True)
class PreflightReport:
    checks: list[str]
    warnings: list[str]
    errors: list[str]

    @property
    def ok(self) -> bool:
        return not self.errors


def _extract_links(html_body: str) -> list[str]:
    return [
        match.group(1).strip()
        for match in re.finditer(r'href\s*=\s*["\']([^"\']+)["\']', html_body, flags=re.IGNORECASE)
    ]


def _validate_rendered_template(html_body: str) -> tuple[list[str], list[str]]:
    warnings: list[str] = []
    errors: list[str] = []

    unresolved_vars = re.findall(r"\{\{\s*[^}]+\s*\}\}", html_body)
    unresolved_blocks = re.findall(r"\{\%\s*[^%]+\s*\%\}", html_body)
    if unresolved_blocks:
        warnings.append("В HTML остались Jinja-блоки (возможно это ожидаемо): " + ", ".join(unresolved_blocks[:3]))
    if unresolved_vars:
        warnings.append("В HTML остались переменные Jinja: " + ", ".join(unresolved_vars[:3]))

    links = _extract_links(html_body)
    for link in links:
        parsed = urlparse(link)
        if link.startswith(("#", "cid:")):
            continue
        if parsed.scheme in {"http", "https", "mailto", "tel"}:
            continue
        warnings.append(f"Ссылка может быть некорректной: {link}")

    if "<html" not in html_body.lower() or "<body" not in html_body.lower():
        warnings.append("Шаблон не содержит полные теги <html>/<body>")

    return warnings, errors


def run_preflight(
    *,
    base_dir: Path,
    config_path: Path,
    recipients_path: Path,
    templates_path: Path,
    template_override: str | None = None,
) -> PreflightReport:
    checks: list[str] = []
    warnings: list[str] = []
    errors: list[str] = []

    try:
        config = load_config(config_path)
        checks.append(f"Конфиг загружен: {config_path}")
    except ConfigError as error:
        return PreflightReport(checks=checks, warnings=warnings, errors=[str(error)])

    try:
        recipients = load_recipients(recipients_path)
        checks.append(f"Получатели загружены: {len(recipients)}")
    except RecipientsError as error:
        return PreflightReport(checks=checks, warnings=warnings, errors=[str(error)])

    # Email validation
    email_validation = validate_csv_recipients(str(recipients_path))
    if email_validation.get("error"):
        warnings.append(f"Ошибка валидации email: {email_validation['error']}")
    else:
        invalid_count = email_validation.get("total_invalid", 0)
        warning_count = email_validation.get("total_warnings", 0)
        if invalid_count > 0:
            errors.append(f"Невалидные email-адреса: {invalid_count}")
        if warning_count > 0:
            warnings.append(f"Email-адреса с предупреждениями: {warning_count}")

    template_name = template_override or config.message.template
    renderer = TemplateRenderer(templates_path)
    available_templates = renderer.list_templates()
    if template_name not in available_templates:
        errors.append(f"Шаблон не найден: {template_name}")
        return PreflightReport(checks=checks, warnings=warnings, errors=errors)
    checks.append(f"Шаблон найден: {template_name}")

    try:
        attachment_paths = _resolve_attachment_paths(base_dir, config.message.attachments)
        inline_image_paths = _resolve_inline_image_paths(base_dir, config.message.inline_images)
    except CampaignError as error:
        errors.append(str(error))
        return PreflightReport(checks=checks, warnings=warnings, errors=errors)

    checks.append(f"Вложений: {len(attachment_paths)}")
    checks.append(f"Inline-изображений: {len(inline_image_paths)}")
    checks.append(f"SMTP-аккаунтов: {len(config.smtp_accounts)}")

    first_recipient = recipients[0]
    rendered = renderer.render(
        template_name=template_name,
        recipient=first_recipient,
        context={
            **config.content,
            "subject": config.message.subject,
            "inline_images": _build_inline_context(inline_image_paths),
        },
    )
    checks.append(f"Рендер тестового письма: {first_recipient.email}")

    template_warnings, template_errors = _validate_rendered_template(rendered)
    warnings.extend(template_warnings)
    errors.extend(template_errors)

    return PreflightReport(checks=checks, warnings=warnings, errors=errors)


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


def _build_dedupe_sent_set(
    *,
    history_csv_path: Path,
    template_name: str,
    dedupe_template_scope: bool,
    dedupe_history_days: int,
) -> set[str]:
    if not history_csv_path.exists():
        return set()

    now_utc = datetime.now(timezone.utc)
    cutoff = now_utc - timedelta(days=max(dedupe_history_days, 0))
    sent: set[str] = set()

    with history_csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        for row in reader:
            if (row.get("status") or "").strip().lower() != "sent":
                continue
            recipient = (row.get("recipient") or "").strip().lower()
            if not recipient:
                continue
            row_template = (row.get("template") or "").strip()
            if dedupe_template_scope and row_template != template_name:
                continue

            timestamp_raw = (row.get("timestamp") or "").strip()
            if timestamp_raw:
                try:
                    row_dt = datetime.fromisoformat(timestamp_raw.replace("Z", "+00:00"))
                    if row_dt.tzinfo is None:
                        row_dt = row_dt.replace(tzinfo=timezone.utc)
                    if dedupe_history_days > 0 and row_dt < cutoff:
                        continue
                except ValueError:
                    pass

            sent.add(recipient)

    return sent


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

    preflight = run_preflight(
        base_dir=base_dir,
        config_path=config_path,
        recipients_path=recipients_path,
        templates_path=templates_path,
        template_override=template_override,
    )
    if preflight.errors:
        raise CampaignError("; ".join(preflight.errors))

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
    
    # Rate limiting: calculate minimum delay between sends
    min_delay_from_rate = 0.0
    if config.delivery.rate_limit_per_minute and config.delivery.rate_limit_per_minute > 0:
        min_delay_from_rate = 60.0 / config.delivery.rate_limit_per_minute
    
    final_delay = max(delay_seconds, min_delay_from_rate)

    successful = 0
    failed = 0
    processed = 0
    dedupe_sent_set: set[str] = set()
    if config.delivery.skip_previously_sent:
        dedupe_sent_set = _build_dedupe_sent_set(
            history_csv_path=history_csv_path,
            template_name=template_name,
            dedupe_template_scope=config.delivery.dedupe_template_scope,
            dedupe_history_days=config.delivery.dedupe_history_days,
        )

    def emit(message: str) -> None:
        if progress_callback is not None:
            progress_callback(message)

    logger.info(
        "Старт кампании: dry_run=%s, recipients=%s, template=%s, attachments=%s, inline_images=%s, smtp_accounts=%s, dedupe=%s",
        dry_run,
        len(recipients),
        template_name,
        len(attachment_paths),
        len(inline_image_paths),
        len(mailers),
        config.delivery.skip_previously_sent,
    )
    emit(f"Лог: {log_path}")
    emit(f"История CSV: {history_csv_path}")
    emit(f"История JSONL: {history_jsonl_path}")

    for index, recipient in enumerate(recipients, start=1):
        mailer = mailers[(index - 1) % len(mailers)]
        smtp_account = mailer.settings.from_email
        recipient_key = recipient.email.strip().lower()

        if config.delivery.skip_previously_sent and recipient_key in dedupe_sent_set:
            message = f"[SKIP] {index}/{len(recipients)} дубликат: {recipient.email}"
            logger.info(message)
            emit(message)
            processed += 1
            _append_history(
                csv_path=history_csv_path,
                jsonl_path=history_jsonl_path,
                record={
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "recipient": recipient.email,
                    "status": "skipped-duplicate",
                    "template": template_name,
                    "smtp_account": smtp_account,
                    "dry_run": str(dry_run).lower(),
                    "error": "already_sent",
                },
            )
            continue

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
                _send_with_retry(
                    mailer=mailer,
                    recipient=recipient,
                    message_settings=config.message,
                    html_body=html_body,
                    attachment_paths=attachment_paths,
                    inline_image_paths=inline_image_paths,
                    retry_attempts=config.delivery.retry_attempts,
                    retry_backoff_seconds=config.delivery.retry_backoff_seconds,
                )
                message = (
                    f"[OK] {index}/{len(recipients)} отправлено: {recipient.email} "
                    f"через {smtp_account}"
                )
                successful += 1
                history_status = "sent"
                history_error = ""
                dedupe_sent_set.add(recipient_key)
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

        if final_delay > 0 and index < len(recipients):
            emit(f"Пауза {final_delay:.2f} сек.")
            time.sleep(final_delay)

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
