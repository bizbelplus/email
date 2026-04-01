from __future__ import annotations

import base64
import csv
import json
import logging
import random
import re
import time
import threading
import webbrowser
from concurrent.futures import ThreadPoolExecutor, as_completed
from copy import deepcopy
from dataclasses import dataclass
import mimetypes
from datetime import datetime, timedelta, timezone
from pathlib import Path
from typing import Callable
from urllib.parse import urlparse

from .config import ConfigError, load_config
from .models import MessageSettings, Recipient
from .recipients import RecipientsError, load_recipients
from .renderer import TemplateRenderer
from .smtp_client import SMTPMailer
from .validators import validate_csv_recipients
from .proxy_utils import load_proxies, pick_random_proxy


def _humanize_error_ru(error: Exception | str) -> str:
    raw = str(error)
    low = raw.lower()

    if "network unreachable" in low:
        return "Сеть недоступна для целевого узла (обычно прокси не имеет маршрута к SMTP)"
    if "timed out" in low or "timeout" in low:
        return "Превышено время ожидания соединения"
    if "authentication" in low or "535" in low or "username and password not accepted" in low:
        return "Ошибка авторизации SMTP (логин/пароль или app-password неверные)"
    if "connection refused" in low:
        return "Соединение отклонено сервером"
    if "name or service not known" in low or "nodename nor servname provided" in low:
        return "Не удалось определить DNS-имя SMTP-сервера"
    if "certificate" in low or "ssl" in low:
        return "Ошибка SSL/TLS (сертификат или режим шифрования)"
    if "socket error" in low:
        return f"Ошибка сокета: {raw}"
    return raw


def _is_retryable_send_error(error: Exception | str) -> bool:
    raw = str(error).lower()
    retry_markers = (
        "network unreachable",
        "timed out",
        "timeout",
        "connection reset",
        "connection aborted",
        "temporary",
        "try again",
        "socket error",
    )
    return any(marker in raw for marker in retry_markers)


def _send_with_retry(
    mailer,
    recipient,
    message_settings,
    html_body,
    attachment_paths,
    inline_image_paths,
    retry_attempts: int = 1,
    retry_backoff_seconds: float = 5.0,
    before_attempt: Callable[[int], None] | None = None,
) -> None:
    """Send email with automatic retry on temporary failures."""
    last_error = None
    for attempt in range(retry_attempts):
        if before_attempt is not None:
            before_attempt(attempt + 1)
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
            if attempt < retry_attempts - 1 and _is_retryable_send_error(error):
                time.sleep(retry_backoff_seconds)
            else:
                break
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
    body_text_override: str | None = None,
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
        _ = _collect_random_attachment_files(base_dir, config.message.random_attachments_folder)
    except CampaignError as error:
        errors.append(str(error))
        return PreflightReport(checks=checks, warnings=warnings, errors=errors)

    random_attachments = []
    if config.message.random_attachments_folder:
        try:
            random_attachments = _collect_random_attachment_files(base_dir, config.message.random_attachments_folder)
            checks.append(f"Случайных вложений в папке: {len(random_attachments)} ({config.message.random_attachments_folder})")
        except CampaignError as error:
            errors.append(str(error))
            return PreflightReport(checks=checks, warnings=warnings, errors=errors)

    checks.append(f"Вложений: {len(attachment_paths)}")
    checks.append(f"Inline-изображений: {len(inline_image_paths)}")
    checks.append(f"SMTP-аккаунтов: {len(config.smtp_accounts)}")

    first_recipient = recipients[0]
    preflight_content = dict(config.content)
    if body_text_override is not None:
        preflight_content["body_text"] = body_text_override
        preflight_content["message_text"] = body_text_override

    rendered = renderer.render(
        template_name=template_name,
        recipient=first_recipient,
        context={
            **preflight_content,
            "subject": config.message.subject,
            "inline_images": _build_inline_context(inline_image_paths),
            "smtp": {
                "host": config.smtp_accounts[0].host,
                "port": config.smtp_accounts[0].port,
                "username": config.smtp_accounts[0].username,
                "from_email": config.smtp_accounts[0].from_email,
                "from_name": config.smtp_accounts[0].from_name,
                "use_tls": config.smtp_accounts[0].use_tls,
                "use_ssl": config.smtp_accounts[0].use_ssl,
            },
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


def _copy_message_settings(message: MessageSettings, reply_to: str | None) -> MessageSettings:
    return MessageSettings(
        subject=message.subject,
        template=message.template,
        reply_to=reply_to,
        attachments=list(message.attachments),
        random_attachments_folder=message.random_attachments_folder,
        inline_images=dict(message.inline_images),
    )


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
        "subject",
        "template",
        "smtp_account",
        "proxy",
        "reply_to",
        "dry_run",
        "error",
    ]

    csv_exists = csv_path.exists()
    if csv_exists:
        try:
            with csv_path.open("r", encoding="utf-8-sig", newline="") as handle:
                reader = csv.DictReader(handle)
                existing_headers = reader.fieldnames or []
                needs_migration = "subject" not in existing_headers
                existing_rows = list(reader) if needs_migration else []

            if needs_migration:
                with csv_path.open("w", encoding="utf-8", newline="") as handle:
                    writer = csv.DictWriter(handle, fieldnames=headers)
                    writer.writeheader()
                    for row in existing_rows:
                        writer.writerow({key: row.get(key, "") for key in headers})
        except (OSError, csv.Error):
            pass

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


def _load_reply_to_list(base_dir: Path) -> list[str]:
    candidate = base_dir / "config" / "replyto_emails.txt"
    if not candidate.exists():
        return []
    emails: list[str] = []
    with candidate.open("r", encoding="utf-8") as handle:
        for line in handle:
            email = line.strip()
            if email and "@" in email:
                emails.append(email)
    return emails


def _collect_random_attachment_files(base_dir: Path, folder_path: str | None) -> list[Path]:
    if not folder_path:
        return []

    folder = (base_dir / folder_path).resolve()
    if not folder.exists() or not folder.is_dir():
        raise CampaignError(f"Папка случайных вложений не найдена или не папка: {folder}")

    allowed_ext = {".jpg", ".jpeg", ".png", ".gif", ".bmp", ".webp", ".svg"}
    files = [p for p in folder.iterdir() if p.is_file() and p.suffix.lower() in allowed_ext]
    if not files:
        raise CampaignError(f"В папке случайных вложений нет подходящих файлов: {folder}")
    return files


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
    body_text_override: str | None = None,
    recipient_email: str | None = None,
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
    if recipient_email:
        recipient_candidates = [r for r in recipients if r.email.lower() == recipient_email.lower()]
        if not recipient_candidates:
            raise CampaignError(f"Recipient не найден: {recipient_email}")
        recipient = recipient_candidates[0]
    else:
        recipient = recipients[0]

    html_body = renderer.render(
        template_name=template_name,
        recipient=recipient,
        context={
            **config.content,
            **({"body_text": body_text_override, "message_text": body_text_override} if body_text_override is not None else {}),
            "subject": config.message.subject,
            "inline_images": _build_inline_context(inline_image_paths),
            "smtp": {
                "host": config.smtp_accounts[0].host,
                "port": config.smtp_accounts[0].port,
                "username": config.smtp_accounts[0].username,
                "from_email": config.smtp_accounts[0].from_email,
                "from_name": config.smtp_accounts[0].from_name,
                "use_tls": config.smtp_accounts[0].use_tls,
                "use_ssl": config.smtp_accounts[0].use_ssl,
            },
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
    subject_override: str | None = None,
    subject_mode: str | None = None,
    subject_variants: list[str] | None = None,
    body_text_override: str | None = None,
    body_text_mode: str | None = None,
    body_text_variants: list[str] | None = None,
    random_attachments_folder_override: str | None = None,
    use_proxy: bool = True,
    proxy_file_override: str | None = None,
    delay_override: float | None = None,
    rate_limit_per_minute: int | None = None,
    retry_attempts: int | None = None,
    retry_backoff_seconds: float | None = None,
    parallel_smtp_enabled: bool | None = None,
    parallel_smtp_accounts: int | None = None,
    batch_interval_seconds: float | None = None,
    reply_to_override: str | None = None,
    reply_to_mode_override: str | None = None,
    stop_event: threading.Event | None = None,
    pause_event: threading.Event | None = None,
    runtime_overrides_getter: Callable[[], dict[str, object]] | None = None,
    progress_callback: Callable[[str], None] | None = None,
) -> CampaignSummary:
    try:
        config = load_config(config_path)
        recipients = load_recipients(recipients_path)
    except (ConfigError, RecipientsError) as error:
        raise CampaignError(str(error)) from error

    if rate_limit_per_minute is not None:
        config.delivery.rate_limit_per_minute = max(0, int(rate_limit_per_minute))
    if retry_attempts is not None:
        config.delivery.retry_attempts = max(1, int(retry_attempts))
    if retry_backoff_seconds is not None:
        config.delivery.retry_backoff_seconds = max(0.0, float(retry_backoff_seconds))

    if subject_override:
        config.message.subject = str(subject_override).strip()

    runtime_subject_mode_key = (subject_mode or "fixed").strip().lower()
    runtime_subject_variants = [str(item).strip() for item in (subject_variants or []) if str(item).strip()]
    runtime_subject_campaign_value = config.message.subject
    if runtime_subject_mode_key == "random_campaign" and runtime_subject_variants:
        runtime_subject_campaign_value = random.choice(runtime_subject_variants)

    text_mode = (body_text_mode or "fixed").strip().lower()
    text_variants = [str(item).strip() for item in (body_text_variants or []) if str(item).strip()]

    if text_mode == "random_campaign" and text_variants:
        selected_text = random.choice(text_variants)
        config.content["body_text"] = selected_text
        config.content["message_text"] = selected_text
    elif text_mode == "fixed" and body_text_override is not None:
        selected_text = str(body_text_override)
        config.content["body_text"] = selected_text
        config.content["message_text"] = selected_text

    if parallel_smtp_enabled is not None:
        config.delivery.parallel_smtp_enabled = bool(parallel_smtp_enabled)
    if parallel_smtp_accounts is not None:
        config.delivery.parallel_smtp_accounts = max(1, int(parallel_smtp_accounts))
    if batch_interval_seconds is not None:
        config.delivery.batch_interval_seconds = max(0.0, float(batch_interval_seconds))
    if reply_to_override is not None:
        normalized_reply_to = str(reply_to_override).strip()
        config.message.reply_to = normalized_reply_to or None

    reply_to_mode = (str(reply_to_mode_override).strip().lower() if reply_to_mode_override is not None else "random")
    if reply_to_mode not in {"random", "fixed"}:
        reply_to_mode = "random"

    if random_attachments_folder_override:
        config.message.random_attachments_folder = random_attachments_folder_override

    preflight = run_preflight(
        base_dir=base_dir,
        config_path=config_path,
        recipients_path=recipients_path,
        templates_path=templates_path,
        template_override=template_override,
        body_text_override=body_text_override,
    )
    if preflight.errors:
        raise CampaignError("; ".join(preflight.errors))

    renderer = TemplateRenderer(templates_path)
    available_templates = renderer.list_templates()
    runtime_template_name = template_override or config.message.template
    if runtime_template_name not in available_templates:
        raise CampaignError(f"Шаблон не найден: {runtime_template_name}")

    attachment_paths = _resolve_attachment_paths(base_dir, config.message.attachments)
    inline_image_paths = _resolve_inline_image_paths(base_dir, config.message.inline_images)
    random_attachment_files = []
    if config.message.random_attachments_folder:
        random_attachment_files = _collect_random_attachment_files(base_dir, config.message.random_attachments_folder)
    logger, log_path = _create_logger(base_dir, config.delivery.log_file)
    history_csv_path = _resolve_output_path(base_dir, config.delivery.history_csv)
    history_jsonl_path = _resolve_output_path(base_dir, config.delivery.history_jsonl)
    runtime_smtp_accounts = list(config.smtp_accounts)
    # --- Загрузка списка прокси ---
    proxy_file = (base_dir / (proxy_file_override or "config/proxies.txt"))
    proxies = load_proxies(proxy_file) if proxy_file.exists() else []

    # --- Reply-To list ---
    reply_to_candidates = _load_reply_to_list(base_dir)
    if not reply_to_candidates and config.message.reply_to:
        reply_to_candidates = [config.message.reply_to]
    if reply_to_mode == "fixed" and config.message.reply_to:
        reply_to_candidates = [config.message.reply_to]

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
            template_name=runtime_template_name,
            dedupe_template_scope=config.delivery.dedupe_template_scope,
            dedupe_history_days=config.delivery.dedupe_history_days,
        )

    def emit(message: str) -> None:
        if progress_callback is not None:
            progress_callback(message)

    def _reload_smtp_accounts_runtime() -> None:
        nonlocal runtime_smtp_accounts
        try:
            refreshed = load_config(config_path)
        except Exception as error:  # noqa: BLE001
            emit(f"⚠️ Не удалось перечитать SMTP-аккаунты: {error}")
            return

        new_accounts = list(getattr(refreshed, "smtp_accounts", []) or [])
        if not new_accounts:
            emit("⚠️ SMTP-аккаунты не обновлены: список пуст, продолжаем со старыми")
            return

        runtime_smtp_accounts = new_accounts
        emit(f"🔄 SMTP-аккаунты обновлены: {len(runtime_smtp_accounts)}")

    def _reload_template_and_subject_runtime() -> None:
        nonlocal runtime_template_name, runtime_subject_mode_key, runtime_subject_variants, runtime_subject_campaign_value
        if runtime_overrides_getter is None:
            return
        try:
            options = runtime_overrides_getter() or {}
        except Exception as error:  # noqa: BLE001
            emit(f"⚠️ Не удалось обновить шаблон/тему: {error}")
            return

        new_template = str(options.get("template_override") or "").strip()
        if new_template and new_template != runtime_template_name:
            current_templates = renderer.list_templates()
            if new_template in current_templates:
                runtime_template_name = new_template
                emit(f"🔄 Шаблон обновлён: {runtime_template_name}")
            else:
                emit(f"⚠️ Шаблон не найден, оставлен текущий: {new_template}")

        new_subject_mode = str(options.get("subject_mode") or runtime_subject_mode_key).strip().lower()
        if new_subject_mode not in {"fixed", "random_campaign", "random_recipient"}:
            new_subject_mode = runtime_subject_mode_key

        raw_variants = options.get("subject_variants")
        if isinstance(raw_variants, list):
            new_subject_variants = [str(item).strip() for item in raw_variants if str(item).strip()]
        else:
            new_subject_variants = list(runtime_subject_variants)

        new_fixed_subject = str(options.get("subject_override") or "").strip()

        runtime_subject_mode_key = new_subject_mode
        runtime_subject_variants = new_subject_variants
        if runtime_subject_mode_key == "random_campaign" and runtime_subject_variants:
            runtime_subject_campaign_value = random.choice(runtime_subject_variants)
            emit("🔄 Тема обновлена: режим random_campaign")
        elif runtime_subject_mode_key == "fixed" and new_fixed_subject:
            runtime_subject_campaign_value = new_fixed_subject
            emit("🔄 Тема обновлена: фиксированная")
        elif runtime_subject_mode_key == "random_recipient":
            emit("🔄 Тема обновлена: режим random_recipient")

    def _wait_if_paused_or_stopped() -> bool:
        announced = False
        while pause_event is not None and pause_event.is_set():
            if stop_event is not None and stop_event.is_set():
                return False
            if not announced:
                emit("⏸ Рассылка на паузе")
                announced = True
            time.sleep(0.2)
        if announced:
            _reload_smtp_accounts_runtime()
            _reload_template_and_subject_runtime()
            emit("▶ Рассылка продолжена")
        return not (stop_event is not None and stop_event.is_set())

    def _sleep_with_controls(seconds: float) -> bool:
        if seconds <= 0:
            return True
        end_at = time.monotonic() + seconds
        while True:
            if stop_event is not None and stop_event.is_set():
                return False
            if not _wait_if_paused_or_stopped():
                return False
            now = time.monotonic()
            if now >= end_at:
                return True
            time.sleep(min(0.2, end_at - now))

    def _send_task(recipient_index: int, recipient: Recipient, mailer_settings, proxy_list, reply_to_candidates):
        nonlocal successful, failed, processed
        local_mailer = SMTPMailer(deepcopy(mailer_settings))

        proxy_info = "none"
        send_attempts = 0

        def _apply_proxy_for_attempt(attempt_no: int) -> None:
            nonlocal proxy_info, send_attempts
            send_attempts = attempt_no
            if use_proxy:
                if proxy_list:
                    proxy = pick_random_proxy(proxy_list)
                    local_mailer.settings.proxy_host = proxy.get("proxy_host")
                    local_mailer.settings.proxy_port = proxy.get("proxy_port")
                    local_mailer.settings.proxy_type = proxy.get("proxy_type")
                    local_mailer.settings.proxy_user = proxy.get("proxy_user")
                    local_mailer.settings.proxy_pass = proxy.get("proxy_pass")
                    proxy_info = f"{local_mailer.settings.proxy_type}:{local_mailer.settings.proxy_host}:{local_mailer.settings.proxy_port}"
                elif local_mailer.settings.proxy_host and local_mailer.settings.proxy_port and local_mailer.settings.proxy_type:
                    proxy_info = (
                        f"{local_mailer.settings.proxy_type}:"
                        f"{local_mailer.settings.proxy_host}:"
                        f"{local_mailer.settings.proxy_port}"
                    )
                else:
                    proxy_info = "enabled-but-not-configured"
            else:
                local_mailer.settings.proxy_host = None
                local_mailer.settings.proxy_port = None
                local_mailer.settings.proxy_type = None
                local_mailer.settings.proxy_user = None
                local_mailer.settings.proxy_pass = None
                proxy_info = "none"

        _apply_proxy_for_attempt(1)

        selected_reply_to = random.choice(reply_to_candidates) if reply_to_candidates else config.message.reply_to
        message_settings = _copy_message_settings(config.message, selected_reply_to)
        if runtime_subject_mode_key == "random_recipient" and runtime_subject_variants:
            message_settings.subject = random.choice(runtime_subject_variants)
        else:
            message_settings.subject = runtime_subject_campaign_value

        smtp_account = local_mailer.settings.from_email
        recipient_key = recipient.email.strip().lower()
        recipient_text_context: dict[str, str] = {}
        if text_mode == "random_recipient" and text_variants:
            recipient_text = random.choice(text_variants)
            recipient_text_context = {
                "body_text": recipient_text,
                "message_text": recipient_text,
            }

        html_body = renderer.render(
            template_name=runtime_template_name,
            recipient=recipient,
            context={
                **config.content,
                **recipient_text_context,
                "subject": message_settings.subject,
                "inline_images": _build_inline_context(inline_image_paths),
                "smtp": {
                    "host": local_mailer.settings.host,
                    "port": local_mailer.settings.port,
                    "username": local_mailer.settings.username,
                    "from_email": local_mailer.settings.from_email,
                    "from_name": local_mailer.settings.from_name,
                    "use_tls": local_mailer.settings.use_tls,
                    "use_ssl": local_mailer.settings.use_ssl,
                },
            },
        )

        recipient_attachment_paths = list(attachment_paths)
        if random_attachment_files:
            recipient_attachment_paths.append(random.choice(random_attachment_files))

        reply_to_info = selected_reply_to or "not-set"

        if dry_run:
            message = (
                f"[DRY-RUN] {recipient_index}/{len(recipients)} подготовлено для {recipient.email} "
                f"(subject={message_settings.subject}, smtp={smtp_account}, proxy={proxy_info}, reply_to={reply_to_info})"
            )
            status = "dry-run"
            error_msg = ""
            success_delta = 1
            fail_delta = 0
            sent = False
        else:
            try:
                _send_with_retry(
                    mailer=local_mailer,
                    recipient=recipient,
                    message_settings=message_settings,
                    html_body=html_body,
                    attachment_paths=recipient_attachment_paths,
                    inline_image_paths=inline_image_paths,
                    retry_attempts=config.delivery.retry_attempts,
                    retry_backoff_seconds=config.delivery.retry_backoff_seconds,
                    before_attempt=_apply_proxy_for_attempt,
                )
                message = (
                    f"[OK] {recipient_index}/{len(recipients)} отправлено: {recipient.email} "
                    f"через {smtp_account} subject={message_settings.subject} proxy={proxy_info} reply_to={reply_to_info} attempts={send_attempts}"
                )
                status = "sent"
                error_msg = ""
                success_delta = 1
                fail_delta = 0
                sent = True
            except Exception as error:
                reason_ru = _humanize_error_ru(error)
                message = (
                    f"[ERROR] {recipient_index}/{len(recipients)} {recipient.email}: {reason_ru} "
                    f"(subject={message_settings.subject}, smtp={smtp_account}, proxy={proxy_info}, reply_to={reply_to_info}, attempts={send_attempts})"
                )
                status = "error"
                error_msg = reason_ru
                success_delta = 0
                fail_delta = 1
                sent = False

        return {
            "message": message,
            "status": status,
            "error_msg": error_msg,
            "success_delta": success_delta,
            "fail_delta": fail_delta,
            "sent": sent,
            "recipient": recipient.email,
            "subject": message_settings.subject,
            "smtp_account": smtp_account,
            "proxy": proxy_info,
            "reply_to": selected_reply_to or "",
            "dry_run": str(dry_run).lower(),
        }

    logger.info(
        "Старт кампании: dry_run=%s, recipients=%s, template=%s, attachments=%s, inline_images=%s, smtp_accounts=%s, dedupe=%s, reply_to_mode=%s, reply_to_candidates=%s, config_reply_to=%s",
        dry_run,
        len(recipients),
        runtime_template_name,
        len(attachment_paths),
        len(inline_image_paths),
        len(runtime_smtp_accounts),
        config.delivery.skip_previously_sent,
        reply_to_mode,
        len(reply_to_candidates),
        config.message.reply_to or "",
    )
    emit(f"Лог: {log_path}")
    emit(f"История CSV: {history_csv_path}")
    emit(f"История JSONL: {history_jsonl_path}")

    target_recipients = []
    for index, recipient in enumerate(recipients, start=1):
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
                    "subject": config.message.subject,
                    "template": runtime_template_name,
                    "smtp_account": "",
                    "proxy": "disabled" if not use_proxy else "n/a",
                    "reply_to": config.message.reply_to or "",
                    "dry_run": str(dry_run).lower(),
                    "error": "already_sent",
                },
            )
            continue
        target_recipients.append((index, recipient))

    # Параллельная отправка должна включаться только явным флагом.
    # Иначе пользователь видит параллельный режим даже при выключенном чекбоксе.
    use_parallel = bool(config.delivery.parallel_smtp_enabled) and (len(runtime_smtp_accounts) > 1)
    if use_parallel:
        batch_size = min(config.delivery.parallel_smtp_accounts, len(runtime_smtp_accounts)) if len(runtime_smtp_accounts) > 0 else 1
        batch_interval = max(0.0, config.delivery.batch_interval_seconds)

        for batch_start in range(0, len(target_recipients), batch_size):
            if not _wait_if_paused_or_stopped():
                emit("⏹ Рассылка остановлена пользователем")
                break
            batch = target_recipients[batch_start : batch_start + batch_size]
            futures = {}
            with ThreadPoolExecutor(max_workers=batch_size) as executor:
                for offset, (index, recipient) in enumerate(batch):
                    if not runtime_smtp_accounts:
                        emit("⏹ Нет SMTP-аккаунтов для продолжения рассылки")
                        break
                    mailer_settings = deepcopy(runtime_smtp_accounts[(batch_start + offset) % len(runtime_smtp_accounts)])
                    futures[executor.submit(_send_task, index, recipient, mailer_settings, proxies, reply_to_candidates)] = (index, recipient)

                for future in as_completed(futures):
                    result = future.result()
                    successful += result["success_delta"]
                    failed += result["fail_delta"]
                    processed += 1
                    logger.info(result["message"])
                    emit(result["message"])
                    _append_history(
                        csv_path=history_csv_path,
                        jsonl_path=history_jsonl_path,
                        record={
                            "timestamp": datetime.now(timezone.utc).isoformat(),
                            "recipient": result["recipient"],
                            "status": result["status"],
                            "subject": result["subject"],
                            "template": runtime_template_name,
                            "smtp_account": result["smtp_account"],
                            "proxy": result["proxy"],
                            "reply_to": result["reply_to"],
                            "dry_run": result["dry_run"],
                            "error": result["error_msg"],
                        },
                    )

            if batch_interval > 0 and batch_start + batch_size < len(target_recipients):
                emit(f"Пауза {batch_interval:.2f} сек.")
                if not _sleep_with_controls(batch_interval):
                    emit("⏹ Рассылка остановлена пользователем")
                    break
    else:
        for index, recipient in target_recipients:
            if not _wait_if_paused_or_stopped():
                emit("⏹ Рассылка остановлена пользователем")
                break
            if not runtime_smtp_accounts:
                emit("⏹ Нет SMTP-аккаунтов для продолжения рассылки")
                break
            mailer_settings = deepcopy(runtime_smtp_accounts[(index - 1) % len(runtime_smtp_accounts)])
            result = _send_task(index, recipient, mailer_settings, proxies, reply_to_candidates)
            successful += result["success_delta"]
            failed += result["fail_delta"]
            processed += 1
            logger.info(result["message"])
            emit(result["message"])
            _append_history(
                csv_path=history_csv_path,
                jsonl_path=history_jsonl_path,
                record={
                    "timestamp": datetime.now(timezone.utc).isoformat(),
                    "recipient": result["recipient"],
                    "status": result["status"],
                    "subject": result["subject"],
                    "template": runtime_template_name,
                    "smtp_account": result["smtp_account"],
                    "proxy": result["proxy"],
                    "reply_to": result["reply_to"],
                    "dry_run": result["dry_run"],
                    "error": result["error_msg"],
                },
            )

            if final_delay > 0 and index < len(recipients):
                emit(f"Пауза {final_delay:.2f} сек.")
                if not _sleep_with_controls(final_delay):
                    emit("⏹ Рассылка остановлена пользователем")
                    break

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
