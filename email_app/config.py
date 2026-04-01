from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

import yaml

from .models import AppConfig, DeliverySettings, MessageSettings, SMTPSettings


class ConfigError(ValueError):
    """Raised when config data is invalid."""


def _parse_bool(value: Any, default: bool) -> bool:
    if value in (None, ""):
        return default
    if isinstance(value, bool):
        return value
    return str(value).strip().lower() in {"1", "true", "yes", "on"}


def _require(mapping: dict[str, Any], key: str) -> Any:
    value = mapping.get(key)
    if value in (None, ""):
        raise ConfigError(f"Отсутствует обязательное поле: {key}")
    return value


def _build_smtp_settings(mapping: dict[str, Any]) -> SMTPSettings:

    smtp = SMTPSettings(
        host=str(_require(mapping, "host")),
        port=int(_require(mapping, "port")),
        username=str(_require(mapping, "username")),
        password=str(_require(mapping, "password")),
        from_email=str(_require(mapping, "from_email")),
        from_name=str(_require(mapping, "from_name")),
        use_tls=_parse_bool(mapping.get("use_tls", True), True),
        use_ssl=_parse_bool(mapping.get("use_ssl", False), False),
        timeout_seconds=int(mapping.get("timeout_seconds", 30)),
        proxy_host=mapping.get("proxy_host"),
        proxy_port=int(mapping["proxy_port"]) if mapping.get("proxy_port") else None,
        proxy_type=mapping.get("proxy_type"),
        proxy_user=mapping.get("proxy_user"),
        proxy_pass=mapping.get("proxy_pass"),
    )

    if smtp.use_tls and smtp.use_ssl:
        raise ConfigError("Нельзя одновременно включить use_tls и use_ssl")

    return smtp


def _load_smtp_accounts_csv(accounts_file: Path) -> list[SMTPSettings]:
    with accounts_file.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        accounts: list[SMTPSettings] = []
        for row in reader:
            normalized = {key: (value or "").strip() for key, value in row.items() if key}
            if not any(normalized.values()):
                continue
            normalized["use_tls"] = _parse_bool(normalized.get("use_tls", "true"), True)
            normalized["use_ssl"] = _parse_bool(normalized.get("use_ssl", "false"), False)
            normalized["timeout_seconds"] = int(normalized.get("timeout_seconds", "30") or 30)
            accounts.append(_build_smtp_settings(normalized))
    return accounts


def _load_smtp_accounts_txt(accounts_file: Path) -> list[SMTPSettings]:
    accounts: list[SMTPSettings] = []
    with accounts_file.open("r", encoding="utf-8-sig") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue

            parts = [item.strip() for item in line.split("|")]
            if len(parts) < 6:
                raise ConfigError(
                    "TXT-файл SMTP-аккаунтов должен содержать минимум 6 полей через '|': "
                    "host|port|username|password|from_email|from_name"
                    f" (строка {line_number})"
                )

            mapping: dict[str, Any] = {
                "host": parts[0],
                "port": parts[1],
                "username": parts[2],
                "password": parts[3],
                "from_email": parts[4],
                "from_name": parts[5],
                "use_tls": _parse_bool(parts[6], True) if len(parts) > 6 else True,
                "use_ssl": _parse_bool(parts[7], False) if len(parts) > 7 else False,
                "timeout_seconds": int(parts[8]) if len(parts) > 8 and parts[8] else 30,
            }
            accounts.append(_build_smtp_settings(mapping))
    return accounts


def _load_smtp_accounts(accounts_file: Path) -> list[SMTPSettings]:
    if not accounts_file.exists():
        raise ConfigError(f"Файл SMTP-аккаунтов не найден: {accounts_file}")

    if accounts_file.suffix.lower() == ".txt":
        accounts = _load_smtp_accounts_txt(accounts_file)
    else:
        accounts = _load_smtp_accounts_csv(accounts_file)

    if not accounts:
        raise ConfigError("Файл SMTP-аккаунтов пуст")

    return accounts


def load_config(config_path: str | Path) -> AppConfig:
    path = Path(config_path)
    if not path.exists():
        raise ConfigError(f"Файл конфигурации не найден: {path}")

    raw_data = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    smtp_raw = raw_data.get("smtp") or {}
    message_raw = raw_data.get("message") or {}
    content_raw = raw_data.get("content") or {}
    accounts_file_value = smtp_raw.get("accounts_file")
    if accounts_file_value:
        smtp_accounts = _load_smtp_accounts(path.parent.parent / str(accounts_file_value))
    else:
        smtp_accounts = [_build_smtp_settings(smtp_raw)]

    message = MessageSettings(
        subject=str(_require(message_raw, "subject")),
        template=str(message_raw.get("template")) if message_raw.get("template") else None,
        reply_to=(str(message_raw["reply_to"]) if message_raw.get("reply_to") else None),
        attachments=[str(item) for item in message_raw.get("attachments", [])],
        random_attachments_folder=str(message_raw.get("random_attachments_folder")) if message_raw.get("random_attachments_folder") else None,
        inline_images={
            str(key): str(value)
            for key, value in (message_raw.get("inline_images") or {}).items()
        },
    )

    delivery_raw = raw_data.get("delivery") or {}
    delivery = DeliverySettings(
        delay_seconds=float(delivery_raw.get("delay_seconds", 0.0)),
        log_file=str(delivery_raw.get("log_file", "logs/email_app.log")),
        history_csv=str(delivery_raw.get("history_csv", "history/email_history.csv")),
        history_jsonl=str(delivery_raw.get("history_jsonl", "history/email_history.jsonl")),
        skip_previously_sent=bool(delivery_raw.get("skip_previously_sent", False)),
        dedupe_template_scope=bool(delivery_raw.get("dedupe_template_scope", True)),
        dedupe_history_days=int(delivery_raw.get("dedupe_history_days", 30)),
        scheduled_time=str(delivery_raw["scheduled_time"]) if delivery_raw.get("scheduled_time") else None,
        rate_limit_per_minute=int(delivery_raw.get("rate_limit_per_minute")) if delivery_raw.get("rate_limit_per_minute") else None,
        parallel_smtp_enabled=bool(delivery_raw.get("parallel_smtp_enabled", False)),
        parallel_smtp_accounts=int(delivery_raw.get("parallel_smtp_accounts", 1)),
        batch_interval_seconds=float(delivery_raw.get("batch_interval_seconds", 0.0)),
        retry_attempts=int(delivery_raw.get("retry_attempts", 1)),
        retry_backoff_seconds=float(delivery_raw.get("retry_backoff_seconds", 5.0)),
    )

    if not isinstance(content_raw, dict):
        raise ConfigError("Секция content должна быть объектом")

    return AppConfig(
        smtp_accounts=smtp_accounts,
        message=message,
        delivery=delivery,
        content=content_raw,
    )
