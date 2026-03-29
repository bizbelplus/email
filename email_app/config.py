from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

import yaml

from .models import AppConfig, DeliverySettings, MessageSettings, SMTPSettings


class ConfigError(ValueError):
    """Raised when config data is invalid."""


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
        use_tls=bool(mapping.get("use_tls", True)),
        use_ssl=bool(mapping.get("use_ssl", False)),
        timeout_seconds=int(mapping.get("timeout_seconds", 30)),
    )

    if smtp.use_tls and smtp.use_ssl:
        raise ConfigError("Нельзя одновременно включить use_tls и use_ssl")

    return smtp


def _load_smtp_accounts(accounts_file: Path) -> list[SMTPSettings]:
    if not accounts_file.exists():
        raise ConfigError(f"CSV-файл SMTP-аккаунтов не найден: {accounts_file}")

    with accounts_file.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        accounts: list[SMTPSettings] = []
        for row in reader:
            normalized = {key: (value or "").strip() for key, value in row.items() if key}
            if not any(normalized.values()):
                continue
            normalized["use_tls"] = normalized.get("use_tls", "true").lower() in {"1", "true", "yes", "on"}
            normalized["use_ssl"] = normalized.get("use_ssl", "false").lower() in {"1", "true", "yes", "on"}
            normalized["timeout_seconds"] = int(normalized.get("timeout_seconds", "30") or 30)
            accounts.append(_build_smtp_settings(normalized))

    if not accounts:
        raise ConfigError("CSV-файл SMTP-аккаунтов пуст")

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
        template=str(_require(message_raw, "template")),
        reply_to=(str(message_raw["reply_to"]) if message_raw.get("reply_to") else None),
        attachments=[str(item) for item in message_raw.get("attachments", [])],
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
    )

    if not isinstance(content_raw, dict):
        raise ConfigError("Секция content должна быть объектом")

    return AppConfig(
        smtp_accounts=smtp_accounts,
        message=message,
        delivery=delivery,
        content=content_raw,
    )
