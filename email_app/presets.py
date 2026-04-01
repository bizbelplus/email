from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import yaml


class PresetError(ValueError):
    """Raised when preset data is invalid."""


@dataclass
class CampaignPreset:
    config: str = "config/settings.yaml"
    recipients: str = "recipients.csv"
    templates: str = "templates"
    template: str | None = None
    delay_seconds: float | None = None
    dry_run: bool = True
    attachments_folder: str | None = None
    use_proxy: bool = True
    proxy_file: str | None = None
    rate_limit_per_minute: int | None = None
    retry_attempts: int | None = None
    retry_backoff_seconds: float | None = None
    parallel_smtp_enabled: bool | None = None
    parallel_smtp_accounts: int | None = None
    batch_interval_seconds: float | None = None
    reply_to: str | None = None
    reply_to_mode: str | None = None


def load_preset(path: str | Path) -> CampaignPreset:
    preset_path = Path(path)
    if not preset_path.exists():
        raise PresetError(f"Файл пресета не найден: {preset_path}")

    raw_data = yaml.safe_load(preset_path.read_text(encoding="utf-8")) or {}
    if not isinstance(raw_data, dict):
        raise PresetError("Пресет должен быть объектом YAML")

    def _to_int(value):
        if value in (None, ""):
            return None
        return int(value)

    def _to_float(value):
        if value in (None, ""):
            return None
        return float(value)

    return CampaignPreset(
        config=str(raw_data.get("config", "config/settings.yaml")),
        recipients=str(raw_data.get("recipients", "recipients.csv")),
        templates=str(raw_data.get("templates", "templates")),
        template=(str(raw_data["template"]) if raw_data.get("template") else None),
        delay_seconds=(
            float(raw_data["delay_seconds"])
            if raw_data.get("delay_seconds") not in (None, "")
            else None
        ),
        dry_run=bool(raw_data.get("dry_run", True)),
        attachments_folder=(str(raw_data["attachments_folder"]) if raw_data.get("attachments_folder") else None),
        use_proxy=bool(raw_data.get("use_proxy", True)),
        proxy_file=(str(raw_data["proxy_file"]) if raw_data.get("proxy_file") else None),
        rate_limit_per_minute=_to_int(raw_data.get("rate_limit_per_minute")),
        retry_attempts=_to_int(raw_data.get("retry_attempts")),
        retry_backoff_seconds=_to_float(raw_data.get("retry_backoff_seconds")),
        parallel_smtp_enabled=(
            bool(raw_data.get("parallel_smtp_enabled"))
            if raw_data.get("parallel_smtp_enabled") is not None
            else None
        ),
        parallel_smtp_accounts=_to_int(raw_data.get("parallel_smtp_accounts")),
        batch_interval_seconds=_to_float(raw_data.get("batch_interval_seconds")),
        reply_to=(str(raw_data["reply_to"]) if raw_data.get("reply_to") else None),
        reply_to_mode=(str(raw_data["reply_to_mode"]) if raw_data.get("reply_to_mode") else None),
    )


def save_preset(path: str | Path, preset: CampaignPreset) -> Path:
    preset_path = Path(path)
    preset_path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "config": preset.config,
        "recipients": preset.recipients,
        "templates": preset.templates,
        "template": preset.template,
        "delay_seconds": preset.delay_seconds,
        "dry_run": preset.dry_run,
        "attachments_folder": preset.attachments_folder,
        "use_proxy": preset.use_proxy,
        "proxy_file": preset.proxy_file,
        "rate_limit_per_minute": preset.rate_limit_per_minute,
        "retry_attempts": preset.retry_attempts,
        "retry_backoff_seconds": preset.retry_backoff_seconds,
        "parallel_smtp_enabled": preset.parallel_smtp_enabled,
        "parallel_smtp_accounts": preset.parallel_smtp_accounts,
        "batch_interval_seconds": preset.batch_interval_seconds,
        "reply_to": preset.reply_to,
        "reply_to_mode": preset.reply_to_mode,
    }
    preset_path.write_text(
        yaml.safe_dump(payload, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return preset_path
