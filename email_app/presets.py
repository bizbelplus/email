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


def load_preset(path: str | Path) -> CampaignPreset:
    preset_path = Path(path)
    if not preset_path.exists():
        raise PresetError(f"Файл пресета не найден: {preset_path}")

    raw_data = yaml.safe_load(preset_path.read_text(encoding="utf-8")) or {}
    if not isinstance(raw_data, dict):
        raise PresetError("Пресет должен быть объектом YAML")

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
    }
    preset_path.write_text(
        yaml.safe_dump(payload, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return preset_path
