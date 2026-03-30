from __future__ import annotations

import csv
import json
from dataclasses import dataclass
from pathlib import Path
from typing import Callable, Union, Optional

from .presets import CampaignPreset
from .service import CampaignSummary, run_campaign


class CampaignQueueError(ValueError):
    """Raised when queue file data is invalid."""


@dataclass
class QueueSummary:
    campaigns_total: int
    campaigns_completed: int
    total_processed: int
    total_successful: int
    total_failed: int


def load_campaign_queue(path: Union[str, Path]) -> list[CampaignPreset]:
    queue_path = Path(path)
    if not queue_path.exists():
        raise CampaignQueueError(f"Файл очереди не найден: {queue_path}")

    if queue_path.suffix.lower() == ".csv":
        return _load_campaign_queue_csv(queue_path)
    return _load_campaign_queue_json(queue_path)


def _load_campaign_queue_json(queue_path: Path) -> list[CampaignPreset]:

    raw_data = json.loads(queue_path.read_text(encoding="utf-8"))
    if not isinstance(raw_data, list):
        raise CampaignQueueError("JSON-очередь должна быть массивом кампаний")

    presets: list[CampaignPreset] = []
    for index, item in enumerate(raw_data, start=1):
        if not isinstance(item, dict):
            raise CampaignQueueError(f"Элемент очереди #{index} должен быть объектом")
        presets.append(
            CampaignPreset(
                config=str(item.get("config", "config/settings.yaml")),
                recipients=str(item.get("recipients", "recipients.csv")),
                templates=str(item.get("templates", "templates")),
                template=(str(item["template"]) if item.get("template") else None),
                delay_seconds=(
                    float(item["delay_seconds"])
                    if item.get("delay_seconds") not in (None, "")
                    else None
                ),
                dry_run=bool(item.get("dry_run", True)),
            )
        )

    if not presets:
        raise CampaignQueueError("Очередь кампаний пуста")

    return presets


def _load_campaign_queue_csv(queue_path: Path) -> list[CampaignPreset]:
    with queue_path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        presets = [
            CampaignPreset(
                config=str(row.get("config", "config/settings.yaml")),
                recipients=str(row.get("recipients", "recipients.csv")),
                templates=str(row.get("templates", "templates")),
                template=(str(row["template"]) if row.get("template") else None),
                delay_seconds=(float(row["delay_seconds"]) if row.get("delay_seconds") else None),
                dry_run=str(row.get("dry_run", "true")).strip().lower() in {"1", "true", "yes", "on"},
            )
            for row in reader
        ]

    if not presets:
        raise CampaignQueueError("CSV-очередь кампаний пуста")

    return presets


def save_campaign_queue(path: Union[str, Path], campaigns: list[CampaignPreset]) -> Path:
    queue_path = Path(path)
    queue_path.parent.mkdir(parents=True, exist_ok=True)
    if queue_path.suffix.lower() == ".csv":
        return _save_campaign_queue_csv(queue_path, campaigns)
    return _save_campaign_queue_json(queue_path, campaigns)


def _save_campaign_queue_json(queue_path: Path, campaigns: list[CampaignPreset]) -> Path:
    payload = [
        {
            "config": campaign.config,
            "recipients": campaign.recipients,
            "templates": campaign.templates,
            "template": campaign.template,
            "delay_seconds": campaign.delay_seconds,
            "dry_run": campaign.dry_run,
        }
        for campaign in campaigns
    ]
    queue_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return queue_path


def _save_campaign_queue_csv(queue_path: Path, campaigns: list[CampaignPreset]) -> Path:
    with queue_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=["config", "recipients", "templates", "template", "delay_seconds", "dry_run"],
        )
        writer.writeheader()
        for campaign in campaigns:
            writer.writerow(
                {
                    "config": campaign.config,
                    "recipients": campaign.recipients,
                    "templates": campaign.templates,
                    "template": campaign.template or "",
                    "delay_seconds": campaign.delay_seconds if campaign.delay_seconds is not None else "",
                    "dry_run": str(campaign.dry_run).lower(),
                }
            )
    return queue_path


def run_campaign_queue(
    *,
    base_dir: Path,
    campaigns: list[CampaignPreset],
    progress_callback: Optional[Callable[[str], None]] = None,
) -> QueueSummary:
    def emit(message: str) -> None:
        if progress_callback is not None:
            progress_callback(message)

    completed = 0
    total_processed = 0
    total_successful = 0
    total_failed = 0

    for index, campaign in enumerate(campaigns, start=1):
        emit(
            f"[QUEUE] Запуск кампании {index}/{len(campaigns)} | "
            f"config={campaign.config} | recipients={campaign.recipients}"
        )
        summary: CampaignSummary = run_campaign(
            base_dir=base_dir,
            config_path=base_dir / campaign.config,
            recipients_path=base_dir / campaign.recipients,
            templates_path=base_dir / campaign.templates,
            dry_run=campaign.dry_run,
            template_override=campaign.template,
            delay_override=campaign.delay_seconds,
            progress_callback=progress_callback,
        )
        completed += 1
        total_processed += summary.processed
        total_successful += summary.successful
        total_failed += summary.failed

    return QueueSummary(
        campaigns_total=len(campaigns),
        campaigns_completed=completed,
        total_processed=total_processed,
        total_successful=total_successful,
        total_failed=total_failed,
    )
