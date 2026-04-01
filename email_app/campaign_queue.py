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
                attachments_folder=(str(item["attachments_folder"]) if item.get("attachments_folder") else None),
                use_proxy=bool(item.get("use_proxy", True)),
                proxy_file=(str(item["proxy_file"]) if item.get("proxy_file") else None),
                rate_limit_per_minute=(int(item["rate_limit_per_minute"]) if item.get("rate_limit_per_minute") not in (None, "") else None),
                retry_attempts=(int(item["retry_attempts"]) if item.get("retry_attempts") not in (None, "") else None),
                retry_backoff_seconds=(float(item["retry_backoff_seconds"]) if item.get("retry_backoff_seconds") not in (None, "") else None),
                parallel_smtp_enabled=(bool(item.get("parallel_smtp_enabled")) if item.get("parallel_smtp_enabled") is not None else None),
                parallel_smtp_accounts=(int(item["parallel_smtp_accounts"]) if item.get("parallel_smtp_accounts") not in (None, "") else None),
                batch_interval_seconds=(float(item["batch_interval_seconds"]) if item.get("batch_interval_seconds") not in (None, "") else None),
                reply_to=(str(item["reply_to"]) if item.get("reply_to") else None),
                reply_to_mode=(str(item["reply_to_mode"]) if item.get("reply_to_mode") else None),
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
                attachments_folder=(str(row["attachments_folder"]) if row.get("attachments_folder") else None),
                use_proxy=str(row.get("use_proxy", "true")).strip().lower() in {"1", "true", "yes", "on"},
                proxy_file=(str(row["proxy_file"]) if row.get("proxy_file") else None),
                rate_limit_per_minute=(int(row["rate_limit_per_minute"]) if row.get("rate_limit_per_minute") else None),
                retry_attempts=(int(row["retry_attempts"]) if row.get("retry_attempts") else None),
                retry_backoff_seconds=(float(row["retry_backoff_seconds"]) if row.get("retry_backoff_seconds") else None),
                parallel_smtp_enabled=(str(row["parallel_smtp_enabled"]).strip().lower() in {"1", "true", "yes", "on"} if row.get("parallel_smtp_enabled") else None),
                parallel_smtp_accounts=(int(row["parallel_smtp_accounts"]) if row.get("parallel_smtp_accounts") else None),
                batch_interval_seconds=(float(row["batch_interval_seconds"]) if row.get("batch_interval_seconds") else None),
                reply_to=(str(row["reply_to"]) if row.get("reply_to") else None),
                reply_to_mode=(str(row["reply_to_mode"]) if row.get("reply_to_mode") else None),
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
            "attachments_folder": campaign.attachments_folder,
            "use_proxy": campaign.use_proxy,
            "proxy_file": campaign.proxy_file,
            "rate_limit_per_minute": campaign.rate_limit_per_minute,
            "retry_attempts": campaign.retry_attempts,
            "retry_backoff_seconds": campaign.retry_backoff_seconds,
            "parallel_smtp_enabled": campaign.parallel_smtp_enabled,
            "parallel_smtp_accounts": campaign.parallel_smtp_accounts,
            "batch_interval_seconds": campaign.batch_interval_seconds,
            "reply_to": campaign.reply_to,
            "reply_to_mode": campaign.reply_to_mode,
        }
        for campaign in campaigns
    ]
    queue_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return queue_path


def _save_campaign_queue_csv(queue_path: Path, campaigns: list[CampaignPreset]) -> Path:
    with queue_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "config",
                "recipients",
                "templates",
                "template",
                "delay_seconds",
                "dry_run",
                "attachments_folder",
                "use_proxy",
                "proxy_file",
                "rate_limit_per_minute",
                "retry_attempts",
                "retry_backoff_seconds",
                "parallel_smtp_enabled",
                "parallel_smtp_accounts",
                "batch_interval_seconds",
                "reply_to",
                "reply_to_mode",
            ],
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
                    "attachments_folder": campaign.attachments_folder or "",
                    "use_proxy": str(campaign.use_proxy).lower(),
                    "proxy_file": campaign.proxy_file or "",
                    "rate_limit_per_minute": campaign.rate_limit_per_minute if campaign.rate_limit_per_minute is not None else "",
                    "retry_attempts": campaign.retry_attempts if campaign.retry_attempts is not None else "",
                    "retry_backoff_seconds": campaign.retry_backoff_seconds if campaign.retry_backoff_seconds is not None else "",
                    "parallel_smtp_enabled": "" if campaign.parallel_smtp_enabled is None else str(campaign.parallel_smtp_enabled).lower(),
                    "parallel_smtp_accounts": campaign.parallel_smtp_accounts if campaign.parallel_smtp_accounts is not None else "",
                    "batch_interval_seconds": campaign.batch_interval_seconds if campaign.batch_interval_seconds is not None else "",
                    "reply_to": campaign.reply_to or "",
                    "reply_to_mode": campaign.reply_to_mode or "",
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
            random_attachments_folder_override=campaign.attachments_folder,
            use_proxy=campaign.use_proxy,
            proxy_file_override=campaign.proxy_file,
            delay_override=campaign.delay_seconds,
            rate_limit_per_minute=campaign.rate_limit_per_minute,
            retry_attempts=campaign.retry_attempts,
            retry_backoff_seconds=campaign.retry_backoff_seconds,
            parallel_smtp_enabled=campaign.parallel_smtp_enabled,
            parallel_smtp_accounts=campaign.parallel_smtp_accounts,
            batch_interval_seconds=campaign.batch_interval_seconds,
            reply_to_override=campaign.reply_to,
            reply_to_mode_override=campaign.reply_to_mode,
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
