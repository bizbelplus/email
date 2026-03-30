from __future__ import annotations

import csv
import json
from collections import Counter
from dataclasses import dataclass
from pathlib import Path


class StatsError(ValueError):
    """Raised when history data is invalid."""


@dataclass
class HistoryStats:
    total: int
    sent: int
    dry_run: int
    errors: int
    unique_recipients: int
    top_templates: list[tuple[str, int]]
    top_smtp_accounts: list[tuple[str, int]]


@dataclass
class HistoryRecord:
    timestamp: str
    recipient: str
    status: str
    template: str
    smtp_account: str
    dry_run: str
    error: str


def load_history_records(history_csv_path: str | Path) -> list[HistoryRecord]:
    path = Path(history_csv_path)
    if not path.exists():
        raise StatsError(f"Файл истории не найден: {path}")

    with path.open("r", encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        return [
            HistoryRecord(
                timestamp=row.get("timestamp", ""),
                recipient=row.get("recipient", ""),
                status=row.get("status", ""),
                template=row.get("template", ""),
                smtp_account=row.get("smtp_account", ""),
                dry_run=row.get("dry_run", ""),
                error=row.get("error", ""),
            )
            for row in reader
        ]


def filter_history_records(
    records: list[HistoryRecord],
    *,
    status: str | None = None,
    template_query: str | None = None,
    smtp_query: str | None = None,
) -> list[HistoryRecord]:
    status_filter = (status or "").strip().lower()
    template_filter = (template_query or "").strip().lower()
    smtp_filter = (smtp_query or "").strip().lower()

    filtered: list[HistoryRecord] = []
    for record in records:
        if status_filter and status_filter != "all" and record.status.lower() != status_filter:
            continue
        if template_filter and template_filter not in record.template.lower():
            continue
        if smtp_filter and smtp_filter not in record.smtp_account.lower():
            continue
        filtered.append(record)
    return filtered


def export_history_records(path: str | Path, records: list[HistoryRecord]) -> Path:
    export_path = Path(path)
    export_path.parent.mkdir(parents=True, exist_ok=True)

    payload = [
        {
            "timestamp": record.timestamp,
            "recipient": record.recipient,
            "status": record.status,
            "template": record.template,
            "smtp_account": record.smtp_account,
            "dry_run": record.dry_run,
            "error": record.error,
        }
        for record in records
    ]

    if export_path.suffix.lower() == ".json":
        export_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
        return export_path

    with export_path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(
            handle,
            fieldnames=[
                "timestamp",
                "recipient",
                "status",
                "template",
                "smtp_account",
                "dry_run",
                "error",
            ],
        )
        writer.writeheader()
        writer.writerows(payload)
    return export_path


def summarize_history_records(records: list[HistoryRecord]) -> HistoryStats:
    if not records:
        return HistoryStats(
            total=0,
            sent=0,
            dry_run=0,
            errors=0,
            unique_recipients=0,
            top_templates=[],
            top_smtp_accounts=[],
        )

    template_counter = Counter(record.template for record in records if record.template)
    smtp_counter = Counter(record.smtp_account for record in records if record.smtp_account)
    recipients = {record.recipient for record in records if record.recipient}

    return HistoryStats(
        total=len(records),
        sent=sum(1 for record in records if record.status == "sent"),
        dry_run=sum(1 for record in records if record.status == "dry-run"),
        errors=sum(1 for record in records if record.status == "error"),
        unique_recipients=len(recipients),
        top_templates=template_counter.most_common(5),
        top_smtp_accounts=smtp_counter.most_common(5),
    )


def load_history_stats(history_csv_path: str | Path) -> HistoryStats:
    return summarize_history_records(load_history_records(history_csv_path))
