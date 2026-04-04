from __future__ import annotations

import csv
from pathlib import Path

from .models import Recipient


class RecipientsError(ValueError):
    """Raised when recipients data is invalid."""


def _load_recipients_csv(path: Path) -> list[Recipient]:
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        if not reader.fieldnames or "email" not in reader.fieldnames:
            raise RecipientsError("В CSV должен быть столбец email")

        recipients: list[Recipient] = []
        for row in reader:
            email = (row.get("email") or "").strip()
            if not email:
                continue
            payload = {key: (value or "").strip() for key, value in row.items() if key}
            recipients.append(Recipient(email=email, data=payload))
    return recipients


def _load_recipients_txt(path: Path) -> list[Recipient]:
    recipients: list[Recipient] = []
    with path.open("r", encoding="utf-8-sig") as handle:
        for raw_line in handle:
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            recipients.append(Recipient(email=line, data={"email": line}))
    return recipients


def load_recipients(csv_path: str | Path) -> list[Recipient]:
    path = Path(csv_path)
    if not path.exists():
        raise RecipientsError(f"Файл получателей не найден: {path}")

    if path.suffix.lower() == ".txt":
        recipients = _load_recipients_txt(path)
    else:
        recipients = _load_recipients_csv(path)

    if not recipients:
        raise RecipientsError("Список получателей пуст")

    return recipients
