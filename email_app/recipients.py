from __future__ import annotations

import csv
from pathlib import Path

from .models import Recipient


class RecipientsError(ValueError):
    """Raised when recipients data is invalid."""


def load_recipients(csv_path: str | Path) -> list[Recipient]:
    path = Path(csv_path)
    if not path.exists():
        raise RecipientsError(f"CSV-файл не найден: {path}")

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

    if not recipients:
        raise RecipientsError("Список получателей пуст")

    return recipients
