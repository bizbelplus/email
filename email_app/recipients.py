from __future__ import annotations

import csv
import io
from pathlib import Path

from .models import Recipient


class RecipientsError(ValueError):
    """Raised when recipients data is invalid."""


def _normalize_column_name(name: str) -> str:
    text = str(name).strip().lower().replace("\u00a0", " ")
    # Normalize common Cyrillic letters that visually look like Latin.
    confusables = {
        "а": "a",
        "е": "e",
        "о": "o",
        "р": "p",
        "с": "c",
        "у": "y",
        "х": "x",
        "к": "k",
        "м": "m",
        "т": "t",
        "в": "b",
        "і": "i",
    }
    return "".join(confusables.get(ch, ch) for ch in text)


def _score_decoded_csv(text: str) -> int:
    if not text:
        return -100

    score = 0
    if "\n" in text:
        score += 2
    if ";" in text or "," in text:
        score += 2
    if "@" in text:
        score += 3
    if "\ufffd" in text:
        score -= 5
    if "\x00" in text:
        score -= 8

    first_line = text.splitlines()[0] if text.splitlines() else ""
    if _resolve_email_column([p.strip() for p in first_line.split(";") + first_line.split(",") if p.strip()]):
        score += 6

    return score


def _read_text_with_fallback(path: Path) -> str:
    encodings = ("utf-8-sig", "utf-8", "cp1251", "cp866", "koi8-r", "utf-16", "utf-16-le", "utf-16-be")
    payload = path.read_bytes()

    best_text: str | None = None
    best_score = -10_000
    last_error: Exception | None = None

    for enc in encodings:
        try:
            decoded = payload.decode(enc)
        except UnicodeDecodeError as error:
            last_error = error
            continue
        except Exception as error:  # noqa: BLE001
            last_error = error
            continue

        score = _score_decoded_csv(decoded)
        if score > best_score:
            best_score = score
            best_text = decoded

    if best_text is not None:
        return best_text
    if last_error is not None:
        raise last_error
    raise RecipientsError(f"Не удалось прочитать файл получателей: {path}")


def _detect_csv_delimiter(text: str) -> str:
    sample = text[:8192]
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=",;")
        if dialect.delimiter in {",", ";"}:
            return str(dialect.delimiter)
    except Exception:
        pass
    return ";" if sample.count(";") > sample.count(",") else ","


def _resolve_email_column(fieldnames: list[str]) -> str | None:
    aliases = {"email", "mail", "e-mail", "почта"}
    lowered = {_normalize_column_name(name): name for name in fieldnames}
    for alias in aliases:
        if alias in lowered:
            return lowered[alias]
    for raw in fieldnames:
        low = _normalize_column_name(raw)
        if "mail" in low or "почт" in low:
            return raw
    return None


def _load_recipients_csv(path: Path) -> list[Recipient]:
    text = _read_text_with_fallback(path)
    delimiter = _detect_csv_delimiter(text)
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)

    fieldnames = [str(name).strip() for name in (reader.fieldnames or []) if str(name).strip()]
    email_key = _resolve_email_column(fieldnames)
    if not fieldnames or not email_key:
        raise RecipientsError("В CSV должен быть столбец email/mail")

    recipients: list[Recipient] = []
    for row in reader:
        normalized = {str(key).strip(): str(value or "").strip() for key, value in row.items() if key}
        email = normalized.get(email_key, "").strip()
        if not email:
            continue
        normalized["email"] = email
        recipients.append(Recipient(email=email, data=normalized))
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
