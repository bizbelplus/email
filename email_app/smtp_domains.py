from __future__ import annotations

from pathlib import Path
from typing import Any

CONN_LABELS = ["SSL/TLS", "STARTTLS", "plain"]
CONN_LABEL_TO_INT = {"plain": 0, "STARTTLS": 1, "SSL/TLS": 2}
_CONN_INT_TO_LABEL = {value: key for key, value in CONN_LABEL_TO_INT.items()}


def _conn_type_to_flags(conn_type: int | str) -> tuple[bool, bool]:
    if isinstance(conn_type, str):
        key = conn_type.strip()
        if key.isdigit():
            conn_type = int(key)
        else:
            normalized = key.lower()
            if normalized in {"ssl", "ssl/tls", "tls_ssl"}:
                return False, True
            if normalized in {"starttls", "tls"}:
                return True, False
            return False, False

    if conn_type == 2:
        return False, True
    if conn_type == 1:
        return True, False
    return False, False


def _flags_to_conn_type(use_tls: bool, use_ssl: bool) -> int:
    if use_ssl:
        return 2
    if use_tls:
        return 1
    return 0


def _flags_to_label(use_tls: bool, use_ssl: bool) -> str:
    return _CONN_INT_TO_LABEL[_flags_to_conn_type(use_tls, use_ssl)]


def parse_smtp_domains_file(path: Path) -> dict[str, dict[str, Any]]:
    domains: dict[str, dict[str, Any]] = {}
    if not path.exists():
        return domains

    with path.open("r", encoding="utf-8-sig") as handle:
        for line_number, raw_line in enumerate(handle, start=1):
            line = raw_line.strip()
            if not line or line.startswith("#"):
                continue
            if not line.startswith("@") or ";" not in line:
                continue

            domain_part, payload = line.split(";", 1)
            domain = domain_part.lstrip("@").strip().lower()
            if not domain:
                continue

            parts = [item.strip() for item in payload.split(":") if item.strip()]
            if len(parts) < 2:
                raise ValueError(f"Некорректная строка SMTP-доменов {line_number}: {line}")

            host = parts[0]
            port = int(parts[1])
            conn_token = parts[-1] if len(parts) >= 3 else "1"
            use_tls, use_ssl = _conn_type_to_flags(conn_token)
            domains[domain] = {
                "host": host,
                "port": port,
                "use_tls": use_tls,
                "use_ssl": use_ssl,
            }
    return domains


def save_smtp_domains_file(path: Path, domains: dict[str, dict[str, Any]]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    lines = [
        "# Формат: @domain;smtp.host:port:conn_type",
        "# conn_type: 0=plain, 1=STARTTLS, 2=SSL/TLS",
    ]
    for domain, settings in sorted(domains.items()):
        normalized = domain.strip().lstrip("@").lower()
        if not normalized:
            continue
        host = str(settings.get("host", "")).strip()
        port = int(settings.get("port", 587))
        use_tls = bool(settings.get("use_tls", True))
        use_ssl = bool(settings.get("use_ssl", False))
        conn_type = _flags_to_conn_type(use_tls, use_ssl)
        lines.append(f"@{normalized};{host}:{port}:{conn_type}")
    path.write_text("\n".join(lines) + "\n", encoding="utf-8")


def load_domains(base_dir: Path) -> dict[str, dict[str, Any]]:
    config_path = base_dir / "config" / "smtp_domains.txt"
    root_path = base_dir / "smtp_domains.txt"
    if config_path.exists():
        return parse_smtp_domains_file(config_path)
    if root_path.exists():
        return parse_smtp_domains_file(root_path)
    return {}


def get_smtp_defaults_for_email(
    email: str,
    domains: dict[str, dict[str, Any]] | None = None,
) -> dict[str, Any]:
    domains = domains or {}
    text = str(email).strip().lower()
    domain = text.split("@", 1)[1] if "@" in text else text
    if domain in domains:
        return dict(domains[domain])

    if domain == "gmail.com":
        return {"host": "smtp.gmail.com", "port": 587, "use_tls": True, "use_ssl": False}
    if domain in {"outlook.com", "hotmail.com", "live.com"}:
        return {"host": "smtp.office365.com", "port": 587, "use_tls": True, "use_ssl": False}
    if domain == "yandex.ru":
        return {"host": "smtp.yandex.ru", "port": 465, "use_tls": False, "use_ssl": True}
    if domain == "mail.ru":
        return {"host": "smtp.mail.ru", "port": 465, "use_tls": False, "use_ssl": True}
    if domain:
        return {"host": f"smtp.{domain}", "port": 587, "use_tls": True, "use_ssl": False}
    return {"host": "", "port": 587, "use_tls": True, "use_ssl": False}
