from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass(slots=True)
class SMTPSettings:
    host: str
    port: int
    username: str
    password: str
    from_email: str
    from_name: str
    use_tls: bool = True
    use_ssl: bool = False
    timeout_seconds: int = 30


@dataclass(slots=True)
class MessageSettings:
    subject: str
    template: str
    reply_to: str | None = None
    attachments: list[str] = field(default_factory=list)
    inline_images: dict[str, str] = field(default_factory=dict)


@dataclass(slots=True)
class DeliverySettings:
    delay_seconds: float = 0.0
    log_file: str = "logs/email_app.log"
    history_csv: str = "history/email_history.csv"
    history_jsonl: str = "history/email_history.jsonl"


@dataclass(slots=True)
class AppConfig:
    smtp_accounts: list[SMTPSettings]
    message: MessageSettings
    delivery: DeliverySettings
    content: dict[str, Any]

    @property
    def smtp(self) -> SMTPSettings:
        return self.smtp_accounts[0]


@dataclass(slots=True)
class Recipient:
    email: str
    data: dict[str, str]

    @property
    def name(self) -> str:
        return self.data.get("name", "").strip()
