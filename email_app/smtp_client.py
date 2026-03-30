from __future__ import annotations


import mimetypes
import smtplib
import ssl
from email.message import EmailMessage
from pathlib import Path
from typing import Optional
import socket

try:
    import socks  # type: ignore
except ImportError:
    socks = None

from .models import MessageSettings, Recipient, SMTPSettings


class SMTPMailer:
    def __init__(self, settings: SMTPSettings) -> None:
        self.settings = settings

    def _build_message(
        self,
        recipient: Recipient,
        message_settings: MessageSettings,
        html_body: str,
        attachment_paths: Optional[list[Path]] = None,
        inline_image_paths: Optional[dict[str, Path]] = None,
    ) -> EmailMessage:
        message = EmailMessage()
        message["Subject"] = message_settings.subject
        message["From"] = f"{self.settings.from_name} <{self.settings.from_email}>"
        message["To"] = recipient.email
        if message_settings.reply_to:
            message["Reply-To"] = message_settings.reply_to
        message.set_content("Для просмотра письма используйте HTML-совместимый клиент.")
        message.add_alternative(html_body, subtype="html")
        html_part = message.get_payload()[-1]
        for cid, inline_path in (inline_image_paths or {}).items():
            content = inline_path.read_bytes()
            mime_type, _ = mimetypes.guess_type(inline_path.name)
            maintype, subtype = (mime_type or "application/octet-stream").split("/", maxsplit=1)
            html_part.add_related(
                content,
                maintype=maintype,
                subtype=subtype,
                cid=f"<{cid}>",
                filename=inline_path.name,
                disposition="inline",
            )
        for attachment_path in attachment_paths or []:
            content = attachment_path.read_bytes()
            mime_type, _ = mimetypes.guess_type(attachment_path.name)
            maintype, subtype = (mime_type or "application/octet-stream").split("/", maxsplit=1)
            message.add_attachment(
                content,
                maintype=maintype,
                subtype=subtype,
                filename=attachment_path.name,
            )
        return message

    def _open(self) -> smtplib.SMTP:
        timeout = self.settings.timeout_seconds
        # Прокси поддержка
        proxy_host = getattr(self.settings, "proxy_host", None)
        proxy_port = getattr(self.settings, "proxy_port", None)
        proxy_type = getattr(self.settings, "proxy_type", None)
        proxy_user = getattr(self.settings, "proxy_user", None)
        proxy_pass = getattr(self.settings, "proxy_pass", None)

        if proxy_host and proxy_port and proxy_type:
            if not socks:
                raise RuntimeError("Для поддержки прокси установите пакет PySocks: pip install PySocks")
            _type = {
                "socks5": socks.SOCKS5,
                "socks4": socks.SOCKS4,
                "http": socks.HTTP,
                "https": socks.HTTP,  # HTTP(S) реализуется одинаково
            }.get(str(proxy_type).lower())
            if not _type:
                raise ValueError(f"Неизвестный тип прокси: {proxy_type}")
            socks.set_default_proxy(_type, proxy_host, proxy_port, True if proxy_user else False, proxy_user, proxy_pass)
            socket.socket = socks.socksocket

        if self.settings.use_ssl:
            return smtplib.SMTP_SSL(
                host=self.settings.host,
                port=self.settings.port,
                timeout=timeout,
                context=ssl.create_default_context(),
            )
        return smtplib.SMTP(host=self.settings.host, port=self.settings.port, timeout=timeout)

    def send(
        self,
        recipient: Recipient,
        message_settings: MessageSettings,
        html_body: str,
        attachment_paths: Optional[list[Path]] = None,
        inline_image_paths: Optional[dict[str, Path]] = None,
    ) -> None:
        message = self._build_message(
            recipient,
            message_settings,
            html_body,
            attachment_paths,
            inline_image_paths,
        )
        with self._open() as server:
            server.ehlo()
            if self.settings.use_tls:
                server.starttls(context=ssl.create_default_context())
                server.ehlo()
            server.login(self.settings.username, self.settings.password)
            server.send_message(message)
