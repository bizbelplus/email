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


_SOCKS_TYPE_MAP = {
    "socks5": "SOCKS5",
    "socks4": "SOCKS4",
    "http": "HTTP",
    "https": "HTTP",
}


def _make_socks_smtp(
    host: str,
    port: int,
    timeout: float,
    use_ssl: bool,
    ssl_context: ssl.SSLContext,
    proxy_type_str: str,
    proxy_host: str,
    proxy_port: int,
    proxy_user: str | None,
    proxy_pass: str | None,
) -> smtplib.SMTP:
    """Создаёт SMTP-соединение через SOCKS-прокси с передачей hostname (не IP).

    Прямое использование socks.socksocket().set_proxy() + connect(hostname) гарантирует,
    что DNS-резолвинг происходит на стороне прокси-сервера (rdns=True), а не на клиенте.
    Это критично: если клиент резолвит hostname в IP и передаёт IP прокси, прокси
    может вернуть 0x03 «Network unreachable», даже если hostname он обслуживает нормально.
    """
    if socks is None:
        raise RuntimeError("Для поддержки прокси установите пакет PySocks: pip install PySocks")

    socks_const = getattr(socks, _SOCKS_TYPE_MAP.get(proxy_type_str.lower(), ""), None)
    if socks_const is None:
        raise ValueError(f"Неизвестный тип прокси: {proxy_type_str}")

    # Создаём socks-сокет и явно указываем прокси на экземпляре (thread-safe, без глобального патча)
    raw_sock = socks.socksocket()
    raw_sock.settimeout(timeout)
    raw_sock.set_proxy(
        proxy_type=socks_const,
        addr=proxy_host,
        port=proxy_port,
        rdns=True,  # DNS резолвится на стороне прокси — hostname передаётся как есть
        username=str(proxy_user) if proxy_user else None,
        password=str(proxy_pass) if proxy_pass else None,
    )
    # Подключение с hostname — прокси сам резолвит DNS, нет локального getaddrinfo
    raw_sock.connect((host, port))

    if use_ssl:
        ssl_sock = ssl_context.wrap_socket(raw_sock, server_hostname=host)
        # SMTP_SSL с уже подключённым SSL-сокетом
        server = smtplib.SMTP_SSL.__new__(smtplib.SMTP_SSL)
        smtplib.SMTP.__init__(server)
        server.sock = ssl_sock
        server.file = ssl_sock.makefile("rb")
        server._host = host
        code, msg = server.getreply()
        if code != 220:
            server.close()
            raise smtplib.SMTPConnectError(code, msg)
        return server

    # Обычный SMTP (STARTTLS) — передаём уже подключённый сокет
    server = smtplib.SMTP.__new__(smtplib.SMTP)
    smtplib.SMTP.__init__(server)
    server.sock = raw_sock
    server.file = raw_sock.makefile("rb")
    server._host = host
    code, msg = server.getreply()
    if code != 220:
        server.close()
        raise smtplib.SMTPConnectError(code, msg)
    return server


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
        proxy_host = getattr(self.settings, "proxy_host", None)
        proxy_port = getattr(self.settings, "proxy_port", None)
        proxy_type = getattr(self.settings, "proxy_type", None)
        proxy_user = getattr(self.settings, "proxy_user", None)
        proxy_pass = getattr(self.settings, "proxy_pass", None)

        if proxy_host and proxy_port and proxy_type:
            # Используем socks.socksocket напрямую — hostname передаётся в прокси «как есть»,
            # DNS резолвится на стороне прокси-сервера. Это исправляет ошибку 0x03
            # «Network unreachable», которая возникала, когда клиент резолвил IP локально.
            return _make_socks_smtp(
                host=str(self.settings.host),
                port=int(self.settings.port),
                timeout=float(timeout),
                use_ssl=bool(self.settings.use_ssl),
                ssl_context=ssl.create_default_context(),
                proxy_type_str=str(proxy_type),
                proxy_host=str(proxy_host),
                proxy_port=int(proxy_port),
                proxy_user=proxy_user or None,
                proxy_pass=proxy_pass or None,
            )

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
