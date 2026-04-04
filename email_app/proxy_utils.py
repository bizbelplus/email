import random
from pathlib import Path


def _normalize_proxy_type(value: str | None, default: str = "socks5") -> str:
    normalized = str(value or "").strip().lower()
    if normalized in {"socks5", "socks4", "http", "https"}:
        return normalized
    return default

def load_proxies(
    proxy_file: str | Path,
    *,
    default_proxy_type: str = "socks5",
    default_proxy_user: str | None = None,
    default_proxy_pass: str | None = None,
) -> list[dict]:
    proxies = []
    default_type = _normalize_proxy_type(default_proxy_type)
    with open(proxy_file, encoding="utf-8-sig") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue

            proxy: dict

            # Формат: host:port@user:pass (тип берется из default_proxy_type)
            if "@" in line:
                host_port, auth = line.split("@", 1)
                hp_parts = [item.strip() for item in host_port.split(":", 1)]
                auth_parts = [item.strip() for item in auth.split(":", 1)]
                if len(hp_parts) != 2 or len(auth_parts) != 2:
                    continue
                proxy = {
                    'proxy_host': hp_parts[0],
                    'proxy_port': int(hp_parts[1]),
                    'proxy_type': default_type,
                    'proxy_user': auth_parts[0] or None,
                    'proxy_pass': auth_parts[1] or None,
                }
                proxies.append(proxy)
                continue

            parts = [item.strip() for item in line.split(':')]

            # Формат: host:port:type[:user:pass]
            if len(parts) >= 3 and parts[2]:
                proxy = {
                    'proxy_host': parts[0],
                    'proxy_port': int(parts[1]),
                    'proxy_type': _normalize_proxy_type(parts[2], default=default_type),
                    'proxy_user': parts[3] if len(parts) > 3 and parts[3] else None,
                    'proxy_pass': parts[4] if len(parts) > 4 and parts[4] else None,
                }
                proxies.append(proxy)
                continue

            # Формат: host:port (тип и auth берутся из дефолтов)
            if len(parts) == 2:
                proxy = {
                    'proxy_host': parts[0],
                    'proxy_port': int(parts[1]),
                    'proxy_type': default_type,
                    'proxy_user': default_proxy_user,
                    'proxy_pass': default_proxy_pass,
                }
                proxies.append(proxy)
                continue
    return proxies

def pick_random_proxy(proxies: list[dict]) -> dict:
    return random.choice(proxies) if proxies else {}
