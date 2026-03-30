import random
from pathlib import Path

def load_proxies(proxy_file: str | Path) -> list[dict]:
    proxies = []
    with open(proxy_file, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
            parts = line.split(':')
            if len(parts) < 3:
                continue
            proxy = {
                'proxy_host': parts[0],
                'proxy_port': int(parts[1]),
                'proxy_type': parts[2],
                'proxy_user': parts[3] if len(parts) > 3 else None,
                'proxy_pass': parts[4] if len(parts) > 4 else None,
            }
            proxies.append(proxy)
    return proxies

def pick_random_proxy(proxies: list[dict]) -> dict:
    return random.choice(proxies) if proxies else {}
