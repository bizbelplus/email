from __future__ import annotations

import sys
from pathlib import Path

try:
    from .main import main
except ImportError:
    package_root = Path(__file__).resolve().parent
    project_root = package_root.parent
    if str(project_root) not in sys.path:
        sys.path.insert(0, str(project_root))
    from email_app.main import main


if __name__ == "__main__":
    raise SystemExit(main())
