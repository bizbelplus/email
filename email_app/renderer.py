from __future__ import annotations

from pathlib import Path
from typing import Any

from jinja2 import Environment, FileSystemLoader, StrictUndefined, select_autoescape

from .models import Recipient


class TemplateRenderer:
    def __init__(self, template_dir: str | Path) -> None:
        self.template_dir = Path(template_dir)
        self.environment = Environment(
            loader=FileSystemLoader(self.template_dir),
            autoescape=select_autoescape(["html", "xml"]),
            undefined=StrictUndefined,
            trim_blocks=True,
            lstrip_blocks=True,
        )

    def list_templates(self) -> list[str]:
        if not self.template_dir.exists():
            return []
        return sorted(
            file_path.name
            for file_path in self.template_dir.iterdir()
            if file_path.is_file() and file_path.suffix.lower() in {".html", ".htm", ".xml"}
        )

    def render(self, template_name: str, recipient: Recipient, context: dict[str, Any]) -> str:
        template = self.environment.get_template(template_name)
        merged_context = {**context, "recipient": {**recipient.data, "email": recipient.email}}
        return template.render(**merged_context)
