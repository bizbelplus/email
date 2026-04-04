from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any, Union

from jinja2 import Environment, FileSystemLoader, Undefined, UndefinedError, meta, select_autoescape

from .models import Recipient


class TemplateRenderError(ValueError):
    """Raised when a Jinja template cannot be rendered with provided context."""


class TemplateRenderer:
    def __init__(self, template_dir: Union[str, Path]) -> None:
        self.template_dir = Path(template_dir)
        self.environment = Environment(
            loader=FileSystemLoader(self.template_dir),
            autoescape=select_autoescape(["html", "xml"]),
            undefined=Undefined,
            trim_blocks=True,
            lstrip_blocks=True,
        )

    def list_templates(self) -> list[str]:
        if not self.template_dir.exists():
            return []
        return sorted(
            file_path.name
            for file_path in self.template_dir.iterdir()
            if file_path.is_file() and file_path.suffix.lower() in {".html", ".htm", ".xml", ".txt"}
        )

    def render(self, template_name: str, recipient: Recipient, context: dict[str, Any]) -> str:
        template = self.environment.get_template(template_name)
        recipient_data = defaultdict(str)
        recipient_data.update({**recipient.data, "email": recipient.email})
        merged_context = defaultdict(str)
        merged_context.update(context)
        merged_context["recipient"] = recipient_data
        try:
            return template.render(merged_context)
        except UndefinedError as error:
            raise TemplateRenderError(str(error)) from error

    def extract_template_variables(self, template_name: str) -> set[str]:
        source_text = self.environment.loader.get_source(self.environment, template_name)[0]
        import re
        vars = set(re.findall(r"\{\{\s*([^\s\}]+)\s*\}\}", source_text))
        return vars

    def find_undeclared_variables(self, template_name: str) -> set[str]:
        source_text = self.environment.loader.get_source(self.environment, template_name)[0]
        ast = self.environment.parse(source_text)
        return set(meta.find_undeclared_variables(ast))
