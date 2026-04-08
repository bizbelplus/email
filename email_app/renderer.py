from __future__ import annotations

from collections import defaultdict
from pathlib import Path
from typing import Any, Union
import random
import re

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
            rendered = template.render(merged_context)
            rendered = self.render_legacy_text(rendered, recipient_data)
            return rendered
        except UndefinedError as error:
            raise TemplateRenderError(str(error)) from error

    def render_legacy_text(self, text: str, recipient_data: dict[str, Any] | None = None) -> str:
        """Processes legacy placeholders and spintax in plain text fragments.

        Supports %field% placeholders (with exact-match priority) and {A|B|C} spintax.
        """
        data = defaultdict(str)
        if recipient_data:
            data.update(recipient_data)
        result = self._apply_legacy_placeholders(text, data)
        result = self._expand_spintax(result)
        return result

    def _apply_legacy_placeholders(self, text: str, recipient_data: dict[str, Any]) -> str:
        """Back-compat placeholders: resolves %field% by exact key, then alias fallback."""

        raw_map: dict[str, str] = {}
        norm_map: dict[str, str] = {}
        for key, value in dict(recipient_data).items():
            k = str(key or "").strip()
            if not k:
                continue
            v = str(value or "").strip()
            raw_map[k] = v
            norm_map[k.lower()] = v

        # Fallback aliases only when exact placeholder key is absent.
        alias_map: dict[str, list[str]] = {
            "name": ["имя", "фио", "full_name"],
            "имя": ["name", "фио", "full_name"],
            "фио": ["name", "имя", "full_name"],
            "company": ["компания", "organization", "org"],
            "компания": ["company", "organization", "org"],
            "phone": ["телефон", "номер", "mobile"],
            "телефон": ["phone", "номер", "mobile"],
        }

        def _resolve_placeholder(token: str) -> str:
            key = str(token or "").strip()
            if not key:
                return ""

            # 1) Exact key lookup (case-insensitive)
            exact = norm_map.get(key.lower())
            if exact is not None:
                return exact

            # 2) Alias fallback
            for alt in alias_map.get(key.lower(), []):
                candidate = norm_map.get(alt.lower())
                if candidate is not None:
                    return candidate
            return ""

        return re.sub(
            r"%\s*([^%\r\n]+?)\s*%",
            lambda m: _resolve_placeholder(m.group(1)),
            text,
            flags=re.IGNORECASE,
        )

    def _expand_spintax(self, text: str) -> str:
        """Expands legacy spintax blocks: {A|B|C}. Supports nested blocks."""

        def _split_top_level(content: str) -> list[str]:
            parts: list[str] = []
            start = 0
            depth = 0
            i = 0
            while i < len(content):
                ch = content[i]
                nxt = content[i + 1] if i + 1 < len(content) else ""
                if ch == "{" and nxt not in {"{", "%", "#"}:
                    depth += 1
                elif ch == "}" and depth > 0:
                    depth -= 1
                elif ch == "|" and depth == 0:
                    parts.append(content[start:i])
                    start = i + 1
                i += 1
            parts.append(content[start:])
            return parts

        def _expand_once(src: str) -> tuple[str, bool]:
            out: list[str] = []
            i = 0
            changed = False
            while i < len(src):
                ch = src[i]
                nxt = src[i + 1] if i + 1 < len(src) else ""
                if ch != "{" or nxt in {"{", "%", "#"}:
                    out.append(ch)
                    i += 1
                    continue

                depth = 1
                j = i + 1
                while j < len(src):
                    cur = src[j]
                    cur_next = src[j + 1] if j + 1 < len(src) else ""
                    if cur == "{" and cur_next not in {"{", "%", "#"}:
                        depth += 1
                    elif cur == "}":
                        depth -= 1
                        if depth == 0:
                            break
                    j += 1

                if j >= len(src) or src[j] != "}":
                    out.append(ch)
                    i += 1
                    continue

                block = src[i + 1:j]
                variants = _split_top_level(block)
                if len(variants) <= 1:
                    out.append(src[i:j + 1])
                    i = j + 1
                    continue

                out.append(random.choice(variants))
                changed = True
                i = j + 1

            return "".join(out), changed

        result = text
        for _ in range(20):
            result, changed = _expand_once(result)
            if not changed:
                break
        return result

    def extract_template_variables(self, template_name: str) -> set[str]:
        source_text = self.environment.loader.get_source(self.environment, template_name)[0]
        import re
        vars = set(re.findall(r"\{\{\s*([^\s\}]+)\s*\}\}", source_text))
        return vars

    def find_undeclared_variables(self, template_name: str) -> set[str]:
        source_text = self.environment.loader.get_source(self.environment, template_name)[0]
        ast = self.environment.parse(source_text)
        return set(meta.find_undeclared_variables(ast))
