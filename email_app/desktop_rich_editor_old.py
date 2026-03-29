from __future__ import annotations

import argparse
import base64
import mimetypes
import re
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QFont, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QFileDialog,
    QFontComboBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSpinBox,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)


class DesktopRichEditorWindow(QMainWindow):
    def __init__(self, template_path: Path) -> None:
        super().__init__()
        self.template_path = template_path.resolve()
        self.source_mode = False
        self.original_head = '<meta charset="utf-8">'
        self.original_html_attrs = ' lang="ru"'
        self.original_body_attrs = ''

        self.setWindowTitle(f"Редактор письма — {self.template_path.name}")
        self.resize(1200, 860)

        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.setTabStopDistance(32)

        self.source_editor = QTextEdit()
        self.source_editor.setAcceptRichText(False)
        self.source_editor.hide()

        self.status_label = QLabel("Готово")

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.setContentsMargins(8, 8, 8, 8)
        layout.setSpacing(8)
        layout.addWidget(self.editor, 1)
        layout.addWidget(self.source_editor, 1)

        footer = QHBoxLayout()
        footer.addWidget(QLabel(f"Файл: {self.template_path}"), 1)
        footer.addWidget(self.status_label)
        layout.addLayout(footer)

        self.setCentralWidget(central)
        self._build_toolbar()
        self.load_template()

    def _build_toolbar(self) -> None:
        toolbar = QToolBar("Форматирование")
        toolbar.setMovable(False)
        self.addToolBar(toolbar)

        def add_action(text: str, callback, checkable: bool = False) -> QAction:
            action = QAction(text, self)
            action.setCheckable(checkable)
            action.triggered.connect(callback)
            toolbar.addAction(action)
            return action

        add_action("Ж", self.toggle_bold, checkable=True)
        add_action("К", self.toggle_italic, checkable=True)
        add_action("Ч", self.toggle_underline, checkable=True)
        toolbar.addSeparator()

        self.font_combo = QFontComboBox()
        self.font_combo.currentFontChanged.connect(self.change_font_family)
        toolbar.addWidget(self.font_combo)

        self.font_size = QSpinBox()
        self.font_size.setRange(8, 96)
        self.font_size.setValue(16)
        self.font_size.valueChanged.connect(self.change_font_size)
        toolbar.addWidget(self.font_size)

        toolbar.addSeparator()
        add_action("Цвет", self.change_text_color)
        add_action("Фон", self.change_bg_color)
        toolbar.addSeparator()
        add_action("Ссылка", self.insert_link)
        add_action("Фото", self.insert_image)
        add_action("CID", self.insert_cid_image)
        add_action("Таблица", self.insert_table)
        add_action("Переменная", self.insert_variable)
        toolbar.addSeparator()
        add_action("HTML", self.toggle_source_mode, checkable=True)
        add_action("Перезагрузить", self.reload_current_view)
        add_action("Сохранить", self.save_template)

    def current_editor(self) -> QTextEdit:
        return self.source_editor if self.source_mode else self.editor

    def set_status(self, text: str) -> None:
        self.status_label.setText(text)

    def load_template(self) -> None:
        if not self.template_path.exists():
            self.template_path.parent.mkdir(parents=True, exist_ok=True)
            self.template_path.write_text(self.default_template_html(), encoding="utf-8")

        raw_html = self.template_path.read_text(encoding="utf-8")
        self.original_html_attrs = self._extract_tag_attrs(raw_html, "html") or ' lang="ru"'
        self.original_body_attrs = self._extract_tag_attrs(raw_html, "body") or ""
        self.original_head = self._extract_head(raw_html) or '<meta charset="utf-8">'
        body_html = self._extract_body(raw_html) or "<p><br></p>"
        self.editor.setHtml(body_html)
        self.source_editor.setPlainText(raw_html)
        self.set_status("Шаблон загружен")

    def reload_current_view(self) -> None:
        if self.source_mode:
            self.source_editor.setPlainText(self.template_path.read_text(encoding="utf-8"))
        else:
            self.load_template()

    def toggle_bold(self) -> None:
        weight = QFont.Bold if self.editor.fontWeight() != QFont.Bold else QFont.Normal
        self.editor.setFontWeight(weight)

    def toggle_italic(self) -> None:
        self.editor.setFontItalic(not self.editor.fontItalic())

    def toggle_underline(self) -> None:
        self.editor.setFontUnderline(not self.editor.fontUnderline())

    def change_font_family(self, font: QFont) -> None:
        self.editor.setCurrentFont(font)

    def change_font_size(self, size: int) -> None:
        self.editor.setFontPointSize(float(size))

    def change_text_color(self) -> None:
        color = QColorDialog.getColor(parent=self, title="Цвет текста")
        if color.isValid():
            self.editor.setTextColor(color)

    def change_bg_color(self) -> None:
        color = QColorDialog.getColor(parent=self, title="Цвет фона")
        if color.isValid():
            fmt = QTextCharFormat()
            fmt.setBackground(color)
            self.editor.mergeCurrentCharFormat(fmt)

    def insert_link(self) -> None:
        url, ok = QInputDialog.getText(self, "Ссылка", "Введите URL:", text="https://")
        if not ok or not url.strip():
            return
        cursor = self.editor.textCursor()
        selected = cursor.selectedText() or url.strip()
        cursor.insertHtml(f'<a href="{url.strip()}">{selected}</a>')

    def insert_image(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Выберите фото",
            str(self.template_path.parent),
            "Images (*.png *.jpg *.jpeg *.gif *.webp *.bmp *.svg)",
        )
        if not path:
            return
        file_path = Path(path)
        mime = mimetypes.guess_type(file_path.name)[0] or "application/octet-stream"
        encoded = base64.b64encode(file_path.read_bytes()).decode("ascii")
        self.editor.textCursor().insertHtml(
            f'<img src="data:{mime};base64,{encoded}" alt="{file_path.name}" style="max-width:100%;height:auto;border:0;">'
        )

    def insert_cid_image(self) -> None:
        alias, ok = QInputDialog.getText(self, "CID-картинка", "Введите ключ inline_images:", text="hero")
        if not ok or not alias.strip():
            return
        self.editor.textCursor().insertHtml(
            f'<img src="cid:{{{{ inline_images.{alias.strip()}.cid }}}}" alt="{alias.strip()}" style="max-width:100%;height:auto;border:0;">'
        )

    def insert_table(self) -> None:
        rows, ok_rows = QInputDialog.getInt(self, "Таблица", "Строк:", value=2, minValue=1, maxValue=20)
        if not ok_rows:
            return
        cols, ok_cols = QInputDialog.getInt(self, "Таблица", "Столбцов:", value=2, minValue=1, maxValue=10)
        if not ok_cols:
            return
        cells = []
        for _ in range(rows):
            row = "".join('<td style="border:1px solid #cbd5e1;padding:8px;">Текст</td>' for _ in range(cols))
            cells.append(f"<tr>{row}</tr>")
        html = '<table style="border-collapse:collapse;width:100%;margin:12px 0;">' + ''.join(cells) + '</table>'
        self.editor.textCursor().insertHtml(html)

    def insert_variable(self) -> None:
        value, ok = QInputDialog.getText(
            self,
            "Переменная",
            "Введите переменную, например recipient.name или headline:",
            text="recipient.name",
        )
        if not ok or not value.strip():
            return
        self.editor.textCursor().insertText("{{ " + value.strip() + " }}")

    def toggle_source_mode(self, checked: bool) -> None:
        self.source_mode = checked
        if checked:
            self.source_editor.setPlainText(self.build_full_html_from_editor())
            self.editor.hide()
            self.source_editor.show()
            self.set_status("Режим HTML-кода")
        else:
            raw_html = self.source_editor.toPlainText()
            self.original_html_attrs = self._extract_tag_attrs(raw_html, "html") or self.original_html_attrs
            self.original_body_attrs = self._extract_tag_attrs(raw_html, "body") or self.original_body_attrs
            self.original_head = self._extract_head(raw_html) or self.original_head
            self.editor.setHtml(self._extract_body(raw_html) or "<p><br></p>")
            self.source_editor.hide()
            self.editor.show()
            self.set_status("Визуальный режим")

    def build_full_html_from_editor(self) -> str:
        editor_html = self.editor.toHtml()
        body_html = self._extract_body(editor_html) or self.editor.toHtml()
        return (
            "<!doctype html>\n"
            f"<html{self.original_html_attrs}>\n"
            "  <head>\n"
            f"{self.original_head}\n"
            "  </head>\n"
            f"  <body{self.original_body_attrs}>\n"
            f"{body_html}\n"
            "  </body>\n"
            "</html>\n"
        )

    def save_template(self) -> None:
        html = self.source_editor.toPlainText() if self.source_mode else self.build_full_html_from_editor()
        self.template_path.write_text(html.rstrip() + "\n", encoding="utf-8")
        QMessageBox.information(self, "Редактор письма", f"Шаблон сохранён:\n{self.template_path}")
        self.set_status("Шаблон сохранён")

    @staticmethod
    def _extract_head(html: str) -> str:
        match = re.search(r"<head[^>]*>(.*?)</head>", html, flags=re.IGNORECASE | re.DOTALL)
        return (match.group(1).strip() if match else "")

    @staticmethod
    def _extract_body(html: str) -> str:
        match = re.search(r"<body[^>]*>(.*?)</body>", html, flags=re.IGNORECASE | re.DOTALL)
        return (match.group(1).strip() if match else html.strip())

    @staticmethod
    def _extract_tag_attrs(html: str, tag_name: str) -> str:
        match = re.search(rf"<{tag_name}([^>]*)>", html, flags=re.IGNORECASE | re.DOTALL)
        return match.group(1).rstrip() if match else ""

    @staticmethod
    def default_template_html() -> str:
                return """<!doctype html>
<html lang=\"ru\">
    <head>
        <meta charset=\"utf-8\">
        <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">
        <title>{{ subject }}</title>
    </head>
    <body style=\"margin:0;padding:24px;background:#f3f4f6;font-family:Arial,sans-serif;color:#111827;\">
        <table role=\"presentation\" width=\"100%\" cellspacing=\"0\" cellpadding=\"0\" style=\"max-width:640px;margin:0 auto;background:#ffffff;border-radius:16px;overflow:hidden;\">
            <tr>
                <td style=\"padding:32px 40px;background:linear-gradient(135deg,#2563eb,#7c3aed);color:#ffffff;\">
                    <h1 style=\"margin:0 0 8px;font-size:28px;line-height:1.2;\">{{ headline }}</h1>
                    <p style=\"margin:0;font-size:16px;line-height:1.6;\">{{ preheader }}</p>
                </td>
            </tr>
            <tr>
                <td style=\"padding:40px;\">
                    <p>Здравствуйте{% if recipient.name %}, {{ recipient.name }}{% endif %}!</p>
                    <p>{{ body_intro }}</p>
                    <p>{{ body_text }}</p>
                    <p>
                        <a href=\"{{ cta_url }}\" style=\"display:inline-block;background:#2563eb;color:#ffffff;text-decoration:none;padding:14px 22px;border-radius:10px;font-weight:700;\">{{ cta_label }}</a>
                    </p>
                </td>
            </tr>
        </table>
    </body>
</html>
"""


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Desktop rich editor for email templates")
    parser.add_argument("--template", required=True, help="Absolute path to template file")
    args = parser.parse_args(argv)

    app = QApplication(sys.argv if argv is None else [sys.argv[0], *argv])
    window = DesktopRichEditorWindow(Path(args.template))
    window.show()
    return app.exec()


if __name__ == "__main__":
    raise SystemExit(main())
