"""Modern desktop rich HTML editor with advanced formatting tools."""

from __future__ import annotations

import argparse
import base64
import mimetypes
import re
import sys
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtGui import QAction, QColor, QFont, QIcon, QTextCharFormat, QTextCursor
from PySide6.QtWidgets import (
    QApplication,
    QColorDialog,
    QComboBox,
    QFileDialog,
    QFontComboBox,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSpinBox,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)


class ModernDesktopRichEditor(QMainWindow):
    """Modern desktop rich HTML editor with dark theme and advanced tools."""

    def __init__(self, template_path: Path) -> None:
        super().__init__()
        self.template_path = template_path.resolve()
        self.source_mode = False

        self.setWindowTitle(f"📝 HTML Editor — {self.template_path.name}")
        self.setStyleSheet(self._get_dark_stylesheet())
        self.resize(1400, 900)

        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        layout = QVBoxLayout(central)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Main editor area
        self.editor = QTextEdit()
        self.editor.setAcceptRichText(True)
        self.editor.setTabStopDistance(32)
        self.editor.setStyleSheet(self._get_editor_stylesheet())

        # Source code editor
        self.source_editor = QTextEdit()
        self.source_editor.setAcceptRichText(False)
        self.source_editor.setStyleSheet(self._get_editor_stylesheet())
        self.source_editor.setFont(QFont("Courier", 10))
        self.source_editor.hide()

        # Status bar
        self.status_label = QLabel("✓ Готово")
        self.char_count_label = QLabel("Символов: 0")
        self.mode_label = QLabel("📝 Редактор")

        # Footer
        footer_layout = QHBoxLayout()
        footer_layout.setContentsMargins(12, 8, 12, 8)
        footer_layout.addWidget(QLabel(f"📄 {self.template_path}"), 1)
        footer_layout.addWidget(self.mode_label)
        footer_layout.addWidget(self.char_count_label)
        footer_layout.addWidget(self.status_label)

        layout.addWidget(self.editor, 1)
        layout.addWidget(self.source_editor, 1)

        footer_widget = QWidget()
        footer_widget.setLayout(footer_layout)
        footer_widget.setStyleSheet("border-top: 1px solid #404040; background-color: #1e1e1e;")
        layout.addWidget(footer_widget)

        self._build_toolbars()
        self.doctype = "<!doctype html>"
        self.html_open_tag = "<html lang=\"ru\">"
        self.body_start_tag = "<body>"
        self.body_end_tag = "</body>"
        self.head_html = ""
        self.load_template()
        self.update_char_count()
        self.editor.textChanged.connect(self.update_char_count)

    def _get_dark_stylesheet(self) -> str:
        """Return modern dark theme stylesheet."""
        return """
        QMainWindow {
            background-color: #1e1e1e;
            color: #e0e0e0;
        }
        QToolBar {
            background-color: #2d2d30;
            border-bottom: 1px solid #3e3e42;
            spacing: 2px;
            padding: 4px;
        }
        QToolBar::separator {
            background-color: #505050;
            width: 1px;
            margin: 0 4px;
        }
        QPushButton {
            background-color: #007acc;
            color: #ffffff;
            border: none;
            border-radius: 3px;
            padding: 4px 8px;
            font-weight: bold;
        }
        QPushButton:hover {
            background-color: #0098ff;
        }
        QPushButton:pressed {
            background-color: #005a9e;
        }
        QComboBox, QSpinBox, QFontComboBox {
            background-color: #3c3c3c;
            color: #e0e0e0;
            border: 1px solid #505050;
            border-radius: 3px;
            padding: 4px;
        }
        QComboBox::drop-down {
            border: none;
        }
        QLabel {
            color: #e0e0e0;
        }
        """

    def _get_editor_stylesheet(self) -> str:
        """Return editor stylesheet."""
        return """
        QTextEdit {
            background-color: #252526;
            color: #d4d4d4;
            border: 1px solid #3e3e42;
            font-family: 'Segoe UI', Arial;
            font-size: 11pt;
            selection-background-color: #007acc;
        }
        """

    def _build_toolbars(self) -> None:
        """Построить тулбары с русскими подписями и расширенными настройками."""
        # Форматирование
        fmt_toolbar = self._create_toolbar("Форматирование", "fmt")
        self._add_button(fmt_toolbar, "Жирный", self.toggle_bold, "Ctrl+B")
        self._add_button(fmt_toolbar, "Курсив", self.toggle_italic, "Ctrl+I")
        self._add_button(fmt_toolbar, "Подчёркнутый", self.toggle_underline, "Ctrl+U")
        fmt_toolbar.addSeparator()

        self.font_combo = QFontComboBox()
        self.font_combo.setMaximumWidth(150)
        self.font_combo.currentFontChanged.connect(self.change_font_family)
        fmt_toolbar.addWidget(QLabel("Шрифт:"))
        fmt_toolbar.addWidget(self.font_combo)

        self.font_size = QSpinBox()
        self.font_size.setRange(8, 96)
        self.font_size.setValue(14)
        self.font_size.setMaximumWidth(70)
        self.font_size.valueChanged.connect(self.change_font_size)
        fmt_toolbar.addWidget(QLabel("Размер:"))
        fmt_toolbar.addWidget(self.font_size)

        # Межстрочный интервал
        fmt_toolbar.addWidget(QLabel("Межстрочный:"))
        self.line_spacing = QSpinBox()
        self.line_spacing.setRange(80, 300)
        self.line_spacing.setValue(120)
        self.line_spacing.setMaximumWidth(70)
        self.line_spacing.valueChanged.connect(self.change_line_spacing)
        fmt_toolbar.addWidget(self.line_spacing)

        fmt_toolbar.addSeparator()
        self._add_button(fmt_toolbar, "Цвет текста", self.change_text_color, "Ctrl+Shift+C")
        self._add_button(fmt_toolbar, "Цвет выделения", self.change_bg_color, "Ctrl+Shift+H")
        self._add_button(fmt_toolbar, "Фон письма", self.change_page_bg)

        # Undo/Redo
        fmt_toolbar.addSeparator()
        self._add_button(fmt_toolbar, "↺ Отмена", self.editor.undo, "Ctrl+Z")
        self._add_button(fmt_toolbar, "↻ Повтор", self.editor.redo, "Ctrl+Y")

        # Списки и выравнивание
        list_toolbar = self._create_toolbar("Списки и выравнивание", "list")
        self._add_button(list_toolbar, "Маркированный список", self.insert_bullet_list)
        self._add_button(list_toolbar, "Нумерованный список", self.insert_numbered_list)
        list_toolbar.addSeparator()
        self._add_button(list_toolbar, "Выровнять влево", self.align_left)
        self._add_button(list_toolbar, "По центру", self.align_center)
        self._add_button(list_toolbar, "Выровнять вправо", self.align_right)

        # Вставка
        insert_toolbar = self._create_toolbar("Вставка", "insert")
        self._add_button(insert_toolbar, "Ссылка", self.insert_link, "Ctrl+L")
        self._add_button(insert_toolbar, "Изображение", self.insert_image)
        self._add_button(insert_toolbar, "Таблица", self.insert_table)
        self._add_button(insert_toolbar, "Горизонтальная линия", self.insert_horizontal_rule)
        self._add_button(insert_toolbar, "Блок кода", self.insert_code_block)
        insert_toolbar.addSeparator()
        self._add_button(insert_toolbar, "Переменная", self.insert_variables)
        self._add_button(insert_toolbar, "Быстрый блок", self.insert_quick_block)

        # Вид
        view_toolbar = self._create_toolbar("Вид", "view")
        self._add_button(view_toolbar, "Сохранить", self.save_template, "Ctrl+S")
        self._add_button(view_toolbar, "Перезагрузить", self.load_template, "Ctrl+R")
        view_toolbar.addSeparator()
        self._add_button(view_toolbar, "Исходный код", self.toggle_source_mode, "Ctrl+`")

    def change_line_spacing(self, value: int) -> None:
        """Изменить межстрочный интервал (проценты)."""
        cursor = self.editor.textCursor()
        block_fmt = cursor.blockFormat()
        block_fmt.setLineHeight(value, block_fmt.ProportionalHeight)
        cursor.setBlockFormat(block_fmt)

    def change_page_bg(self) -> None:
        """Изменить цвет фона письма."""
        color = QColorDialog.getColor(Qt.white, self, "Выбрать цвет фона письма")
        if color.isValid():
            self.editor.setStyleSheet(self._get_editor_stylesheet() + f"\nQTextEdit {{ background-color: {color.name()}; }}")

    def insert_code_block(self) -> None:
        """Вставить блок кода."""
        code, ok = QInputDialog.getMultiLineText(self, "Вставить код", "Код:")
        if ok and code:
            html = f'<pre style="background:#222;color:#e0e0e0;padding:8px;border-radius:4px;font-family:monospace;">{code}</pre>'
            self.editor.insertHtml(html)

    def insert_quick_block(self) -> None:
        """Вставить быстрый шаблон блока (например, кнопка, alert, info)."""
        blocks = {
            "Кнопка (ссылка)": '<a href="https://example.com" style="display:inline-block;padding:10px 24px;background:#007acc;color:#fff;text-decoration:none;border-radius:4px;">Текст кнопки</a>',
            "Важное сообщение": '<div style="background:#ffe0e0;color:#a00;padding:12px;border-radius:4px;">Важное сообщение!</div>',
            "Инфо-блок": '<div style="background:#e0f7fa;color:#00796b;padding:12px;border-radius:4px;">Информационный блок</div>',
        }
        items = list(blocks.keys())
        block, ok = QInputDialog.getItem(self, "Быстрый блок", "Выберите блок:", items, 0, False)
        if ok and block:
            self.editor.insertHtml(blocks[block])

    def _create_toolbar(self, name: str, object_name: str) -> QToolBar:
        """Create and add a toolbar."""
        toolbar = QToolBar(name)
        toolbar.setObjectName(object_name)
        toolbar.setMovable(False)
        self.addToolBar(toolbar)
        return toolbar

    def _add_button(self, toolbar: QToolBar, text: str, callback, shortcut: str = None) -> QPushButton:
        """Add styled button to toolbar."""
        button = QPushButton(text)
        button.setMaximumHeight(28)
        button.clicked.connect(callback)
        if shortcut:
            button.setShortcut(shortcut)
            button.setToolTip(f"{text} ({shortcut})")
        toolbar.addWidget(button)
        return button

    def toggle_bold(self) -> None:
        """Toggle bold formatting."""
        fmt = self.editor.currentCharFormat()
        fmt.setFontWeight(700 if fmt.fontWeight() < 700 else 400)
        self.editor.setCurrentCharFormat(fmt)

    def toggle_italic(self) -> None:
        """Toggle italic formatting."""
        fmt = self.editor.currentCharFormat()
        fmt.setFontItalic(not fmt.fontItalic())
        self.editor.setCurrentCharFormat(fmt)

    def toggle_underline(self) -> None:
        """Toggle underline formatting."""
        fmt = self.editor.currentCharFormat()
        fmt.setFontUnderline(not fmt.fontUnderline())
        self.editor.setCurrentCharFormat(fmt)

    def change_font_family(self, font: QFont) -> None:
        """Change font family."""
        fmt = self.editor.currentCharFormat()
        fmt.setFont(font)
        self.editor.setCurrentCharFormat(fmt)

    def change_font_size(self, size: int) -> None:
        """Change font size."""
        fmt = self.editor.currentCharFormat()
        fmt.setFontPointSize(size)
        self.editor.setCurrentCharFormat(fmt)

    def change_text_color(self) -> None:
        """Change text color."""
        color = QColorDialog.getColor(Qt.white, self, "Выбрать цвет текста")
        if color.isValid():
            fmt = self.editor.currentCharFormat()
            fmt.setForeground(color)
            self.editor.setCurrentCharFormat(fmt)

    def change_bg_color(self) -> None:
        """Change background color."""
        color = QColorDialog.getColor(Qt.white, self, "Выбрать цвет фона")
        if color.isValid():
            fmt = self.editor.currentCharFormat()
            fmt.setBackground(color)
            self.editor.setCurrentCharFormat(fmt)

    def _parse_template(self, content: str) -> None:
        """Вынуть doctype/html/head/body из шаблона."""
        doctype_match = re.search(r"<!doctype[^>]*>", content, flags=re.IGNORECASE)
        self.doctype = doctype_match.group(0) if doctype_match else "<!doctype html>"

        html_match = re.search(r"(<html[^>]*>)", content, flags=re.IGNORECASE)
        self.html_open_tag = html_match.group(1) if html_match else "<html lang=\"ru\">"

        body_match_full = re.search(r"(<body[^>]*>)(.*?)(</body>)", content, flags=re.DOTALL | re.IGNORECASE)
        if body_match_full:
            self.body_start_tag = body_match_full.group(1)
            self.body_end_tag = body_match_full.group(3)
            self.body_html = body_match_full.group(2)
        else:
            self.body_start_tag = "<body>"
            self.body_end_tag = "</body>"
            self.body_html = content

        head_match = re.search(r"(<head[^>]*>.*?</head>)", content, flags=re.DOTALL | re.IGNORECASE)
        self.head_html = head_match.group(1) if head_match else "<head></head>"

    def _apply_plain_text_to_body(self, plain_text: str) -> str:
        """Применить текст из визуального редактора к телу, сохраняя оригинальную структуру тегов."""
        parts = re.split(r"(<[^>]+>)", self.body_html)
        text_indices = [i for i, part in enumerate(parts) if not part.startswith("<")]
        if not text_indices:
            return self.body_html

        # Для каждого текстового сегмента распределяем новый текст пропорционально по длине исходного
        total_orig_len = sum(len(parts[i]) for i in text_indices)
        if total_orig_len == 0:
            return self.body_html

        normalized_target = plain_text.replace("\r\n", "\n").replace("\r", "\n")
        pos = 0
        for i in text_indices:
            orig_len = len(parts[i])
            if orig_len == 0:
                continue
            take = min(orig_len, max(0, len(normalized_target) - pos))
            parts[i] = normalized_target[pos: pos + take]
            pos += take

        if pos < len(normalized_target):
            # если осталось доп. содержимого — добавляем в последний текстовый сегмент
            parts[text_indices[-1]] += normalized_target[pos:]

        return "".join(parts)

    def _assemble_template(self, body_html: str | None = None) -> str:
        """Собрать финальный HTML из сохранённых блоков + отредактированного body."""
        if body_html is None:
            rendered = self.editor.toHtml()
            body_match = re.search(r"<body[^>]*>(.*?)</body>", rendered, flags=re.DOTALL | re.IGNORECASE)
            body_html = body_match.group(1) if body_match else rendered

        return (
            f"{self.doctype}\n"
            f"{self.html_open_tag}\n"
            f"{self.head_html}\n"
            f"{self.body_start_tag}{body_html}{self.body_end_tag}\n"
            f"</html>"
        )

    def insert_link(self) -> None:
        """Insert hyperlink."""
        text, ok = QInputDialog.getText(self, "Вставить ссылку", "Текст ссылки:")
        if not ok or not text:
            return
        url, ok = QInputDialog.getText(self, "Вставить ссылку", "URL:")
        if not ok or not url:
            return
        fmt = self.editor.currentCharFormat()
        fmt.setAnchor(True)
        fmt.setAnchorHref(url)
        fmt.setForeground(QColor(0, 100, 200))
        fmt.setFontUnderline(True)
        cursor = self.editor.textCursor()
        cursor.insertText(text, fmt)

    def insert_image(self) -> None:
        """Insert image."""
        path, _ = QFileDialog.getOpenFileName(
            self, "Выбрать изображение", "", "Images (*.png *.jpg *.jpeg *.gif *.svg)"
        )
        if not path:
            return
        cid = Path(path).stem
        html = f'<img src="cid:{cid}" alt="{cid}" style="max-width: 100%; height: auto;">'
        self.editor.insertHtml(html)
        self.status_label.setText(f"✓ Изображение вставлено: {cid}")

    def insert_table(self) -> None:
        """Insert table."""
        rows, ok1 = QInputDialog.getInt(self, "Вставить таблицу", "Строк:", 2, 1, 20)
        if not ok1:
            return
        cols, ok2 = QInputDialog.getInt(self, "Вставить таблицу", "Столбцов:", 2, 1, 20)
        if not ok2:
            return

        html = '<table border="1" style="border-collapse: collapse; width: 100%;">'
        for _ in range(rows):
            html += "<tr>"
            for _ in range(cols):
                html += '<td style="padding: 8px; border: 1px solid #ccc;">Ячейка</td>'
            html += "</tr>"
        html += "</table>"
        self.editor.insertHtml(html)
        self.status_label.setText(f"✓ Таблица {rows}x{cols} вставлена")

    def insert_bullet_list(self) -> None:
        """Insert bullet list."""
        html = "<ul><li>Пункт 1</li><li>Пункт 2</li><li>Пункт 3</li></ul>"
        self.editor.insertHtml(html)

    def insert_numbered_list(self) -> None:
        """Insert numbered list."""
        html = "<ol><li>Первый</li><li>Второй</li><li>Третий</li></ol>"
        self.editor.insertHtml(html)

    def insert_horizontal_rule(self) -> None:
        """Insert horizontal rule."""
        self.editor.insertHtml("<hr style='border: none; height: 2px; background: #ccc; margin: 16px 0;'>")

    def insert_variables(self) -> None:
        """Insert template variables."""
        variables = ["{{ name }}", "{{ email }}", "{{ date }}", "{{ company }}", "{{ custom_field }}"]
        var, ok = QInputDialog.getItem(self, "Вставить переменную", "Переменная:", variables, 0, True)
        if ok and var:
            self.editor.insertPlainText(var)

    def align_left(self) -> None:
        """Align text left."""
        self.editor.alignment()
        cursor = self.editor.textCursor()
        cursor.setPosition(0)
        self.editor.setAlignment(Qt.AlignmentFlag.AlignLeft)

    def align_center(self) -> None:
        """Align text center."""
        self.editor.setAlignment(Qt.AlignmentFlag.AlignCenter)

    def align_right(self) -> None:
        """Align text right."""
        self.editor.setAlignment(Qt.AlignmentFlag.AlignRight)

    def toggle_source_mode(self) -> None:
        """Toggle between WYSIWYG and source code modes."""
        self.source_mode = not self.source_mode

        if self.source_mode:
            # Switch to source mode
            current_template = self._assemble_template()
            self.source_editor.setPlainText(current_template)
            self.editor.hide()
            self.source_editor.show()
            self.mode_label.setText("</> Исходный код")
            self.status_label.setText("✓ Режим редактирования HTML")
        else:
            # Switch back to WYSIWYG
            source = self.source_editor.toPlainText()
            try:
                self._parse_template(source)
                self.editor.setHtml(self.body_html)
                self.status_label.setText("✓ Режим WYSIWYG")
            except Exception as e:
                self.status_label.setText(f"✗ Ошибка: {e}")
            self.source_editor.hide()
            self.editor.show()
            self.mode_label.setText("📝 Редактор")

    def load_template(self) -> None:
        """Load template from file."""
        if not self.template_path.exists():
            self.status_label.setText(f"✗ Файл не найден: {self.template_path}")
            return

        content = self.template_path.read_text(encoding="utf-8")

        # Save raw original structure to preserve metadata and theme-related attributes
        doctype_match = re.search(r"<!doctype[^>]*>", content, flags=re.IGNORECASE)
        self.doctype = doctype_match.group(0) if doctype_match else "<!doctype html>"

        html_match = re.search(r"(<html[^>]*>)", content, flags=re.IGNORECASE)
        self.html_open_tag = html_match.group(1) if html_match else "<html lang=\"ru\">"

        body_match_full = re.search(r"(<body[^>]*>)(.*?)(</body>)", content, flags=re.DOTALL | re.IGNORECASE)
        if body_match_full:
            self.body_start_tag = body_match_full.group(1)
            self.body_end_tag = body_match_full.group(3)
        else:
            self.body_start_tag = "<body>"
            self.body_end_tag = "</body>"

        self._parse_template(content)

        # Set full HTML content (Qt ограничен, но в editor видно форматирование максимально)
        self.editor.setHtml(self.body_html)
        self.source_editor.setPlainText(content)
        self.status_label.setText(f"✓ Загружено: {self.template_path}")
        self.update_char_count()

    def save_template(self) -> None:
        """Save template to file."""
        if self.source_mode:
            content = self.source_editor.toPlainText()
            self._parse_template(content)
        else:
            # Сохраняем WYSIWYG-результат, но сохраняем head/html атрибуты из оригинала
            rendered = self.editor.toHtml()
            body_match = re.search(r"<body[^>]*>(.*?)</body>", rendered, flags=re.DOTALL | re.IGNORECASE)
            body_content = body_match.group(1) if body_match else rendered
            content = self._assemble_template(body_content)
            self._parse_template(content)

        self.template_path.write_text(content, encoding="utf-8")
        self.source_editor.setPlainText(content)
        self.status_label.setText(f"✓ Сохранено: {self.template_path}" )

    def update_char_count(self) -> None:
        """Update character count."""
        text = self.editor.toPlainText()
        count = len(text)
        self.char_count_label.setText(f"Символов: {count}")


def main() -> None:
    """Main entry point."""
    parser = argparse.ArgumentParser(description="Modern HTML Email Template Editor")
    parser.add_argument("--template", type=Path, required=True, help="Path to HTML template")
    args = parser.parse_args()

    app = QApplication(sys.argv)
    window = ModernDesktopRichEditor(args.template)
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
