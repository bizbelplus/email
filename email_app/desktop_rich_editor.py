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
        """Build modern toolbar with categorized tools."""
        # Formatting toolbar
        fmt_toolbar = self._create_toolbar("Форматирование", "fmt")
        self._add_button(fmt_toolbar, "𝐁 Bold", self.toggle_bold, "Ctrl+B")
        self._add_button(fmt_toolbar, "𝘐 Italic", self.toggle_italic, "Ctrl+I")
        self._add_button(fmt_toolbar, "𝗨 Underline", self.toggle_underline, "Ctrl+U")
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

        fmt_toolbar.addSeparator()
        self._add_button(fmt_toolbar, "🎨 Color", self.change_text_color, "Ctrl+Shift+C")
        self._add_button(fmt_toolbar, "🖌️ Highlight", self.change_bg_color, "Ctrl+Shift+H")

        # List & Alignment toolbar
        list_toolbar = self._create_toolbar("Списки & Выравнивание", "list")
        self._add_button(list_toolbar, "• Bullet", self.insert_bullet_list)
        self._add_button(list_toolbar, "1. Numbered", self.insert_numbered_list)
        list_toolbar.addSeparator()
        self._add_button(list_toolbar, "⬅ Align Left", self.align_left)
        self._add_button(list_toolbar, "⬇ Align Center", self.align_center)
        self._add_button(list_toolbar, "➡ Align Right", self.align_right)

        # Insert toolbar
        insert_toolbar = self._create_toolbar("Вставить", "insert")
        self._add_button(insert_toolbar, "🔗 Link", self.insert_link, "Ctrl+L")
        self._add_button(insert_toolbar, "🖼️ Image", self.insert_image)
        self._add_button(insert_toolbar, "📊 Table", self.insert_table)
        self._add_button(insert_toolbar, "━━ Line", self.insert_horizontal_rule)
        insert_toolbar.addSeparator()
        self._add_button(insert_toolbar, "⚙️ Variables", self.insert_variables)

        # View toolbar
        view_toolbar = self._create_toolbar("Вид", "view")
        self._add_button(view_toolbar, "💾 Save", self.save_template, "Ctrl+S")
        self._add_button(view_toolbar, "🔄 Reload", self.load_template, "Ctrl+R")
        view_toolbar.addSeparator()
        self._add_button(view_toolbar, "</> Source", self.toggle_source_mode, "Ctrl+`")

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
            html = self.editor.toHtml()
            self.source_editor.setPlainText(html)
            self.editor.hide()
            self.source_editor.show()
            self.mode_label.setText("</> Исходный код")
            self.status_label.setText("✓ Режим редактирования HTML")
        else:
            # Switch back to WYSIWYG
            html = self.source_editor.toPlainText()
            try:
                self.editor.setHtml(html)
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
        self.editor.setHtml(content)
        self.source_editor.setPlainText(content)
        self.status_label.setText(f"✓ Загружено: {self.template_path}")
        self.update_char_count()

    def save_template(self) -> None:
        """Save template to file."""
        content = self.source_editor.toPlainText() if self.source_mode else self.editor.toHtml()
        self.template_path.write_text(content, encoding="utf-8")
        self.status_label.setText(f"✓ Сохранено: {self.template_path}")

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
