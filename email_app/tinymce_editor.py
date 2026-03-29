from __future__ import annotations

import json
import threading
import time
import webbrowser
from http import HTTPStatus
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import urlparse


class RichEditorError(Exception):
    """Ошибка визуального редактора шаблонов."""


class RichTemplateEditorServer:
    def __init__(self, base_dir: Path) -> None:
        self.base_dir = base_dir.resolve()
        self._template_path: Path | None = None
        self._server: ThreadingHTTPServer | None = None
        self._thread: threading.Thread | None = None
        self._lock = threading.Lock()
        self._port: int | None = None

    def open_template(self, template_path: Path) -> str:
        template_path = template_path.resolve()
        with self._lock:
            self._template_path = template_path

        if self._server is None:
            self._start_server()

        if self._port is None:
            raise RichEditorError("Не удалось определить порт визуального редактора")

        url = f"http://127.0.0.1:{self._port}/?t={int(time.time() * 1000)}"
        webbrowser.open(url)
        return url

    def _current_template_path(self) -> Path:
        with self._lock:
            if self._template_path is None:
                raise RichEditorError("Шаблон для визуального редактора не выбран")
            return self._template_path

    def _load_template(self) -> dict[str, str]:
        template_path = self._current_template_path()
        if template_path.exists():
            html = template_path.read_text(encoding="utf-8")
        else:
            html = self._default_template_html()

        try:
            relative = template_path.relative_to(self.base_dir)
            path_text = relative.as_posix()
        except ValueError:
            path_text = str(template_path)

        return {
            "path": path_text,
            "html": html,
        }

    def _save_template(self, html: str) -> Path:
        template_path = self._current_template_path()
        template_path.parent.mkdir(parents=True, exist_ok=True)
        template_path.write_text(html.rstrip() + "\n", encoding="utf-8")
        return template_path

    def _start_server(self) -> None:
        editor = self

        class Handler(BaseHTTPRequestHandler):
            def do_GET(self) -> None:  # noqa: N802
                parsed = urlparse(self.path)
                if parsed.path == "/":
                    self._send_response(HTTPStatus.OK, EDITOR_HTML.encode("utf-8"), "text/html; charset=utf-8")
                    return
                if parsed.path == "/api/template":
                    payload = json.dumps(editor._load_template(), ensure_ascii=False).encode("utf-8")
                    self._send_response(HTTPStatus.OK, payload, "application/json; charset=utf-8")
                    return
                self._send_response(HTTPStatus.NOT_FOUND, b"Not found", "text/plain; charset=utf-8")

            def do_POST(self) -> None:  # noqa: N802
                parsed = urlparse(self.path)
                if parsed.path != "/api/template":
                    self._send_response(HTTPStatus.NOT_FOUND, b"Not found", "text/plain; charset=utf-8")
                    return

                length = int(self.headers.get("Content-Length", "0"))
                raw_body = self.rfile.read(length)
                try:
                    payload = json.loads(raw_body.decode("utf-8"))
                except json.JSONDecodeError:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "Некорректный JSON"})
                    return

                html = str(payload.get("html", "")).strip()
                if not html:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": "HTML не должен быть пустым"})
                    return

                try:
                    saved_path = editor._save_template(html)
                except RichEditorError as error:
                    self._send_json(HTTPStatus.BAD_REQUEST, {"error": str(error)})
                    return

                self._send_json(HTTPStatus.OK, {"ok": True, "saved_path": str(saved_path)})

            def log_message(self, format: str, *args: object) -> None:  # noqa: A003
                return

            def _send_json(self, status: HTTPStatus, payload: dict[str, object]) -> None:
                body = json.dumps(payload, ensure_ascii=False).encode("utf-8")
                self._send_response(status, body, "application/json; charset=utf-8")

            def _send_response(self, status: HTTPStatus, body: bytes, content_type: str) -> None:
                self.send_response(status)
                self.send_header("Content-Type", content_type)
                self.send_header("Content-Length", str(len(body)))
                self.end_headers()
                self.wfile.write(body)

        self._server = ThreadingHTTPServer(("127.0.0.1", 0), Handler)
        self._port = self._server.server_address[1]
        self._thread = threading.Thread(target=self._server.serve_forever, daemon=True)
        self._thread.start()

    @staticmethod
    def _default_template_html() -> str:
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


EDITOR_HTML = """<!doctype html>
<html lang=\"ru\">
  <head>
    <meta charset=\"utf-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
    <title>Визуальный редактор шаблона</title>
    <style>
      :root {
        --bg: #0f172a;
        --panel: #111827;
        --panel-2: #1f2937;
        --text: #f9fafb;
        --muted: #94a3b8;
        --border: rgba(148, 163, 184, 0.25);
        --accent: #2563eb;
      }
      * { box-sizing: border-box; }
      body {
        margin: 0;
        font-family: Inter, Arial, sans-serif;
        background: var(--bg);
        color: var(--text);
      }
      .layout {
        display: grid;
        grid-template-rows: auto auto 1fr auto;
        min-height: 100vh;
      }
      .header,
      .toolbar,
      .footer {
        padding: 12px 16px;
        background: var(--panel);
        border-bottom: 1px solid var(--border);
      }
      .footer {
        border-bottom: 0;
        border-top: 1px solid var(--border);
        color: var(--muted);
        font-size: 13px;
      }
      .header h1 {
        margin: 0 0 4px;
        font-size: 18px;
      }
      .header p {
        margin: 0;
        color: var(--muted);
        font-size: 13px;
      }
      .toolbar {
        display: flex;
        flex-wrap: wrap;
        gap: 8px;
        align-items: center;
        background: var(--panel-2);
      }
      .toolbar button,
      .toolbar select,
      .toolbar input[type=\"color\"] {
        border: 1px solid var(--border);
        background: #0b1220;
        color: var(--text);
        border-radius: 10px;
        padding: 8px 10px;
        font-size: 13px;
      }
      .toolbar button { cursor: pointer; }
      .toolbar button.primary { background: var(--accent); border-color: var(--accent); }
      .workspace {
        display: grid;
        grid-template-columns: 1fr 0;
        min-height: 0;
      }
      .workspace.source-open { grid-template-columns: 1fr 420px; }
      .editor-pane,
      .source-pane { min-height: 0; }
      .editor-pane {
        background: #e5e7eb;
        padding: 12px;
      }
      #editorSurface {
        min-height: calc(100vh - 240px);
        background: white;
        color: #111827;
        padding: 24px;
        outline: none;
        border-radius: 12px;
        overflow: auto;
      }
      #editorSurface:focus { box-shadow: inset 0 0 0 2px rgba(37, 99, 235, 0.35); }
      .source-pane {
        display: flex;
        flex-direction: column;
        border-left: 1px solid var(--border);
        background: #020617;
      }
      .source-pane.hidden { display: none; }
      .source-pane-header {
        padding: 12px 14px;
        font-size: 13px;
        color: var(--muted);
        border-bottom: 1px solid var(--border);
      }
      #sourceEditor {
        flex: 1;
        width: 100%;
        border: 0;
        resize: none;
        padding: 14px;
        font-family: ui-monospace, SFMono-Regular, Menlo, monospace;
        font-size: 13px;
        line-height: 1.5;
        color: #e5e7eb;
        background: #020617;
      }
      .path { color: #93c5fd; font-weight: 600; }
      @media (max-width: 980px) {
        .workspace,
        .workspace.source-open { grid-template-columns: 1fr; }
        .source-pane {
          border-left: 0;
          border-top: 1px solid var(--border);
          min-height: 260px;
        }
      }
    </style>
  </head>
  <body>
    <div class=\"layout\">
      <div class=\"header\">
        <h1>Визуальный редактор письма</h1>
        <p>Офлайн-редактор. Текущий файл: <span id=\"templatePath\" class=\"path\">—</span></p>
      </div>
      <div class=\"toolbar\">
        <button type=\"button\" data-cmd=\"bold\"><b>B</b></button>
        <button type=\"button\" data-cmd=\"italic\"><i>I</i></button>
        <button type=\"button\" data-cmd=\"underline\"><u>U</u></button>
        <button type=\"button\" data-cmd=\"strikeThrough\"><s>S</s></button>
        <select id=\"fontName\">
          <option value=\"\">Шрифт</option>
          <option value=\"Arial\">Arial</option>
          <option value=\"Verdana\">Verdana</option>
          <option value=\"Tahoma\">Tahoma</option>
          <option value=\"Georgia\">Georgia</option>
          <option value=\"Times New Roman\">Times New Roman</option>
          <option value=\"Courier New\">Courier New</option>
        </select>
        <select id=\"fontSize\">
          <option value=\"\">Размер</option>
          <option value=\"2\">13</option>
          <option value=\"3\">16</option>
          <option value=\"4\">18</option>
          <option value=\"5\">24</option>
          <option value=\"6\">32</option>
        </select>
        <input id=\"foreColor\" type=\"color\" value=\"#111827\" title=\"Цвет текста\">
        <input id=\"backColor\" type=\"color\" value=\"#ffffff\" title=\"Фон текста\">
        <button type=\"button\" data-cmd=\"justifyLeft\">Лево</button>
        <button type=\"button\" data-cmd=\"justifyCenter\">Центр</button>
        <button type=\"button\" data-cmd=\"justifyRight\">Право</button>
        <button type=\"button\" data-cmd=\"insertUnorderedList\">• Список</button>
        <button type=\"button\" data-cmd=\"insertOrderedList\">1. Список</button>
        <button id=\"linkButton\" type=\"button\">Ссылка</button>
        <button id=\"imageButton\" type=\"button\">Фото</button>
        <button id=\"cidButton\" type=\"button\">CID-картинка</button>
        <button id=\"tableButton\" type=\"button\">Таблица</button>
        <button id=\"varsButton\" type=\"button\">Переменные</button>
        <button id=\"sourceButton\" type=\"button\">HTML-код</button>
        <button id=\"reloadButton\" type=\"button\">Перезагрузить</button>
        <button id=\"saveButton\" type=\"button\" class=\"primary\">Сохранить</button>
      </div>
      <div id=\"workspace\" class=\"workspace\">
        <div class=\"editor-pane\">
          <div id=\"editorSurface\" contenteditable=\"true\" spellcheck=\"true\"></div>
        </div>
        <div id=\"sourcePane\" class=\"source-pane hidden\">
          <div class=\"source-pane-header\">Режим исходного HTML</div>
          <textarea id=\"sourceEditor\"></textarea>
        </div>
      </div>
      <div id=\"status\" class=\"footer\">Загрузка редактора...</div>
      <input id=\"imagePicker\" type=\"file\" accept=\"image/*\" style=\"display:none\">
    </div>

    <script>
      const workspace = document.getElementById('workspace');
      const sourcePane = document.getElementById('sourcePane');
      const sourceEditor = document.getElementById('sourceEditor');
      const editorSurface = document.getElementById('editorSurface');
      const statusNode = document.getElementById('status');
      const imagePicker = document.getElementById('imagePicker');
      let sourceMode = false;
      let currentHeadHtml = '<meta charset="utf-8">';
      let currentHtmlAttributes = ' lang="ru"';
      let currentBodyAttributes = '';
      let savedRange = null;

      function setStatus(message, isError = false) {
        statusNode.textContent = message;
        statusNode.style.color = isError ? '#fca5a5' : '#94a3b8';
      }

      function escapeAttribute(value) {
        return String(value).replace(/&/g, '&amp;').replace(/"/g, '&quot;');
      }

      function attributesToString(element) {
        if (!element || !element.attributes) {
          return '';
        }
        return Array.from(element.attributes)
          .map((attribute) => ` ${attribute.name}="${escapeAttribute(attribute.value)}"`)
          .join('');
      }

      function focusEditor() {
        editorSurface.focus();
      }

      function rememberSelection() {
        if (sourceMode) {
          return;
        }
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0) {
          return;
        }
        const range = selection.getRangeAt(0);
        if (editorSurface.contains(range.commonAncestorContainer)) {
          savedRange = range.cloneRange();
        }
      }

      function restoreSelection() {
        focusEditor();
        const selection = window.getSelection();
        if (!selection) {
          return;
        }
        selection.removeAllRanges();
        if (savedRange) {
          selection.addRange(savedRange);
        }
      }

      function execCommand(command, value = null) {
        if (sourceMode) {
          setStatus('Сначала выйдите из режима HTML-кода.', true);
          return;
        }
        restoreSelection();
        document.execCommand('styleWithCSS', false, true);
        document.execCommand(command, false, value);
        rememberSelection();
        focusEditor();
      }

      function insertHtml(html) {
        restoreSelection();
        document.execCommand('insertHTML', false, html);
        rememberSelection();
      }

      function exportDocumentHtml() {
        return '<!doctype html>\n<html' + currentHtmlAttributes + '>\n<head>\n' + currentHeadHtml + '\n</head>\n<body' + currentBodyAttributes + '>\n' + editorSurface.innerHTML + '\n</body>\n</html>\n';
      }

      function setEditorHtml(html) {
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        currentHeadHtml = (doc.head && doc.head.innerHTML.trim()) || '<meta charset="utf-8">';
        currentHtmlAttributes = attributesToString(doc.documentElement) || ' lang="ru"';
        currentBodyAttributes = attributesToString(doc.body);
        editorSurface.innerHTML = (doc.body && doc.body.innerHTML.trim()) || '<p><br></p>';
        focusEditor();
        setStatus('Редактор готов. Форматирование доступно.');
      }

      async function loadTemplate() {
        try {
          setStatus('Загрузка шаблона...');
          const response = await fetch('/api/template');
          const payload = await response.json();
          if (!response.ok) {
            throw new Error(payload.error || 'Ошибка загрузки шаблона');
          }
          document.getElementById('templatePath').textContent = payload.path;
          setEditorHtml(payload.html);
          sourceEditor.value = payload.html;
        } catch (error) {
          setStatus(error.message, true);
        }
      }

      async function saveTemplate() {
        try {
          const html = sourceMode ? sourceEditor.value : exportDocumentHtml();
          setStatus('Сохранение шаблона...');
          const response = await fetch('/api/template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ html }),
          });
          const payload = await response.json();
          if (!response.ok) {
            throw new Error(payload.error || 'Ошибка сохранения шаблона');
          }
          setStatus('Шаблон сохранён: ' + payload.saved_path);
        } catch (error) {
          setStatus(error.message, true);
        }
      }

      function insertLink() {
        const url = window.prompt('Введите URL', 'https://');
        if (!url) {
          return;
        }
        execCommand('createLink', url);
      }

      function insertCidImage() {
        const alias = window.prompt('Введите ключ картинки для inline_images', 'hero');
        if (!alias) {
          return;
        }
        insertHtml('<img src="cid:{{ inline_images.' + alias + '.cid }}" alt="' + alias + '" style="max-width:100%;height:auto;border:0;">');
        setStatus('CID-картинка вставлена.');
      }

      function insertTable() {
        const rows = Number(window.prompt('Сколько строк?', '2') || '0');
        const cols = Number(window.prompt('Сколько столбцов?', '2') || '0');
        if (!rows || !cols) {
          return;
        }
        const cells = Array.from({ length: rows }, () => '<tr>' + Array.from({ length: cols }, () => '<td style="border:1px solid #cbd5e1;padding:8px;">Текст</td>').join('') + '</tr>').join('');
        insertHtml('<table style="border-collapse:collapse;width:100%;margin:12px 0;">' + cells + '</table>');
      }

      function insertVariable() {
        const value = window.prompt('Переменная: subject, recipient.email, recipient.name, headline, preheader, body_intro, body_text, cta_label, cta_url', 'recipient.name');
        if (!value) {
          return;
        }
        insertHtml('{{ ' + value + ' }}');
      }

      function toggleSource() {
        sourceMode = !sourceMode;
        if (sourceMode) {
          sourceEditor.value = exportDocumentHtml();
          workspace.classList.add('source-open');
          sourcePane.classList.remove('hidden');
          editorSurface.style.display = 'none';
          setStatus('Режим HTML-кода включён.');
        } else {
          setEditorHtml(sourceEditor.value);
          workspace.classList.remove('source-open');
          sourcePane.classList.add('hidden');
          editorSurface.style.display = '';
          setStatus('Возврат из режима HTML-кода.');
        }
      }

      document.querySelectorAll('[data-cmd]').forEach((button) => {
        button.addEventListener('mousedown', (event) => event.preventDefault());
        button.addEventListener('click', () => execCommand(button.dataset.cmd));
      });

      document.getElementById('fontName').addEventListener('change', (event) => {
        if (event.target.value) {
          execCommand('fontName', event.target.value);
        }
      });

      document.getElementById('fontSize').addEventListener('change', (event) => {
        if (event.target.value) {
          execCommand('fontSize', event.target.value);
        }
      });

      document.getElementById('foreColor').addEventListener('input', (event) => {
        execCommand('foreColor', event.target.value);
      });

      document.getElementById('backColor').addEventListener('input', (event) => {
        execCommand('hiliteColor', event.target.value);
      });

      document.getElementById('linkButton').addEventListener('mousedown', (event) => event.preventDefault());
      document.getElementById('linkButton').addEventListener('click', insertLink);
      document.getElementById('cidButton').addEventListener('mousedown', (event) => event.preventDefault());
      document.getElementById('cidButton').addEventListener('click', insertCidImage);
      document.getElementById('tableButton').addEventListener('mousedown', (event) => event.preventDefault());
      document.getElementById('tableButton').addEventListener('click', insertTable);
      document.getElementById('varsButton').addEventListener('mousedown', (event) => event.preventDefault());
      document.getElementById('varsButton').addEventListener('click', insertVariable);
      document.getElementById('sourceButton').addEventListener('click', toggleSource);
      document.getElementById('reloadButton').addEventListener('click', loadTemplate);
      document.getElementById('saveButton').addEventListener('click', saveTemplate);
      document.getElementById('imageButton').addEventListener('mousedown', (event) => event.preventDefault());
      document.getElementById('imageButton').addEventListener('click', () => imagePicker.click());

      imagePicker.addEventListener('change', (event) => {
        const file = event.target.files && event.target.files[0];
        if (!file) {
          return;
        }
        const reader = new FileReader();
        reader.onload = () => {
          insertHtml('<img src="' + reader.result + '" alt="' + file.name.replace(/"/g, '') + '" style="max-width:100%;height:auto;border:0;">');
          setStatus('Фото вставлено в письмо.');
          imagePicker.value = '';
        };
        reader.readAsDataURL(file);
      });

      editorSurface.addEventListener('mouseup', rememberSelection);
      editorSurface.addEventListener('keyup', rememberSelection);
      editorSurface.addEventListener('focus', rememberSelection);
      editorSurface.addEventListener('input', rememberSelection);

      loadTemplate();
    </script>
  </body>
</html>
"""
