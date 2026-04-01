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


EDITOR_HTML = r"""<!doctype html>
<html lang="ru">
  <head>
    <meta charset="utf-8">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Визуальный редактор шаблона</title>
    <style>
      * { box-sizing: border-box; margin: 0; padding: 0; }
      html, body { height: 100%; font-family: Arial, sans-serif; background: #0f172a; color: #f9fafb; }
      .layout { display: flex; flex-direction: column; height: 100vh; }
      .header { padding: 12px 16px; background: #111827; border-bottom: 1px solid rgba(148,163,184,0.25); }
      .header h1 { font-size: 18px; margin-bottom: 4px; }
      .header p { font-size: 12px; color: #94a3b8; }
      .toolbar { padding: 8px 12px; background: #1f2937; border-bottom: 1px solid rgba(148,163,184,0.25); display: flex; flex-wrap: wrap; gap: 6px; overflow-y: auto; max-height: 80px; }
      .toolbar button, .toolbar select { padding: 6px 10px; border: 1px solid rgba(148,163,184,0.25); background: #0b1220; color: #f9fafb; border-radius: 6px; font-size: 12px; cursor: pointer; }
      .toolbar button:hover { background: #111827; }
      .workspace { display: flex; flex: 1; overflow: hidden; }
      .editor-pane { flex: 1; background: #e5e7eb; padding: 12px; display: flex; flex-direction: column; }
      #editorFrame { flex: 1; border: 0; background: #ffffff; border-radius: 8px; }
      .source-pane { display: none; flex-direction: column; width: 400px; border-left: 1px solid rgba(148,163,184,0.25); background: #020617; }
      .source-pane.visible { display: flex; }
      .source-header { padding: 10px 12px; font-size: 12px; color: #94a3b8; border-bottom: 1px solid rgba(148,163,184,0.25); }
      #sourceEditor { flex: 1; border: 0; padding: 12px; font-family: monospace; font-size: 12px; color: #e5e7eb; background: #020617; resize: none; }
      .footer { padding: 10px 16px; background: #111827; border-top: 1px solid rgba(148,163,184,0.25); font-size: 12px; color: #94a3b8; display: flex; justify-content: space-between; }
      .dirty { color: #fbbf24; }
      .emoji-picker-overlay { display: none; position: fixed; top: 0; left: 0; right: 0; bottom: 0; background: rgba(0,0,0,0.5); z-index: 9999; }
      .emoji-picker-overlay.visible { display: flex; align-items: center; justify-content: center; }
      .emoji-picker { background: #111827; border: 1px solid rgba(148,163,184,0.25); border-radius: 12px; padding: 16px; box-shadow: 0 20px 60px rgba(0,0,0,0.8); max-width: 400px; max-height: 500px; overflow-y: auto; }
      .emoji-grid { display: grid; grid-template-columns: repeat(8, 1fr); gap: 8px; margin-bottom: 12px; }
      .emoji-btn { background: #0b1220; border: 1px solid rgba(148,163,184,0.25); border-radius: 8px; padding: 8px; font-size: 24px; cursor: pointer; transition: all 0.2s; }
      .emoji-btn:hover { background: #1f2937; transform: scale(1.1); }
      .emoji-category { margin-bottom: 16px; }
      .emoji-category-title { font-size: 12px; color: #94a3b8; margin-bottom: 8px; padding-bottom: 4px; border-bottom: 1px solid rgba(148,163,184,0.15); }
      .emoji-close { text-align: center; }
      .emoji-close button { padding: 8px 16px; background: #0b1220; border: 1px solid rgba(148,163,184,0.25); color: #f9fafb; border-radius: 6px; cursor: pointer; font-size: 12px; }
      .emoji-close button:hover { background: #1f2937; }
      @media (max-width: 980px) { .source-pane { width: 100%; } .workspace { flex-direction: column; } }
    </style>
  </head>
  <body>
    <div class="layout">
      <div class="header">
        <h1>Визуальный редактор письма</h1>
        <p>Файл: <span id="templatePath" style="color:#93c5fd;font-weight:600;">—</span></p>
      </div>
      <div class="toolbar">
        <button data-cmd="bold" title="Жирный"><b>B</b></button>
        <button data-cmd="italic" title="Курсив"><i>I</i></button>
        <button data-cmd="underline" title="Подчёркивание"><u>U</u></button>
        <button data-cmd="strikeThrough"><s>S</s></button>
        <select data-cmd="formatBlock" title="Тип блока">
          <option value="">Блок</option>
          <option value="p">Параграф</option>
          <option value="h1">H1</option>
          <option value="h2">H2</option>
          <option value="h3">H3</option>
        </select>
        <select id="fontSelect" title="Шрифт">
          <option value="">Шрифт</option>
          <option value="Arial">Arial</option>
          <option value="'Helvetica Neue'">Helvetica</option>
          <option value="'Times New Roman'">Times New Roman</option>
          <option value="Georgia">Georgia</option>
          <option value="Garamond">Garamond</option>
          <option value="Verdana">Verdana</option>
          <option value="'Trebuchet MS'">Trebuchet MS</option>
          <option value="Tahoma">Tahoma</option>
          <option value="'Courier New'">Courier New</option>
          <option value="Consolas">Consolas</option>
          <option value="'Comic Sans MS'">Comic Sans</option>
          <option value="Impact">Impact</option>
          <option value="'Palatino Linotype'">Palatino</option>
          <option value="'Book Antiqua'">Book Antiqua</option>
          <option value="'Lucida Console'">Lucida Console</option>
          <option value="'Lucida Sans'">Lucida Sans</option>
          <option value="Calibri">Calibri</option>
          <option value="Cambria">Cambria</option>
          <option value="Candara">Candara</option>
          <option value="Century">Century</option>
          <option value="'Franklin Gothic'">Franklin Gothic</option>
          <option value="'Century Gothic'">Century Gothic</option>
          <option value="Segoe UI">Segoe UI</option>
          <option value="'Apple System'">System</option>
          <option value="Optima">Optima</option>
          <option value="Didot">Didot</option>
        </select>
        <input id="colorPicker" type="color" title="Цвет текста" value="#000000" style="width:36px;height:36px;border:1px solid rgba(148,163,184,0.25);cursor:pointer;border-radius:6px;">
        <input id="bgColorPicker" type="color" title="Цвет фона" value="#ffffff" style="width:36px;height:36px;border:1px solid rgba(148,163,184,0.25);cursor:pointer;border-radius:6px;">
        <button data-cmd="justifyLeft">⬅</button>
        <button data-cmd="justifyCenter">⬆⬇</button>
        <button data-cmd="justifyRight">➡</button>
        <button data-cmd="insertUnorderedList">•</button>
        <button data-cmd="insertOrderedList">1.</button>
        <button id="insertLink">Ссылка</button>
        <button id="insertImage">Фото</button>
        <button id="insertCid">CID</button>
        <button id="insertEmoji">😀</button>
        <button id="insertVar">{{}}</button>
        <button id="toggleSource">HTML</button>
        <button id="reloadBtn">🔄</button>
        <button id="saveBtn" style="background:#2563eb;border-color:#2563eb;">💾 Сохранить</button>
      </div>
      <div class="workspace">
        <div class="editor-pane">
          <iframe id="editorFrame" title="Визуальный редактор"></iframe>
        </div>
        <div id="sourcePane" class="source-pane">
          <div class="source-header">HTML-код</div>
          <textarea id="sourceEditor"></textarea>
        </div>
      </div>
      <div class="footer">
        <span id="status">Инициализация...</span>
        <span id="unsaved"></span>
      </div>
      <div id="emojiOverlay" class="emoji-picker-overlay">
        <div class="emoji-picker" id="emojiPicker"></div>
      </div>
    </div>
    <input id="filePicker" type="file" accept="image/*" style="display:none">
    <script>
      const frame = document.getElementById('editorFrame');
      const sourcePane = document.getElementById('sourcePane');
      const sourceEditor = document.getElementById('sourceEditor');
      const statusNode = document.getElementById('status');
      const unsavedNode = document.getElementById('unsaved');
      let hasChanges = false;
      let sourceMode = false;
      let originalHtml = '';
      let originalHead = '';

      function setStatus(msg, isError) {
        statusNode.textContent = msg;
        statusNode.style.color = isError ? '#fca5a5' : '#94a3b8';
      }

      function markDirty() {
        if (!hasChanges) {
          hasChanges = true;
          unsavedNode.textContent = '● Несохранённые изменения';
          unsavedNode.className = 'dirty';
        }
      }

      function markClean() {
        hasChanges = false;
        unsavedNode.textContent = '';
      }

      function initFrame() {
        try {
          const fdoc = frame.contentDocument;
          fdoc.designMode = 'on';
          fdoc.body.style.margin = '24px';
          fdoc.body.style.fontFamily = 'Arial, sans-serif';
          fdoc.body.style.fontSize = '16px';
          fdoc.body.style.color = '#111827';
          fdoc.body.contentEditable = true;
          fdoc.body.addEventListener('input', function() {
            markDirty();
          });
          frame.contentWindow.focus();
        } catch (e) {
          console.error('initFrame error:', e);
        }
      }

      function getHtml() {
        try {
          const fdoc = frame.contentDocument;
          const bodyHtml = fdoc.body.innerHTML || '<p></p>';
          
          if (!originalHead) {
            return '<!doctype html>\n<html lang="ru">\n<head>\n<meta charset="utf-8">\n</head>\n<body>\n' + bodyHtml + '\n</body>\n</html>';
          }
          
          return '<!doctype html>\n<html lang="ru">\n<head>\n' + originalHead + '\n</head>\n<body>\n' + bodyHtml + '\n</body>\n</html>';
        } catch {
          return originalHtml || sourceEditor.value;
        }
      }

      async function loadTemplate() {
        try {
          setStatus('Загрузка...');
          const res = await fetch('/api/template');
          const data = await res.json();
          if (!res.ok) throw new Error(data.error || 'Ошибка загрузки');
          
          originalHtml = data.html;
          document.getElementById('templatePath').textContent = data.path;
          sourceEditor.value = data.html;
          
          const parser = new DOMParser();
          const doc = parser.parseFromString(data.html, 'text/html');
          
          if (doc.head) {
            originalHead = doc.head.innerHTML;
          }
          
          frame.onload = function() {
            initFrame();
            if (doc.body) {
              frame.contentDocument.body.innerHTML = doc.body.innerHTML;
            }
            markClean();
            setStatus('Готов к редактированию');
          };
          frame.srcdoc = '<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>';
        } catch (err) {
          setStatus(err.message, true);
        }
      }

      async function saveTemplate() {
        try {
          const html = sourceMode ? sourceEditor.value : getHtml();
          setStatus('Сохранение...');
          const res = await fetch('/api/template', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ html: html }),
          });
          const data = await res.json();
          if (!res.ok) throw new Error(data.error || 'Ошибка сохранения');
          
          originalHtml = html;
          const parser = new DOMParser();
          const doc = parser.parseFromString(html, 'text/html');
          if (doc.head) {
            originalHead = doc.head.innerHTML;
          }
          
          markClean();
          setStatus('Сохранено: ' + data.saved_path);
        } catch (err) {
          setStatus(err.message, true);
        }
      }

      function execCmd(cmd, val) {
        try {
          frame.contentWindow.focus();
          frame.contentDocument.execCommand(cmd, false, val || null);
          markDirty();
        } catch (err) {
          console.error('Команда не поддерживается:', cmd);
        }
      }

      document.querySelectorAll('[data-cmd]').forEach(function(btn) {
        btn.addEventListener('click', function() {
          const cmd = btn.getAttribute('data-cmd');
          if (btn.tagName === 'SELECT') {
            if (btn.value) execCmd(cmd, btn.value);
          } else {
            execCmd(cmd);
          }
        });
      });

      document.getElementById('insertLink').addEventListener('click', function() {
        const url = prompt('URL:');
        if (url) execCmd('createLink', url);
      });

      document.getElementById('insertImage').addEventListener('click', function() {
        document.getElementById('filePicker').click();
      });

      document.getElementById('filePicker').addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (!file) return;
        const reader = new FileReader();
        reader.onload = function() {
          execCmd('insertImage', reader.result);
          setStatus('Фото вставлено');
        };
        reader.readAsDataURL(file);
        e.target.value = '';
      });

      document.getElementById('insertCid').addEventListener('click', function() {
        const alias = prompt('Ключ картинки (например: hero)');
        if (alias) {
          frame.contentWindow.focus();
          const img = '<img src="cid:{{ inline_images.' + alias + '.cid }}" alt="' + alias + '" style="max-width:100%;height:auto;">';
          frame.contentDocument.execCommand('insertHTML', false, img);
          markDirty();
        }
      });

      document.getElementById('insertVar').addEventListener('click', function() {
        const vars = ['subject', 'recipient.name', 'recipient.email', 'headline', 'preheader', 'body_intro', 'body_text', 'cta_label', 'cta_url'];
        const msg = 'Переменная:\n' + vars.join('\n');
        const v = prompt(msg);
        if (v) {
          frame.contentWindow.focus();
          frame.contentDocument.execCommand('insertHTML', false, '{{ ' + v + ' }}');
          markDirty();
        }
      });

      document.getElementById('fontSelect').addEventListener('change', function() {
        if (this.value) {
          execCmd('fontName', this.value);
          this.value = '';
        }
      });

      document.getElementById('colorPicker').addEventListener('input', function() {
        execCmd('foreColor', this.value);
      });

      document.getElementById('bgColorPicker').addEventListener('input', function() {
        execCmd('backColor', this.value);
      });

      document.getElementById('insertEmoji').addEventListener('click', function() {
        showEmojiPicker();
      });

      function showEmojiPicker() {
        const overlay = document.getElementById('emojiOverlay');
        const picker = document.getElementById('emojiPicker');
        picker.innerHTML = '';

        const categories = {
          'Улыбки': ['😀', '😃', '😄', '😁', '😆', '😅', '🤣', '😂', '🙂', '🙃', '😉', '😊', '😇', '🥰', '😍', '🤩', '😘', '😗', '😚', '😙', '🥲', '😋', '😛', '😜', '🤪', '😌', '😔', '😑', '😐', '😏', '🥱'],
          'Любовь': ['❤️', '🧡', '💛', '💚', '💙', '💜', '🖤', '🤍', '🤎', '💔', '💕', '💞', '💓', '💗', '💖', '💘', '💝', '💟', '💌'],
          'Жесты': ['👋', '🤚', '🖐️', '✋', '🖖', '👌', '🤌', '🤏', '✌️', '🤞', '🫰', '🤟', '🤘', '🤙', '👍', '👎', '✊', '👊', '🤛', '🤜', '👏', '🙌', '👐', '🤲', '🤝', '🤜', '💪', '🦾', '🦿', '🦵', '🦶'],
          'Звезды': ['⭐', '🌟', '✨', '💫', '🌠', '🔆', '☀️', '🌤️', '⛅', '🌥️', '☁️', '🌦️', '🌧️', '⛈️', '🌩️', '🌨️', '❄️', '☃️', '⛄', '🌬️', '💨', '💧', '💦'],
          'Дела': ['📢', '📣', '📯', '🔔', '🔕', '📻', '📱', '📞', '☎️', '📟', '📠', '📧', '📨', '📩', '📤', '📥', '📦', '📫', '📪', '📬', '📭', '📮', '✏️', '✒️', '🖋️', '🖊️', '📝'],
          'Объекты': ['🎁', '🎀', '🎈', '🎉', '🎊', '🎎', '🏆', '🏅', '⚽', '⚾', '🥎', '🎾', '🏐', '🏈', '🏉', '🎯', '🎳', '🎣', '🎬', '🎤', '🎧', '🎼', '🎹', '🎸', '🎺', '🎷'],
          'Еда': ['🍕', '🍔', '🍟', '🌭', '🥪', '🌮', '🌯', '🥙', '🧆', '🍗', '🍖', '🌶️', '🍝', '🍜', '🍲', '🍛', '🍣', '🍱', '🥟', '🦪', '🍤', '🍙', '🍚', '🍘', '🍥', '🥠', '🥮', '🍢', '🍡', '🍧', '🍨', '🍦', '🍰', '🎂', '🧁', '🍮', '🍭', '🍬', '🍫', '🍿', '🍩', '🍪', '🌰', '🍯', '🥛', '☕', '🍵', '🍶', '🍾', '🍷', '🍸', '🍹', '🍺', '🍻'],
          'Путешествия': ['✈️', '🚀', '🛸', '🚁', '🛶', '⛵', '🚤', '🛳️', '⛴️', '🛥️', '🚢', '🚧', '🚨', '🚔', '🚍', '🚘', '🚖', '🚡', '🚠', '🎡', '🎢', '🏰', '🎠', '⛲', '⛺', '🏖️', '🏝️', '🌋', '⛰️', '🏔️', '🗻', '🗽', '🗼', '🏛️', '⌚', '📱', '📲', '💻', '⌨️', '🖥️', '🖨️', '🖱️', '🖲️', '🕹️'],
          'Животные': ['🐶', '🐱', '🐭', '🐹', '🐰', '🦊', '🐻', '🐼', '🐨', '🐯', '🦁', '🐮', '🐷', '🐽', '🐸', '🐵', '🙈', '🙉', '🙊', '🐒', '🐔', '🐧', '🐦', '🐤', '🦆', '🦅', '🦉', '🦇', '🐺', '🐗', '🐴', '🦄', '🐝', '🪱', '🐛', '🦋', '🐌', '🐞', '🐜', '🪰', '🐢', '🐍', '🦎', '🦖', '🦕', '🐙', '🦑', '🦐', '🦞', '🦀', '🐡', '🐠', '🐟', '🐬', '🐳', '🐋', '🦈', '🐊', '🐅', '🐆', '🦓', '🦍', '🦧', '🐘', '🦛', '🦏', '🐪', '🐫', '🦒', '🦘', '🐃', '🐂', '🐄', '🐎', '🐖', '🐏', '🐑', '🧔', '🐐'],
          'Природа': ['🌲', '🌳', '🌴', '🌵', '🌾', '🌿', '☘️', '🍀', '🍁', '🍂', '🍃', '🌺', '🌻', '🌹', '🥀', '🌷', '🌱', '🌲', '🏵️', '💐', '🌞', '🌝', '🌛', '🌜', '⭐', '🌟', '✨', '⚡', '☄️', '💥', '🔥', '🌪️', '🌈', '☀️', '🌤️', '⛅', '🌥️', '☁️'],
        };

        for (const [category, emojis] of Object.entries(categories)) {
          const catDiv = document.createElement('div');
          catDiv.className = 'emoji-category';
          
          const title = document.createElement('div');
          title.className = 'emoji-category-title';
          title.textContent = category;
          catDiv.appendChild(title);
          
          const grid = document.createElement('div');
          grid.className = 'emoji-grid';
          
          emojis.forEach(function(emoji) {
            const btn = document.createElement('button');
            btn.className = 'emoji-btn';
            btn.textContent = emoji;
            btn.addEventListener('click', function() {
              frame.contentWindow.focus();
              frame.contentDocument.execCommand('insertHTML', false, emoji);
              markDirty();
              overlay.classList.remove('visible');
            });
            grid.appendChild(btn);
          });
          
          catDiv.appendChild(grid);
          picker.appendChild(catDiv);
        }

        const closeDiv = document.createElement('div');
        closeDiv.className = 'emoji-close';
        const closeBtn = document.createElement('button');
        closeBtn.textContent = 'Закрыть';
        closeBtn.addEventListener('click', function() {
          overlay.classList.remove('visible');
        });
        closeDiv.appendChild(closeBtn);
        picker.appendChild(closeDiv);

        overlay.classList.add('visible');
      }

      document.getElementById('emojiOverlay').addEventListener('click', function(e) {
        if (e.target === this) {
          this.classList.remove('visible');
        }
      });

      document.getElementById('toggleSource').addEventListener('click', function() {
        sourceMode = !sourceMode;
        if (sourceMode) {
          sourceEditor.value = getHtml();
          sourcePane.classList.add('visible');
          frame.style.display = 'none';
          setStatus('HTML-режим включен');
        } else {
          try {
            const parser = new DOMParser();
            const doc = parser.parseFromString(sourceEditor.value, 'text/html');
            originalHtml = sourceEditor.value;
            if (doc.head) {
              originalHead = doc.head.innerHTML;
            }
            frame.contentDocument.body.innerHTML = doc.body.innerHTML;
          } catch (err) {
            setStatus('Ошибка в HTML', true);
            return;
          }
          sourcePane.classList.remove('visible');
          frame.style.display = '';
          markClean();
          setStatus('Вернулись к визуальному редактору');
        }
      });

      document.getElementById('reloadBtn').addEventListener('click', loadTemplate);
      document.getElementById('saveBtn').addEventListener('click', saveTemplate);

      loadTemplate();
    </script>
  </body>
</html>
"""
