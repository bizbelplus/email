from __future__ import annotations

from .tinymce_editor import EDITOR_HTML, RichEditorError, RichTemplateEditorServer

'''
<html lang=\"ru\">
  <head>
    <meta charset=\"utf-8\">
    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
    <title>Визуальный редактор шаблона</title>
    <link rel=\"stylesheet\" href=\"https://cdn.jsdelivr.net/npm/tinymce@7/skins/ui/oxide/skin.min.css\">
    <style>
      :root {
        --bg: #0f172a;
        --panel: #111827;
        --text: #f9fafb;
        --muted: #94a3b8;
        --border: rgba(148, 163, 184, 0.25);
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
        grid-template-rows: auto 1fr auto;
        min-height: 100vh;
      }
      .header,
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
      .workspace {
        display: grid;
        grid-template-columns: 1fr 0;
        min-height: 0;
      }
      .workspace.source-open {
        grid-template-columns: 1fr 420px;
      }
      .editor-pane,
      .source-pane {
        min-height: 0;
      }
      .editor-pane {
        background: #e5e7eb;
        padding: 12px;
      }
      #editorArea {
        visibility: hidden;
      }
      .source-pane {
        display: flex;
        flex-direction: column;
        border-left: 1px solid var(--border);
        background: #020617;
      }
      .source-pane.hidden {
        display: none;
      }
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
      .path {
        color: #93c5fd;
        font-weight: 600;
      }
      @media (max-width: 980px) {
        .workspace,
        .workspace.source-open {
          grid-template-columns: 1fr;
        }
        .source-pane {
          border-left: 0;
          border-top: 1px solid var(--border);
          min-height: 260px;
        }
      }
    </style>
    <script src=\"https://cdn.tiny.cloud/1/no-api-key/tinymce/7/tinymce.min.js\" referrerpolicy=\"origin\"></script>
  </head>
  <body>
    <div class=\"layout\">
      <div class=\"header\">
        <h1>Визуальный редактор письма</h1>
        <p>Редактирование как в Word: жирный, курсив, шрифты, цвет, ссылки, фото, таблицы. Текущий файл: <span id=\"templatePath\" class=\"path\">—</span></p>
      </div>
      <div id=\"workspace\" class=\"workspace\">
        <div class=\"editor-pane\">
          <textarea id=\"editorArea\"></textarea>
        </div>
        <div id=\"sourcePane\" class=\"source-pane hidden\">
          <div class=\"source-pane-header\">Режим исходного HTML. Здесь удобно править Jinja-блоки и условные конструкции.</div>
          <textarea id=\"sourceEditor\"></textarea>
        </div>
      </div>
      <div id=\"status\" class=\"footer\">Загрузка редактора...</div>
    </div>

    <script>
      const workspace = document.getElementById('workspace');
      const sourcePane = document.getElementById('sourcePane');
      const sourceEditor = document.getElementById('sourceEditor');
      const statusNode = document.getElementById('status');
      let sourceMode = false;
      let currentHeadHtml = '<meta charset="utf-8">';
      let currentHtmlAttributes = ' lang="ru"';
      let currentBodyAttributes = '';
      let editorReadyPromise = null;

      function setStatus(message, isError = false) {
        statusNode.textContent = message;
        statusNode.style.color = isError ? '#fca5a5' : '#94a3b8';
      }

      function escapeAttribute(value) {
        return String(value)
          .replace(/&/g, '&amp;')
          .replace(/"/g, '&quot;');
      }

      function attributesToString(element) {
        if (!element || !element.attributes) {
          return '';
        }
        return Array.from(element.attributes)
          .map((attribute) => ` ${attribute.name}="${escapeAttribute(attribute.value)}"`)
          .join('');
      }

      function getEditor() {
        return window.tinymce ? window.tinymce.get('editorArea') : null;
      }

      async function ensureEditor() {
        if (editorReadyPromise) {
          return editorReadyPromise;
        }

        editorReadyPromise = new Promise((resolve, reject) => {
          if (!window.tinymce) {
            reject(new Error('TinyMCE не загрузился. Проверьте доступ к интернету.'));
            return;
          }

          tinymce.init({
            selector: '#editorArea',
            menubar: false,
            height: window.innerHeight - 170,
            branding: false,
            promotion: false,
            plugins: 'lists link image table code autoresize',
            toolbar: 'undo redo | blocks fontfamily fontsize | bold italic underline forecolor backcolor | alignleft aligncenter alignright | bullist numlist | link image table | cidimage placeholders | htmlsource reloadtemplate savetemplate | removeformat',
            font_family_formats: 'Arial=Arial,Helvetica,sans-serif;Verdana=Verdana,Geneva,sans-serif;Tahoma=Tahoma,Geneva,sans-serif;Georgia=Georgia,serif;Times New Roman=Times New Roman,Times,serif;Courier New=Courier New,Courier,monospace',
            content_style: 'body { font-family: Arial, sans-serif; font-size: 16px; color: #111827; padding: 24px; }',
            automatic_uploads: false,
            file_picker_types: 'image',
            file_picker_callback: (callback, value, meta) => {
              if (meta.filetype !== 'image') {
                return;
              }
              const input = document.createElement('input');
              input.type = 'file';
              input.accept = 'image/*';
              input.onchange = () => {
                const file = input.files && input.files[0];
                if (!file) {
                  return;
                }
                const reader = new FileReader();
                reader.onload = () => {
                  callback(reader.result, { alt: file.name });
                  setStatus('Фото вставлено в письмо.');
                };
                reader.readAsDataURL(file);
              };
              input.click();
            },
            setup: (editor) => {
              editor.ui.registry.addButton('cidimage', {
                text: 'CID-картинка',
                onAction: () => insertCidImage(),
              });
              editor.ui.registry.addMenuButton('placeholders', {
                text: 'Переменные',
                fetch: (callback) => {
                  callback([
                    { type: 'menuitem', text: 'subject', onAction: () => insertPlaceholder('{{ subject }}') },
                    { type: 'menuitem', text: 'recipient.email', onAction: () => insertPlaceholder('{{ recipient.email }}') },
                    { type: 'menuitem', text: 'recipient.name', onAction: () => insertPlaceholder('{{ recipient.name }}') },
                    { type: 'menuitem', text: 'headline', onAction: () => insertPlaceholder('{{ headline }}') },
                    { type: 'menuitem', text: 'preheader', onAction: () => insertPlaceholder('{{ preheader }}') },
                    { type: 'menuitem', text: 'body_intro', onAction: () => insertPlaceholder('{{ body_intro }}') },
                    { type: 'menuitem', text: 'body_text', onAction: () => insertPlaceholder('{{ body_text }}') },
                    { type: 'menuitem', text: 'cta_label', onAction: () => insertPlaceholder('{{ cta_label }}') },
                    { type: 'menuitem', text: 'cta_url', onAction: () => insertPlaceholder('{{ cta_url }}') },
                  ]);
                },
              });
              editor.ui.registry.addButton('htmlsource', {
                text: 'HTML-код',
                onAction: () => toggleSource(),
              });
              editor.ui.registry.addButton('reloadtemplate', {
                text: 'Перезагрузить',
                onAction: () => loadTemplate(),
              });
              editor.ui.registry.addButton('savetemplate', {
                text: 'Сохранить',
                onAction: () => saveTemplate(),
              });
              editor.on('init', () => {
                document.getElementById('editorArea').style.visibility = 'visible';
                resolve(editor);
              });
            },
          }).catch(reject);
        });

        return editorReadyPromise;
      }

      function exportDocumentHtml() {
        const editor = getEditor();
        const bodyHtml = editor ? editor.getContent() : '';
        return '<!doctype html>\n<html' + currentHtmlAttributes + '>\n<head>\n' + currentHeadHtml + '\n</head>\n<body' + currentBodyAttributes + '>\n' + bodyHtml + '\n</body>\n</html>\n';
      }

      async function setEditorHtml(html) {
        const editor = await ensureEditor();
        const parser = new DOMParser();
        const doc = parser.parseFromString(html, 'text/html');
        currentHeadHtml = (doc.head && doc.head.innerHTML.trim()) || '<meta charset="utf-8">';
        currentHtmlAttributes = attributesToString(doc.documentElement) || ' lang="ru"';
        currentBodyAttributes = attributesToString(doc.body);
        editor.setContent((doc.body && doc.body.innerHTML.trim()) || '<p><br></p>');
        editor.focus();
        setStatus('Редактор готов. Можно форматировать письмо.');
      }

      async function loadTemplate() {
        try {
          setStatus('Загрузка шаблона...');
          await ensureEditor();
          const response = await fetch('/api/template');
          const payload = await response.json();
          if (!response.ok) {
            throw new Error(payload.error || 'Ошибка загрузки шаблона');
          }
          document.getElementById('templatePath').textContent = payload.path;
          await setEditorHtml(payload.html);
          sourceEditor.value = payload.html;
        } catch (error) {
          setStatus(error.message, true);
        }
      }

      function insertPlaceholder(value) {
        const editor = getEditor();
        if (!editor) {
          return;
        }
        editor.insertContent(value);
        editor.focus();
      }

      function insertCidImage() {
        const editor = getEditor();
        if (!editor) {
          return;
        }
        const alias = window.prompt('Введите ключ картинки для inline_images', 'hero');
        if (!alias) {
          return;
        }
        editor.insertContent('<img src="cid:{{ inline_images.' + alias + '.cid }}" alt="' + alias + '" style="max-width:100%;height:auto;border:0;">');
        editor.focus();
        setStatus('CID-картинка вставлена.');
      }

      async function toggleSource() {
        sourceMode = !sourceMode;
        const editor = await ensureEditor();
        const container = editor.getContainer();
        if (sourceMode) {
          sourceEditor.value = exportDocumentHtml();
          workspace.classList.add('source-open');
          sourcePane.classList.remove('hidden');
          container.style.display = 'none';
          setStatus('Режим HTML-кода включён.');
        } else {
          await setEditorHtml(sourceEditor.value);
          workspace.classList.remove('source-open');
          sourcePane.classList.add('hidden');
          container.style.display = '';
          setStatus('Возврат из режима HTML-кода.');
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

      ensureEditor().then(loadTemplate).catch((error) => {
        setStatus(error.message || 'Не удалось загрузить визуальный редактор.', true);
      });
    </script>
  </body>
</html>
    __OLD_RICH_EDITOR__
          case 'fontName':
            handled = applyTextStyle('font-family', value);
            break;
          case 'fontSize':
            handled = applyTextStyle('font-size', fontSizes[String(value)] || '16px');
            break;
          case 'foreColor':
            handled = applyTextStyle('color', value);
            break;
          case 'hiliteColor':
            handled = applyTextStyle('background-color', value);
            break;
          case 'createLink':
            handled = wrapSelection('a', {}, { href: value, target: '_blank' });
            break;
          case 'insertHTML':
            handled = insertHtmlAtSelection(value || '');
            break;
          case 'justifyLeft':
            handled = applyAlignment('left');
            break;
          case 'justifyCenter':
            handled = applyAlignment('center');
            break;
          case 'justifyRight':
            handled = applyAlignment('right');
            break;
          case 'insertUnorderedList':
            handled = insertList(false);
            break;
          case 'insertOrderedList':
            handled = insertList(true);
            break;
          case 'formatBlock':
            handled = applyFormatBlock(value);
            break;
          case 'removeFormat':
            handled = removeFormatting();
            break;
          case 'undo':
          case 'redo':
            document.execCommand(command, false, value);
            handled = true;
            break;
          default:
            document.execCommand(command, false, value);
            handled = true;
            break;
        }

        if (handled) {
          rememberSelection();
        }
      }

      function insertHtml(html) {
        if (sourceMode) {
          sourceEditor.setRangeText(html, sourceEditor.selectionStart, sourceEditor.selectionEnd, 'end');
          return;
        }
        runCommand('insertHTML', html);
      }

      function insertPlaceholder(value) {
        if (!value) {
          return;
        }
        insertHtml(value);
      }

      function insertLink() {
        const url = window.prompt('Введите URL ссылки', 'https://');
        if (!url) {
          return;
        }
        runCommand('createLink', url);
      }

      function insertTable() {
        const rows = Number(window.prompt('Сколько строк?', '2') || '0');
        const cols = Number(window.prompt('Сколько столбцов?', '2') || '0');
        if (!rows || !cols) {
          return;
        }
        const cells = Array.from({ length: rows }, () => (
          '<tr>' + Array.from({ length: cols }, () => '<td style="border:1px solid #cbd5e1;padding:8px;">Текст</td>').join('') + '</tr>'
        )).join('');
        insertHtml('<table style="border-collapse:collapse;width:100%;margin:12px 0;">' + cells + '</table>');
      }

      function insertCidImage() {
        const alias = window.prompt('Введите ключ картинки для inline_images', 'hero');
        if (!alias) {
          return;
        }
        insertHtml('<img src="cid:{{ inline_images.' + alias + '.cid }}" alt="' + alias + '" style="max-width:100%;height:auto;border:0;">');
      }

      function insertBase64Image(event) {
        const file = event.target.files[0];
        if (!file) {
          return;
        }
        const reader = new FileReader();
        reader.onload = () => {
          insertHtml('<img src="' + reader.result + '" alt="' + file.name.replace(/"/g, '') + '" style="max-width:100%;height:auto;border:0;">');
          setStatus('Изображение вставлено как base64. Для боевой рассылки лучше заменить на CID-картинку.');
          event.target.value = '';
        };
        reader.readAsDataURL(file);
      }

      function toggleSource() {
        sourceMode = !sourceMode;
        if (sourceMode) {
          rememberSelection();
          sourceEditor.value = exportDocumentHtml();
          workspace.classList.add('source-open');
          sourcePane.classList.remove('hidden');
          setStatus('Режим HTML-кода включён.');
        } else {
          setEditorHtml(sourceEditor.value);
          workspace.classList.remove('source-open');
          sourcePane.classList.add('hidden');
          setStatus('Возврат из режима HTML-кода.');
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

      editorSurface.addEventListener('focus', () => {
        setStatus('Редактор активен. Можно печатать.');
      });

      editorSurface.addEventListener('click', () => {
        focusEditor();
        rememberSelection();
      });

      editorSurface.addEventListener('keyup', rememberSelection);
      editorSurface.addEventListener('mouseup', rememberSelection);
      editorSurface.addEventListener('input', rememberSelection);

      document.querySelectorAll('button').forEach((button) => {
        button.addEventListener('mousedown', (event) => {
          if (button.id !== 'imagePickerButton') {
            event.preventDefault();
          }
        });
      });

      imagePickerButton.addEventListener('click', () => {
        imagePicker.click();
      });

      loadTemplate();
    </script>
  </body>
</html>
'''
