# Email app на Python


Современное и безопасное SMTP-приложение для **легитимной** рассылки HTML-писем по вашему списку получателей.
Поддерживает прокси, пул SMTP-аккаунтов, вложения, inline-изображения, персонализацию, лимиты, антидубли, статистику, визуальный редактор и многое другое.


**Возможности:**
- SMTP (TLS/SSL)
- Прокси (SOCKS5, HTTP(S), с авторизацией)
- Пул SMTP-аккаунтов (TXT/CSV, ротация)
- HTML-шаблоны через Jinja2
- Персонализация по CSV и переменным
- Вложения и inline-изображения (cid:)
- Встроенный визуальный редактор (modern GUI)
- Preflight-чеклист и валидация шаблонов
- Антидубли (skip_previously_sent, dedupe)
- Лимиты отправки (rate_limit_per_minute и др.)
- Логирование и история отправок (CSV, JSONL)
- Статистика, фильтрация, экспорт
- Импорт/экспорт очередей кампаний (JSON/CSV)
- Пресеты кампаний (YAML)
- HTML preview и предпросмотр в GUI
- Поддержка Reply-To, BCC, delay, retry

> 📖 **Подробное пошаговое руководство:** [GUIDE.md](GUIDE.md)


## Быстрый старт

1. Скопируйте config/settings.example.yaml → config/settings.yaml и заполните свои SMTP-данные.
2. Подготовьте recipients.csv (минимум столбец email, остальные — для персонализации).
3. Запустите dry-run для проверки:
    ```bash
    python -m email_app --dry-run
    ```
4. Для реальной отправки:
    ```bash
    python -m email_app
    ```
5. Для запуска GUI:
    ```bash
    python -m email_app --modern-gui
    ```

## Пример конфига (YAML)

```yaml
smtp:
   host: smtp.example.com
   port: 587
   username: user@example.com
   password: CHANGE_ME
   from_email: user@example.com
   from_name: Example Sender
   use_tls: true
   use_ssl: false
   timeout_seconds: 30
   # proxy_host: 127.0.0.1
   # proxy_port: 1080
   # proxy_type: socks5  # или http, https
   # proxy_user: null
   # proxy_pass: null
   # accounts_file: config/smtp_accounts.example.txt

message:
   subject: "Пример HTML-письма"
   template: newsletter.html
   reply_to: support@example.com
   attachments:
      - files/example.txt
   inline_images:
      hero: files/example-inline.svg

delivery:
   delay_seconds: 0.5
   log_file: logs/email_app.log
   history_csv: history/email_history.csv
   history_jsonl: history/email_history.jsonl
   skip_previously_sent: false
   dedupe_template_scope: true
   dedupe_history_days: 30
   scheduled_time: null
   rate_limit_per_minute: 20
   retry_attempts: 2
   retry_backoff_seconds: 10

content:
   headline: "Добро пожаловать"
   preheader: "Короткое описание письма"
   body_intro: "Это пример безопасного SMTP-приложения на Python для легитимной рассылки opt-in писем."
   body_text: "Вы можете менять тему, шаблон, inline-изображения и персональные поля через CSV и YAML-конфиг."
   cta_label: "Открыть сайт"
   cta_url: "https://example.com"
```

## Пример CSV получателей

```csv
email,name,company
user@example.com,Иван,Example LLC
```

## Пример пула SMTP-аккаунтов (TXT)

```txt
# host|port|username|password|from_email|from_name|use_tls|use_ssl|timeout_seconds
smtp.example.com|587|mailer1@example.com|CHANGE_ME|mailer1@example.com|Mailer One|true|false|30
smtp.example.com|587|mailer2@example.com|CHANGE_ME|mailer2@example.com|Mailer Two|true|false|30
```

Также поддерживается CSV:

```csv
host,port,username,password,from_email,from_name,use_tls,use_ssl,timeout_seconds
smtp.example.com,587,mailer1@example.com,CHANGE_ME,mailer1@example.com,Mailer One,true,false,30
smtp.example.com,587,mailer2@example.com,CHANGE_ME,mailer2@example.com,Mailer Two,true,false,30
```

## Пример шаблона (Jinja2)

```html
<p>Здравствуйте, {{ recipient.name }}!</p>
<p>{{ body_intro }}</p>
<a href="{{ cta_url }}">{{ cta_label }}</a>
```

## Пример вложения inline-изображения

```yaml
message:
   inline_images:
      hero: files/my-logo.png
```
```html
<img src="cid:{{ inline_images.hero.cid }}" alt="Логотип">
```

## Пример запуска с прокси

```bash
python -m email_app --config config/settings.yaml
# или с указанием прокси в YAML (см. выше)
```

## Советы и лучшие практики

- Используйте только для легитимных рассылок (opt-in).
- Перед запуском всегда делайте dry-run.
- Проверяйте лимиты SMTP-провайдера и не превышайте их.
- Используйте антидубли и историю отправок для чистоты базы.
- Для массовых рассылок используйте пул SMTP и прокси.
- Храните шаблоны и пресеты для повторного использования.
- В случае ошибок смотрите логи и историю.

---

- [email_app/main.py](email_app/main.py) — CLI-вход
- [email_app/config.py](email_app/config.py) — загрузка YAML-конфига
- [email_app/recipients.py](email_app/recipients.py) — загрузка CSV
- [email_app/renderer.py](email_app/renderer.py) — рендер HTML
- [email_app/smtp_client.py](email_app/smtp_client.py) — SMTP-отправка
- [email_app/service.py](email_app/service.py) — orchestration, задержки и логирование
- [email_app/gui.py](email_app/gui.py) — desktop-интерфейс
- [email_app/modern_gui.py](email_app/modern_gui.py) — modern-интерфейс на `customtkinter`
- [templates/newsletter.html](templates/newsletter.html) — пример шаблона
- [config/settings.example.yaml](config/settings.example.yaml) — пример конфига
- [config/smtp_accounts.example.txt](config/smtp_accounts.example.txt) — пример пула SMTP-аккаунтов (TXT)
- [config/smtp_accounts.example.csv](config/smtp_accounts.example.csv) — пример пула SMTP-аккаунтов (CSV)
- [recipients.csv](recipients.csv) — пример базы получателей
- [files/example.txt](files/example.txt) — пример вложения
- [files/example-inline.svg](files/example-inline.svg) — пример inline-изображения
- [presets/example.yaml](presets/example.yaml) — пример пресета кампании
- [presets/queue.example.json](presets/queue.example.json) — пример очереди кампаний
- [presets/queue.example.csv](presets/queue.example.csv) — пример очереди кампаний в CSV


## Установка и документация

> **Требуется Python 3.10 или новее** (используются современные типы и dataclass).

1. Создайте виртуальное окружение:
   ```bash
   python3.10 -m venv .venv && source .venv/bin/activate
   ```
2. Установите зависимости:
   ```bash
   pip install -r requirements.txt
   ```
3. Скопируйте пример конфига:
   ```bash
   cp config/settings.example.yaml config/settings.yaml
   ```
4. Заполните свои SMTP-данные и проверьте recipients.csv.
5. Подробнее — см. [GUIDE.md](GUIDE.md).

---

## Запуск

### macOS — двойной клик

Откройте папку проекта в Finder и дважды кликните **`EmailApp.app`**.
При первом запуске macOS спросит разрешение — нажмите **Открыть**.

### Windows — двойной клик

Дважды кликните **`launch.bat`** в проводнике.

### Из терминала

Проверка без отправки:

`python -m email_app --dry-run`

Реальная отправка:

`python -m email_app`

Запуск GUI:

`python -m email_app --gui`

Запуск modern GUI:

`python -m email_app --modern-gui`

HTML-предпросмотр в браузере:

`python -m email_app --preview`

Только preflight-проверка без отправки:

`python -m email_app --preflight`

Боевая отправка без интерактивного подтверждения:

`python -m email_app --yes`

Экспорт текущей кампании в очередь:

`python -m email_app --export-queue presets/my-queue.json`

Сохранить пресет:

`python -m email_app --save-preset presets/my-campaign.yaml`

Запустить по пресету:

`python -m email_app --preset presets/my-campaign.yaml`

Запустить очередь из JSON:

`python -m email_app --queue-file presets/queue.example.json`

Запустить очередь из CSV:

`python -m email_app --queue-file presets/queue.example.csv`

Показать статистику:

`python -m email_app --show-stats`

Показать отфильтрованную статистику:

`python -m email_app --show-stats --status-filter sent --template-filter newsletter`

Экспортировать отфильтрованную статистику:

`python -m email_app --show-stats --export-stats history/filtered.json`

Запуск с другим шаблоном из CLI:

`python -m email_app --template newsletter.html --delay-seconds 1.5`

## Формат CSV

Обязателен столбец `email`.

Пример:

```csv
email,name,company
user@example.com,Иван,Example LLC
```

Любые дополнительные поля можно использовать в HTML-шаблоне.

## Вложения и логирование

- Вложения задаются в `message.attachments`
- Inline-картинки задаются в `message.inline_images`
- Пауза между письмами задается в `delivery.delay_seconds`
- Лог пишется в путь `delivery.log_file`
- История отправок пишется в `delivery.history_csv` и `delivery.history_jsonl`
- В GUI можно выбрать шаблон и запустить `dry-run` без реальной отправки
- В GUI доступен встроенный HTML preview, плюс кнопка открытия в браузере
- В GUI и modern GUI есть встроенный редактор HTML-шаблонов
- В GUI и modern GUI есть визуальный редактор письма: шрифты, цвета, ссылки, таблицы, картинки и режим HTML-кода
- Перед запуском выполняется preflight: проверка шаблона, файлов и предупреждения по ссылкам/переменным
- Для боевой отправки есть подтверждение запуска
- Можно включить антидубли через `delivery.skip_previously_sent`

Для шаблонов используйте ссылки вида `cid:hero`, а в конфиге задавайте:

```yaml
message:
   inline_images:
      hero: files/example-inline.svg
```

В шаблоне это можно использовать так:

```html
<img src="cid:{{ inline_images.hero.cid }}" alt="Hero">
```

## Несколько SMTP-аккаунтов

По умолчанию используется один аккаунт из секции `smtp`.

Если нужен пул аккаунтов, укажите в конфиге `smtp.accounts_file` и используйте пример [config/smtp_accounts.example.csv](config/smtp_accounts.example.csv).

Письма будут распределяться по аккаунтам по кругу.

## Пресеты кампаний

Пресет хранит:
- путь к конфигу
- путь к CSV получателей
- путь к шаблонам
- выбранный шаблон
- задержку
- флаг `dry_run`

В GUI есть кнопки сохранения и загрузки пресетов.

## История отправок

После запуска приложение пишет журнал результатов в два файла:

- CSV для Excel/таблиц
- JSONL для последующей обработки скриптами

Для каждой записи сохраняются:
- время
- получатель
- статус
- шаблон
- SMTP-аккаунт
- dry-run или нет
- текст ошибки, если была

## Очередь кампаний JSON/CSV

Очередь может храниться в JSON или CSV с полями:
- `config`
- `recipients`
- `templates`
- `template`
- `delay_seconds`
- `dry_run`

Примеры есть в [presets/queue.example.json](presets/queue.example.json) и [presets/queue.example.csv](presets/queue.example.csv).

Эту очередь можно импортировать, запускать и экспортировать из CLI и из GUI.

## Фильтрация и экспорт статистики

Статистика поддерживает фильтры по:
- статусу
- части имени шаблона
- части SMTP-аккаунта

Результат можно экспортировать в CSV или JSON из CLI и из GUI.

## Пример переменных в шаблоне

- `{{ subject }}`
- `{{ recipient.email }}`
- `{{ recipient.name }}`
- `{{ headline }}`
- `{{ body_intro }}`
- `{{ body_text }}`
- `{{ cta_label }}`
- `{{ cta_url }}`

## Что можно добавить дальше

Если захотите, следующим сообщением могу добавить:
- тёмный HTML preview-шаблон
- быстрые пресеты фильтров статистики
- CSV-импорт нескольких очередей за раз

## Использование рандомных прокси

Для рандомизации прокси создайте файл `config/proxies.txt`:

```
185.252.215.173:24091:socks5:111111:111111
185.252.215.173:40238:socks5:111111:111111
185.252.215.173:39878:socks5:111111:111111
185.252.215.173:30808:socks5:111111:111111
185.252.215.173:33674:socks5:111111:111111
185.252.215.173:39117:socks5:111111:111111
```

- Формат: host:port:type[:user:pass]
- Для каждого письма прокси выбирается случайно из этого списка.
- Если файл отсутствует — используется стандартная логика (прокси из конфига или из аккаунта).

Прокси подставляются независимо от SMTP-аккаунта!

На Windows клонируешь:
git clone <url>
cd Email (или имя папки)
Если ещё нет venv:
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
