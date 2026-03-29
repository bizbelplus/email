# Email app на Python

Это минимальное и безопасное SMTP-приложение для **легитимной** отправки HTML-писем по вашему списку получателей.

Поддерживается:
- SMTP
- несколько SMTP-аккаунтов из CSV
- HTML-шаблоны через Jinja2
- персонализация по CSV
- `dry-run` режим для проверки
- TLS/SSL
- desktop GUI на `tkinter`
- modern GUI на `customtkinter`
- вложения
- inline-картинки через `cid:`
- задержка между письмами
- логирование в файл
- история отправок в CSV и JSONL
- HTML-предпросмотр перед отправкой
- пресеты кампаний в YAML
- встроенный редактор HTML-шаблонов
- визуальный редактор письма с форматированием как в Word
- preflight-чеклист перед запуском кампании
- валидация шаблона и ссылок перед отправкой
- антидубли (пропуск ранее отправленных адресов)
- импорт очереди кампаний из JSON/CSV
- страница статистики по истории отправок
- фильтрация и экспорт статистики
- встроенный HTML preview внутри GUI

Не добавлял функции, которые могут использоваться для обхода ограничений почтовых сервисов или нежелательных рассылок, вроде прокси-ротации и спинтакса/рандомизации текста.

> 📖 **Подробное пошаговое руководство:** [GUIDE.md](GUIDE.md)

## Структура

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
- [config/smtp_accounts.example.csv](config/smtp_accounts.example.csv) — пример пула SMTP-аккаунтов
- [recipients.csv](recipients.csv) — пример базы получателей
- [files/example.txt](files/example.txt) — пример вложения
- [files/example-inline.svg](files/example-inline.svg) — пример inline-изображения
- [presets/example.yaml](presets/example.yaml) — пример пресета кампании
- [presets/queue.example.json](presets/queue.example.json) — пример очереди кампаний
- [presets/queue.example.csv](presets/queue.example.csv) — пример очереди кампаний в CSV

## Установка

> **Требуется Python 3.10 или новее** (используются `@dataclass(slots=True)` и `X | Y` типы).

1. Создайте виртуальное окружение на Python 3.10+:

   `python3.10 -m venv .venv && source .venv/bin/activate`

2. Установите зависимости:

   `pip install -r requirements.txt`

3. Скопируйте [config/settings.example.yaml](config/settings.example.yaml) в `config/settings.yaml`.
4. Укажите ваши SMTP-данные.
5. Подготовьте [recipients.csv](recipients.csv).
6. При необходимости обновите список вложений в конфиге.

См. подробный разбор каждого шага в [GUIDE.md](GUIDE.md).

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
