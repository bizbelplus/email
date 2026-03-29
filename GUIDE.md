# Руководство пользователя — Email App

Пошаговая инструкция по работе с приложением: от установки до запуска кампаний, GUI, очередей и статистики.

---

## Содержание

1. [Системные требования](#1-системные-требования)
2. [Установка](#2-установка)
3. [Настройка конфига SMTP](#3-настройка-конфига-smtp)
4. [Подготовка списка получателей](#4-подготовка-списка-получателей)
5. [Первый запуск — dry-run](#5-первый-запуск--dry-run)
6. [Реальная отправка из командной строки](#6-реальная-отправка-из-командной-строки)
7. [Графический интерфейс (GUI)](#7-графический-интерфейс-gui)
8. [HTML-предпросмотр письма](#8-html-предпросмотр-письма)
9. [Редактор шаблонов в GUI](#9-редактор-шаблонов-в-gui)
10. [Inline-изображения (cid:)](#10-inline-изображения-cid)
11. [Пресеты кампаний](#11-пресеты-кампаний)
12. [Несколько SMTP-аккаунтов](#12-несколько-smtp-аккаунтов)
13. [Очередь кампаний JSON / CSV](#13-очередь-кампаний-json--csv)
14. [История отправок](#14-история-отправок)
15. [Статистика и фильтрация](#15-статистика-и-фильтрация)
16. [Вложения к письмам](#16-вложения-к-письмам)
17. [Кастомизация HTML-шаблона](#17-кастомизация-html-шаблона)
18. [Полный список CLI-аргументов](#18-полный-список-cli-аргументов)
19. [Типичные ошибки и решения](#19-типичные-ошибки-и-решения)

---

## 1. Системные требования

| Требование | Версия |
|---|---|
| Python | **3.10 или новее** |
| ОС | macOS / Linux / Windows |
| Pip | любая актуальная |

> ⚠️ Python 3.9 и старше **не поддерживается** — используются `@dataclass(slots=True)` и `X | Y` типы из Python 3.10.

---

## 2. Установка

```bash
# 1. Перейти в папку проекта
cd /путь/до/Email

# 2. Создать виртуальное окружение на Python 3.10+
python3.10 -m venv .venv          # или python3.11, python3.12, python3.13

# Если Python 3.10+ доступен только через conda:
/Users/yourname/miniconda3/bin/python -m venv .venv

# 3. Активировать окружение
source .venv/bin/activate          # macOS / Linux
# .venv\Scripts\activate           # Windows

# 4. Установить зависимости
pip install -r requirements.txt

# 5. Создать конфиг из примера
cp config/settings.example.yaml config/settings.yaml
```

После установки вывод `python -m email_app --help` должен показать список аргументов.

---

## 3. Настройка конфига SMTP

Откройте `config/settings.yaml` и заполните блок `smtp`:

```yaml
smtp:
  host: smtp.gmail.com        # SMTP-сервер вашего провайдера
  port: 587                   # 587 для TLS, 465 для SSL
  username: you@gmail.com
  password: ВАШ_ПАРОЛЬ        # для Gmail — App Password, не основной пароль
  from_email: you@gmail.com
  from_name: Имя Отправителя
  use_tls: true
  use_ssl: false
  timeout_seconds: 30
```

**Gmail:** включите двухфакторную аутентификацию, затем создайте [App Password](https://myaccount.google.com/apppasswords) и вставьте его как `password`.

**Yandex:** используйте `smtp.yandex.ru`, порт `465`, `use_ssl: true`.

**Mail.ru:** используйте `smtp.mail.ru`, порт `465`, `use_ssl: true`.

---

## 4. Подготовка списка получателей

Файл `recipients.csv` должен содержать минимум столбец `email`.
Дополнительные столбцы автоматически становятся переменными в шаблоне.

```csv
email,name,company
ivan@example.com,Иван,Example LLC
maria@example.com,Мария,Tech Corp
```

Любые поля CSV доступны в шаблоне как `{{ recipient.name }}`, `{{ recipient.company }}` и т.д.

---

## 5. Первый запуск — dry-run

`dry-run` рендерит письма и логирует всё как обычно, но **не отправляет** сообщения.
Используйте это для проверки конфига и шаблона:

```bash
# Активируйте venv (если ещё не активирован)
source .venv/bin/activate

# Запуск в режиме dry-run
python -m email_app --dry-run
```

В терминале появится:

```
[dry-run] → ivan@example.com  ✓
[dry-run] → maria@example.com ✓
Итог: обработано 2, успешно 2, ошибок 0
```

Лог пишется в `logs/email_app.log`, история — в `history/email_history.csv`.

---

## 6. Реальная отправка из командной строки

После успешного dry-run снимите флаг:

```bash
python -m email_app
```

Дополнительные опции:

```bash
# Использовать другой шаблон
python -m email_app --template promo.html

# Задержка 2 секунды между письмами
python -m email_app --delay-seconds 2

# Другой список получателей
python -m email_app --recipients lists/vip.csv

# Другой конфиг
python -m email_app --config config/production.yaml
```

---

## 7. Графический интерфейс (GUI)

Доступны два варианта GUI.

### Стандартный (tkinter)

```bash
python -m email_app --gui
```

### Modern (customtkinter — тёмная тема)

```bash
python -m email_app --modern-gui
```

**Кнопки в GUI:**

| Кнопка | Действие |
|---|---|
| **Запустить** | Реальная отправка кампании |
| **Dry-run** | Тестовый прогон без отправки |
| **Preview** | Открыть встроенный предпросмотр первого письма |
| **Редактор шаблонов** | Встроенный HTML-редактор |
| **Загрузить пресет** | Открыть сохранённый YAML-пресет |
| **Сохранить пресет** | Сохранить текущие настройки в YAML |
| **Очередь JSON** | Загрузить и запустить очередь кампаний из JSON/CSV |
| **Экспорт JSON/CSV** | Сохранить текущую кампанию в файл очереди |
| **Статистика** | Открыть окно статистики по истории отправок |

В правой части GUI находятся поля:
- **Конфиг** — путь к `settings.yaml`
- **Получатели** — путь к CSV
- **Шаблоны** — папка с HTML-шаблонами
- **Шаблон** — имя конкретного HTML-файла
- **Задержка** — паузы между письмами в секундах

---

## 8. HTML-предпросмотр письма

Рендерит письмо для первого получателя из CSV:

```bash
# Открыть в браузере
python -m email_app --preview
```

В GUI нажмите кнопку **Preview** — откроется встроенное окно с:
- рендером HTML (через tkhtmlview)
- вкладкой с исходным кодом письма
- кнопкой «Открыть в браузере»

---

## 9. Редактор шаблонов в GUI

1. В GUI нажмите **Редактор шаблонов**.
2. Выберите HTML-файл из папки `templates/`.
3. Отредактируйте код в текстовом поле.
4. Нажмите **Сохранить** — файл перезапишется.

Нажмите **Preview** сразу после сохранения, чтобы увидеть результат.

---

## 10. Inline-изображения (cid:)

Inline-изображения встраиваются прямо в письмо (не как вложение) и отображаются без внешних запросов.

**Шаг 1.** Добавьте файл в конфиг:

```yaml
message:
  inline_images:
    hero: files/my-logo.png   # ключ: путь к файлу
    banner: files/banner.jpg
```

**Шаг 2.** Используйте в шаблоне:

```html
<img src="cid:{{ inline_images.hero.cid }}" alt="Логотип" width="200">
```

В предпросмотре `cid:` автоматически заменяется на `data:image/...;base64,...` для отображения в браузере.

---

## 11. Пресеты кампаний

Пресет — YAML-файл с настройками запуска. Удобно переключаться между кампаниями без редактирования конфига.

**Сохранить пресет из CLI:**

```bash
python -m email_app \
  --config config/promo.yaml \
  --recipients lists/subscribers.csv \
  --template promo.html \
  --delay-seconds 1 \
  --save-preset presets/promo-campaign.yaml
```

**Загрузить пресет:**

```bash
python -m email_app --preset presets/promo-campaign.yaml
```

**Содержимое файла пресета:**

```yaml
config: config/promo.yaml
recipients: lists/subscribers.csv
templates: templates
template: promo.html
delay_seconds: 1.0
dry_run: false
```

В GUI — кнопки **Загрузить пресет** и **Сохранить пресет**.

---

## 12. Несколько SMTP-аккаунтов

Можно использовать пул из нескольких аккаунтов — письма будут распределяться по кругу (round-robin).

**Шаг 1.** Создайте `config/smtp_accounts.csv` (см. пример `config/smtp_accounts.example.csv`):

```csv
host,port,username,password,from_email,from_name,use_tls,use_ssl
smtp.gmail.com,587,acc1@gmail.com,pass1,acc1@gmail.com,Sender One,true,false
smtp.gmail.com,587,acc2@gmail.com,pass2,acc2@gmail.com,Sender Two,true,false
```

**Шаг 2.** В конфиге укажите путь к файлу вместо отдельных полей:

```yaml
smtp:
  accounts_file: config/smtp_accounts.csv
```

---

## 13. Очередь кампаний JSON / CSV

Очередь позволяет запустить несколько кампаний подряд одной командой.

### Формат JSON (`presets/queue.json`):

```json
[
  {
    "config": "config/settings.yaml",
    "recipients": "lists/group1.csv",
    "template": "newsletter.html",
    "delay_seconds": 1.0,
    "dry_run": false
  },
  {
    "config": "config/settings.yaml",
    "recipients": "lists/group2.csv",
    "template": "promo.html",
    "delay_seconds": 2.0,
    "dry_run": false
  }
]
```

### Формат CSV (`presets/queue.csv`):

```csv
config,recipients,templates,template,delay_seconds,dry_run
config/settings.yaml,lists/group1.csv,templates,newsletter.html,1.0,false
config/settings.yaml,lists/group2.csv,templates,promo.html,2.0,false
```

### Запуск очереди:

```bash
python -m email_app --queue-file presets/queue.json
# или
python -m email_app --queue-file presets/queue.csv
```

### Экспорт текущей кампании в очередь:

```bash
python -m email_app --export-queue presets/queue.json
```

В GUI — кнопки **Очередь JSON** (запуск) и **Экспорт JSON/CSV** (сохранение).

---

## 14. История отправок

После каждого запуска (включая dry-run) приложение пишет историю в два файла:

| Файл | Формат | Назначение |
|---|---|---|
| `history/email_history.csv` | CSV | Таблица для Excel / сводные отчёты |
| `history/email_history.jsonl` | JSONL | Скрипты, автоматизация |

**Столбцы:**

```
timestamp, recipient, status, template, smtp_account, dry_run, error
```

Пример строки в CSV:

```
2025-03-29T14:00:00,ivan@example.com,sent,newsletter.html,acc1@gmail.com,false,
```

Путь к файлам настраивается в конфиге:

```yaml
delivery:
  history_csv: history/email_history.csv
  history_jsonl: history/email_history.jsonl
```

---

## 15. Статистика и фильтрация

### CLI

```bash
# Общая статистика
python -m email_app --show-stats

# Фильтр по статусу
python -m email_app --show-stats --status-filter sent
python -m email_app --show-stats --status-filter error
python -m email_app --show-stats --status-filter dry-run

# Фильтр по шаблону (подстрока)
python -m email_app --show-stats --template-filter newsletter

# Фильтр по SMTP-аккаунту (подстрока)
python -m email_app --show-stats --smtp-filter gmail

# Несколько фильтров сразу
python -m email_app --show-stats --status-filter sent --template-filter promo

# Экспорт в CSV
python -m email_app --show-stats --export-stats history/report.csv

# Экспорт в JSON
python -m email_app --show-stats --export-stats history/report.json

# Статистика из другого файла
python -m email_app --show-stats --history-csv history/old_history.csv
```

### GUI

Нажмите кнопку **Статистика** — откроется окно с:
- Сводкой (всего, отправлено, dry-run, ошибки, уникальных получателей)
- Выпадающим фильтром по статусу
- Полями поиска по шаблону и SMTP-аккаунту
- Таблицей записей
- Кнопками **Экспорт CSV** и **Экспорт JSON**

---

## 16. Вложения к письмам

Добавьте пути к файлам в конфиг:

```yaml
message:
  attachments:
    - files/report.pdf
    - files/terms.docx
```

Все получатели получат одни и те же вложения.

---

## 17. Кастомизация HTML-шаблона

Шаблоны находятся в папке `templates/` и используют синтаксис [Jinja2](https://jinja.palletsprojects.com/).

**Доступные переменные:**

| Переменная | Источник |
|---|---|
| `{{ recipient.email }}` | CSV-столбец `email` |
| `{{ recipient.name }}` | CSV-столбец `name` |
| `{{ recipient.* }}` | любой столбец из CSV |
| `{{ subject }}` | `message.subject` из конфига |
| `{{ headline }}` | `content.headline` из конфига |
| `{{ body_intro }}` | `content.body_intro` из конфига |
| `{{ body_text }}` | `content.body_text` из конфига |
| `{{ cta_label }}` | `content.cta_label` из конфига |
| `{{ cta_url }}` | `content.cta_url` из конфига |
| `{{ inline_images.КЛЮЧ.cid }}` | ключ из `message.inline_images` |

**Пример персонализации:**

```html
<p>Привет, {{ recipient.name }}!</p>
<p>{{ body_intro }}</p>
<a href="{{ cta_url }}">{{ cta_label }}</a>
```

**Несколько шаблонов** можно хранить в папке `templates/` и переключаться между ними через `--template` или GUI.

---

## 18. Полный список CLI-аргументов

```
python -m email_app [аргументы]

Основные:
  --config PATH          Путь к YAML-конфигу          (по умолч.: config/settings.yaml)
  --recipients PATH      Путь к CSV получателей        (по умолч.: recipients.csv)
  --templates PATH       Папка с HTML-шаблонами        (по умолч.: templates)
  --template NAME        Имя HTML-файла шаблона
  --delay-seconds N      Задержка между письмами (сек)
  --dry-run              Тестовый прогон без отправки

GUI:
  --gui                  Запустить стандартный GUI (tkinter)
  --modern-gui           Запустить modern GUI (customtkinter)

Предпросмотр:
  --preview              Сгенерировать preview и открыть в браузере

Пресеты:
  --preset PATH          Загрузить параметры из YAML-пресета
  --save-preset PATH     Сохранить текущие параметры в YAML-пресет и выйти

Очередь:
  --queue-file PATH      Запустить очередь кампаний из JSON или CSV
  --export-queue PATH    Экспортировать текущую кампанию в файл очереди и выйти

Статистика:
  --show-stats           Показать статистику по истории и выйти
  --history-csv PATH     Путь к history CSV (для --show-stats)
  --status-filter STR    Фильтр по статусу: sent / dry-run / error
  --template-filter STR  Фильтр по подстроке имени шаблона
  --smtp-filter STR      Фильтр по подстроке SMTP-аккаунта
  --export-stats PATH    Экспортировать отфильтрованную статистику в CSV или JSON
```

---

## 19. Типичные ошибки и решения

### `TypeError: dataclass() got an unexpected keyword argument 'slots'`

**Причина:** используется Python < 3.10.

**Решение:** создайте venv на Python 3.10+:

```bash
python3.10 -m venv .venv
```

---

### `SMTPAuthenticationError`

**Причина:** неверный пароль или не разрешён доступ по SMTP.

**Gmail:** создайте App Password в Google Account → Безопасность → Пароли приложений.

---

### `ConnectionRefusedError` / `TimeoutError`

**Причина:** неверный host/port или SMTP-сервер заблокирован.

**Решение:** проверьте значения `smtp.host` и `smtp.port`, убедитесь, что порт открыт (587 для TLS, 465 для SSL).

---

### Письма уходят, но изображения не отображаются

**Причина:** email-клиент блокирует inline-изображения.

**Решение:** это нормальное поведение почтовых клиентов — они требуют разрешения пользователя на показ изображений. Убедитесь, что вы используете `cid:` (inline), а не внешние ссылки.

---

### `FileNotFoundError: config/settings.yaml`

**Причина:** конфиг не был создан.

**Решение:**

```bash
cp config/settings.example.yaml config/settings.yaml
```

---

### GUI не открывается (нет окна)

**macOS:** убедитесь, что запускаете Python с правом на GUI. Conda-окружения обычно работают. Если запускаете через SSH — потребуется X11 или VNC.

---

*Версия руководства соответствует состоянию проекта на момент последнего обновления.*
