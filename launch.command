#!/bin/zsh
# Лаунчер Email App — двойной клик в Finder запускает GUI

# Переходим в папку скрипта (где бы он ни лежал)
cd "$(dirname "$0")"

# Активируем виртуальное окружение
source .venv/bin/activate

# Запускаем modern GUI
python -m email_app --modern-gui
