#!/usr/bin/env bash
set -e
cd "$(dirname "$0")"

if ! command -v python3 >/dev/null; then
    echo "Python 3 не найден. Установите python3."
    exit 1
fi

if [ ! -d ".venv" ]; then
    echo "Создаю виртуальное окружение..."
    python3 -m venv .venv
    . .venv/bin/activate
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
else
    . .venv/bin/activate
fi

python avpo.py