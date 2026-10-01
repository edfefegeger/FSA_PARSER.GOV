#!/usr/bin/env bash
# Запуск парсера ФСА в Linux. При первом запуске сам создаёт venv и ставит зависимости.
set -e
cd "$(dirname "$(readlink -f "$0")")"

if [ ! -x .venv/bin/python ]; then
    echo "Первый запуск: создаю виртуальное окружение..."
    if ! python3 -m venv .venv 2>/dev/null; then
        rm -rf .venv
        echo "Не удалось создать venv. Установите пакет:  sudo apt install python3-venv"
        exit 1
    fi
    .venv/bin/pip install --quiet --default-timeout 100 --retries 10 --upgrade pip
    .venv/bin/pip install --quiet --default-timeout 100 --retries 10 -r requirements.txt
fi

.venv/bin/python fsa_parser.py "$@"
