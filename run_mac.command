#!/usr/bin/env bash
# Автоматический запуск (macOS): создаёт venv, ставит зависимости и запускает приложение.

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

PY_BIN="python3"
if ! command -v "$PY_BIN" >/dev/null 2>&1; then
  echo "Не найден python3. Установите Python 3.10+."
  exit 1
fi

VENV_DIR=".venv"
if [ ! -d "$VENV_DIR" ]; then
  "$PY_BIN" -m venv "$VENV_DIR"
fi

source "$VENV_DIR/bin/activate"
python -m pip install --upgrade pip
pip install -r requirements.txt

python TechDirRentMan/main.py
