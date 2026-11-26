#!/usr/bin/env bash
cd "$(dirname "$0")"
if [ ! -d .venv ]; then
  echo "Creando entorno e instalando librerías (solo la primera vez)..."
  python3 -m venv .venv && . .venv/bin/activate && pip install --upgrade pip && pip install -r requirements.txt || exit 1
else
  . .venv/bin/activate
fi
python app.py
