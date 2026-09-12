#!/bin/bash
# Doble clic para abrir el Generador de Contratos en el navegador.
cd "$(dirname "$0")" || exit 1

if [ ! -d .venv ]; then
  python3 -m venv .venv && .venv/bin/pip install -r requirements.txt || exit 1
fi

(sleep 2 && open "http://127.0.0.1:5050") &
.venv/bin/python app.py
