@echo off
REM Doble clic para abrir el Generador de Contratos en Windows.
setlocal
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo Preparando el entorno de Python. Esto tarda un par de minutos la primera vez...
  py -3 -m venv .venv 2>nul
  if not exist ".venv\Scripts\python.exe" python -m venv .venv 2>nul
  if not exist ".venv\Scripts\python.exe" (
    echo.
    echo No se encontro Python. Instalalo desde https://www.python.org/downloads/
    echo y marca la casilla "Add python.exe to PATH". Despues vuelve a dar doble clic aqui.
    pause
    exit /b 1
  )
  ".venv\Scripts\python.exe" -m pip install --quiet --no-cache-dir -r requirements.txt
)

set ORIENS_ABRIR=1
".venv\Scripts\python.exe" app.py
pause
