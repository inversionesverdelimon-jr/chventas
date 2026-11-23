@echo off
REM CH Ventas - inicia el tablero en http://localhost:8050
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creando entorno de Python por primera vez...
  py -3 -m venv .venv || python -m venv .venv
  ".venv\Scripts\python.exe" -m pip install --upgrade pip
  ".venv\Scripts\python.exe" -m pip install -r requirements.txt
)
".venv\Scripts\python.exe" app.py %*
pause
