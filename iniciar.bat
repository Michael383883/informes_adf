@echo off
chcp 65001 >nul
cd /d "%~dp0"
where python >nul 2>nul || (echo No encuentro Python. Instala Python 3.10+ desde python.org y marca "Add to PATH". & pause & exit /b 1)
if not exist .venv (
  echo Creando entorno e instalando librerias, solo la primera vez...
  python -m venv .venv || (pause & exit /b 1)
  call .venv\Scripts\activate.bat
  python -m pip install --upgrade pip
  pip install -r requirements.txt || (pause & exit /b 1)
) else (
  call .venv\Scripts\activate.bat
)
python app.py
pause
