@echo off
setlocal

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  echo [INFO] Creating virtual environment...
  python -m venv .venv
)

call ".venv\Scripts\activate.bat"
python -m pip install --upgrade pip
pip install -r requirements.txt

echo [INFO] Building exe with PyInstaller...
pyinstaller --noconfirm --clean --windowed --name 对账助手 --onedir gui_app.py

echo [DONE] Build finished.
echo Output: dist\对账助手\对账助手.exe
pause
