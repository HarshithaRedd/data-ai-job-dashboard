@echo off
cd /d "%~dp0"
if not exist "venv\Scripts\python.exe" (
  echo Run SETUP_ONCE.bat first.
  pause
  exit /b 1
)
call venv\Scripts\activate.bat
python main.py
pause
