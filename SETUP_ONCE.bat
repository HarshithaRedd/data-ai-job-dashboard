@echo off
cd /d "%~dp0"
echo Creating the Python environment...
python -m venv venv
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt
python -m playwright install chromium
echo.
echo Setup complete.
pause
