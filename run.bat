@echo off
rem Audit Suite launcher: creates a local Python environment on first run, then opens the app in the browser.
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
    echo First run: setting up the Python environment. This takes a few minutes...
    python -m venv .venv || (echo Python 3.10+ is required: https://www.python.org/downloads/ & pause & exit /b 1)
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt || (echo Package install failed. & pause & exit /b 1)
)
".venv\Scripts\python.exe" -m streamlit run app.py
pause
