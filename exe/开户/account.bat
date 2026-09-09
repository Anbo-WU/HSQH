@echo off
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
chcp 65001 >nul
if not exist "%~dp0.venv\Scripts\python.exe" (
    echo Python environment not found: "%~dp0.venv\Scripts\python.exe"
    echo Please create the account .venv and install requirements.txt first.
    exit /b 1
)
"%~dp0.venv\Scripts\python.exe" "%~dp0run.py" %*
exit /b %errorlevel%
