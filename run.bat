@echo off
cd /d "%~dp0"

if exist ".venv\Scripts\python.exe" (
    set PYTHON=".venv\Scripts\python.exe"
) else (
    set PYTHON=python
)

%PYTHON% app.py
if errorlevel 1 (
    echo.
    echo ERROR: app failed to start. See message above.
    pause
)
