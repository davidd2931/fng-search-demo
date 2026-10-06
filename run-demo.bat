@echo off
setlocal
cd /d "%~dp0"

set "PYTHON="
where py >nul 2>&1 && set "PYTHON=py -3"
if not defined PYTHON where python >nul 2>&1 && set "PYTHON=python"

if not defined PYTHON (
    echo Python 3 was not found.
    echo Install Python 3.10 or later from https://www.python.org/downloads/windows/
    echo During setup, select "Add python.exe to PATH", then run this file again.
    pause
    exit /b 1
)

if not exist ".venv\Scripts\python.exe" (
    echo First run: creating a private Python environment for this demo...
    %PYTHON% -m venv .venv
    if errorlevel 1 goto setup_error
)

".venv\Scripts\python.exe" -c "import customtkinter, numpy, openpyxl, pandas, pyarrow" >nul 2>&1
if errorlevel 1 (
    echo First run: installing the demo's Python packages...
    ".venv\Scripts\python.exe" -m pip install --upgrade pip
    if errorlevel 1 goto setup_error
    ".venv\Scripts\python.exe" -m pip install -r requirements.txt
    if errorlevel 1 goto setup_error
)

echo Starting BOM Search Demo...
".venv\Scripts\python.exe" src\app.py
if errorlevel 1 pause
exit /b

:setup_error
echo Setup did not finish. Check your internet connection and try again.
echo If this is a work computer, your IT team may need to allow Python package downloads.
pause
exit /b 1
