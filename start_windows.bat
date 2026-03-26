@echo off
cd /d "%~dp0"

where python3 >nul 2>&1
if %errorlevel%==0 (
    set PYTHON_BIN=python3
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set PYTHON_BIN=python
    ) else (
        echo Python wurde nicht gefunden.
        pause
        exit /b 1
    )
)

%PYTHON_BIN% -c "import openpyxl" 2>nul
if %errorlevel% neq 0 (
    echo Installiere openpyxl...
    %PYTHON_BIN% -m pip install openpyxl
)

%PYTHON_BIN% Skript_V6-5.py
pause
