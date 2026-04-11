@echo off
REM ============================================================
REM  Mewsie's Multi-XML Toolbox — One-Click Launcher
REM ============================================================

cd /d "%~dp0"

echo +------------------------------------------------+
echo [     Mewsie's Multi-XML Toolbox Launcher        ]
echo +------------------------------------------------+
echo.
echo Working directory: %CD%
echo.

:: -------------------------------
:: Check Python
:: -------------------------------
echo [1/4] Checking Python...

where python >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found in PATH.
    pause
    exit /b
)

python --version
echo Python OK
echo.

:: -------------------------------
:: Check pip
:: -------------------------------
echo [2/4] Checking pip...

python -m pip --version
if errorlevel 1 (
    echo ERROR: pip is not working.
    pause
    exit /b
)

echo pip OK
echo.

:: -------------------------------
:: Check openpyxl
:: -------------------------------
echo [3/4] Checking openpyxl...

python -c "import openpyxl" 2>error.log
if errorlevel 1 (
    echo openpyxl not found. Installing...
    python -m pip install openpyxl
    if errorlevel 1 (
        echo ERROR: Failed to install openpyxl
        type error.log
        pause
        exit /b
    )
) else (
    echo openpyxl OK
)

del error.log >nul 2>&1
echo.

:: -------------------------------
:: Check tkinter
:: -------------------------------
echo [4/4] Checking tkinter...

python -c "import tkinter" 2>error.log
if errorlevel 1 (
    echo ERROR: tkinter is missing!
    echo.
    echo Fix: Reinstall Python and enable "tcl/tk and IDLE"
    echo.
    type error.log
    pause
    exit /b
) else (
    echo tkinter OK
)

del error.log >nul 2>&1
echo.

:: -------------------------------
:: Launch app
:: -------------------------------
echo Launching MewBox...
echo.

python mewbox.py

echo.
echo ===============================
echo   MewBox exited
echo ===============================
pause