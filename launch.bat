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

REM ── Check Python ────────────────────────────────────────────
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH!
    echo Please install Python from https://python.org
    echo Make sure to check "Add to PATH" during installation.
    echo After installing Python, restart your computer and try again.
    pause
    exit /b 1
)

REM ── Check pip ───────────────────────────────────────────────
echo Checking pip...
python -m pip --version >nul 2>&1
if errorlevel 1 (
    echo Pip not found. Trying to install pip...
    python -m ensurepip --default-pip >nul 2>&1
    python -m pip --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Pip is still not working!
        echo Please right-click this script and choose "Run as administrator".
        pause
        exit /b 1
    )
    echo Pip installed successfully.
) else (
    echo Pip is working.
)

REM ── Dependencies ────────────────────────────────────────────
echo Checking required dependencies...

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing openpyxl...
    python -m pip install openpyxl --no-warn-script-location >nul 2>&1
    python -c "import openpyxl" >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Failed to install openpyxl!
        echo Try right-clicking this script and selecting "Run as administrator".
        pause
        exit /b 1
    )
    echo openpyxl installed successfully.
) else (
    echo openpyxl OK.
)

REM ── Check main script ────────────────────────────────────────
if not exist "mewbox.py" (
    echo ERROR: mewbox.py not found in %CD%
    echo Make sure this batch file is in the same folder as mewbox.py.
    pause
    exit /b 1
)

REM ── Create default output folders ───────────────────────────
if not exist "libconfig" mkdir "libconfig"
if not exist "reports"   mkdir "reports"
if not exist "MyShop"    mkdir "MyShop"
if not exist "csv_exports"    mkdir "csv_exports"
REM ── Launch ───────────────────────────────────────────────────
echo.
echo All checks passed. Starting Mewsie's Multi-XML Toolbox...
echo.
python mewbox.py
if errorlevel 1 (
    echo.
    echo ERROR: The application crashed or failed to start.
    echo.
    echo Troubleshooting:
    echo   1. Make sure Python 3.9 or newer is installed.
    echo   2. Run this script as administrator and try again.
    echo   3. Run manually: python mewbox.py
    pause
)