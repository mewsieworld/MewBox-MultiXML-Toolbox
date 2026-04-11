@echo off
REM ============================================================
REM  Mewsie's Multi-XML Toolbox — One-Click Launcher
REM ============================================================

echo +------------------------------------------------+
echo [     Mewsie's Multi-XML Toolbox Launcher        ]
echo +------------------------------------------------+
echo.
echo Working directory: %CD%
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