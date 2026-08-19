@echo off
chcp 65001 >nul
set "PYTHONUTF8=1"
cd /d "%~dp0"

if "%~1"=="" (
    python "%~dp0run_adjustment.py"
) else (
    python "%~dp0run_adjustment.py" "%~f1"
)

echo.
pause
