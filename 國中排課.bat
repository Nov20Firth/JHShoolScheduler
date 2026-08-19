@echo off
setlocal
title Junior High Timetable Scheduler v0.4.0
chcp 65001 >nul
set "PYTHONUTF8=1"
cd /d "%~dp0"

python --version >nul 2>nul
if errorlevel 1 (
    echo Python was not found. Please install Python and add it to PATH.
    echo.
    pause
    exit /b 1
)

if "%~1"=="" (
    python -X utf8 "%~dp0run_schedule.py"
) else if "%~2"=="" (
    python -X utf8 "%~dp0run_schedule.py" "%~f1"
) else (
    python -X utf8 "%~dp0run_schedule.py" "%~f1" "%~f2"
)

set "SCHEDULE_EXIT=%ERRORLEVEL%"
echo.
if not "%SCHEDULE_EXIT%"=="0" echo Scheduling did not finish. Please review the error above.
pause
exit /b %SCHEDULE_EXIT%
