@echo off
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo Python is not installed or not in PATH
    pause
    exit /b
)
python UI/main.py
pause 