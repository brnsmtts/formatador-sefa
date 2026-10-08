@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 (
    python servir_local.py
) else (
    py -3 servir_local.py
)
if errorlevel 1 pause
