@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\pythonw.exe" (
    echo The local Python environment has not been created.
    echo Read docs\WINDOWS_SETUP.md and approve Stage 1 setup first.
    echo This launcher does not install or download anything.
    pause
    exit /b 1
)
start "Belsito Plumbing Email Assistant" ".venv\Scripts\pythonw.exe" -m belsito_mail
