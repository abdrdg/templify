@echo off
REM Activate venv and run invitation_sender.py
setlocal
set VENV_DIR=%~dp0venv
if exist "%VENV_DIR%\Scripts\activate.bat" (
    call "%VENV_DIR%\Scripts\activate.bat"
) else (
    echo Virtual environment not found in %VENV_DIR%\Scripts
    exit /b 1
)
python "%~dp0invitation_sender.py"
