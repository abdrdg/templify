@echo off
REM Setup and run invitation_generator.py with proper environment checks
setlocal
set SCRIPT_DIR=%~dp0
set VENV_DIR=%SCRIPT_DIR%venv
set PYTHON_EXE=%VENV_DIR%\Scripts\python.exe
set PIP_EXE=%VENV_DIR%\Scripts\pip.exe

echo =============================================
echo Templify - Invitation Generator Setup
echo =============================================

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://python.org
    pause
    exit /b 1
)

REM Check if virtual environment exists
if not exist "%VENV_DIR%\Scripts\activate.bat" (
    echo Creating virtual environment...
    python -m venv "%VENV_DIR%"
    if errorlevel 1 (
        echo ERROR: Failed to create virtual environment
        pause
        exit /b 1
    )
    echo Virtual environment created successfully!
) else (
    echo Virtual environment found.
)

REM Activate virtual environment
echo Activating virtual environment...
call "%VENV_DIR%\Scripts\activate.bat"
if errorlevel 1 (
    echo ERROR: Failed to activate virtual environment
    pause
    exit /b 1
)

REM Check if requirements.txt exists and install/update packages
if exist "%SCRIPT_DIR%requirements.txt" (
    echo Installing/updating required packages...
    "%PIP_EXE%" install -r "%SCRIPT_DIR%requirements.txt" --upgrade
    if errorlevel 1 (
        echo WARNING: Some packages may not have installed correctly
        echo Continuing anyway...
    ) else (
        echo Packages installed successfully!
    )
) else (
    echo WARNING: requirements.txt not found, skipping package installation
)

echo.
echo Starting Invitation Generator...
echo =============================================
"%PYTHON_EXE%" "%SCRIPT_DIR%invitation_generator.py"

if errorlevel 1 (
    echo.
    echo Application encountered an error.
    pause
)

echo Application closed.
