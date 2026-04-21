@echo off
title SUI Converter - Build Tool
echo ============================================
echo  Michigan MiUI SUI Converter - EXE Builder
echo ============================================
echo.

REM Check Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Please install Python 3.8+ from python.org
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

echo [1/3] Installing required packages...
pip install pandas openpyxl pyinstaller --quiet
if errorlevel 1 (
    echo ERROR: Package installation failed. Check your internet connection.
    pause
    exit /b 1
)

echo [2/3] Building SUI_Converter.exe...
pyinstaller --onefile --windowed --name "SUI_Converter" --icon NONE sui_converter.py
if errorlevel 1 (
    echo ERROR: Build failed. See output above for details.
    pause
    exit /b 1
)

echo [3/3] Done!
echo.
echo Your executable is ready at:
echo   dist\SUI_Converter.exe
echo.
echo You can distribute that single file to anyone on Windows.
echo No Python or other software required on their machine.
echo.
pause
