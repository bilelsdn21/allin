@echo off
cd /d "%~dp0"
where python >nul 2>nul
if errorlevel 1 (
    echo ============================================================
    echo  Python is NOT installed on this PC.
    echo  1. Open the Microsoft Store
    echo  2. Search "Python 3.11"  (or newer) and install it
    echo  3. Double-click this SETUP.bat again
    echo ============================================================
    pause
    exit /b 1
)
python setup_check.py
