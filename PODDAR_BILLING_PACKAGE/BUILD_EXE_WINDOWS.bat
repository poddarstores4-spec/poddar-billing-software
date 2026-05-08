@echo off
echo ================================================
echo   PODDAR STORES - Building .exe for Windows
echo ================================================
echo.

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found!
    echo Please install Python 3.10+ from https://python.org
    pause
    exit
)

echo Installing required libraries...
pip install pyinstaller reportlab pillow num2words

echo.
echo Building PODDAR_STORES_BILLING.exe ...
pyinstaller --onefile --windowed --name "PODDAR_STORES_BILLING" ^
    --add-data "assets/sign.png;assets" ^
    --add-data "data;data" ^
    --hidden-import num2words ^
    --hidden-import reportlab ^
    --hidden-import PIL ^
    main.py

echo.
echo ================================================
echo   BUILD COMPLETE!
echo   File: dist\PODDAR_STORES_BILLING.exe
echo ================================================
pause
