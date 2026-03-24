@echo off
echo ============================================
echo   ACERT - Windows EXE Builder
echo ============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ first.
    pause
    exit /b 1
)

echo [1/4] Installing dependencies...
pip install flask pypdf reportlab pandas openpyxl pyinstaller --quiet
if errorlevel 1 (
    echo [ERROR] Failed to install dependencies.
    pause
    exit /b 1
)

echo [2/4] Cleaning old build files...
if exist "dist\ACERT.exe" del /f /q "dist\ACERT.exe"
if exist "build" rmdir /s /q "build"

echo [3/4] Building .exe ...
pyinstaller ACERT.spec --noconfirm
if errorlevel 1 (
    echo [ERROR] Build failed. Please check the error above.
    pause
    exit /b 1
)

echo [4/4] Done!
echo.
echo ============================================
echo   Output: dist\ACERT.exe
echo   Double-click ACERT.exe to run the app.
echo ============================================
echo.
pause
