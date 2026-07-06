@echo off
chcp 65001 >nul
title Build IIS Config Tool to EXE

echo ============================================================
echo    BUILD IIS CONFIG TOOL TO EXE
echo ============================================================
echo.

REM Tìm Python
set "PYTHON_CMD=python"

REM Thử tìm Python trong PATH
python --version >nul 2>&1
if errorlevel 1 (
    echo [CANH BAO] Khong tim thay Python trong PATH, dang tim...
    
    REM Thử các đường dẫn phổ biến
    if exist "C:\Python312\python.exe" set "PYTHON_CMD=C:\Python312\python.exe"
    if exist "C:\Python311\python.exe" set "PYTHON_CMD=C:\Python311\python.exe"
    if exist "C:\Python310\python.exe" set "PYTHON_CMD=C:\Python310\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python312\python.exe" set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python312\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python311\python.exe" set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python311\python.exe"
    if exist "%LOCALAPPDATA%\Programs\Python\Python310\python.exe" set "PYTHON_CMD=%LOCALAPPDATA%\Programs\Python\Python310\python.exe"
    
    REM Kiểm tra lại
    "%PYTHON_CMD%" --version >nul 2>&1
    if errorlevel 1 (
        echo [LOI] Khong tim thay Python!
        echo.
        echo Vui long:
        echo 1. Mo CMD MOI va chay lai
        echo 2. Hoac cai dat Python tu: https://www.python.org/downloads/
        echo.
        pause
        exit /b 1
    )
)

echo Tim thay Python: 
"%PYTHON_CMD%" --version

REM Kiểm tra PyInstaller
echo.
echo Kiem tra PyInstaller...
"%PYTHON_CMD%" -c "import PyInstaller" >nul 2>&1
if errorlevel 1 (
    echo PyInstaller chua duoc cai dat. Dang cai dat...
    "%PYTHON_CMD%" -m pip install pyinstaller
    if errorlevel 1 (
        echo [LOI] Khong the cai dat PyInstaller!
        pause
        exit /b 1
    )
)

REM Kiểm tra các thư viện cần thiết
echo Kiem tra cac thu vien...
"%PYTHON_CMD%" -m pip install -r requirements.txt
if errorlevel 1 (
    echo [LOI] Khong the cai dat cac thu vien!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo    BAT DAU BUILD
echo ============================================================
echo.

REM Build bằng PyInstaller
"%PYTHON_CMD%" -m PyInstaller --clean --noconfirm iis_config_tool.spec

if errorlevel 1 (
    echo.
    echo [LOI] Build that bai!
    pause
    exit /b 1
)

echo.
echo ============================================================
echo    BUILD THANH CONG!
echo ============================================================
echo.
echo File EXE: dist\iis_config_tool.exe
echo.
echo Ban co the copy thu muc dist\ sang may khach hang.
echo.

pause
