@echo off
chcp 65001 >nul
title Xoa Windows Service

echo ============================================================
echo    TOOL XOA WINDOWS SERVICE
echo ============================================================
echo.
echo [CANH BAO] Tool nay se XOA TAT CA service trong danh sach Excel!
echo Sau khi xoa, chay CauHinhService.bat de tao lai.
echo.
echo Nhan phim bat ky de tiep tuc, hoac dong cua so de huy...
echo.
pause

REM Lấy đường dẫn thư mục chứa file .bat
set "SCRIPT_DIR=%~dp0"
set "EXCEL_FILE=%SCRIPT_DIR%..\ExcelCauHinh\Settup AXE.xlsx"
set "SHEET_NAME=service"

REM Kiểm tra file Excel có tồn tại không
if not exist "%EXCEL_FILE%" (
    echo [LOI] Khong tim thay file Excel!
    echo Duong dan: %EXCEL_FILE%
    echo.
    pause
    exit /b 1
)

REM Kiểm tra Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [LOI] Python chua duoc cai dat!
    pause
    exit /b 1
)

REM Kiểm tra thư viện
echo Dang kiem tra cac thu vien can thiet...
python -c "import pandas; import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Dang cai dat thu vien...
    pip install -r "%SCRIPT_DIR%requirements.txt"
)

echo.
echo ============================================================
echo    BAT DAU XOA SERVICE
echo ============================================================
echo File Excel: %EXCEL_FILE%
echo Sheet: %SHEET_NAME%
echo ============================================================
echo.

REM Chạy tool với --clean (chỉ xóa, không tạo)
cd /d "%SCRIPT_DIR%"
python service_config_tool.py "%EXCEL_FILE%" --sheet "%SHEET_NAME%" --clean

echo.
echo ============================================================
echo    De tao lai service, chay: CauHinhService.bat
echo ============================================================
echo.

pause
