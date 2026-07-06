@echo off
chcp 65001 >nul
title Xoa Database SQL Server

echo ============================================================
echo    TOOL XOA DATABASE SQL SERVER
echo ============================================================
echo.
echo [CANH BAO] Tool nay se XOA TAT CA database trong danh sach Excel!
echo Sau khi xoa, chay CauHinhSQL.bat de restore lai.
echo.
echo Nhan phim bat ky de tiep tuc, hoac dong cua so de huy...
echo.
pause

REM Lấy đường dẫn thư mục chứa file .bat
set "SCRIPT_DIR=%~dp0"
set "EXCEL_FILE=%SCRIPT_DIR%..\ExcelCauHinh\Settup AXE.xlsx"
set "SHEET_NAME=SQL"

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
    echo [LOI] Python chua duoc cai dat hoac khong co trong PATH!
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
echo    BAT DAU XOA DATABASE
echo ============================================================
echo File Excel: %EXCEL_FILE%
echo Sheet: %SHEET_NAME%
echo ============================================================
echo.

REM Chạy tool với --clean (chỉ xóa, không restore)
cd /d "%SCRIPT_DIR%"
python sql_config_tool.py "%EXCEL_FILE%" --sheet "%SHEET_NAME%" --clean

echo.
echo ============================================================
echo    De restore lai, chay: CauHinhSQL.bat
echo ============================================================
echo.

pause
