@echo off
chcp 65001 >nul
title Cau hinh SQL Server tu Excel (Dry Run)

echo ============================================================
echo    TOOL CAU HINH SQL SERVER TU FILE EXCEL (DRY RUN)
echo ============================================================
echo.
echo [CANH BAO] Che do DRY RUN - Chi hien thi lenh, KHONG thuc thi!
echo.

REM Lấy đường dẫn thư mục chứa file .bat
set "SCRIPT_DIR=%~dp0"
set "EXCEL_FILE=%SCRIPT_DIR%..\ExcelCauHinh\Settup AXE.xlsx"
set "SHEET_NAME=SQL"
set "EXE_FILE=%SCRIPT_DIR%dist\sql_config_tool.exe"

REM Kiểm tra file Excel có tồn tại không
if not exist "%EXCEL_FILE%" (
    echo [LOI] Khong tim thay file Excel!
    echo Duong dan: %EXCEL_FILE%
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo    XEM TRUOC CAC LENH SE CHAY (DRY RUN)
echo ============================================================
echo File Excel: %EXCEL_FILE%
echo Sheet: %SHEET_NAME%
echo ============================================================
echo.

REM Ưu tiên chạy file .exe nếu có (không cần Python)
if exist "%EXE_FILE%" (
    echo [INFO] Su dung phien ban EXE (khong can Python)
    cd /d "%SCRIPT_DIR%"
    "%EXE_FILE%" "%EXCEL_FILE%" --sheet "%SHEET_NAME%" --dry-run
    goto :check_result
)

REM Nếu không có .exe, kiểm tra Python
echo [INFO] Khong tim thay file EXE, su dung Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [LOI] Python chua duoc cai dat hoac khong co trong PATH!
    echo.
    echo HUONG DAN:
    echo 1. Cai dat Python tu https://www.python.org/downloads/
    echo 2. HOAC build tool thanh file EXE bang cach chay: build_exe.bat
    echo    (Chi can build 1 lan, sau do co the chay tren bat ky may nao)
    echo.
    pause
    exit /b 1
)

REM Kiểm tra các thư viện cần thiết
echo Dang kiem tra cac thu vien can thiet...
python -c "import pandas; import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo [CANH BAO] Chua cai dat cac thu vien can thiet!
    echo Dang cai dat...
    pip install -r "%SCRIPT_DIR%requirements.txt"
    if errorlevel 1 (
        echo [LOI] Khong the cai dat cac thu vien!
        echo Vui long chay: pip install -r requirements.txt
        echo.
        pause
        exit /b 1
    )
)

REM Chạy tool ở chế độ dry-run bằng Python
cd /d "%SCRIPT_DIR%"
python sql_config_tool.py "%EXCEL_FILE%" --sheet "%SHEET_NAME%" --dry-run

:check_result

echo.
echo ============================================================
echo    KET THUC DRY RUN
echo ============================================================
echo.
echo De thuc thi cau hinh, chay file: CauHinhSQL.bat
echo.

pause
