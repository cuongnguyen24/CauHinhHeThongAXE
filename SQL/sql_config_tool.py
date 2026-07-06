"""
Tool tự động cấu hình SQL Server từ file Excel
Đọc file Excel sheet 'Cài đặt SQL' và thực thi restore database + tạo user

Cấu trúc Excel (sheet 'Cài đặt SQL'):
- Rows đầu: Thông tin chung (đường dẫn backup, database, user/password)
- Rows giữa: SQL chuẩn bị (DECLARE biến, CREATE TABLE tạm) ở cột cuối
- Rows cuối: Mỗi row = 1 database cần restore, cột cuối chứa script SQL đầy đủ
"""

import os
import sys
import subprocess
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import io
import re
import tempfile

# Fix encoding for Windows console
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except:
            pass


class SQLConfigTool:
    def __init__(self, excel_path: str, sqlcmd_path: str = None):
        """
        Khởi tạo tool cấu hình SQL Server
        """
        self.excel_path = excel_path
        self.sqlcmd_path = sqlcmd_path or self._find_sqlcmd()
        
    def _find_sqlcmd(self) -> str:
        """Tự động tìm đường dẫn sqlcmd.exe"""
        common_paths = [
            r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\sqlcmd.exe",
            r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\180\Tools\Binn\sqlcmd.exe",
            r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\130\Tools\Binn\sqlcmd.exe",
            r"C:\Program Files\Microsoft SQL Server\150\Tools\Binn\sqlcmd.exe",
            r"C:\Program Files\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\sqlcmd.exe",
            r"C:\Program Files (x86)\Microsoft SQL Server\Client SDK\ODBC\110\Tools\Binn\sqlcmd.exe",
        ]
        
        for path in common_paths:
            if os.path.exists(path):
                return path
        
        try:
            result = subprocess.run(['where', 'sqlcmd'], capture_output=True, text=True)
            if result.returncode == 0:
                return result.stdout.strip().split('\n')[0]
        except:
            pass
        
        return "sqlcmd"
    
    def read_excel(self, sheet_name: str = None) -> pd.DataFrame:
        """Đọc file Excel và trả về DataFrame"""
        try:
            excel_file = pd.ExcelFile(self.excel_path)
            sheets = excel_file.sheet_names
            
            print(f"Đang đọc file Excel: {self.excel_path}")
            print(f"Tìm thấy {len(sheets)} sheet(s): {', '.join(sheets)}")
            
            target_sheet = sheet_name
            if target_sheet and target_sheet not in sheets:
                for s in sheets:
                    if sheet_name.lower() in s.lower() or s.lower() in sheet_name.lower():
                        target_sheet = s
                        break
            
            if target_sheet and target_sheet in sheets:
                print(f"\nĐọc sheet: '{target_sheet}'")
                df = pd.read_excel(self.excel_path, sheet_name=target_sheet, header=None)
                return df
            else:
                print(f"\nCảnh báo: Không tìm thấy sheet '{sheet_name}'")
                return None
                
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {str(e)}")
            raise
    
    def parse_excel_structure(self, df: pd.DataFrame) -> Dict:
        """
        Parse cấu trúc Excel sheet 'Cài đặt SQL'
        
        Cấu trúc:
        - Rows đầu: Thông tin chung (đường dẫn backup, database, user/password, server)
        - SQL chuẩn bị: DECLARE, CREATE TABLE (ở cột cuối)
        - Mỗi database: 1 row với thông tin + script SQL ở cột cuối
        """
        result = {
            'server_name': 'localhost',  # Mặc định localhost
            'backup_path': '',
            'database_path': '',
            'user_name': '',
            'password': '',
            'preparation_sql': [],     # SQL chuẩn bị (DECLARE, CREATE TABLE)
            'declare_sql': '',         # Riêng phần DECLARE để tái sử dụng
            'databases': [],           # Thông tin từng database
        }
        
        num_rows = df.shape[0]
        num_cols = df.shape[1]
        
        # === PARSE THÔNG TIN CHUNG ===
        for row_idx in range(min(12, num_rows)):
            col0 = str(df.iloc[row_idx, 0]).strip() if num_cols > 0 and pd.notna(df.iloc[row_idx, 0]) else ''
            col1 = str(df.iloc[row_idx, 1]).strip() if num_cols > 1 and pd.notna(df.iloc[row_idx, 1]) else ''
            
            col0_lower = col0.lower()
            
            if 'server' in col0_lower and ('name' in col0_lower or 'tên' in col0_lower or 'địa chỉ' in col0_lower):
                result['server_name'] = col1 if col1 else 'localhost'
            elif 'đường dẫn' in col0_lower and 'bak' in col0_lower:
                result['backup_path'] = col1
            elif 'đường dẫn' in col0_lower and ('mdf' in col0_lower or 'database' in col0_lower):
                result['database_path'] = col1
            elif 'user' in col0_lower and ('mới' in col0_lower or 'tên' in col0_lower):
                result['user_name'] = col1
            elif 'password' in col0_lower or 'mật khẩu' in col0_lower:
                result['password'] = col1
        
        # Nếu chưa tìm được user/password, thử từ section "Tạo User"
        if not result['user_name']:
            for row_idx in range(min(12, num_rows)):
                col0 = str(df.iloc[row_idx, 0]).strip() if num_cols > 0 and pd.notna(df.iloc[row_idx, 0]) else ''
                if col0.lower() == 'user':
                    if row_idx + 1 < num_rows:
                        user_val = str(df.iloc[row_idx + 1, 0]).strip() if pd.notna(df.iloc[row_idx + 1, 0]) else ''
                        pass_val = str(df.iloc[row_idx + 1, 1]).strip() if num_cols > 1 and pd.notna(df.iloc[row_idx + 1, 1]) else ''
                        if user_val and user_val != 'nan':
                            result['user_name'] = user_val
                        if pass_val and pass_val != 'nan':
                            result['password'] = pass_val
        
        # === PARSE SQL TỪ CÁC Ô ===
        for row_idx in range(num_rows):
            for col_idx in range(num_cols):
                if pd.notna(df.iloc[row_idx, col_idx]):
                    value = str(df.iloc[row_idx, col_idx]).strip()
                    value_upper = value.upper()
                    
                    # Tìm SQL chuẩn bị (DECLARE, CREATE TABLE #)
                    if value_upper.startswith('DECLARE '):
                        result['declare_sql'] = value
                        result['preparation_sql'].append(value)
                    elif value_upper.startswith('CREATE TABLE #'):
                        result['preparation_sql'].append(value)
                    
                    # Tìm SQL restore (INSERT INTO #FileList ...)
                    elif value_upper.startswith('INSERT INTO #FILELIST'):
                        db_info = {
                            'bak_name': '',
                            'bak_path': '',
                            'db_name': '',
                            'mdf_path': '',
                            'ldf_path': '',
                            'sql_script': value,
                        }
                        
                        if num_cols > 0 and pd.notna(df.iloc[row_idx, 0]):
                            db_info['bak_name'] = str(df.iloc[row_idx, 0]).strip()
                        if num_cols > 1 and pd.notna(df.iloc[row_idx, 1]):
                            db_info['bak_path'] = str(df.iloc[row_idx, 1]).strip()
                        if num_cols > 2 and pd.notna(df.iloc[row_idx, 2]):
                            db_info['db_name'] = str(df.iloc[row_idx, 2]).strip()
                        if num_cols > 3 and pd.notna(df.iloc[row_idx, 3]):
                            db_info['mdf_path'] = str(df.iloc[row_idx, 3]).strip()
                        if num_cols > 4 and pd.notna(df.iloc[row_idx, 4]):
                            db_info['ldf_path'] = str(df.iloc[row_idx, 4]).strip()
                        
                        # === AUTO-FIX: Nếu tên database chứa .bak, sửa lại ===
                        db_name_raw = db_info['db_name']
                        if db_name_raw.lower().endswith('.bak'):
                            clean_name = db_name_raw[:-4]  # Bỏ .bak
                            db_info['db_name_original'] = db_name_raw
                            db_info['db_name'] = clean_name
                            
                            # Sửa SQL script: thay tên DB có .bak thành tên sạch
                            sql = db_info['sql_script']
                            # RESTORE DATABASE name.bak → RESTORE DATABASE [clean_name]
                            sql = sql.replace(f'DATABASE {db_name_raw} ', f'DATABASE [{clean_name}] ')
                            sql = sql.replace(f'DATABASE {db_name_raw}\n', f'DATABASE [{clean_name}]\n')
                            # EXEC name.bak.sys. → EXEC [clean_name].sys.
                            sql = sql.replace(f'{db_name_raw}.sys.', f'[{clean_name}].sys.')
                            # name.bak.mdf → clean_name.mdf (trong path)
                            sql = sql.replace(f'{db_name_raw}.mdf', f'{clean_name}.mdf')
                            sql = sql.replace(f'{db_name_raw}.ldf', f'{clean_name}.ldf')
                            
                            db_info['sql_script'] = sql
                            
                            # Sửa đường dẫn mdf/ldf
                            if db_info['mdf_path'].endswith(f'{db_name_raw}.mdf'):
                                db_info['mdf_path'] = db_info['mdf_path'].replace(f'{db_name_raw}.mdf', f'{clean_name}.mdf')
                            if db_info['ldf_path'].endswith(f'{db_name_raw}.ldf'):
                                db_info['ldf_path'] = db_info['ldf_path'].replace(f'{db_name_raw}.ldf', f'{clean_name}.ldf')
                            
                            print(f"  ⚠ Auto-fix: Tên DB '{db_name_raw}' → '{clean_name}' (bỏ .bak)")
                        
                        result['databases'].append(db_info)
        
        return result
    
    def build_clean_script(self, config: Dict) -> str:
        """
        Xây dựng script SQL chỉ XÓA database (không restore).
        """
        script_parts = []
        
        script_parts.append("-- ============================================")
        script_parts.append("-- SCRIPT XÓA DATABASE")
        script_parts.append("-- Generated by SQL Config Tool")
        script_parts.append("-- ============================================")
        script_parts.append("")
        
        if config['databases']:
            for db_info in config['databases']:
                db_name = db_info.get('db_name', '').strip()
                if db_name and db_name != 'nan':
                    script_parts.append(f"IF DB_ID('{db_name}') IS NOT NULL")
                    script_parts.append("BEGIN")
                    script_parts.append(f"    ALTER DATABASE [{db_name}] SET SINGLE_USER WITH ROLLBACK IMMEDIATE;")
                    script_parts.append(f"    DROP DATABASE [{db_name}];")
                    script_parts.append(f"    PRINT N'Da xoa database: {db_name}';")
                    script_parts.append("END")
                    script_parts.append("ELSE")
                    script_parts.append("BEGIN")
                    script_parts.append(f"    PRINT N'Database {db_name} khong ton tai - bo qua';")
                    script_parts.append("END")
                    script_parts.append("")
            
            # Xóa login user nếu muốn clean hoàn toàn
            user_name = config.get('user_name', '')
            if user_name and user_name != 'nan':
                script_parts.append(f"-- Xóa login (bỏ comment nếu muốn xóa luôn login)")
                script_parts.append(f"-- IF EXISTS (SELECT * FROM sys.server_principals WHERE name = '{user_name}')")
                script_parts.append(f"-- BEGIN")
                script_parts.append(f"--     DROP LOGIN [{user_name}];")
                script_parts.append(f"--     PRINT N'Da xoa login: {user_name}';")
                script_parts.append(f"-- END")
                script_parts.append("")
        
        script_parts.append("PRINT N'';")
        script_parts.append("PRINT N'=== HOAN TAT XOA DATABASE ===';")
        
        return '\n'.join(script_parts)
    
    def build_restore_script(self, config: Dict) -> str:
        """
        Xây dựng script SQL để RESTORE database.
        Mỗi database = 1 batch riêng (phân cách bằng GO).
        Lỗi ở 1 DB không ảnh hưởng DB khác.
        """
        script_parts = []
        
        script_parts.append("-- ============================================")
        script_parts.append("-- SCRIPT CẤU HÌNH SQL SERVER TỰ ĐỘNG")
        script_parts.append("-- Generated by SQL Config Tool")
        script_parts.append("-- ============================================")
        script_parts.append("")
        
        # === BATCH 1: Tạo SQL Login ===
        user_name = config.get('user_name', '')
        password = config.get('password', '')
        
        if user_name and user_name != 'nan':
            script_parts.append("-- === BƯỚC 1: Tạo SQL Login ===")
            script_parts.append(f"IF NOT EXISTS (SELECT * FROM sys.server_principals WHERE name = '{user_name}')")
            script_parts.append("BEGIN")
            pwd = password if password and password != 'nan' else '123456'
            script_parts.append(f"    CREATE LOGIN [{user_name}] WITH PASSWORD = '{pwd}';")
            script_parts.append(f"    PRINT N'Da tao login: {user_name}';")
            script_parts.append("END")
            script_parts.append("ELSE")
            script_parts.append("BEGIN")
            script_parts.append(f"    PRINT N'Login {user_name} da ton tai - bo qua';")
            script_parts.append("END")
            script_parts.append("GO")
            script_parts.append("")
        
        # === BATCH 2: Tạo bảng tạm ===
        create_table_sql = None
        for sql in config['preparation_sql']:
            if sql.upper().startswith('CREATE TABLE #'):
                create_table_sql = sql
                break
        
        if create_table_sql:
            script_parts.append("-- === BƯỚC 2: Tạo bảng tạm ===")
            match = re.search(r'CREATE TABLE\s+(#\w+)', create_table_sql, re.IGNORECASE)
            if match:
                temp_table = match.group(1)
                script_parts.append(f"IF OBJECT_ID('tempdb..{temp_table}') IS NOT NULL DROP TABLE {temp_table};")
            script_parts.append(create_table_sql)
            script_parts.append("GO")
            script_parts.append("")
        
        # === BATCH 3+: Restore từng database ===
        declare_sql = config.get('declare_sql', '')
        
        if config['databases']:
            for i, db_info in enumerate(config['databases']):
                db_name = db_info.get('db_name', f'Database_{i+1}')
                sql_script = db_info.get('sql_script', '')
                
                if not sql_script:
                    continue
                
                script_parts.append(f"-- === Restore [{i+1}/{len(config['databases'])}]: {db_name} ===")
                script_parts.append(f"PRINT N'';")
                script_parts.append(f"PRINT N'=== [{i+1}/{len(config['databases'])}] Dang restore: {db_name} ===';")
                
                # Khai báo lại biến trong mỗi batch (vì GO reset biến)
                if declare_sql:
                    script_parts.append(declare_sql)
                
                script_parts.append(sql_script)
                
                script_parts.append(f"PRINT N'=== Hoan tat restore: {db_name} ===';")
                script_parts.append("GO")
                script_parts.append("")
        
        script_parts.append("-- === KẾT THÚC ===")
        script_parts.append("PRINT N'';")
        script_parts.append("PRINT N'=== HOAN TAT CAU HINH SQL SERVER ===';")
        
        return '\n'.join(script_parts)
    
    def run_sql_script(self, script: str, server: str = None, 
                       dry_run: bool = False) -> Tuple[bool, str]:
        """
        Chạy toàn bộ script SQL bằng sqlcmd (sử dụng file tạm)
        KHÔNG dùng flag -b để 1 batch lỗi không dừng toàn bộ
        """
        if dry_run:
            print("\n[DRY RUN] Script SQL sẽ được chạy:")
            print("-" * 60)
            lines = script.split('\n')
            for line in lines:
                if len(line) > 120:
                    print(f"  {line[:120]}...")
                else:
                    print(f"  {line}")
            print("-" * 60)
            return True, "Dry run mode - không thực thi"
        
        # Ghi script ra file tạm
        script_dir = os.path.dirname(os.path.abspath(__file__))
        temp_sql_path = os.path.join(script_dir, '_temp_config_script.sql')
        
        try:
            with open(temp_sql_path, 'w', encoding='utf-8-sig') as f:
                f.write(script)
            
            print(f"\n📄 Đã tạo script SQL tạm: {temp_sql_path}")
            
            # Xây dựng lệnh sqlcmd (KHÔNG dùng -b)
            cmd_parts = [self.sqlcmd_path]
            
            if server:
                cmd_parts.extend(['-S', server])
            
            cmd_parts.append('-E')  # Windows Authentication
            # ODBC Driver 18+ mặc định mã hóa và không tin chứng chỉ tự ký trên localhost
            cmd_parts.append('-C')
            cmd_parts.extend(['-i', temp_sql_path])
            
            print(f"🚀 Đang chạy sqlcmd...")
            print(f"   Server: {server or 'localhost'}")
            print(f"   Lệnh: {' '.join(cmd_parts)}")
            print("")
            
            result = subprocess.run(
                cmd_parts,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore',
                timeout=600  # Timeout 10 phút (restore có thể lâu)
            )
            
            output = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""
            
            # Phân tích output - chỉ đếm lỗi SQL Server thực sự
            error_count = 0
            success_count = 0
            real_errors = []  # Lưu các lỗi thực sự để hiển thị
            
            if output:
                print("--- Output từ sqlcmd ---")
                for line in output.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                    line_lower = line.lower()
                    
                    # === PHÁT HIỆN LỖI THỰC SỰ ===
                    # Format lỗi SQL Server: "Msg XXX, Level YY" hoặc "ERROR ( message:... )"
                    is_real_error = False
                    
                    # Pattern 1: Msg XXX, Level YY (lỗi SQL Server)
                    if re.search(r'msg\s+\d+,\s*level\s+\d+', line_lower):
                        is_real_error = True
                    # Pattern 2: ERROR ( message:... )
                    elif re.search(r'error\s*\(', line_lower):
                        is_real_error = True
                    # Pattern 3: Các từ khóa lỗi cụ thể
                    elif any(kw in line_lower for kw in [
                        'cannot open', 'cannot find', 'invalid', 'syntax error',
                        'must declare', 'terminating abnormally', 'does not exist',
                        'already exists', 'failed', 'access denied', 'permission denied'
                    ]):
                        is_real_error = True
                    
                    # === PHÁT HIỆN THÀNH CÔNG ===
                    is_success = False
                    if any(kw in line_lower for kw in [
                        'successfully processed', 'restore database successfully',
                        'hoan tat', 'da tao', 'da ton tai', 'da xoa',
                        'processed', 'successfully'
                    ]):
                        is_success = True
                    
                    # === HIỂN THỊ VÀ ĐẾM ===
                    if is_real_error:
                        print(f"  ✗ {line}")
                        error_count += 1
                        real_errors.append(line)
                    elif is_success:
                        print(f"  ✓ {line}")
                        success_count += 1
                    elif line.startswith('==='):
                        print(f"  📌 {line}")
                    else:
                        # Các dòng thông thường (không phải lỗi, không phải thành công)
                        print(f"  {line}")
                
                print("--- Kết thúc output ---")
            
            if stderr:
                print(f"\n⚠ Stderr: {stderr[:500]}")
                if 'certificate chain was issued by an authority that is not trusted' in stderr.lower():
                    print(
                        "\nGợi ý SSL: ODBC Driver 18 cần tin chứng chỉ SQL Server. "
                        "Tool đã thêm -C cho sqlcmd; nếu vẫn lỗi, cài chứng chỉ hợp lệ hoặc chạy lại bản tool mới."
                    )
                error_count += 1
            
            # Chỉ coi là có lỗi nếu có lỗi thực sự HOẶC returncode != 0
            has_error = (error_count > 0) or (result.returncode != 0)
            
            # Tóm tắt: đếm database restore thành công từ output
            db_restore_success = output.count('=== Hoan tat restore:') if output else 0
            db_restore_started = output.count('=== [') if output else 0
            
            if db_restore_started > 0:
                summary = f"{db_restore_success}/{db_restore_started} database restore thành công"
            else:
                summary = f"{success_count} thành công"
            
            if error_count > 0:
                summary += f", {error_count} lỗi"
            
            if has_error and real_errors:
                # Hiển thị các lỗi thực sự
                print(f"\n⚠ Các lỗi phát hiện:")
                for err in real_errors[:5]:  # Chỉ hiển thị 5 lỗi đầu
                    print(f"  - {err}")
            
            # Nếu tất cả database đã restore thành công, coi là thành công
            if db_restore_started > 0 and db_restore_success == db_restore_started and error_count == 0:
                return True, summary
            elif has_error:
                return False, summary
            else:
                return True, summary
                
        except subprocess.TimeoutExpired:
            return False, "Timeout: Script chạy quá 10 phút"
        except Exception as e:
            return False, str(e)
        finally:
            try:
                if os.path.exists(temp_sql_path):
                    os.remove(temp_sql_path)
            except:
                pass
    
    def process_excel(self, dry_run: bool = False, sheet_name: str = None, clean: bool = False):
        """
        Xử lý file Excel và thực thi cấu hình SQL Server
        
        Args:
            dry_run: Chỉ hiển thị script, không thực thi
            sheet_name: Tên sheet cần đọc
            clean: Xóa database cũ trước khi restore
        """
        print("=" * 60)
        print("SQL Server Configuration Tool")
        print("=" * 60)
        
        if clean:
            print("⚠ CHẾ ĐỘ CLEAN: Sẽ XÓA database cũ trước khi restore!")
            print("")
        
        # 1. Đọc Excel
        df = self.read_excel(sheet_name=sheet_name)
        if df is None:
            print("Lỗi: Không đọc được sheet Excel")
            return
        
        # 2. Parse cấu trúc
        config = self.parse_excel_structure(df)
        
        # 3. Hiển thị thông tin
        print("\n" + "=" * 60)
        print("THÔNG TIN CẤU HÌNH ĐÃ ĐỌC")
        print("=" * 60)
        print(f"SQL Server: {config['server_name']}")
        print(f"Đường dẫn backup (.bak): {config['backup_path']}")
        print(f"Đường dẫn database (.mdf): {config['database_path']}")
        print(f"User SQL: {config['user_name']}")
        print(f"Password: {'*' * len(config['password']) if config['password'] else '(không có)'}")
        print(f"SQL chuẩn bị: {len(config['preparation_sql'])} lệnh")
        print(f"Database cần restore: {len(config['databases'])}")
        
        if config['databases']:
            print("\nDanh sách database:")
            for i, db in enumerate(config['databases'], 1):
                print(f"  {i}. {db['db_name']} (từ {db['bak_name']})")
                print(f"     BAK: {db['bak_path']}")
                print(f"     MDF: {db['mdf_path']}")
                print(f"     LDF: {db['ldf_path']}")
        
        if not config['preparation_sql'] and not config['databases']:
            print("\n⚠ Không tìm thấy lệnh SQL nào trong Excel!")
            return
        
        # 4. Xây dựng script
        print("\n" + "=" * 60)
        
        if clean:
            print("XÂY DỰNG SCRIPT XÓA DATABASE")
            print("=" * 60)
            script = self.build_clean_script(config)
        else:
            print("XÂY DỰNG SCRIPT RESTORE")
            print("=" * 60)
            script = self.build_restore_script(config)
        
        total_lines = len(script.split('\n'))
        print(f"Đã tạo script SQL: {total_lines} dòng")
        
        if not clean:
            print(f"Mỗi database = 1 batch riêng (lỗi 1 DB không ảnh hưởng DB khác)")
        
        # 5. Chạy script
        print("\n" + "=" * 60)
        if clean:
            print("BẮT ĐẦU XÓA DATABASE")
        else:
            print("BẮT ĐẦU CẤU HÌNH SQL SERVER")
        print("=" * 60)
        
        server = config.get('server_name', 'localhost')
        success, summary = self.run_sql_script(script, server=server, dry_run=dry_run)
        
        # 6. Kết quả
        print("\n" + "=" * 60)
        if clean:
            if success:
                print(f"✓ XÓA DATABASE THÀNH CÔNG ({summary})")
                print(f"\nĐể restore lại, chạy: CauHinhSQL.bat")
            else:
                print(f"⚠ XÓA DATABASE HOÀN TẤT VỚI LỖI ({summary})")
        else:
            if success:
                print(f"✓ HOÀN TẤT CẤU HÌNH THÀNH CÔNG ({summary})")
            else:
                print(f"⚠ CẤU HÌNH HOÀN TẤT VỚI LỖI ({summary})")
                print(f"\nGợi ý:")
                print(f"  - Kiểm tra file .bak có tồn tại không")
                print(f"  - Kiểm tra đường dẫn mdf/ldf có hợp lệ không")
                print(f"  - Nếu DB đã tồn tại, chạy CauHinhSQL_Clean.bat để xóa trước")
                print(f"  - Kiểm tra quyền trên SQL Server")
        print("=" * 60)


def main():
    """Hàm main để chạy tool"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Tool cấu hình SQL Server từ file Excel')
    parser.add_argument('excel_file', nargs='?', help='Đường dẫn đến file Excel')
    parser.add_argument('--dry-run', action='store_true', help='Chỉ hiển thị script, không thực thi')
    parser.add_argument('--clean', action='store_true', help='Xóa database cũ trước khi restore')
    parser.add_argument('--sqlcmd', default=None, help='Đường dẫn đến sqlcmd.exe')
    parser.add_argument('--sheet', default=None, help='Tên sheet cụ thể cần đọc')
    parser.add_argument('--server', default=None, help='Tên SQL Server (mặc định: localhost)')
    
    args = parser.parse_args()
    
    # Nếu không có excel_file, dùng đường dẫn mặc định
    if not args.excel_file:
        default_excel = str(Path(__file__).resolve().parent.parent / "ExcelCauHinh" / "Settup AXE.xlsx")
        if os.path.exists(default_excel):
            args.excel_file = default_excel
            if not args.sheet:
                args.sheet = "SQL"
        else:
            print("Lỗi: Không tìm thấy file Excel mặc định")
            print(f"Đường dẫn mong đợi: {default_excel}")
            sys.exit(1)
    
    if not os.path.exists(args.excel_file):
        print(f"Lỗi: Không tìm thấy file {args.excel_file}")
        sys.exit(1)
    
    tool = SQLConfigTool(args.excel_file, args.sqlcmd)
    tool.process_excel(dry_run=args.dry_run, sheet_name=args.sheet, clean=args.clean)


if __name__ == '__main__':
    main()
