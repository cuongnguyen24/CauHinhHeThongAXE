"""
Tool tự động cấu hình IIS từ file Excel
Đọc file Excel và thực thi các lệnh appcmd để cấu hình IIS
"""

import os
import sys
import subprocess
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import json
import io
import re

# Fix encoding for Windows console
if sys.platform == 'win32':
    if sys.stdout.encoding != 'utf-8':
        try:
            sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
            sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')
        except:
            pass

class IISConfigTool:
    def __init__(self, excel_path: str, appcmd_path: str = r"c:\Windows\system32\inetsrv\appcmd.exe"):
        """
        Khởi tạo tool cấu hình IIS
        
        Args:
            excel_path: Đường dẫn đến file Excel
            appcmd_path: Đường dẫn đến appcmd.exe
        """
        self.excel_path = excel_path
        self.appcmd_path = appcmd_path
        self.config = {}
        
    def read_excel(self, sheet_name: str = None) -> Dict:
        """
        Đọc file Excel và parse các sheet thành cấu hình
        
        Args:
            sheet_name: Tên sheet cụ thể cần đọc (None = đọc tất cả)
        """
        try:
            # Đọc tất cả các sheet
            excel_file = pd.ExcelFile(self.excel_path)
            sheets = excel_file.sheet_names
            
            print(f"Đang đọc file Excel: {self.excel_path}")
            print(f"Tìm thấy {len(sheets)} sheet(s): {', '.join(sheets)}")
            
            # Nếu chỉ định sheet cụ thể
            if sheet_name:
                if sheet_name in sheets:
                    print(f"\nĐọc sheet: '{sheet_name}'")
                    df = pd.read_excel(self.excel_path, sheet_name=sheet_name)
                    return {sheet_name: df}
                else:
                    # Tìm sheet có tên tương tự
                    sheet_lower = sheet_name.lower()
                    for s in sheets:
                        if sheet_lower in s.lower() or s.lower() in sheet_lower:
                            print(f"\nTìm thấy sheet tương tự: '{s}' (tìm '{sheet_name}')")
                            df = pd.read_excel(self.excel_path, sheet_name=s)
                            return {s: df}

                    if 'iis' in sheet_lower:
                        for s in sheets:
                            if 'iis' in s.lower():
                                print(f"\nTim thay sheet IIS: '{s}' (tim '{sheet_name}')")
                                df = pd.read_excel(self.excel_path, sheet_name=s)
                                return {s: df}

                    raise ValueError(
                        f"Khong tim thay sheet '{sheet_name}'. "
                        f"Cac sheet hien co: {', '.join(sheets)}"
                    )
            
            # Đọc từng sheet
            data = {}
            for s in sheets:
                df = pd.read_excel(self.excel_path, sheet_name=s)
                data[s] = df
                print(f"\nSheet '{s}': {len(df)} dòng")
            
            return data
            
        except Exception as e:
            print(f"Lỗi khi đọc file Excel: {str(e)}")
            raise
    
    def parse_general_info(self, df: pd.DataFrame) -> Dict:
        """
        Parse thông tin chung từ sheet
        Tìm các giá trị trong các cột "Nội dung" và "Giá trị cần nhập"
        """
        config = {}
        
        # Tìm cột "Nội dung" và "Giá trị cần nhập"
        noi_dung_col = None
        gia_tri_col = None
        
        for col in df.columns:
            col_lower = str(col).lower()
            if 'nội dung' in col_lower or 'content' in col_lower:
                noi_dung_col = col
            elif 'giá trị' in col_lower or 'value' in col_lower or 'cần nhập' in col_lower:
                gia_tri_col = col
        
        # Parse theo cấu trúc cột
        if noi_dung_col and gia_tri_col:
            for idx, row in df.iterrows():
                if pd.notna(row[noi_dung_col]) and pd.notna(row[gia_tri_col]):
                    noi_dung = str(row[noi_dung_col]).strip()
                    gia_tri = str(row[gia_tri_col]).strip()
                    
                    noi_dung_lower = noi_dung.lower()
                    
                    # Tên Application chuẩn
                    if 'tên application' in noi_dung_lower and ('chuẩn' in noi_dung_lower or 'standard' in noi_dung_lower):
                        config['app_name'] = gia_tri
                    # Đường dẫn gốc
                    elif 'đường dẫn gốc' in noi_dung_lower or 'root path' in noi_dung_lower:
                        config['root_path'] = gia_tri if gia_tri.endswith('\\') else gia_tri + '\\'
                    # Tên site chuẩn
                    elif 'tên site' in noi_dung_lower and ('chuẩn' in noi_dung_lower or 'standard' in noi_dung_lower):
                        config['site_name'] = gia_tri
                    # Domain Web
                    elif 'domain web' in noi_dung_lower or ('domain' in noi_dung_lower and 'web' in noi_dung_lower):
                        config['domain'] = gia_tri
                    # Đường dẫn thư mục vật lý
                    elif 'đường dẫn thư mục' in noi_dung_lower and 'vật lý' in noi_dung_lower:
                        config['physical_path'] = gia_tri
                    # Đường dẫn CMD
                    elif 'đường dẫn cmd' in noi_dung_lower or 'cmd path' in noi_dung_lower:
                        config['appcmd_path'] = gia_tri
        
        # Nếu không tìm thấy cột chuẩn, thử parse theo cách cũ
        if not config:
            for idx, row in df.iterrows():
                row_values = {}
                for col in df.columns:
                    if pd.notna(row[col]):
                        row_values[str(col).lower()] = str(row[col]).strip()
                
                for key, value in row_values.items():
                    if 'tên application' in key and ('chuẩn' in key or 'standard' in key):
                        config['app_name'] = value
                    elif 'đường dẫn gốc' in key:
                        config['root_path'] = value if value.endswith('\\') else value + '\\'
                    elif 'tên site' in key and ('chuẩn' in key or 'standard' in key):
                        config['site_name'] = value
                    elif 'domain web' in key:
                        config['domain'] = value
                    elif 'đường dẫn thư mục' in key and 'vật lý' in key:
                        config['physical_path'] = value
        
        return config
    
    def parse_app_pools(self, df: pd.DataFrame) -> List[Dict]:
        """
        Parse thông tin Application Pools
        Tìm các dòng chứa thông tin app pool
        """
        app_pools = []
        
        for idx, row in df.iterrows():
            pool = {}
            row_data = {}
            
            # Lưu tất cả giá trị của dòng
            for col in df.columns:
                if pd.notna(row[col]):
                    col_name = str(col).lower()
                    value = str(row[col]).strip()
                    row_data[col_name] = value
            
            # Tìm các trường
            for key, value in row_data.items():
                if 'tên application pool' in key or 'application pool name' in key:
                    pool['name'] = value
                elif 'clr version' in key or 'runtime version' in key or '.net clr' in key:
                    pool['runtime_version'] = value.replace('v', '') if value.startswith('v') else value
                    if not pool['runtime_version'].startswith('v'):
                        pool['runtime_version'] = 'v' + pool['runtime_version']
                elif 'pipeline mode' in key or 'managed pipeline' in key:
                    pool['pipeline_mode'] = value
            
            # Nếu tìm thấy tên app pool, thêm vào danh sách
            if pool.get('name'):
                # Set giá trị mặc định nếu thiếu
                if 'runtime_version' not in pool:
                    pool['runtime_version'] = 'v4.0'
                if 'pipeline_mode' not in pool:
                    pool['pipeline_mode'] = 'Integrated'
                app_pools.append(pool)
        
        return app_pools
    
    def parse_websites(self, df: pd.DataFrame) -> List[Dict]:
        """
        Parse thông tin Websites
        """
        websites = []
        
        for idx, row in df.iterrows():
            site = {}
            for col in df.columns:
                if pd.notna(row[col]):
                    value = str(row[col]).strip()
                    
                    if 'tên website' in str(col).lower() or 'website name' in str(col).lower():
                        site['name'] = value
                    elif 'physical path' in str(col).lower() or 'đường dẫn' in str(col).lower():
                        if 'physical' in str(col).lower() or 'vật lý' in str(col).lower():
                            site['physical_path'] = value
                    elif 'port' in str(col).lower():
                        try:
                            site['port'] = int(value)
                        except:
                            pass
                    elif 'type' in str(col).lower() and 'http' in value.lower():
                        site['type'] = 'http'
                    elif 'application pool' in str(col).lower():
                        site['app_pool'] = value
            
            if site.get('name'):
                websites.append(site)
        
        return websites
    
    def parse_applications(self, df: pd.DataFrame) -> List[Dict]:
        """
        Parse thông tin Applications
        Tìm các dòng chứa thông tin application
        """
        applications = []
        
        for idx, row in df.iterrows():
            app = {}
            row_data = {}
            
            # Lưu tất cả giá trị của dòng
            for col in df.columns:
                if pd.notna(row[col]):
                    col_name = str(col).lower()
                    value = str(row[col]).strip()
                    row_data[col_name] = value
            
            # Tìm các trường
            for key, value in row_data.items():
                if ('tên application' in key or 'application name' in key) and 'pool' not in key:
                    app['name'] = value
                elif 'site name' in key:
                    app['site_name'] = value
                elif 'physical path' in key or ('đường dẫn' in key and 'vật lý' in key):
                    app['physical_path'] = value
                elif 'path' in key and ('/' in value or value.startswith('/')):
                    app['path'] = value if value.startswith('/') else '/' + value
            
            # Nếu có tên application, tạo path mặc định nếu chưa có
            if app.get('name'):
                if 'path' not in app:
                    app['path'] = '/' + app['name']
                if app.get('site_name'):
                    applications.append(app)
        
        return applications
    
    def run_appcmd(self, command: str, dry_run: bool = False) -> Tuple[bool, str]:
        """
        Thực thi lệnh appcmd
        
        Args:
            command: Lệnh appcmd cần chạy
            dry_run: Nếu True, chỉ in lệnh không chạy thực tế
            
        Returns:
            (success, output): Kết quả thực thi
        """
        if dry_run:
            print(f"[DRY RUN] {command}")
            return True, "Dry run mode"
        
        try:
            # Làm sạch lệnh: loại bỏ khoảng trắng thừa
            # Thay thế nhiều khoảng trắng liên tiếp bằng một khoảng trắng
            command_cleaned = re.sub(r'\s+', ' ', command.strip())
            
            # Tách lệnh appcmd và các tham số
            parts = command_cleaned.split()
            if parts[0].endswith('appcmd') or parts[0].endswith('appcmd.exe'):
                # Thay thế appcmd bằng đường dẫn đầy đủ
                parts[0] = self.appcmd_path
            
            result = subprocess.run(
                parts,
                capture_output=True,
                text=True,
                shell=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            # Nếu returncode == 0, coi như thành công (kể cả khi không có output)
            if result.returncode == 0:
                output = result.stdout.strip() if result.stdout else ""
                # Nếu không có output, có thể là lệnh thành công nhưng không có thông báo
                return True, output if output else "Thành công (không có output)"
            else:
                # Có lỗi, lấy thông báo lỗi
                error_msg = result.stderr.strip() if result.stderr else result.stdout.strip() if result.stdout else "Lỗi không xác định"
                return False, error_msg
                
        except Exception as e:
            return False, str(e)
    
    def create_app_pool(self, pool: Dict, dry_run: bool = False) -> bool:
        """
        Tạo Application Pool
        """
        name = pool.get('name')
        runtime_version = pool.get('runtime_version', 'v4.0')
        pipeline_mode = pool.get('pipeline_mode', 'Integrated')
        
        command = f'{self.appcmd_path} add apppool /name:{name} /managedRuntimeVersion:{runtime_version} /managedPipelineMode:{pipeline_mode}'
        
        print(f"\nTạo Application Pool: {name}")
        success, output = self.run_appcmd(command, dry_run)
        
        if success:
            print(f"✓ Đã tạo Application Pool: {name}")
        else:
            if "already exists" in output.lower() or "đã tồn tại" in output.lower():
                print(f"⚠ Application Pool {name} đã tồn tại")
                return True
            else:
                print(f"✗ Lỗi: {output}")
        
        return success
    
    def create_website(self, site: Dict, dry_run: bool = False) -> bool:
        """
        Tạo Website
        """
        name = site.get('name')
        physical_path = site.get('physical_path')
        port = site.get('port', 80)
        
        command = f'{self.appcmd_path} add site /name:"{name}" /physicalPath:{physical_path} /bindings:http/*:{port}:'
        
        print(f"\nTạo Website: {name}")
        success, output = self.run_appcmd(command, dry_run)
        
        if success:
            print(f"✓ Đã tạo Website: {name}")
        else:
            if "already exists" in output.lower() or "đã tồn tại" in output.lower():
                print(f"⚠ Website {name} đã tồn tại")
                return True
            else:
                print(f"✗ Lỗi: {output}")
        
        return success
    
    def create_application(self, app: Dict, dry_run: bool = False) -> bool:
        """
        Tạo Application
        """
        name = app.get('name')
        site_name = app.get('site_name')
        physical_path = app.get('physical_path')
        path = app.get('path', f'/{name}')
        
        command = f'{self.appcmd_path} add app /site.name:"{site_name}" /path:{path} /physicalPath:{physical_path}'
        
        print(f"\nTạo Application: {name} tại {path}")
        success, output = self.run_appcmd(command, dry_run)
        
        if success:
            print(f"✓ Đã tạo Application: {name}")
        else:
            if "already exists" in output.lower() or "đã tồn tại" in output.lower():
                print(f"⚠ Application {name} đã tồn tại")
                return True
            else:
                print(f"✗ Lỗi: {output}")
        
        return success
    
    def set_app_pool_for_site(self, site_name: str, app_pool_name: str, dry_run: bool = False) -> bool:
        """
        Gán Application Pool cho Website
        """
        command = f'{self.appcmd_path} set app {site_name}/ /applicationPool:{app_pool_name}'
        
        print(f"\nGán Application Pool {app_pool_name} cho Site {site_name}")
        success, output = self.run_appcmd(command, dry_run)
        
        if success:
            print(f"✓ Đã gán Application Pool cho Site")
        else:
            print(f"✗ Lỗi: {output}")
        
        return success
    
    def set_app_pool_for_app(self, site_name: str, app_path: str, app_pool_name: str, dry_run: bool = False) -> bool:
        """
        Gán Application Pool cho Application
        """
        command = f'{self.appcmd_path} set app {site_name}{app_path} /applicationPool:{app_pool_name}'
        
        print(f"\nGán Application Pool {app_pool_name} cho Application {app_path}")
        success, output = self.run_appcmd(command, dry_run)
        
        if success:
            print(f"✓ Đã gán Application Pool cho Application")
        else:
            print(f"✗ Lỗi: {output}")
        
        return success
    
    def extract_appcmd_commands(self, df: pd.DataFrame) -> List[str]:
        """
        Trích xuất các lệnh appcmd trực tiếp từ Excel nếu có
        Chỉ lấy các lệnh hợp lệ (có chứa 'add' hoặc 'set')
        """
        commands = []
        
        for idx, row in df.iterrows():
            for col in df.columns:
                if pd.notna(row[col]):
                    value = str(row[col]).strip()
                    # Tìm các dòng chứa lệnh appcmd hợp lệ
                    # Lệnh hợp lệ phải có: appcmd và (add hoặc set)
                    if 'appcmd' in value.lower():
                        # Kiểm tra xem có phải là lệnh hợp lệ không (có add hoặc set)
                        if ' add ' in value.lower() or ' set ' in value.lower():
                            # Làm sạch lệnh: loại bỏ khoảng trắng thừa
                            command = re.sub(r'\s+', ' ', value.strip())
                            
                            # Thay thế đường dẫn appcmd nếu cần
                            if 'c:\\Windows\\system32\\inetsrv\\appcmd' in command:
                                command = command.replace('c:\\Windows\\system32\\inetsrv\\appcmd', self.appcmd_path)
                            elif command.lower().startswith('appcmd'):
                                # Nếu chỉ có 'appcmd', thêm đường dẫn đầy đủ
                                command = command.replace('appcmd', self.appcmd_path, 1)
                            
                            if ' add site ' in command.lower() and '/name:"' in command:
                                command = re.sub(r'/name:"([^"]+)"', r'/name:\1', command)
                            
                            # Sửa lệnh add app - loại bỏ dấu ngoặc kép thừa trong /site.name:"..."
                            # Thay /site.name:"SITE_NAME" thành /site.name:SITE_NAME
                            if ' add app ' in command.lower() and '/site.name:"' in command:
                                command = re.sub(r'/site\.name:"([^"]+)"', r'/site.name:\1', command)
                            
                            # Sửa format lệnh set app - loại bỏ dấu ngoặc kép thừa và đảm bảo chỉ có 1 cặp
                            if ' set app ' in command.lower():
                                # Tìm và trích xuất identifier (loại bỏ tất cả dấu ngoặc kép cũ)
                                # Pattern: set app ["']?IDENTIFIER["']? /applicationPool
                                match = re.search(r'set app\s+["\']*([^"\'\s/]+)(/.*?)\s*["\']*\s*/applicationPool', command, re.IGNORECASE)
                                if match:
                                    site_name = match.group(1).strip()
                                    app_path = match.group(2).strip().rstrip('"').rstrip("'")
                                    identifier = f"{site_name}{app_path}"
                                    # Thay thế toàn bộ phần set app với format đúng (chỉ 1 cặp dấu ngoặc kép)
                                    old_pattern = r'set app\s+["\']*[^"\'\s/]+/.*?\s*["\']*\s*/applicationPool'
                                    new_pattern = f'set app "{identifier}" /applicationPool'
                                    command = re.sub(old_pattern, new_pattern, command, flags=re.IGNORECASE)
                            
                            # Chỉ thêm nếu là lệnh hợp lệ (có ít nhất 3 phần: appcmd, add/set, và tham số)
                            parts = command.split()
                            if len(parts) >= 3:
                                commands.append(command)
        
        return commands
    
    def process_excel(self, dry_run: bool = False, sheet_name: str = None):
        """
        Xử lý file Excel và thực thi cấu hình
        
        Args:
            dry_run: Chế độ dry run (chỉ hiển thị, không thực thi)
            sheet_name: Tên sheet cụ thể cần đọc (None = đọc tất cả)
        """
        print("=" * 60)
        print("IIS Configuration Tool")
        print("=" * 60)
        
        # Đọc Excel
        excel_data = self.read_excel(sheet_name=sheet_name)
        
        # Tìm sheet chứa thông tin cấu hình
        general_info = {}
        app_pools = []
        websites = []
        applications = []
        appcmd_commands = []
        
        # Parse từ các sheet
        for sheet_name, df in excel_data.items():
            sheet_lower = sheet_name.lower()
            
            # Trích xuất lệnh appcmd trực tiếp nếu có
            commands = self.extract_appcmd_commands(df)
            appcmd_commands.extend(commands)
            
            # Nếu sheet là "Cài đặt IIS", parse tất cả từ đó
            if 'cài đặt iis' in sheet_lower:
                general_info.update(self.parse_general_info(df))
                app_pools.extend(self.parse_app_pools(df))
                websites.extend(self.parse_websites(df))
                applications.extend(self.parse_applications(df))
            elif 'thông tin chung' in sheet_lower or 'general' in sheet_lower:
                general_info.update(self.parse_general_info(df))
            elif 'app pool' in sheet_lower or 'application pool' in sheet_lower:
                pools = self.parse_app_pools(df)
                app_pools.extend(pools)
            elif 'website' in sheet_lower and 'application' not in sheet_lower:
                sites = self.parse_websites(df)
                websites.extend(sites)
            elif 'application' in sheet_lower and 'pool' not in sheet_lower:
                apps = self.parse_applications(df)
                applications.extend(apps)
        
        # Nếu không tìm thấy theo tên sheet, thử parse từ sheet đầu tiên
        if not appcmd_commands and not app_pools and not websites and not applications:
            print("\nKhông tìm thấy sheet theo tên chuẩn, đang thử parse từ sheet đầu tiên...")
            first_sheet = list(excel_data.values())[0]
            # Thử parse tất cả từ sheet đầu tiên
            general_info.update(self.parse_general_info(first_sheet))
            app_pools.extend(self.parse_app_pools(first_sheet))
            websites.extend(self.parse_websites(first_sheet))
            applications.extend(self.parse_applications(first_sheet))
            commands = self.extract_appcmd_commands(first_sheet)
            appcmd_commands.extend(commands)
        
        print("\n" + "=" * 60)
        print("THÔNG TIN CẤU HÌNH ĐÃ ĐỌC")
        print("=" * 60)
        print(f"General Info: {general_info}")
        print(f"Application Pools: {len(app_pools)}")
        print(f"Websites: {len(websites)}")
        print(f"Applications: {len(applications)}")
        print(f"Appcmd Commands (direct): {len(appcmd_commands)}")
        
        # Thực thi cấu hình
        print("\n" + "=" * 60)
        print("BẮT ĐẦU CẤU HÌNH IIS")
        print("=" * 60)
        
        # Ưu tiên: Nếu có lệnh appcmd trực tiếp, chạy chúng trước
        if appcmd_commands:
            print("\n--- Chạy các lệnh appcmd trực tiếp từ Excel ---")
            for cmd in appcmd_commands:
                print(f"\nLệnh: {cmd}")
                success, output = self.run_appcmd(cmd, dry_run)
                if success:
                    if output and output != "Thành công (không có output)":
                        print(f"✓ Thành công: {output}")
                    else:
                        print(f"✓ Thành công")
                else:
                    # Kiểm tra các trường hợp đặc biệt
                    output_lower = output.lower()
                    
                    # Xử lý lỗi "already exists" - thử dùng set thay vì add
                    if "duplicate" in output_lower or "already exists" in output_lower or "đã tồn tại" in output_lower:
                        # Nếu là lệnh add, thử chuyển sang set
                        if " add " in cmd.lower():
                            # Thử dùng set thay vì add
                            cmd_set = cmd.replace(" add ", " set ", 1)
                            print(f"  → Đã tồn tại, thử cập nhật...")
                            success_set, output_set = self.run_appcmd(cmd_set, dry_run)
                            if success_set:
                                print(f"✓ Đã cập nhật thành công")
                            else:
                                print(f"⚠ Đã tồn tại (bỏ qua)")
                        else:
                            print(f"⚠ Đã tồn tại (bỏ qua)")
                    
                    # Xử lý lỗi "cannot find" - thử format không có dấu ngoặc kép
                    elif "cannot find" in output_lower or "not found" in output_lower:
                        # Thử sửa format lệnh set app - loại bỏ tất cả dấu ngoặc kép
                        if " set app " in cmd.lower():
                            # Trích xuất identifier (loại bỏ tất cả dấu ngoặc kép)
                            match = re.search(r'set app\s+["\']*([^"\'\s/]+)(/.*?)\s*["\']*\s*/applicationPool', cmd, re.IGNORECASE)
                            if match:
                                site_name = match.group(1).strip()
                                app_path = match.group(2).strip().rstrip('"').rstrip("'")
                                identifier = f"{site_name}{app_path}"
                                # Tạo lệnh mới với format không có dấu ngoặc kép
                                new_cmd = re.sub(
                                    r'set app\s+["\']*[^"\'\s/]+/.*?\s*["\']*\s*/applicationPool',
                                    f'set app {identifier} /applicationPool',
                                    cmd,
                                    flags=re.IGNORECASE
                                )
                                print(f"  → Thử lại với format không có dấu ngoặc kép...")
                                success_retry, output_retry = self.run_appcmd(new_cmd, dry_run)
                                if success_retry:
                                    print(f"✓ Thành công")
                                else:
                                    print(f"✗ Lỗi: Không tìm thấy đối tượng - {output}")
                            else:
                                print(f"✗ Lỗi: Không tìm thấy đối tượng - {output}")
                        else:
                            print(f"✗ Lỗi: Không tìm thấy đối tượng - {output}")
                    else:
                        print(f"✗ Lỗi: {output}")
        
        # Nếu không có lệnh trực tiếp, tạo từ cấu hình
        if not appcmd_commands:
            # 1. Tạo Application Pools
            for pool in app_pools:
                self.create_app_pool(pool, dry_run)
            
            # 2. Tạo Websites
            for site in websites:
                self.create_website(site, dry_run)
            
            # 3. Tạo Applications
            for app in applications:
                self.create_application(app, dry_run)
            
            # 4. Gán Application Pool cho Site
            for site in websites:
                app_pool = site.get('app_pool') or general_info.get('app_pool')
                if app_pool:
                    self.set_app_pool_for_site(site['name'], app_pool, dry_run)
            
            # 5. Gán Application Pool cho Applications
            app_pool_name = general_info.get('app_pool') or (app_pools[0]['name'] if app_pools else None)
            for app in applications:
                if app_pool_name:
                    app_path = app.get('path', f"/{app['name']}")
                    self.set_app_pool_for_app(app['site_name'], app_path, app_pool_name, dry_run)
        
        print("\n" + "=" * 60)
        print("HOÀN TẤT CẤU HÌNH")
        print("=" * 60)


def main():
    """
    Hàm main để chạy tool
    """
    import argparse
    
    parser = argparse.ArgumentParser(description='Tool cấu hình IIS từ file Excel')
    parser.add_argument('excel_file', nargs='?', help='Đường dẫn đến file Excel')
    parser.add_argument('--dry-run', action='store_true', help='Chỉ hiển thị lệnh, không thực thi')
    parser.add_argument('--appcmd', default=r'c:\Windows\system32\inetsrv\appcmd.exe', 
                       help='Đường dẫn đến appcmd.exe')
    parser.add_argument('--sheet', default=None, help='Tên sheet cụ thể cần đọc (mặc định: đọc tất cả)')
    
    args = parser.parse_args()
    
    # Nếu không có excel_file, dùng đường dẫn mặc định
    if not args.excel_file:
        default_excel = str(Path(__file__).resolve().parent.parent / "ExcelCauHinh" / "Settup AXE.xlsx")
        if os.path.exists(default_excel):
            args.excel_file = default_excel
            if not args.sheet:
                args.sheet = "IIS"
        else:
            print("Lỗi: Không tìm thấy file Excel mặc định")
            print(f"Đường dẫn mong đợi: {default_excel}")
            sys.exit(1)
    
    if not os.path.exists(args.excel_file):
        print(f"Lỗi: Không tìm thấy file {args.excel_file}")
        sys.exit(1)
    
    tool = IISConfigTool(args.excel_file, args.appcmd)
    tool.process_excel(dry_run=args.dry_run, sheet_name=args.sheet)


if __name__ == '__main__':
    main()
