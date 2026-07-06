"""
Tool tự động cấu hình Windows Service từ file Excel
Đọc file Excel sheet 'Cài đặt service' và thực thi các lệnh sc create để tạo service
"""

import os
import sys
import subprocess
import pandas as pd
from pathlib import Path
from typing import Dict, List, Optional, Tuple
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


class ServiceConfigTool:
    def __init__(self, excel_path: str):
        """
        Khởi tạo tool cấu hình Windows Service
        
        Args:
            excel_path: Đường dẫn đến file Excel
        """
        self.excel_path = excel_path
        self.sc_path = r"C:\Windows\System32\sc.exe"
        
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
        Parse cấu trúc Excel sheet 'Cài đặt service'
        
        Cấu trúc:
        - Mỗi service có 2 rows:
          + Row 1: "Tên Service XXX" | "TênService" | ... | "sc create ..."
          + Row 2: "Đường dẫn gốc chứa app Service XXX" | "Đường dẫn exe"
        """
        result = {
            'services': [],  # Danh sách service cần tạo
        }
        
        num_rows = df.shape[0]
        num_cols = df.shape[1]
        
        i = 0
        while i < num_rows:
            row = df.iloc[i]
            col0 = str(row[0]).strip() if num_cols > 0 and pd.notna(row[0]) else ''
            
            # Tìm dòng bắt đầu với "Tên Service"
            if 'tên service' in col0.lower():
                service_info = {
                    'name': '',
                    'display_name': '',
                    'exe_path': '',
                    'sc_command': '',
                }
                
                # Lấy tên service từ cột 1
                if num_cols > 1 and pd.notna(row[1]):
                    service_info['name'] = str(row[1]).strip()
                    service_info['display_name'] = service_info['name']
                
                # Lấy lệnh sc create từ cột cuối (thường là cột 3)
                for col_idx in range(num_cols - 1, -1, -1):
                    if pd.notna(row[col_idx]):
                        val = str(row[col_idx]).strip()
                        if val.lower().startswith('sc create'):
                            service_info['sc_command'] = val
                            break
                
                # Lấy đường dẫn exe từ row tiếp theo
                if i + 1 < num_rows:
                    next_row = df.iloc[i + 1]
                    if num_cols > 1 and pd.notna(next_row[1]):
                        service_info['exe_path'] = str(next_row[1]).strip()
                
                # Nếu có đủ thông tin, thêm vào danh sách
                if service_info['name'] and service_info['exe_path']:
                    result['services'].append(service_info)
                    i += 2  # Bỏ qua row đường dẫn
                    continue
            
            i += 1
        
        return result
    
    def extract_sc_command(self, service_info: Dict) -> str:
        """
        Trích xuất hoặc tạo lệnh sc create từ thông tin service
        
        Nếu có sẵn sc_command trong Excel, dùng nó.
        Nếu không, tạo lệnh mới từ name và exe_path.
        """
        if service_info.get('sc_command'):
            return service_info['sc_command']
        
        # Tạo lệnh sc create mới
        name = service_info.get('name', '')
        exe_path = service_info.get('exe_path', '')
        
        if not name or not exe_path:
            return None
        
        # Tạo lệnh sc create
        # sc create ServiceName binPath= "C:\Path\To\Service.exe" start= auto
        cmd = f'sc create {name} binPath= "{exe_path}" start= auto'
        
        # Thêm display name nếu có
        display_name = service_info.get('display_name', name)
        if display_name and display_name != name:
            cmd += f' DisplayName= "{display_name}"'
        
        return cmd
    
    def run_sc_command(self, command: str, dry_run: bool = False) -> Tuple[bool, str]:
        """
        Thực thi lệnh sc (Service Control)
        
        Args:
            command: Lệnh sc cần chạy (vd: "sc create ...")
            dry_run: Nếu True, chỉ in lệnh không chạy thực tế
            
        Returns:
            (success, output): Kết quả thực thi
        """
        if dry_run:
            print(f"[DRY RUN] {command}")
            return True, "Dry run mode"
        
        try:
            # Chạy lệnh qua shell để xử lý đúng dấu ngoặc kép và khoảng trắng
            # Thay thế "sc" bằng đường dẫn đầy đủ nếu cần
            cmd = command
            if cmd.lower().startswith('sc '):
                cmd = self.sc_path + ' ' + cmd[3:]
            
            result = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore',
                shell=True
            )
            
            output = result.stdout.strip() if result.stdout else ""
            stderr = result.stderr.strip() if result.stderr else ""
            
            # sc.exe thường trả về lỗi qua stderr
            if result.returncode == 0:
                return True, output if output else "Thành công"
            else:
                error_msg = stderr if stderr else output if output else "Lỗi không xác định"
                return False, error_msg
                
        except Exception as e:
            return False, str(e)
    
    def check_service_exists(self, service_name: str) -> bool:
        """Kiểm tra service đã tồn tại chưa"""
        try:
            result = subprocess.run(
                [self.sc_path, 'query', service_name],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            return result.returncode == 0
        except:
            return False
    
    def delete_service(self, service_name: str, dry_run: bool = False) -> Tuple[bool, str]:
        """Xóa service (dừng và xóa)"""
        if dry_run:
            print(f"[DRY RUN] sc stop {service_name}")
            print(f"[DRY RUN] sc delete {service_name}")
            return True, "Dry run mode"
        
        try:
            # Dừng service trước
            stop_result = subprocess.run(
                [self.sc_path, 'stop', service_name],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            # Xóa service
            delete_result = subprocess.run(
                [self.sc_path, 'delete', service_name],
                capture_output=True,
                text=True,
                encoding='utf-8',
                errors='ignore'
            )
            
            if delete_result.returncode == 0:
                return True, "Đã xóa service"
            else:
                error_msg = delete_result.stderr.strip() if delete_result.stderr else delete_result.stdout.strip()
                return False, error_msg
                
        except Exception as e:
            return False, str(e)
    
    def process_excel(self, dry_run: bool = False, sheet_name: str = None, clean: bool = False):
        """
        Xử lý file Excel và thực thi cấu hình Windows Service
        
        Args:
            dry_run: Chỉ hiển thị lệnh, không thực thi
            sheet_name: Tên sheet cần đọc
            clean: Xóa service cũ trước khi tạo mới
        """
        print("=" * 60)
        print("Windows Service Configuration Tool")
        print("=" * 60)
        
        if clean:
            print("⚠ CHẾ ĐỘ CLEAN: Sẽ XÓA service cũ trước khi tạo mới!")
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
        print(f"Số lượng service: {len(config['services'])}")
        
        if config['services']:
            print("\nDanh sách service:")
            for i, svc in enumerate(config['services'], 1):
                print(f"  {i}. {svc['name']}")
                print(f"     Đường dẫn: {svc['exe_path']}")
                if svc.get('sc_command'):
                    cmd_display = svc['sc_command'][:80] + "..." if len(svc['sc_command']) > 80 else svc['sc_command']
                    print(f"     Lệnh: {cmd_display}")
        
        if not config['services']:
            print("\n⚠ Không tìm thấy service nào trong Excel!")
            return
        
        # 4. Clean mode: Chỉ xóa service cũ (không tạo mới)
        if clean:
            print("\n" + "=" * 60)
            print("XÓA SERVICE CŨ")
            print("=" * 60)
            
            success_count = 0
            error_count = 0
            
            for svc in config['services']:
                service_name = svc['name']
                if self.check_service_exists(service_name):
                    print(f"\nĐang xóa service: {service_name}")
                    success, output = self.delete_service(service_name, dry_run)
                    if success:
                        print(f"✓ Đã xóa service: {service_name}")
                        success_count += 1
                    else:
                        print(f"✗ Lỗi khi xóa: {output}")
                        error_count += 1
                else:
                    print(f"\n⚠ Service {service_name} không tồn tại - bỏ qua")
            
            # Kết thúc nếu chỉ clean
            print("\n" + "=" * 60)
            if success_count > 0 or error_count == 0:
                print(f"✓ HOÀN TẤT XÓA SERVICE ({success_count} service đã xóa)")
                print(f"\nĐể tạo lại service, chạy: CauHinhService.bat")
            else:
                print(f"⚠ XÓA SERVICE HOÀN TẤT ({success_count} thành công, {error_count} lỗi)")
            print("=" * 60)
            return
        
        # 5. Tạo service mới (chỉ khi không phải clean mode)
        print("\n" + "=" * 60)
        print("BẮT ĐẦU CẤU HÌNH SERVICE")
        print("=" * 60)
        
        success_count = 0
        error_count = 0
        
        for svc in config['services']:
            service_name = svc['name']
            exe_path = svc['exe_path']
            
            print(f"\n--- Service: {service_name} ---")
            
            # Kiểm tra file exe có tồn tại không
            if not dry_run and not os.path.exists(exe_path):
                print(f"✗ Lỗi: File không tồn tại: {exe_path}")
                error_count += 1
                continue
            
            # Kiểm tra service đã tồn tại chưa
            if not clean and self.check_service_exists(service_name):
                print(f"⚠ Service đã tồn tại - bỏ qua")
                print(f"   Để tạo lại, chạy với --clean để xóa trước")
                continue
            
            # Trích xuất hoặc tạo lệnh sc create
            sc_cmd = self.extract_sc_command(svc)
            if not sc_cmd:
                print(f"✗ Lỗi: Không thể tạo lệnh sc create")
                error_count += 1
                continue
            
            # Hiển thị lệnh
            cmd_display = sc_cmd[:100] + "..." if len(sc_cmd) > 100 else sc_cmd
            print(f"Lệnh: {cmd_display}")
            
            # Chạy lệnh
            success, output = self.run_sc_command(sc_cmd, dry_run)
            
            if success:
                success_count += 1
                if output and output != "Thành công":
                    print(f"✓ Thành công: {output}")
                else:
                    print(f"✓ Thành công")
            else:
                error_count += 1
                output_lower = output.lower()
                
                # Xử lý lỗi "already exists"
                if "already exists" in output_lower or "đã tồn tại" in output_lower:
                    print(f"⚠ Service đã tồn tại (bỏ qua)")
                    success_count += 1
                    error_count -= 1
                else:
                    print(f"✗ Lỗi: {output}")
        
        # 6. Kết quả
        print("\n" + "=" * 60)
        if success_count == len(config['services']) and error_count == 0:
            print(f"✓ HOÀN TẤT CẤU HÌNH THÀNH CÔNG ({success_count}/{len(config['services'])} service)")
        else:
            print(f"⚠ CẤU HÌNH HOÀN TẤT ({success_count}/{len(config['services'])} thành công, {error_count} lỗi)")
            if error_count > 0:
                print(f"\nGợi ý:")
                print(f"  - Kiểm tra file .exe có tồn tại không")
                print(f"  - Kiểm tra quyền Administrator (cần để tạo service)")
                print(f"  - Nếu service đã tồn tại, chạy với --clean để xóa trước")
        print("=" * 60)


def main():
    """Hàm main để chạy tool"""
    import argparse
    
    parser = argparse.ArgumentParser(description='Tool cấu hình Windows Service từ file Excel')
    parser.add_argument('excel_file', nargs='?', help='Đường dẫn đến file Excel')
    parser.add_argument('--dry-run', action='store_true', help='Chỉ hiển thị lệnh, không thực thi')
    parser.add_argument('--clean', action='store_true', help='Xóa service cũ trước khi tạo mới')
    parser.add_argument('--sheet', default=None, help='Tên sheet cụ thể cần đọc')
    
    args = parser.parse_args()
    
    # Nếu không có excel_file, dùng đường dẫn mặc định
    if not args.excel_file:
        default_excel = r"C:\Users\Admin\Desktop\ToolAXE\CauHinhHeThongAXE\SERVICE\ExcelCauHinh\Settup AXE.xlsx"
        if os.path.exists(default_excel):
            args.excel_file = default_excel
            if not args.sheet:
                args.sheet = "Cài đặt service"
        else:
            print("Lỗi: Không tìm thấy file Excel mặc định")
            print(f"Đường dẫn mong đợi: {default_excel}")
            sys.exit(1)
    
    if not os.path.exists(args.excel_file):
        print(f"Lỗi: Không tìm thấy file {args.excel_file}")
        sys.exit(1)
    
    tool = ServiceConfigTool(args.excel_file)
    tool.process_excel(dry_run=args.dry_run, sheet_name=args.sheet, clean=args.clean)


if __name__ == '__main__':
    main()
