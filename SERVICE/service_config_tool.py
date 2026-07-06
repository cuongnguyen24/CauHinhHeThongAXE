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
import shutil
from datetime import datetime
import xml.etree.ElementTree as ET

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

    def _find_sheet_name(self, sheet_hint: str) -> Optional[str]:
        excel_file = pd.ExcelFile(self.excel_path)
        sheets = excel_file.sheet_names
        hint = sheet_hint.lower()

        for sheet in sheets:
            if sheet.lower() == hint:
                return sheet
        for sheet in sheets:
            sheet_lower = sheet.lower()
            if hint in sheet_lower or sheet_lower in hint:
                return sheet
        for sheet in sheets:
            if 'service' in hint and 'service' in sheet.lower():
                return sheet
            if 'webconfig' in hint and 'webconfig' in sheet.lower():
                return sheet
        return None

    def _read_sheet_raw(self, sheet_hint: str) -> Optional[pd.DataFrame]:
        sheet_name = self._find_sheet_name(sheet_hint)
        if not sheet_name:
            print(f"Canh bao: Khong tim thay sheet '{sheet_hint}'")
            return None
        return pd.read_excel(self.excel_path, sheet_name=sheet_name, header=None)

    def _cell_text(self, value) -> str:
        if pd.isna(value):
            return ''
        text = str(value).strip()
        if text.lower() == 'nan':
            return ''
        return text.strip().strip('"').strip("'")

    def _parse_add_snippets(self, df: pd.DataFrame, excel_rows: List[int]) -> List[Dict]:
        snippets = []
        for excel_row in excel_rows:
            row_idx = excel_row - 1
            if row_idx < 0 or row_idx >= len(df):
                continue
            for value in df.iloc[row_idx].tolist():
                text = self._cell_text(value)
                if '<add ' not in text.lower():
                    continue
                try:
                    element = ET.fromstring(text)
                    if element.tag.lower().endswith('add'):
                        snippets.append(dict(element.attrib))
                except ET.ParseError as exc:
                    print(f"Canh bao: Khong doc duoc XML snippet o dong {excel_row}: {exc}")
        return snippets

    def _load_webconfig_sheet(self) -> Dict:
        df = self._read_sheet_raw('Webconfig')
        if df is None:
            return {
                'connection_strings': [],
                'export_app_settings': [],
                'ocr_app_settings': [],
            }

        return {
            'connection_strings': self._parse_add_snippets(df, [4, 5, 6, 7, 8, 9]),
            'export_app_settings': self._parse_add_snippets(df, [16]),
            'ocr_app_settings': self._parse_add_snippets(df, [17, 19, 20]),
        }

    def _ensure_config_section(self, root: ET.Element, section_name: str) -> ET.Element:
        section = root.find(section_name)
        if section is None:
            section = ET.Element(section_name)
            root.insert(0, section)
        return section

    def _upsert_add_node(self, section: ET.Element, match_attr: str, attrs: Dict) -> bool:
        match_value = attrs.get(match_attr)
        if not match_value:
            return False

        target = None
        for child in section.findall('add'):
            if child.get(match_attr) == match_value:
                target = child
                break

        if target is None:
            ET.SubElement(section, 'add', attrs)
        else:
            for key, value in attrs.items():
                target.set(key, value)
        return True

    def _update_xml_config(self, config_path: str, connection_strings: List[Dict],
                           app_settings: List[Dict], dry_run: bool, label: str) -> bool:
        if not connection_strings and not app_settings:
            print(f"\nKhong co cau hinh config cho {label}")
            return True

        print(f"\nCap nhat config cho {label}: {config_path}")
        print(f"  ConnectionStrings: {len(connection_strings)}")
        print(f"  AppSettings: {len(app_settings)}")

        if dry_run:
            for item in connection_strings:
                print(f"  [DRY RUN] connectionString: {item.get('name')}")
            for item in app_settings:
                print(f"  [DRY RUN] appSetting: {item.get('key')}")
            return True

        if not os.path.exists(config_path):
            print(f"  Loi: Khong tim thay file config: {config_path}")
            return False

        try:
            backup_path = f"{config_path}.bak_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            shutil.copy2(config_path, backup_path)

            tree = ET.parse(config_path)
            root = tree.getroot()
            connection_section = self._ensure_config_section(root, 'connectionStrings')
            app_section = self._ensure_config_section(root, 'appSettings')

            for item in connection_strings:
                self._upsert_add_node(connection_section, 'name', item)
            for item in app_settings:
                self._upsert_add_node(app_section, 'key', item)

            if hasattr(ET, 'indent'):
                ET.indent(tree, space='  ')
            tree.write(config_path, encoding='utf-8', xml_declaration=True)
            print(f"  Da backup: {backup_path}")
            print(f"  Da cap nhat: {config_path}")
            return True
        except Exception as exc:
            print(f"  Loi cap nhat config: {exc}")
            return False

    def update_service_configs(self, services: List[Dict], dry_run: bool = False) -> bool:
        webconfig_data = self._load_webconfig_sheet()
        ok = True

        for service in services:
            name = service.get('name', '')
            exe_path = service.get('exe_path', '')
            if not exe_path:
                continue

            config_path = exe_path + '.config'
            name_lower = name.lower()
            if 'export' in name_lower:
                app_settings = webconfig_data['export_app_settings']
                label = name or 'ServiceExport'
            elif 'ocr' in name_lower:
                app_settings = webconfig_data['ocr_app_settings']
                label = name or 'ServiceOCR'
            else:
                print(f"\nBo qua cap nhat config cho service khong xac dinh loai: {name}")
                continue

            if not self._update_xml_config(
                config_path=config_path,
                connection_strings=webconfig_data['connection_strings'],
                app_settings=app_settings,
                dry_run=dry_run,
                label=label
            ):
                ok = False

        return ok
        
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
        
        print("\n" + "=" * 60)
        print("CAP NHAT CONFIG CHO SERVICE")
        print("=" * 60)
        self.update_service_configs(config['services'], dry_run=dry_run)

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
        default_excel = str(Path(__file__).resolve().parent.parent / "ExcelCauHinh" / "Settup AXE.xlsx")
        if os.path.exists(default_excel):
            args.excel_file = default_excel
            if not args.sheet:
                args.sheet = "service"
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
