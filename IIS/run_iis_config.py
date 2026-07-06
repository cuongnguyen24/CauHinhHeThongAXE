"""
Script helper để chạy IIS Configuration Tool
Có thể tìm file Excel tự động hoặc nhập đường dẫn
"""

import os
import sys
import glob
from pathlib import Path
from iis_config_tool import IISConfigTool

def find_excel_files(directory: str = ".") -> list:
    """
    Tìm tất cả file Excel trong thư mục
    """
    excel_extensions = ['*.xlsx', '*.xls', '*.xlsm']
    files = []
    
    for ext in excel_extensions:
        files.extend(glob.glob(os.path.join(directory, ext)))
        files.extend(glob.glob(os.path.join(directory, '**', ext), recursive=True))
    
    return files

def main():
    """
    Hàm main với interactive mode
    """
    print("=" * 60)
    print("IIS Configuration Tool - Interactive Mode")
    print("=" * 60)
    
    # Tìm file Excel
    print("\nĐang tìm file Excel...")
    excel_files = find_excel_files(".")
    
    if excel_files:
        print(f"\nTìm thấy {len(excel_files)} file Excel:")
        for i, file in enumerate(excel_files, 1):
            print(f"  {i}. {file}")
        
        choice = input("\nChọn file Excel (số) hoặc nhập đường dẫn: ").strip()
        
        try:
            # Nếu là số, chọn từ danh sách
            idx = int(choice) - 1
            if 0 <= idx < len(excel_files):
                excel_path = excel_files[idx]
            else:
                print("Số không hợp lệ!")
                return
        except ValueError:
            # Nếu không phải số, coi như đường dẫn
            excel_path = choice
            if not os.path.exists(excel_path):
                print(f"File không tồn tại: {excel_path}")
                return
    else:
        excel_path = input("\nNhập đường dẫn đến file Excel: ").strip()
        if not os.path.exists(excel_path):
            print(f"File không tồn tại: {excel_path}")
            return
    
    # Hỏi có muốn dry-run không
    dry_run_choice = input("\nChạy ở chế độ Dry Run? (y/n, mặc định: n): ").strip().lower()
    dry_run = dry_run_choice == 'y'
    
    if dry_run:
        print("\n⚠ CHẾ ĐỘ DRY RUN - Chỉ hiển thị lệnh, không thực thi")
    
    # Chạy tool
    try:
        tool = IISConfigTool(excel_path)
        tool.process_excel(dry_run=dry_run)
        
        if not dry_run:
            print("\n✅ Hoàn tất! Vui lòng kiểm tra cấu hình IIS.")
        else:
            print("\n✅ Dry run hoàn tất! Xem lại các lệnh trên.")
            
    except Exception as e:
        print(f"\n❌ Lỗi: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == '__main__':
    main()
