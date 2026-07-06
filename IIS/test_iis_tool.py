"""
Script test để kiểm tra IIS Configuration Tool
Tạo file Excel mẫu và test các chức năng
"""

import os
import pandas as pd
from pathlib import Path

def create_sample_excel():
    """
    Tạo file Excel mẫu dựa trên cấu trúc từ hình ảnh
    """
    output_file = "IIS_Config_Sample.xlsx"
    
    # 1. Thông tin chung
    general_data = {
        'Nội dung': [
            'Tên Application chuẩn (theo dự án)',
            'Đường dẫn gốc chứa source',
            'Tên site chuẩn (theo dự án)',
            'Domain Web',
            'Đường dẫn thư mục chứa file vật lý',
            'Đường dẫn CMD'
        ],
        'Giá trị cần nhập': [
            'AXE_Test',
            'E:\\SourceCodeAXE\\',
            'AXE_Test',
            'http://localhost:1221',
            'C:\\Users\\Admin\\Desktop\\Test',
            'c:\\Windows\\system32\\inetsrv\\appcmd'
        ],
        'Ghi chú': [
            '',
            'Bắt buộc phải có dấu \\ ở cuối',
            '',
            'Phải tự cập nhật thủ công',
            'Phải tự cập nhật thủ công',
            ''
        ]
    }
    df_general = pd.DataFrame(general_data)
    
    # 2. Application Pools
    app_pool_data = {
        'Tên Application Pool': ['AXE_Test-Web'],
        '.Net CLR Version': ['v4.0'],
        'Managed Pipeline Mode': ['Integrated'],
        'appcmd command': [
            'c:\\Windows\\system32\\inetsrv\\appcmd add apppool /name:AXE_Test-Web /managedRuntimeVersion:v4.0 /managedPipelineMode:Integrated'
        ]
    }
    df_app_pool = pd.DataFrame(app_pool_data)
    
    # 3. Websites
    website_data = {
        'Tên Website/Application': ['AXE_Test_Web'],
        'Physical Path': ['E:\\SourceCodeAXE\\AXE-Dash'],
        'Type': ['http'],
        'Port': [1221],
        'Host': [''],
        'Tên Application Pool': ['AXE_Test-Web'],
        'appcmd command': [
            'c:\\Windows\\system32\\inetsrv\\appcmd add site /name:"AXE_Test_Web" /physicalPath:E:\\SourceCodeAXE\\AXE-Dash /bindings:http/*:1221:'
        ]
    }
    df_website = pd.DataFrame(website_data)
    
    # 4. Applications
    app_data = {
        'Tên Application': ['login', 'admin', 'resumable', 'uploader', 'storage', 'doc'],
        'Site Name': ['AXE_Test_Web'] * 6,
        'Physical Path': [
            'E:\\SourceCodeAXE\\AXE-Acc',
            'E:\\SourceCodeAXE\\AXE-Admin',
            'E:\\SourceCodeAXE\\AXE-ResumUploader',
            'E:\\SourceCodeAXE\\AXE-Uploader',
            'C:\\Users\\Admin\\Desktop\\Test',
            'E:\\SourceCodeAXE\\AXE-SoHoa'
        ],
        'Sử dụng': ['x'] * 6,
        'Phân hệ': ['Mặc định', 'Mặc định', 'Mặc định', 'Mặc định', 'Mặc định', 'Kho lưu trữ, mượn trả'],
        'appcmd command': [
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/login /physicalPath:E:\\SourceCodeAXE\\AXE-Acc',
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/admin /physicalPath:E:\\SourceCodeAXE\\AXE-Admin',
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/resumable /physicalPath:E:\\SourceCodeAXE\\AXE-ResumUploader',
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/uploader /physicalPath:E:\\SourceCodeAXE\\AXE-Uploader',
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/storage /physicalPath:C:\\Users\\Admin\\Desktop\\Test',
            'c:\\Windows\\system32\\inetsrv\\appcmd add app /site.name:"AXE_Test_Web" /path:/doc /physicalPath:E:\\SourceCodeAXE\\AXE-SoHoa'
        ]
    }
    df_app = pd.DataFrame(app_data)
    
    # 5. Update App Pool for Site
    app_pool_site_data = {
        'Site Name': ['AXE_Test_Web'],
        'Path': [''],
        'Tên Application Pool': ['AXE_Test-Web'],
        'appcmd command': [
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/ /applicationPool:AXE_Test-Web'
        ]
    }
    df_app_pool_site = pd.DataFrame(app_pool_site_data)
    
    # 6. Update App Pool for Applications
    app_pool_app_data = {
        'Site Name': ['AXE_Test_Web'] * 6,
        'Application': ['login', 'admin', 'resumable', 'uploader', 'storage', 'doc'],
        'Tên Application Pool': ['AXE_Test-Web'] * 6,
        'appcmd command': [
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/login /applicationPool:AXE_Test-Web',
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/admin /applicationPool:AXE_Test-Web',
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/resumable /applicationPool:AXE_Test-Web',
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/uploader /applicationPool:AXE_Test-Web',
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/storage /applicationPool:AXE_Test-Web',
            'c:\\Windows\\system32\\inetsrv\\appcmd set app AXE_Test_Web/doc /applicationPool:AXE_Test-Web'
        ]
    }
    df_app_pool_app = pd.DataFrame(app_pool_app_data)
    
    # Ghi vào Excel
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df_general.to_excel(writer, sheet_name='Thông tin chung', index=False)
        df_app_pool.to_excel(writer, sheet_name='Application Pools', index=False)
        df_website.to_excel(writer, sheet_name='Websites', index=False)
        df_app.to_excel(writer, sheet_name='Applications', index=False)
        df_app_pool_site.to_excel(writer, sheet_name='App Pool Site', index=False)
        df_app_pool_app.to_excel(writer, sheet_name='App Pool Applications', index=False)
    
    print(f"✅ Đã tạo file Excel mẫu: {output_file}")
    return output_file

if __name__ == '__main__':
    import sys
    import io
    # Set UTF-8 encoding for output
    if sys.stdout.encoding != 'utf-8':
        sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
    
    print("Tao file Excel mau de test...")
    excel_file = create_sample_excel()
    print(f"\nFile da duoc tao: {excel_file}")
    print("\nBan co the su dung file nay de test IIS Configuration Tool:")
    print(f"  python iis_config_tool.py {excel_file} --dry-run")
