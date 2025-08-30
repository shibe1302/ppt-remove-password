import os
import re
import zipfile
import tempfile
import shutil
from pathlib import Path

def remove_pptx_password(pptx_file_path, output_path=None):
    """
    Xóa password khỏi file PPTX bằng cách xóa thẻ p:modifyVerifier
    
    Args:
        pptx_file_path (str): Đường dẫn tới file PPTX
        output_path (str, optional): Đường dẫn file output. Nếu None thì ghi đè file gốc
    
    Returns:
        bool: True nếu thành công, False nếu thất bại
    """
    
    # Kiểm tra file tồn tại
    if not os.path.exists(pptx_file_path):
        print(f"Lỗi: File {pptx_file_path} không tồn tại!")
        return False
    
    # Xác định đường dẫn output
    if output_path is None:
        output_path = pptx_file_path
    
    # Tạo thư mục tạm
    with tempfile.TemporaryDirectory() as temp_dir:
        try:
            print(f"Đang xử lý file: {pptx_file_path}")
            
            # Bước 1: Đổi tên file thành .zip (copy sang thư mục tạm)
            temp_zip_path = os.path.join(temp_dir, "temp_pptx.zip")
            shutil.copy2(pptx_file_path, temp_zip_path)
            print("✓ Đã copy file và đổi extension thành .zip")
            
            # Bước 2: Giải nén file zip
            extract_dir = os.path.join(temp_dir, "extracted")
            with zipfile.ZipFile(temp_zip_path, 'r') as zip_ref:
                zip_ref.extractall(extract_dir)
            print("✓ Đã giải nén file")
            
            # Bước 3: Tìm file ppt/presentation.xml
            presentation_xml_path = os.path.join(extract_dir, "ppt", "presentation.xml")
            
            if not os.path.exists(presentation_xml_path):
                print("Lỗi: Không tìm thấy file ppt/presentation.xml")
                return False
            
            print("✓ Tìm thấy file ppt/presentation.xml")
            
            # Bước 4: Đọc và xử lý XML
            with open(presentation_xml_path, 'r', encoding='utf-8') as file:
                xml_content = file.read()
            
            print("✓ Đã đọc nội dung XML")
            
            # Pattern để xóa thẻ p:modifyVerifier
            pattern = r'<p:modifyVerifier[^>]*\s*/>'
            
            # Đếm số thẻ tìm thấy trước khi xóa
            matches = re.findall(pattern, xml_content, flags=re.DOTALL | re.IGNORECASE)
            
            if matches:
                print(f"✓ Tìm thấy {len(matches)} thẻ p:modifyVerifier")
                # Xóa thẻ p:modifyVerifier
                cleaned_xml = re.sub(pattern, '', xml_content, flags=re.DOTALL | re.IGNORECASE)
                
                # Lưu lại file XML
                with open(presentation_xml_path, 'w', encoding='utf-8') as file:
                    file.write(cleaned_xml)
                
                print("✓ Đã xóa thẻ p:modifyVerifier và lưu file")
            else:
                print("⚠ Không tìm thấy thẻ p:modifyVerifier trong file")
            
            # Bước 5: Nén lại thành zip
            new_zip_path = os.path.join(temp_dir, "new_pptx.zip")
            
            with zipfile.ZipFile(new_zip_path, 'w', zipfile.ZIP_DEFLATED) as zip_ref:
                for root, dirs, files in os.walk(extract_dir):
                    for file in files:
                        file_path = os.path.join(root, file)
                        # Tạo đường dẫn tương đối trong zip
                        arcname = os.path.relpath(file_path, extract_dir)
                        zip_ref.write(file_path, arcname)
            
            print("✓ Đã nén lại thành file zip")
            
            # Bước 6: Đổi tên và copy về vị trí mong muốn
            shutil.copy2(new_zip_path, output_path)
            print(f"✓ Đã lưu file kết quả: {output_path}")
            
            return True
            
        except Exception as e:
            print(f"Lỗi trong quá trình xử lý: {str(e)}")
            return False

def batch_remove_password(input_folder, output_folder=None):
    """
    Xử lý hàng loạt các file PPTX trong một thư mục
    
    Args:
        input_folder (str): Thư mục chứa file PPTX
        output_folder (str, optional): Thư mục output. Nếu None thì ghi đè file gốc
    """
    
    input_path = Path(input_folder)
    
    if not input_path.exists():
        print(f"Lỗi: Thư mục {input_folder} không tồn tại!")
        return
    
    # Tìm tất cả file .pptx
    pptx_files = list(input_path.glob("*.pptx"))
    
    if not pptx_files:
        print("Không tìm thấy file PPTX nào trong thư mục!")
        return
    
    print(f"Tìm thấy {len(pptx_files)} file PPTX")
    
    success_count = 0
    
    for pptx_file in pptx_files:
        print(f"\n--- Xử lý: {pptx_file.name} ---")
        
        # Xác định đường dẫn output
        if output_folder:
            output_path = Path(output_folder) / pptx_file.name
            output_path.parent.mkdir(parents=True, exist_ok=True)
        else:
            output_path = pptx_file
        
        if remove_pptx_password(str(pptx_file), str(output_path)):
            success_count += 1
            print(f"✓ Hoàn thành: {pptx_file.name}")
        else:
            print(f"✗ Thất bại: {pptx_file.name}")
    
    print(f"\n=== Kết quả ===")
    print(f"Tổng số file: {len(pptx_files)}")
    print(f"Thành công: {success_count}")
    print(f"Thất bại: {len(pptx_files) - success_count}")

def main():
    """
    Hàm main để chạy chương trình
    """
    print("=== CÔNG CỤ XÓA PASSWORD PPTX ===")
    print("1. Xử lý 1 file")
    print("2. Xử lý hàng loạt")
    
    choice = input("Chọn tùy chọn (1/2): ").strip()
    
    if choice == "1":
        # Xử lý 1 file
        file_path = input("Nhập đường dẫn file PPTX: ").strip().strip('"')
        
        if not file_path:
            print("Vui lòng nhập đường dẫn file!")
            return
        
        # Hỏi có muốn tạo file mới không
        create_new = input("Tạo file mới? (y/n, mặc định là ghi đè): ").strip().lower()
        
        if create_new == 'y':
            output_path = input("Nhập đường dẫn file output (hoặc Enter để tự động): ").strip().strip('"')
            if not output_path:
                # Tạo tên file mới
                file_path_obj = Path(file_path)
                output_path = str(file_path_obj.parent / f"{file_path_obj.stem}_no_password{file_path_obj.suffix}")
        else:
            output_path = None
        
        if remove_pptx_password(file_path, output_path):
            print("\n🎉 Xử lý thành công!")
        else:
            print("\n❌ Xử lý thất bại!")
    
    elif choice == "2":
        # Xử lý hàng loạt
        input_folder = input("Nhập đường dẫn thư mục chứa file PPTX: ").strip().strip('"')
        
        if not input_folder:
            print("Vui lòng nhập đường dẫn thư mục!")
            return
        
        # Hỏi thư mục output
        output_folder = input("Nhập thư mục output (hoặc Enter để ghi đè): ").strip().strip('"')
        if not output_folder:
            output_folder = None
        
        batch_remove_password(input_folder, output_folder)
    
    else:
        print("Lựa chọn không hợp lệ!")

if __name__ == "__main__":
    import sys
    
    # Nếu có tham số command line
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else None
        
        print(f"Đang xử lý file: {input_file}")
        if remove_pptx_password(input_file, output_file):
            print("🎉 Xử lý thành công!")
        else:
            print("❌ Xử lý thất bại!")
    else:
        # Chạy interactive mode
        main()