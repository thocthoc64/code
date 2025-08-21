from pdf2image import convert_from_path
import os

def convert_pdfs_to_images(folder_path, output_folder_path, poppler_path):
    # Lấy tất cả các file PDF trong thư mục
    pdf_files = [f for f in os.listdir(folder_path) if f.lower().endswith('.pdf')]
    
    # Đảm bảo danh sách PDF không rỗng
    if not pdf_files:
        print("Không có file PDF nào trong thư mục.")
        return

    # Tạo thư mục đầu ra nếu chưa tồn tại
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    for pdf_file in pdf_files:
        pdf_path = os.path.join(folder_path, pdf_file)
        
        try:
            # Chuyển đổi các trang PDF thành danh sách ảnh
            images = convert_from_path(pdf_path, poppler_path=poppler_path)
            
            # Lưu từng trang PDF dưới dạng file ảnh JPG
            for i, image in enumerate(images):
                output_image_path = os.path.join(output_folder_path, f"{os.path.splitext(pdf_file)[0]}_page_{i + 1}.jpg")
                image.save(output_image_path, 'JPEG')
                print(f"Đã chuyển trang {i + 1} của {pdf_file} thành file ảnh: {output_image_path}")
        except Exception as e:
            print(f"Lỗi khi chuyển đổi file {pdf_file} thành ảnh: {e}")

# Đường dẫn tới thư mục chứa PDF, thư mục chứa file ảnh đầu ra và Poppler
folder_path = 'D:\\AN THINH NAM\\San_Pham\\Antech\\04_Tai_Lieu' # folder chứa file pdf
output_folder_path = 'D:\\AN THINH NAM\\San_Pham\\Antech\\04_Tai_Lieu\\img' # folder chứa hình ảnh
poppler_path = 'D:\\Tools\\poppler-24.07.0\\Library\\bin'  # Cập nhật đường dẫn Poppler của bạn

convert_pdfs_to_images(folder_path, output_folder_path, poppler_path)
