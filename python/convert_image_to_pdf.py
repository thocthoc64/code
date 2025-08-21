from PIL import Image
import os

def convert_each_image_to_pdf(folder_path, output_folder_path):
    # Lấy tất cả các file trong thư mục với đuôi .jpg hoặc .jpeg
    image_files = [f for f in os.listdir(folder_path) if f.lower().endswith(('.jpg', '.png', '.jpeg'))]
    
    # Đảm bảo danh sách ảnh không rỗng
    if not image_files:
        print("Không có file ảnh hợp lệ nào trong thư mục.")
        return

    # Tạo thư mục đầu ra nếu chưa tồn tại
    if not os.path.exists(output_folder_path):
        os.makedirs(output_folder_path)

    for image_file in image_files:
        image_path = os.path.join(folder_path, image_file)
        output_pdf_path = os.path.join(output_folder_path, os.path.splitext(image_file)[0] + '.pdf')
        
        try:
            # Mở file ảnh
            image = Image.open(image_path)
            # Chuyển đổi ảnh về chế độ RGB (tránh trường hợp ảnh có kênh alpha)
            image = image.convert('RGB')
            # Lưu ảnh thành file PDF
            image.save(output_pdf_path)
            print(f"Đã chuyển đổi ảnh {image_file} thành file PDF: {output_pdf_path}")
        except Exception as e:
            print(f"Lỗi khi chuyển đổi ảnh {image_file} thành PDF: {e}")

# Đường dẫn tới thư mục chứa ảnh và thư mục chứa file PDF đầu ra
folder_path = 'D:\\AN THINH NAM\\San_Pham\\Antech\\04_Tai_Lieu\\img\\Antech AC Seal'
output_folder_path = 'D:\\AN THINH NAM\\San_Pham\\Antech\\04_Tai_Lieu\\img\\Antech AC Seal\\pdf'

convert_each_image_to_pdf(folder_path, output_folder_path)
