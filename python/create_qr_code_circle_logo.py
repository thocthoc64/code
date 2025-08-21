import qrcode
from PIL import Image, ImageDraw, ImageOps
import os

utm_url = "https://chongthamantech.com/san-pham/mang-gia-cuong-antech/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_waterbar&utm_content=general"

# 1️⃣ Tạo QR code
qr = qrcode.QRCode(
    version=1,
    error_correction=qrcode.constants.ERROR_CORRECT_H,  # Mức chịu lỗi cao để dán logo
    box_size=10,
    border=4,
)
qr.add_data(utm_url)
qr.make(fit=True)

qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

# 2️⃣ Đường dẫn logo
logo_path = r"D:/ANTECH/Gốc/Logo bao bi_Antech/Avatar Cover.jpg"

if os.path.exists(logo_path):
    logo = Image.open(logo_path).convert("RGBA")  # Chuyển sang RGBA để hỗ trợ mask

    # 3️⃣ Resize logo (20% kích thước QR)
    qr_width, qr_height = qr_img.size
    logo_size = int(qr_width * 0.2)
    logo.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)

    # 4️⃣ Bo tròn logo
    mask = Image.new("L", logo.size, 0)  # L = 8-bit mask
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0, logo.size[0], logo.size[1]), fill=255)
    logo.putalpha(mask)  # Áp dụng mask bo tròn cho logo

    # 5️⃣ Căn giữa và dán logo vào QR
    pos = ((qr_width - logo.width) // 2, (qr_height - logo.height) // 2)
    qr_img.paste(logo, pos, mask=logo)

    # 6️⃣ Lưu ảnh QR hoàn chỉnh
    qr_img.save("D:/ANTECH/ANTECH/QR_Code/Link_Print/mang_antech.png")
    print("✅ Đã tạo QR code với logo bo tròn không viền!")
else:
    print(f"❌ Không tìm thấy file logo tại: {logo_path}")
