# -*- coding: utf-8 -*-
"""
Tạo QR hàng loạt từ danh sách URL.
Tên file PNG lấy từ utm_campaign (đã chuẩn hoá).
Yêu cầu: pip install qrcode[pil] pillow
"""

import os
import re
import unicodedata
from pathlib import Path
from urllib.parse import urlparse, parse_qs

import qrcode
from PIL import Image, ImageDraw

# ========== CẤU HÌNH ==========
# 1) Đường dẫn logo (bo tròn, dán giữa QR)
logo_path = r"D:/ANTECH/Gốc/Logo bao bi_Antech/Avatar Cover.jpg"

# 2) Thư mục lưu QR đầu ra
output_dir = Path(r"D:/ANTECH/ANTECH/QR_Code/Link_Print")
output_dir.mkdir(parents=True, exist_ok=True)

# 3) Nguồn URL: chọn MỘT trong 2 cách bên dưới
# 3a) Đọc từ file .txt (mỗi dòng một URL)
input_txt = None  # ví dụ: r"D:/ANTECH/urls.txt"  -> để None nếu không dùng

# 3b) Hoặc khai báo trực tiếp trong code
urls = [
    "https://chongthamantech.com/?utm_source=qrcode&utm_medium=print&utm_campaign=trang_chu&utm_content=general",
    "https://chongthamantech.com/san-pham/antech-coat-201/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_coat_201&utm_content=general",
    "https://chongthamantech.com/san-pham/antech-pu-seal/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_pu_seal&utm_content=general",
    "https://chongthamantech.com/san-pham/antech-ac-seal/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_ac_seal&utm_content=general",
    "https://chongthamantech.com/san-pham/antech-coat-eps-2k/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_coat_eps_2k&utm_content=general",
    "https://chongthamantech.com/san-pham/antech-primer/?utm_source=qrcode&utm_medium=print&utm_campaign=antech_primer&utm_content=general",
]

# 4) Tuỳ chỉnh QR & Logo
BOX_SIZE = 10
BORDER = 4
LOGO_SCALE = 0.2  # logo = 20% cạnh QR
ADD_LOGO_STROKE = False  # True nếu muốn viền trắng quanh logo cho dễ đọc
LOGO_STROKE_WIDTH = 6


# ========== HÀM PHỤ ==========
def strip_accents_and_slugify(text: str) -> str:
    """
    Bỏ dấu tiếng Việt, thay khoảng trắng thành '-', giữ a-z0-9-_,
    cắt độ dài cho an toàn tên file.
    """
    text = unicodedata.normalize("NFKD", text)
    text = "".join(ch for ch in text if not unicodedata.combining(ch))
    text = text.lower()
    text = re.sub(r"[^\w\s-]", "", text)      # bỏ ký tự không an toàn
    text = re.sub(r"\s+", "-", text).strip("-")
    text = re.sub(r"-{2,}", "-", text)        # gộp nhiều dấu '-' liên tiếp
    return text[:120] if len(text) > 120 else text or "qr"


def ensure_unique_path(base: Path) -> Path:
    """
    Nếu base đã tồn tại, thêm -1, -2, ... để không ghi đè.
    """
    if not base.exists():
        return base
    stem, suffix = base.stem, base.suffix
    i = 1
    while True:
        candidate = base.with_name(f"{stem}-{i}{suffix}")
        if not candidate.exists():
            return candidate
        i += 1


def pick_filename_from_url(url: str, index: int) -> str:
    """
    Lấy tên từ utm_campaign; nếu thiếu, dùng cuối path hoặc 'qr-{index}'.
    """
    try:
        parsed = urlparse(url)
        q = parse_qs(parsed.query)
        if "utm_campaign" in q and q["utm_campaign"]:
            raw = q["utm_campaign"][0]
        else:
            # fallback: lấy segment path cuối, nếu rỗng thì gen theo index
            raw = Path(parsed.path).name or f"qr-{index+1}"
    except Exception:
        raw = f"qr-{index+1}"
    return strip_accents_and_slugify(raw)


def make_qr_with_center_logo(data: str, logo_path: str) -> Image.Image:
    qr = qrcode.QRCode(
        version=None,  # để fit tự động
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=BOX_SIZE,
        border=BORDER,
    )
    qr.add_data(data)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert("RGB")

    if logo_path and os.path.exists(logo_path):
        logo = Image.open(logo_path).convert("RGBA")

        # Resize logo theo tỉ lệ QR
        qr_w, qr_h = qr_img.size
        logo_size = int(min(qr_w, qr_h) * LOGO_SCALE)
        logo.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)

        # Tạo mask tròn
        mask = Image.new("L", logo.size, 0)
        draw = ImageDraw.Draw(mask)
        draw.ellipse((0, 0, logo.size[0], logo.size[1]), fill=255)

        logo.putalpha(mask)

        # (Tuỳ chọn) Thêm viền trắng giúp tách logo khỏi mô-đun QR
        if ADD_LOGO_STROKE and LOGO_STROKE_WIDTH > 0:
            stroke = Image.new("RGBA", (logo.size[0] + LOGO_STROKE_WIDTH * 2,
                                        logo.size[1] + LOGO_STROKE_WIDTH * 2), (0, 0, 0, 0))
            stroke_mask = Image.new("L", stroke.size, 0)
            d = ImageDraw.Draw(stroke_mask)
            d.ellipse((0, 0, *stroke.size), fill=255)
            stroke_draw = ImageDraw.Draw(stroke)
            stroke_draw.ellipse((0, 0, *stroke.size), fill=(255, 255, 255, 255))
            # dán logo lên giữa stroke
            stroke.paste(logo, (LOGO_STROKE_WIDTH, LOGO_STROKE_WIDTH), logo)
            logo = stroke
            mask = stroke_mask

        # Dán logo vào giữa
        pos = ((qr_w - logo.width) // 2, (qr_h - logo.height) // 2)
        qr_img.paste(logo, pos, mask=logo if logo.mode == "RGBA" else mask)

    return qr_img


def read_urls():
    if input_txt:
        with open(input_txt, "r", encoding="utf-8") as f:
            lines = [ln.strip() for ln in f if ln.strip()]
        return lines
    return urls


# ========== CHẠY ==========
def main():
    url_list = read_urls()
    if not url_list:
        print("❌ Không có URL đầu vào.")
        return

    print(f"🔧 Bắt đầu tạo QR cho {len(url_list)} URL ...")
    ok, fail = 0, 0

    for i, url in enumerate(url_list):
        try:
            filename = pick_filename_from_url(url, i) + ".png"
            out_path = ensure_unique_path(output_dir / filename)

            qr_img = make_qr_with_center_logo(url, logo_path)
            qr_img.save(out_path)
            ok += 1
            print(f"✅ [{ok}] Đã tạo: {out_path.name}")
        except Exception as e:
            fail += 1
            print(f"⚠️ Lỗi URL #{i+1}: {url}\n   ↳ {e}")

    print(f"\n🎉 Hoàn tất! Thành công: {ok} | Lỗi: {fail} | Lưu tại: {output_dir}")

if __name__ == "__main__":
    main()
