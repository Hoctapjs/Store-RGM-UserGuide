# -*- coding: utf-8 -*-
"""
Bảng màu + kích thước, trích xuất từ file mẫu
"Hướng dẫn sử dụng Base - 20.03.2026.pdf" (McDonald's Bravo).

Muốn đổi nhận diện thương hiệu: sửa DUY NHẤT file này.
"""
from pptx.util import Emu, Pt

# ---- Màu lấy từ file mẫu ------------------------------------------------
RED = "DB0007"       # đỏ chủ đạo - tiêu đề, thanh nhấn
RED_DARK = "BF0000"  # đỏ đậm - nền slide chương
YELLOW = "FFBC0C"    # vàng - số thứ tự, điểm nhấn
GRAY = "595956"      # xám - chữ nội dung
WHITE = "FFFFFF"
LIGHT = "F5F5F3"     # nền nhạt cho khối phụ
BORDER = "D8D8D4"

# ---- Font ---------------------------------------------------------------
FONT = "Calibri"
FONT_BOLD = "Calibri"

# ---- Khổ slide 16:9 giống file mẫu (960 x 540 pt) -----------------------
SLIDE_W = Pt(960)
SLIDE_H = Pt(540)

# ---- Cỡ chữ -------------------------------------------------------------
# Cộng thêm vào MỌI cỡ chữ trong deck. Muốn chữ to/nhỏ toàn bộ: sửa số này.
SZ_BOOST = 3


def sz(pt):
    """Cỡ chữ đã cộng SZ_BOOST. Dùng cho mọi chỗ đặt cỡ chữ."""
    return Pt(pt + SZ_BOOST)


SZ_TITLE_HERO = sz(50)   # slide bìa
SZ_HERO_SUB = sz(23)     # phụ đề slide bìa
SZ_TITLE = sz(28)        # tiêu đề slide thường
SZ_SECTION = sz(60)      # slide phân chương - trang trống nên chữ to
SZ_SECTION_NUM = sz(128)  # số chương khổng lồ
SZ_SECTION_SUB = sz(24)   # phụ đề slide chương
SZ_BODY = sz(16)
SZ_SMALL = sz(13)
SZ_CALLOUT = sz(14)
SZ_NUM = sz(15)          # số trong bong bóng

# ---- Logo ---------------------------------------------------------------
# Ảnh nền trang trí cho slide bìa / slide chương. Đã canh sẵn theo khổ 16:9,
# hoạ tiết nằm nửa phải -> chữ tự động thu về nửa trái để không bị che.
# Để None nếu không muốn dùng.
HERO_BG = "images/logo for slide.png"

# Logo nhỏ đóng góc slide nội dung. Để None nếu không muốn.
LOGO_MARK = None          # ví dụ: "images/logo.png"
LOGO_MARK_H = Pt(26)      # chiều cao logo góc

# ---- Lề -----------------------------------------------------------------
M_LEFT = Pt(48)
M_TOP = Pt(38)
CONTENT_TOP = Pt(126)    # mép trên vùng nội dung, dưới tiêu đề
