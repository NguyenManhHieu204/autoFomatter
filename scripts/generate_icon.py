"""
generate_icon.py: Sinh bộ icon ứng dụng chất lượng cao đa độ phân giải
cho phần mềm Auto-Formatter (Chuẩn hóa Thể thức Văn bản Đảng & Hành chính).
Tạo ra tệp assets/app_icon.ico (16, 24, 32, 48, 64, 128, 256) và assets/app_icon.png (512x512).
"""

import os
import math
from PIL import Image, ImageDraw, ImageFilter


def create_app_icon():
    os.makedirs("assets", exist_ok=True)
    size = 1024
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # 1. Vẽ bóng đổ nhẹ cho biểu tượng nền (Drop Shadow)
    shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    s_draw = ImageDraw.Draw(shadow)
    s_draw.rounded_rectangle([70, 70, 954, 954], radius=210, fill=(0, 0, 0, 70))
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=25))
    img.paste(shadow, (0, 15), shadow)

    # 2. Vẽ hình nền squircle đỏ quyền lực / trang nghiêm (Đảng & Nhà nước)
    # Gradient từ đỏ thẫm #991b1b sang đỏ tươi #dc2626 và đỏ cờ #b91c1c
    for i in range(160):
        t = i / 160.0
        # Tọa độ thu nhỏ dần theo gradient
        inset = 60 + i * 0.3
        r = int(185 - t * 50)  # 185 -> 135
        g = int(28 + t * 5)    # 28 -> 33
        b = int(28 + t * 5)    # 28 -> 33
        draw.rounded_rectangle(
            [inset, inset, size - inset, size - inset],
            radius=int(200 - i * 0.2),
            fill=(r, g, b, 255)
        )

    # Viền kim loại vàng hoàng gia (Gold trim border)
    draw.rounded_rectangle([60, 60, size - 60, size - 60], radius=200, outline=(245, 158, 11, 230), width=14)
    draw.rounded_rectangle([68, 68, size - 68, size - 68], radius=192, outline=(251, 191, 36, 160), width=6)

    # 3. Vẽ tờ văn bản hành chính trắng (Document Sheet)
    doc_x0, doc_y0 = 240, 190
    doc_x1, doc_y1 = 784, 834
    fold_size = 100

    # Bóng đổ của tờ văn bản
    doc_shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    ds_draw = ImageDraw.Draw(doc_shadow)
    ds_draw.rounded_rectangle([doc_x0, doc_y0, doc_x1, doc_y1], radius=32, fill=(0, 0, 0, 90))
    doc_shadow = doc_shadow.filter(ImageFilter.GaussianBlur(radius=18))
    img.paste(doc_shadow, (0, 8), doc_shadow)

    # Tờ giấy trắng với góc gấp trên bên phải
    # Polygon thân giấy
    paper_points = [
        (doc_x0 + 32, doc_y0),
        (doc_x1 - fold_size, doc_y0),
        (doc_x1, doc_y0 + fold_size),
        (doc_x1, doc_y1 - 32),
        (doc_x1 - 32, doc_y1),
        (doc_x0 + 32, doc_y1),
        (doc_x0, doc_y1 - 32),
        (doc_x0, doc_y0 + 32),
    ]
    draw.polygon(paper_points, fill=(255, 255, 255, 255))
    # Viền nhẹ cho tờ giấy
    draw.line(paper_points + [paper_points[0]], fill=(226, 232, 240, 255), width=4)

    # Góc gấp (Folded corner flap)
    fold_points = [
        (doc_x1 - fold_size, doc_y0),
        (doc_x1 - fold_size, doc_y0 + fold_size),
        (doc_x1, doc_y0 + fold_size)
    ]
    draw.polygon(fold_points, fill=(203, 213, 225, 255))
    draw.polygon([
        (doc_x1 - fold_size, doc_y0),
        (doc_x1 - fold_size, doc_y0 + fold_size - 4),
        (doc_x1 - 4, doc_y0 + fold_size)
    ], fill=(241, 245, 249, 255))
    draw.line(fold_points, fill=(148, 163, 184, 255), width=4)

    # 4. Chi tiết trên tờ văn bản (Header, Ngôi sao vàng & Thước căn chỉnh thể thức)
    # Biểu tượng Quốc huy / Đảng: Huy hiệu tròn đỏ có ngôi sao vàng 5 cánh ở trên cùng
    crest_center = (512, 315)
    crest_r = 55
    draw.ellipse([
        crest_center[0] - crest_r, crest_center[1] - crest_r,
        crest_center[0] + crest_r, crest_center[1] + crest_r
    ], fill=(220, 38, 38, 255), outline=(245, 158, 11, 255), width=6)

    # Vẽ ngôi sao vàng 5 cánh chuẩn
    def draw_star(center, outer_r, inner_r, fill_color):
        points = []
        for i in range(10):
            r = outer_r if i % 2 == 0 else inner_r
            angle = i * math.pi / 5 - math.pi / 2
            x = center[0] + r * math.cos(angle)
            y = center[1] + r * math.sin(angle)
            points.append((x, y))
        draw.polygon(points, fill=fill_color)

    draw_star(crest_center, outer_r=34, inner_r=14, fill_color=(251, 191, 36, 255))

    # Cột trái Header (Tên cơ quan / Số hiệu văn bản)
    draw.rounded_rectangle([300, 250, 410, 262], radius=6, fill=(100, 116, 139, 255))
    draw.rounded_rectangle([290, 276, 420, 288], radius=6, fill=(148, 163, 184, 255))
    draw.rounded_rectangle([310, 302, 400, 312], radius=5, fill=(203, 213, 225, 255))

    # Cột phải Header (Quốc hiệu / Tiêu ngữ)
    draw.rounded_rectangle([600, 250, 730, 262], radius=6, fill=(100, 116, 139, 255))
    draw.rounded_rectangle([610, 276, 720, 288], radius=6, fill=(148, 163, 184, 255))
    draw.line([(620, 298), (710, 298)], fill=(220, 38, 38, 255), width=4)

    # Đường phân cách / Tiêu đề văn bản (QUYẾT ĐỊNH / NGHỊ QUYẾT)
    draw.rounded_rectangle([390, 410, 634, 432], radius=8, fill=(30, 41, 59, 255))
    draw.rounded_rectangle([360, 448, 664, 464], radius=6, fill=(71, 85, 105, 255))

    # Các đoạn văn bản nội dung được căn chỉnh chuẩn mực (Justified & Tab Indent)
    # Đoạn 1: Lùi đầu dòng
    draw.rounded_rectangle([360, 506, 724, 520], radius=6, fill=(148, 163, 184, 255))
    draw.rounded_rectangle([300, 536, 724, 550], radius=6, fill=(203, 213, 225, 255))
    draw.rounded_rectangle([300, 566, 680, 580], radius=6, fill=(203, 213, 225, 255))

    # Đoạn 2: Lùi đầu dòng
    draw.rounded_rectangle([360, 616, 724, 630], radius=6, fill=(148, 163, 184, 255))
    draw.rounded_rectangle([300, 646, 724, 660], radius=6, fill=(203, 213, 225, 255))
    draw.rounded_rectangle([300, 676, 560, 690], radius=6, fill=(203, 213, 225, 255))

    # Chân trang (Nơi nhận bên trái)
    draw.rounded_rectangle([300, 725, 420, 735], radius=4, fill=(148, 163, 184, 255))
    draw.rounded_rectangle([300, 745, 400, 753], radius=4, fill=(203, 213, 225, 255))
    draw.rounded_rectangle([300, 763, 380, 771], radius=4, fill=(203, 213, 225, 255))

    # 5. Huy hiệu Chuẩn hóa Thành công (Verification Shield / Badge) ở góc phải dưới
    # Tượng trưng cho việc Tự động Định dạng Chuẩn xác 100%
    badge_center = (740, 740)
    badge_r = 110

    # Bóng đổ của Huy hiệu
    b_shadow = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    bs_draw = ImageDraw.Draw(b_shadow)
    bs_draw.ellipse([
        badge_center[0] - badge_r, badge_center[1] - badge_r,
        badge_center[0] + badge_r, badge_center[1] + badge_r
    ], fill=(0, 0, 0, 110))
    b_shadow = b_shadow.filter(ImageFilter.GaussianBlur(radius=15))
    img.paste(b_shadow, (0, 6), b_shadow)

    # Vòng tròn vàng kim loại bên ngoài
    draw.ellipse([
        badge_center[0] - badge_r, badge_center[1] - badge_r,
        badge_center[0] + badge_r, badge_center[1] + badge_r
    ], fill=(245, 158, 11, 255), outline=(251, 191, 36, 255), width=8)

    # Vòng tròn xanh lục bảo chuẩn hóa (Emerald Green Success)
    inner_r = badge_r - 14
    draw.ellipse([
        badge_center[0] - inner_r, badge_center[1] - inner_r,
        badge_center[0] + inner_r, badge_center[1] + inner_r
    ], fill=(22, 163, 74, 255), outline=(187, 247, 208, 200), width=6)

    # Dấu tích V (Checkmark) màu trắng nổi bật
    check_points = [
        (badge_center[0] - 46, badge_center[1] + 2),
        (badge_center[0] - 12, badge_center[1] + 36),
        (badge_center[0] + 46, badge_center[1] - 32)
    ]
    # Vẽ dấu tích dày bằng các nét nối hoặc polygon
    draw.line(check_points, fill=(255, 255, 255, 255), width=24, joint="curve")

    # 6. Xuất ra tệp PNG chất lượng cao (512x512)
    img_512 = img.resize((512, 512), Image.Resampling.LANCZOS)
    png_path = os.path.join("assets", "app_icon.png")
    img_512.save(png_path, "PNG")
    print(f"[OK] Generated: {png_path}")

    # 7. Xuất ra tệp Windows ICO đa kích thước (16, 24, 32, 48, 64, 128, 256)
    ico_path = os.path.join("assets", "app_icon.ico")
    icon_sizes = [(256, 256), (128, 128), (64, 64), (48, 48), (32, 32), (24, 24), (16, 16)]
    img.save(
        ico_path,
        format="ICO",
        sizes=icon_sizes
    )
    print(f"[OK] Generated: {ico_path}")


if __name__ == "__main__":
    create_app_icon()
