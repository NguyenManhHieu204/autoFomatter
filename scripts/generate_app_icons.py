"""
generate_app_icons.py: Tạo bộ icon vector SVG & PNG sắc nét chuẩn hiện đại (Executive Modern)
cho toàn bộ các nút bấm, thẻ thống kê, bảng trạng thái và vùng kéo thả trong Auto-Formatter.
"""

import os
import sys

ICONS_SVG = {
    # 1. Các nút danh sách tệp
    "file_add": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <line x1="12" y1="18" x2="12" y2="12"/>
  <line x1="9" y1="15" x2="15" y2="15"/>
</svg>""",

    "folder_add": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M4 20h16a2 2 0 0 0 2-2V8a2 2 0 0 0-2-2h-7.93a2 2 0 0 1-1.66-.9l-.82-1.2A2 2 0 0 0 7.93 3H4a2 2 0 0 0-2 2v13c0 1.1.9 2 2 2Z"/>
  <line x1="12" y1="10" x2="12" y2="16"/>
  <line x1="9" y1="13" x2="15" y2="13"/>
</svg>""",

    "trash": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#e11d48" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 6h18"/>
  <path d="M19 6v14c0 1-1 2-2 2H7c-1 0-2-1-2-2V6"/>
  <path d="M8 6V4c0-1 1-2 2-2h4c1 0 2 1 2 2v2"/>
  <line x1="10" y1="11" x2="10" y2="17"/>
  <line x1="14" y1="11" x2="14" y2="17"/>
</svg>""",

    "clear_all": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 12a9 9 0 1 0 9-9 9.75 9.75 0 0 0-6.74 2.74L3 8"/>
  <path d="M3 3v5h5"/>
</svg>""",

    "settings": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#475569" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="3"/>
  <path d="M19.4 15a1.65 1.65 0 0 0 .33 1.82l.06.06a2 2 0 0 1 0 2.83 2 2 0 0 1-2.83 0l-.06-.06a1.65 1.65 0 0 0-1.82-.33 1.65 1.65 0 0 0-1 1.51V21a2 2 0 0 1-2 2 2 2 0 0 1-2-2v-.09A1.65 1.65 0 0 0 9 19.4a1.65 1.65 0 0 0-1.82.33l-.06.06a2 2 0 0 1-2.83 0 2 2 0 0 1 0-2.83l.06-.06a1.65 1.65 0 0 0 .33-1.82 1.65 1.65 0 0 0-1.51-1H3a2 2 0 0 1-2-2 2 2 0 0 1 2-2h.09A1.65 1.65 0 0 0 4.6 9a1.65 1.65 0 0 0-.33-1.82l-.06-.06a2 2 0 0 1 0-2.83 2 2 0 0 1 2.83 0l.06.06a1.65 1.65 0 0 0 1.82.33H9a1.65 1.65 0 0 0 1-1.51V3a2 2 0 0 1 2-2 2 2 0 0 1 2 2v.09a1.65 1.65 0 0 0 1 1.51 1.65 1.65 0 0 0 1.82-.33l.06-.06a2 2 0 0 1 2.83 0 2 2 0 0 1 0 2.83l-.06.06a1.65 1.65 0 0 0-.33 1.82V9a1.65 1.65 0 0 0 1.51 1H21a2 2 0 0 1 2 2 2 2 0 0 1-2 2h-.09a1.65 1.65 0 0 0-1.51 1z"/>
</svg>""",

    # 2. Các nút điều khiển chính
    "play": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#ffffff" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <polygon points="6 3 20 12 6 21 6 3"/>
</svg>""",

    "stop": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#dc2626" stroke="#dc2626" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <rect x="5" y="5" width="14" height="14" rx="2" ry="2"/>
</svg>""",

    # 3. Các nút thao tác kết quả
    "ocr_scan": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M3 7V5a2 2 0 0 1 2-2h2"/>
  <path d="M17 3h2a2 2 0 0 1 2 2v2"/>
  <path d="M21 17v2a2 2 0 0 1-2 2h-2"/>
  <path d="M7 21H5a2 2 0 0 1-2-2v-2"/>
  <line x1="7" y1="12" x2="17" y2="12"/>
  <line x1="7" y1="8" x2="14" y2="8"/>
  <line x1="7" y1="16" x2="12" y2="16"/>
</svg>""",

    "file_word": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <path d="M9 13l2 5 1-3 1 3 2-5"/>
</svg>""",

    "file_pdf": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M14 2H6a2 2 0 0 0-2 2v16a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V8z"/>
  <polyline points="14 2 14 8 20 8"/>
  <line x1="8" y1="13" x2="16" y2="13"/>
  <line x1="8" y1="17" x2="12" y2="17"/>
</svg>""",

    "folder_open": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#334155" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M22 19a2 2 0 0 1-2 2H4a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h5l2 3h9a2 2 0 0 1 2 2z"/>
  <polygon points="2 10 22 10 20 21 4 21 2 10" fill="#f1f5f9"/>
</svg>""",

    # 4. Icon cho 4 thẻ thống kê
    "stat_paragraph": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <line x1="21" y1="6" x2="3" y2="6"/>
  <line x1="15" y1="12" x2="3" y2="12"/>
  <line x1="17" y1="18" x2="3" y2="18"/>
  <circle cx="19" cy="15" r="2" fill="#2563eb"/>
</svg>""",

    "stat_whitespace": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#0891b2" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M4 14v2a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2v-2"/>
  <path d="M12 4l1.5 3.5L17 9l-3.5 1.5L12 14l-1.5-3.5L7 9l3.5-1.5L12 4z" fill="#0891b2"/>
</svg>""",

    "stat_punctuation": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#059669" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M12 20h9"/>
  <path d="M16.5 3.5a2.121 2.121 0 0 1 3 3L7 19l-4 1 1-4L16.5 3.5z"/>
</svg>""",

    "stat_warning": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#d97706" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M12 22s8-4 8-10V5l-8-3-8 3v7c0 6 8 10 8 10z"/>
  <line x1="12" y1="8" x2="12" y2="12"/>
  <line x1="12" y1="16" x2="12.01" y2="16"/>
</svg>""",

    # 5. Icon trạng thái bảng (Table Status Badges)
    "status_pending": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#64748b" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10"/>
  <polyline points="12 6 12 12 16 14"/>
</svg>""",

    "status_running": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
  <path d="M21 12a9 9 0 1 1-6.219-8.56"/>
</svg>""",

    "status_success": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#16a34a" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10" fill="#16a34a"/>
  <polyline points="8 12 11 15 16 9" stroke="#ffffff"/>
</svg>""",

    "status_ocr_success": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#9333ea" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10" fill="#9333ea"/>
  <polyline points="8 12 11 15 16 9" stroke="#ffffff"/>
</svg>""",

    "status_scan": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#d97706" stroke="#ffffff" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
  <path d="M23 19a2 2 0 0 1-2 2H3a2 2 0 0 1-2-2V8a2 2 0 0 1 2-2h4l2-3h6l2 3h4a2 2 0 0 1 2 2z" fill="#d97706"/>
  <circle cx="12" cy="13" r="4" fill="#ffffff" stroke="#d97706"/>
</svg>""",

    "status_error": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="#dc2626" stroke="#ffffff" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
  <circle cx="12" cy="12" r="10" fill="#dc2626"/>
  <line x1="15" y1="9" x2="9" y2="15" stroke="#ffffff"/>
  <line x1="9" y1="9" x2="15" y2="15" stroke="#ffffff"/>
</svg>""",

    # 6. Biểu tượng Vùng Kéo thả (Drop Zone Centerpiece)
    "cloud_upload": """<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 24 24" fill="none" stroke="#2563eb" stroke-width="1.8" stroke-linecap="round" stroke-linejoin="round">
  <path d="M4 14.899A7 7 0 1 1 15.71 8h1.79a4.5 4.5 0 0 1 2.5 8.242"/>
  <path d="M12 12v9"/>
  <path d="m16 16-4-4-4 4"/>
</svg>"""
}


def build_icon_files():
    out_dir = os.path.join("assets", "icons")
    os.makedirs(out_dir, exist_ok=True)

    # 1. Ghi tệp SVG
    for name, svg_content in ICONS_SVG.items():
        svg_path = os.path.join(out_dir, f"{name}.svg")
        with open(svg_path, "w", encoding="utf-8") as f:
            f.write(svg_content.strip())

    print(f"[OK] Da tao {len(ICONS_SVG)} tep SVG trong {out_dir}")

    # 2. Render sang PNG qua QSvgRenderer nếu có PyQt6
    try:
        from PyQt6.QtWidgets import QApplication
        from PyQt6.QtCore import QByteArray, QSize, Qt
        from PyQt6.QtGui import QImage, QPainter
        from PyQt6.QtSvg import QSvgRenderer

        app = QApplication.instance()
        if app is None:
            app = QApplication(sys.argv)

        for name, svg_content in ICONS_SVG.items():
            renderer = QSvgRenderer(QByteArray(svg_content.encode("utf-8")))
            if not renderer.isValid():
                continue

            # Render PNG chuẩn 64x64 và 128x128
            for sz in (64, 128):
                img = QImage(sz, sz, QImage.Format.Format_ARGB32_Premultiplied)
                img.fill(Qt.GlobalColor.transparent)
                painter = QPainter(img)
                painter.setRenderHint(QPainter.RenderHint.Antialiasing)
                renderer.render(painter)
                painter.end()

                png_path = os.path.join(out_dir, f"{name}_{sz}.png" if sz != 64 else f"{name}.png")
                img.save(png_path)

        print(f"[OK] Da render toan bo PNG chat luong cao trong {out_dir}")
    except Exception as e:
        print(f"[!] Chu y khi render PNG: {e}")


if __name__ == "__main__":
    build_icon_files()
