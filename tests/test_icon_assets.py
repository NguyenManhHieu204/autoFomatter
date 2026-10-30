import os
import unittest
from PIL import Image
from core.path_utils import get_app_icon_path, get_assets_dir


class TestIconAssets(unittest.TestCase):
    def test_app_icon_files_exist(self):
        assets_dir = get_assets_dir()
        self.assertTrue(os.path.exists(assets_dir), "Thư mục assets phải tồn tại")
        
        ico_path = os.path.join(assets_dir, "app_icon.ico")
        png_path = os.path.join(assets_dir, "app_icon.png")

        self.assertTrue(os.path.exists(ico_path), "Tệp app_icon.ico phải tồn tại")
        self.assertTrue(os.path.exists(png_path), "Tệp app_icon.png phải tồn tại")

    def test_ico_resolution_sizes(self):
        ico_path = os.path.join(get_assets_dir(), "app_icon.ico")
        with Image.open(ico_path) as img:
            # Kiểm tra xem ico có kích thước tiêu chuẩn không
            self.assertGreaterEqual(img.size[0], 16)

    def test_png_high_res(self):
        png_path = os.path.join(get_assets_dir(), "app_icon.png")
        with Image.open(png_path) as img:
            self.assertEqual(img.size, (512, 512), "Tệp PNG phải đạt độ phân giải tối thiểu 512x512")
            self.assertEqual(img.mode, "RGBA", "Tệp PNG phải hỗ trợ kênh trong suốt RGBA")

    def test_path_utils_icon_resolution(self):
        icon_path = get_app_icon_path()
        self.assertTrue(bool(icon_path))
        self.assertTrue(os.path.exists(icon_path))

    def test_ui_icon_factory_all_keys(self):
        import sys
        from PyQt6.QtWidgets import QApplication
        from ui.icon_factory import get_ui_icon, get_ui_pixmap

        app = QApplication.instance() or QApplication(sys.argv)
        keys = [
            'file_add', 'folder_add', 'trash', 'clear_all', 'settings',
            'play', 'stop', 'ocr_scan', 'file_word', 'file_pdf', 'folder_open',
            'stat_paragraph', 'stat_whitespace', 'stat_punctuation', 'stat_warning',
            'status_pending', 'status_running', 'status_success', 'status_ocr_success',
            'status_scan', 'status_error', 'cloud_upload'
        ]
        for k in keys:
            icon = get_ui_icon(k)
            pix = get_ui_pixmap(k, 24)
            self.assertFalse(icon.isNull(), f"Icon {k} không được null")
            self.assertFalse(pix.isNull(), f"Pixmap {k} không được null")


if __name__ == "__main__":
    unittest.main()
