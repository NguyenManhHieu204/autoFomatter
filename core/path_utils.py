"""
core/path_utils.py: Xử lý đường dẫn tệp tài nguyên khi chạy mã nguồn hoặc sau khi đóng gói PyInstaller
"""

import sys
import os


def get_resource_path(relative_path: str) -> str:
    """
    Trả về đường dẫn tuyệt đối đến tệp tài nguyên:
    - Nếu chạy từ file .exe đóng gói bởi PyInstaller: sys._MEIPASS
    - Nếu chạy trực tiếp từ mã nguồn Python: thư mục gốc dự án
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base_path = sys._MEIPASS
    else:
        # Thư mục gốc AutoFormatter
        base_path = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

    return os.path.join(base_path, relative_path)


def get_config_dir() -> str:
    return get_resource_path("config")


def get_templates_dir() -> str:
    return get_resource_path("templates")


def get_assets_dir() -> str:
    return get_resource_path("assets")


def get_app_icon_path() -> str:
    ico = get_resource_path(os.path.join("assets", "app_icon.ico"))
    if os.path.exists(ico):
        return ico
    png = get_resource_path(os.path.join("assets", "app_icon.png"))
    return png if os.path.exists(png) else ""
