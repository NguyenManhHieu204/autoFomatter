"""
ui/icon_factory.py: Nhà máy quản lý và cung cấp Icon cho toàn bộ giao diện ứng dụng.
- Tự động nạp vector SVG hoặc PNG chất lượng cao từ thư mục assets/icons
- Bộ nhớ đệm (Cache) trong RAM giúp truy xuất tức thì mà không phải đọc đĩa nhiều lần
- Đảm bảo hiển thị sắc nét trên mọi độ phân giải màn hình (High-DPI Scaling)
"""

import os
from typing import Dict, Optional
from PyQt6.QtGui import QIcon, QPixmap
from PyQt6.QtCore import QSize, Qt

from core.path_utils import get_assets_dir

_ICON_CACHE: Dict[str, QIcon] = {}
_PIXMAP_CACHE: Dict[str, QPixmap] = {}


def get_ui_icon(name: str) -> QIcon:
    """
    Lấy QIcon chuẩn theo tên định danh:
    Ưu tiên tệp vector .svg để co giãn không vỡ nét, sau đó đến .png
    """
    if name in _ICON_CACHE:
        return _ICON_CACHE[name]

    assets_icons = os.path.join(get_assets_dir(), "icons")
    svg_path = os.path.join(assets_icons, f"{name}.svg")
    png_path = os.path.join(assets_icons, f"{name}.png")

    icon = QIcon()
    if os.path.exists(svg_path):
        icon = QIcon(svg_path)
    elif os.path.exists(png_path):
        icon = QIcon(png_path)

    _ICON_CACHE[name] = icon
    return icon


def get_ui_pixmap(name: str, size: int = 24) -> QPixmap:
    """
    Lấy QPixmap sắc nét theo kích thước chỉ định (sử dụng SmoothTransformation)
    """
    cache_key = f"{name}_{size}"
    if cache_key in _PIXMAP_CACHE:
        return _PIXMAP_CACHE[cache_key]

    assets_icons = os.path.join(get_assets_dir(), "icons")
    svg_path = os.path.join(assets_icons, f"{name}.svg")
    png_path = os.path.join(assets_icons, f"{name}.png")

    pixmap = QPixmap()
    if os.path.exists(svg_path):
        # Nạp trực tiếp từ vector SVG với kích thước chuẩn
        pixmap = QIcon(svg_path).pixmap(size, size)
    elif os.path.exists(png_path):
        pixmap = QPixmap(png_path).scaled(
            size, size,
            Qt.AspectRatioMode.KeepAspectRatio,
            Qt.TransformationMode.SmoothTransformation
        )

    _PIXMAP_CACHE[cache_key] = pixmap
    return pixmap
