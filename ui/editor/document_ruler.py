"""
ui/editor/document_ruler.py: Thước căn lề milimét (Horizontal Document Ruler)
- Hiển thị thang đo milimét chuẩn in ấn A4 (0 đến 210 mm)
- Đánh dấu vùng lề trái (30 mm) và lề phải (15 mm) bằng dải màu bóng mờ
- Các vạch chia 1mm, 5mm và đánh số centimet (1, 2, 3... 21 cm)
- Con trỏ thụt dòng đầu tiên (First Line Indent: 1.27 cm) và lề trái/phải tương tác
"""

from PyQt6.QtWidgets import QWidget
from PyQt6.QtGui import QPainter, QColor, QFont, QPen, QPolygonF
from PyQt6.QtCore import Qt, QPointF, QRectF, pyqtSignal


class HorizontalRuler(QWidget):
    """Thước đo ngang milimét chuẩn in A4 (210mm)"""
    margins_changed = pyqtSignal(float, float, float)  # left_mm, right_mm, first_line_mm

    def __init__(self, parent=None, page_width_mm: float = 210.0):
        super().__init__(parent)
        self.page_width_mm = page_width_mm
        self.left_margin_mm = 30.0
        self.right_margin_mm = 15.0
        self.first_line_indent_mm = 12.7
        self.zoom = 1.0
        self.setFixedHeight(28)
        self.setMouseTracking(True)

    def set_zoom(self, zoom: float):
        self.zoom = max(0.4, min(3.0, zoom))
        self.update()

    def set_margins(self, left_mm: float, right_mm: float, first_line_mm: float = 12.7):
        self.left_margin_mm = left_mm
        self.right_margin_mm = right_mm
        self.first_line_indent_mm = first_line_mm
        self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        w = self.width()
        h = self.height()

        # Nền xám nhạt toàn thước
        painter.fillRect(0, 0, w, h, QColor("#e2e8f0"))

        # Tính toán tỷ lệ pixel trên 1 mm
        px_per_mm = w / self.page_width_mm

        left_px = self.left_margin_mm * px_per_mm
        right_px = (self.page_width_mm - self.right_margin_mm) * px_per_mm

        # Vùng trắng thể hiện chiều rộng văn bản khả dụng (165mm)
        active_rect = QRectF(left_px, 0, right_px - left_px, h)
        painter.fillRect(active_rect, QColor("#ffffff"))

        # Đường viền thước
        painter.setPen(QColor("#cbd5e1"))
        painter.drawLine(0, h - 1, w, h - 1)

        # Vẽ vạch chia milimét
        pen_minor = QPen(QColor("#94a3b8"), 1)
        pen_major = QPen(QColor("#475569"), 1)
        font_num = QFont("Segoe UI", 7, QFont.Weight.Normal)
        painter.setFont(font_num)

        for mm in range(0, int(self.page_width_mm) + 1):
            x = mm * px_per_mm
            if mm % 10 == 0:
                # Vạch 10mm (1cm)
                painter.setPen(pen_major)
                painter.drawLine(int(x), h - 12, int(x), h - 1)
                cm_num = mm // 10
                if 0 < cm_num < int(self.page_width_mm // 10):
                    painter.setPen(QColor("#334155"))
                    painter.drawText(QRectF(x - 10, 2, 20, 12), Qt.AlignmentFlag.AlignCenter, str(cm_num))
            elif mm % 5 == 0:
                # Vạch 5mm
                painter.setPen(pen_minor)
                painter.drawLine(int(x), h - 8, int(x), h - 1)
            else:
                # Vạch 1mm
                if px_per_mm >= 3.0:  # Chỉ vẽ khi zoom đủ lớn
                    painter.setPen(pen_minor)
                    painter.drawLine(int(x), h - 4, int(x), h - 1)

        # Vẽ con trỏ định vị Lề trái (Left Margin Pointer - Hình tam giác chỉ lên)
        painter.setBrush(QColor("#2563eb"))
        painter.setPen(QColor("#1d4ed8"))
        poly_left = QPolygonF([
            QPointF(left_px - 4, h - 1),
            QPointF(left_px + 4, h - 1),
            QPointF(left_px, h - 9)
        ])
        painter.drawPolygon(poly_left)

        # Vẽ con trỏ thụt dòng đầu tiên (First Line Indent - Tam giác chỉ xuống ở 1.27cm từ lề trái)
        indent_px = (self.left_margin_mm + self.first_line_indent_mm) * px_per_mm
        painter.setBrush(QColor("#7c3aed"))
        painter.setPen(QColor("#6d28d9"))
        poly_indent = QPolygonF([
            QPointF(indent_px - 4, 1),
            QPointF(indent_px + 4, 1),
            QPointF(indent_px, 9)
        ])
        painter.drawPolygon(poly_indent)

        # Vẽ con trỏ định vị Lề phải (Right Margin Pointer - Hình tam giác chỉ lên)
        painter.setBrush(QColor("#2563eb"))
        painter.setPen(QColor("#1d4ed8"))
        poly_right = QPolygonF([
            QPointF(right_px - 4, h - 1),
            QPointF(right_px + 4, h - 1),
            QPointF(right_px, h - 9)
        ])
        painter.drawPolygon(poly_right)
