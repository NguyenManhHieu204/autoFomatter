"""
ui/signature_dialog.py: Hộp thoại Ký số Điện tử Văn bản chuẩn Ban Cơ yếu Chính phủ & NĐ 30/2020/NĐ-CP
Tích hợp bộ xem trước PDF tương tác trực quan:
- Cho phép xem từng trang tài liệu PDF
- Bấm chuột hoặc kéo thả trực tiếp để chọn trang và vị trí đặt chữ ký số / con dấu
- Tự động nhận diện (Smart Locator) khối ký (T/M BAN THƯỜNG VỤ, BÍ THƯ, CHỦ TỊCH, GIÁM ĐỐC...)
  để tự động nhảy đến đúng trang và vị trí chuẩn hóa theo Nghị định 30/2020/NĐ-CP
"""

import os
import datetime
from typing import Optional, List, Dict, Any

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton,
    QComboBox, QRadioButton, QButtonGroup, QFileDialog, QGroupBox,
    QFrame, QMessageBox, QLineEdit, QSizePolicy, QScrollArea,
    QSpinBox, QSlider, QSplitter, QWidget, QApplication
)
from PyQt6.QtCore import Qt, QSize, QPoint, QRect, QRectF, pyqtSignal
from PyQt6.QtGui import (
    QIcon, QColor, QFont, QPixmap, QPainter, QPen, QBrush, QImage, QCursor
)
from PyQt6.QtPdf import QPdfDocument

from core.signature_engine import (
    WindowsCertStoreHelper, VGCABridge, PdfDigitalSigner,
    PdfSignatureLocator, CertificateInfo
)
from ui.icon_factory import get_ui_icon, get_ui_pixmap


class PdfInteractiveCanvas(QWidget):
    """
    Vùng vẽ tương tác hiển thị trang PDF và cho phép kéo thả / nhấp chuột
    để đặt ô chữ ký số / con dấu theo thời gian thực.
    """
    position_changed = pyqtSignal(int, float, float)  # page, norm_x, norm_y

    def __init__(self, parent=None):
        super().__init__(parent)
        self.doc = QPdfDocument(self)
        self.current_page = 0
        self.page_count = 0
        self.zoom = 1.0

        # Tọa độ chữ ký được chuẩn hóa (0.0 đến 1.0)
        self.target_page = 0
        self.norm_x = 0.55
        self.norm_y = 0.72
        self.stamp_w_mm = 65.0
        self.stamp_h_mm = 32.0

        # Thông tin chữ ký hiển thị trực quan
        self.sign_type = "personal"
        self.signer_name = ""
        self.signer_title = ""
        self.agency_name = ""
        self.custom_image_path: Optional[str] = None

        # Trạng thái kéo chuột
        self.is_dragging = False
        self.drag_offset_x = 0
        self.drag_offset_y = 0

        # Bộ đệm ảnh hiển thị để tối ưu hiệu năng vẽ
        self._cached_page_image: Optional[QImage] = None
        self._cached_render_page = -1
        self._cached_render_size = QSize()

        self.setMouseTracking(True)
        self.setSizePolicy(QSizePolicy.Policy.Fixed, QSizePolicy.Policy.Fixed)

    def load_pdf(self, pdf_path: str) -> bool:
        """Nạp tệp PDF vào bộ hiển thị"""
        if not os.path.exists(pdf_path):
            return False
        self.doc.load(pdf_path)
        self.page_count = self.doc.pageCount()
        self._invalidate_cache()
        self._update_widget_size()
        self.update()
        return self.page_count > 0

    def set_current_page(self, page_idx: int):
        """Chuyển đến trang xem"""
        if self.page_count > 0:
            self.current_page = max(0, min(self.page_count - 1, page_idx))
            self._invalidate_cache()
            self._update_widget_size()
            self.update()

    def set_zoom(self, zoom: float):
        """Thay đổi tỷ lệ phóng thu"""
        self.zoom = max(0.4, min(2.5, zoom))
        self._invalidate_cache()
        self._update_widget_size()
        self.update()

    def set_stamp_position(self, page_idx: int, norm_x: float, norm_y: float):
        """Cập nhật tọa độ ô ký"""
        self.target_page = max(0, min(max(0, self.page_count - 1), page_idx))
        box_w_ratio = self.stamp_w_mm / 210.0
        box_h_ratio = self.stamp_h_mm / 297.0
        self.norm_x = max(0.01, min(0.99 - box_w_ratio, norm_x))
        self.norm_y = max(0.01, min(0.99 - box_h_ratio, norm_y))
        self.update()

    def set_stamp_size(self, w_mm: float, h_mm: float):
        """Cập nhật kích thước ô ký theo mm"""
        self.stamp_w_mm = max(30.0, min(120.0, w_mm))
        self.stamp_h_mm = max(15.0, min(80.0, h_mm))
        self.update()

    def set_signer_preview_data(
        self, sign_type: str, name: str, title: str, agency: str, img_path: Optional[str]
    ):
        """Cập nhật dữ liệu người ký để hiển thị trực quan"""
        self.sign_type = sign_type
        self.signer_name = name
        self.signer_title = title
        self.agency_name = agency
        self.custom_image_path = img_path
        self.update()

    def _invalidate_cache(self):
        self._cached_page_image = None
        self._cached_render_page = -1

    def _get_page_point_size(self) -> QSize:
        if self.page_count > 0:
            pt = self.doc.pagePointSize(self.current_page)
            return QSize(int(pt.width()), int(pt.height()))
        return QSize(595, 842)

    def _update_widget_size(self):
        pt = self._get_page_point_size()
        disp_w = int(pt.width() * self.zoom)
        disp_h = int(pt.height() * self.zoom)
        self.setFixedSize(disp_w, disp_h)

    def _get_box_rect(self) -> QRect:
        pt = self._get_page_point_size()
        disp_w = int(pt.width() * self.zoom)
        disp_h = int(pt.height() * self.zoom)

        # Tính kích thước ô ký theo tỷ lệ khổ giấy A4
        box_w = max(40, int((self.stamp_w_mm / 210.0) * disp_w))
        box_h = max(25, int((self.stamp_h_mm / 297.0) * disp_h))
        box_x = int(self.norm_x * disp_w)
        box_y = int(self.norm_y * disp_h)
        return QRect(box_x, box_y, box_w, box_h)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform)

        if self.page_count == 0:
            painter.fillRect(self.rect(), QColor("#f1f5f9"))
            painter.setPen(QColor("#64748b"))
            painter.drawText(self.rect(), Qt.AlignmentFlag.AlignCenter, "Chưa nạp tệp PDF")
            return

        pt = self._get_page_point_size()
        disp_w = int(pt.width() * self.zoom)
        disp_h = int(pt.height() * self.zoom)
        target_size = QSize(disp_w, disp_h)

        # 1. Vẽ trang tài liệu PDF
        if (
            self._cached_page_image is None
            or self._cached_render_page != self.current_page
            or self._cached_render_size != target_size
        ):
            # Render ở độ phân giải hiển thị
            self._cached_page_image = self.doc.render(self.current_page, target_size)
            self._cached_render_page = self.current_page
            self._cached_render_size = target_size

        if self._cached_page_image and not self._cached_page_image.isNull():
            painter.drawImage(0, 0, self._cached_page_image)
        else:
            painter.fillRect(self.rect(), Qt.GlobalColor.white)

        # 2. Vẽ viền trang giấy thanh lịch
        painter.setPen(QColor(203, 213, 225))
        painter.drawRect(0, 0, disp_w - 1, disp_h - 1)

        # 3. Nếu đang ở đúng trang ký -> Vẽ khung chữ ký số tương tác
        if self.current_page == self.target_page:
            box_rect = self._get_box_rect()
            is_personal = (self.sign_type == "personal")
            is_combined = (self.sign_type == "combined")

            if is_personal:
                primary_color = QColor(124, 58, 237)
                bg_color = QColor(237, 233, 254, 180)
                badge_text = "Chữ ký cá nhân (NĐ 30)"
            elif is_combined:
                primary_color = QColor(16, 185, 129)
                bg_color = QColor(209, 250, 229, 180)
                badge_text = "Ký kết hợp: Dấu trùm 1/3 chữ ký (Điều 13 NĐ 30)"
            else:
                primary_color = QColor(220, 38, 38)
                bg_color = QColor(254, 226, 226, 180)
                badge_text = "Con dấu cơ quan (NĐ 30)"

            # Nền mờ cho ô chữ ký
            painter.fillRect(box_rect, bg_color)

            # Viền nét đứt sang trọng
            dashed_pen = QPen(primary_color, 2, Qt.PenStyle.DashLine)
            painter.setPen(dashed_pen)
            painter.drawRect(box_rect)

            # Vẽ 4 góc định vị (Corner handles)
            painter.setBrush(primary_color)
            handle_size = 6
            for corner in [
                box_rect.topLeft(),
                box_rect.topRight() - QPoint(handle_size, 0),
                box_rect.bottomLeft() - QPoint(0, handle_size),
                box_rect.bottomRight() - QPoint(handle_size, handle_size)
            ]:
                painter.drawRect(QRect(corner.x(), corner.y(), handle_size, handle_size))

            # Huy hiệu định vị trên đầu ô ký
            badge_rect = QRect(box_rect.x(), max(0, box_rect.y() - 20), box_rect.width(), 19)
            painter.fillRect(badge_rect, primary_color)
            painter.setPen(Qt.GlobalColor.white)
            badge_font = QFont("Segoe UI", 8, QFont.Weight.Bold)
            painter.setFont(badge_font)
            painter.drawText(badge_rect, Qt.AlignmentFlag.AlignCenter, badge_text)

            # Vẽ nội dung mô phỏng bên trong ô ký
            if self.custom_image_path and os.path.exists(self.custom_image_path):
                img = QPixmap(self.custom_image_path)
                scaled_img = img.scaled(
                    box_rect.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation
                )
                img_x = box_rect.x() + (box_rect.width() - scaled_img.width()) // 2
                img_y = box_rect.y() + (box_rect.height() - scaled_img.height()) // 2
                painter.drawPixmap(img_x, img_y, scaled_img)
            else:
                if is_personal:
                    # Chữ ký tay điện tử màu xanh
                    painter.setPen(QColor(30, 64, 175))
                    name_font = QFont("Segoe Script", max(9, int(11 * self.zoom)), QFont.Weight.Bold)
                    painter.setFont(name_font)
                    name_str = self.signer_name or "Chữ ký điện tử"
                    name_rect = QRect(box_rect.x(), box_rect.y() + int(box_rect.height() * 0.15), box_rect.width(), int(box_rect.height() * 0.45))
                    painter.drawText(name_rect, Qt.AlignmentFlag.AlignCenter, name_str)

                    # Dòng thời gian ký
                    painter.setPen(QColor(30, 64, 175))
                    info_font = QFont("Times New Roman", max(7, int(8 * self.zoom)), QFont.Weight.Normal)
                    info_font.setItalic(True)
                    painter.setFont(info_font)
                    now_str = datetime.datetime.now().strftime("%d/%m/%Y %H:%M")
                    info_rect = QRect(box_rect.x(), box_rect.y() + int(box_rect.height() * 0.60), box_rect.width(), int(box_rect.height() * 0.35))
                    painter.drawText(info_rect, Qt.AlignmentFlag.AlignCenter, f"Ký điện tử: {now_str}")
                else:
                    # Con dấu tròn đỏ
                    painter.setPen(QColor(220, 38, 38))
                    d = min(box_rect.width(), box_rect.height()) - 6
                    cx = box_rect.x() + (box_rect.width() - d) // 2
                    cy = box_rect.y() + (box_rect.height() - d) // 2
                    painter.setBrush(Qt.BrushStyle.NoBrush)
                    painter.drawEllipse(cx, cy, d, d)

                    seal_font = QFont("Times New Roman", max(7, int(8 * self.zoom)), QFont.Weight.Bold)
                    painter.setFont(seal_font)
                    org_str = (self.agency_name or "CƠ QUAN").upper()
                    painter.drawText(QRect(cx, cy + int(d * 0.35), d, int(d * 0.3)), Qt.AlignmentFlag.AlignCenter, org_str)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton and self.page_count > 0:
            box_rect = self._get_box_rect()
            # Nếu người dùng bấm vào trong ô ký -> Bắt đầu kéo
            if self.current_page == self.target_page and box_rect.contains(event.pos()):
                self.is_dragging = True
                self.drag_offset_x = event.pos().x() - box_rect.x()
                self.drag_offset_y = event.pos().y() - box_rect.y()
                self.setCursor(Qt.CursorShape.SizeAllCursor)
            else:
                # Bấm ra vị trí khác trên trang: Chuyển ngay tâm ô ký về vị trí bấm!
                self.target_page = self.current_page
                pt = self._get_page_point_size()
                disp_w = int(pt.width() * self.zoom)
                disp_h = int(pt.height() * self.zoom)
                box_w = max(40, int((self.stamp_w_mm / 210.0) * disp_w))
                box_h = max(25, int((self.stamp_h_mm / 297.0) * disp_h))

                new_box_x = event.pos().x() - box_w // 2
                new_box_y = event.pos().y() - box_h // 2
                self.norm_x = max(0.01, min(0.99 - (self.stamp_w_mm / 210.0), new_x := new_box_x / disp_w))
                self.norm_y = max(0.01, min(0.99 - (self.stamp_h_mm / 297.0), new_y := new_box_y / disp_h))

                self.is_dragging = True
                self.drag_offset_x = box_w // 2
                self.drag_offset_y = box_h // 2
                self.position_changed.emit(self.target_page, self.norm_x, self.norm_y)
                self.update()

    def mouseMoveEvent(self, event):
        if self.is_dragging and self.page_count > 0:
            pt = self._get_page_point_size()
            disp_w = int(pt.width() * self.zoom)
            disp_h = int(pt.height() * self.zoom)
            box_w = max(40, int((self.stamp_w_mm / 210.0) * disp_w))
            box_h = max(25, int((self.stamp_h_mm / 297.0) * disp_h))

            new_x = event.pos().x() - self.drag_offset_x
            new_y = event.pos().y() - self.drag_offset_y

            self.norm_x = max(0.01, min(0.99 - (self.stamp_w_mm / 210.0), new_x / disp_w))
            self.norm_y = max(0.01, min(0.99 - (self.stamp_h_mm / 297.0), new_y / disp_h))
            self.position_changed.emit(self.target_page, self.norm_x, self.norm_y)
            self.update()
        else:
            if self.page_count > 0 and self.current_page == self.target_page:
                box_rect = self._get_box_rect()
                if box_rect.contains(event.pos()):
                    self.setCursor(Qt.CursorShape.SizeAllCursor)
                else:
                    self.setCursor(Qt.CursorShape.CrossCursor)
            else:
                self.setCursor(Qt.CursorShape.CrossCursor)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.is_dragging = False
            self.setCursor(Qt.CursorShape.ArrowCursor)


class SignatureDialog(QDialog):
    """
    Hộp thoại Ký số PDF trực quan toàn diện theo Nghị định 30/2020/NĐ-CP
    và chuẩn mã hóa Ban Cơ yếu Chính phủ.
    """
    def __init__(self, parent=None, pdf_path: str = "", default_agency: str = ""):
        super().__init__(parent)
        self.pdf_path = pdf_path
        self.default_agency = default_agency
        self.signed_pdf_path: Optional[str] = None
        self.certs: List[CertificateInfo] = []

        self.setWindowTitle("Ký số Điện tử Văn bản — Ban Cơ yếu Chính phủ & Nghị định 30/2020")
        self.resize(1140, 750)
        self.setMinimumSize(980, 640)
        self.setWindowFlags(self.windowFlags() & ~Qt.WindowType.WindowContextHelpButtonHint)

        self._init_ui()
        self._load_document()
        self._refresh_environment()

    def _init_ui(self):
        main_layout = QHBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(12)

        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ====================== CỘT TRÁI: BỘ XEM PDF VÀ ĐỊNH VỊ ======================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(8)

        # Thanh điều hướng trang & Thu phóng
        nav_toolbar = QFrame()
        nav_toolbar.setStyleSheet(
            "background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 6px; padding: 4px 8px;"
        )
        nav_layout = QHBoxLayout(nav_toolbar)
        nav_layout.setContentsMargins(4, 2, 4, 2)
        nav_layout.setSpacing(8)

        self.btn_prev_page = QPushButton("◀ Trang trước")
        self.btn_prev_page.setObjectName("btnSecondary")
        self.btn_prev_page.clicked.connect(self._prev_page)

        self.lbl_page_info = QLabel("Trang 1 / 1")
        self.lbl_page_info.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 13px;")

        self.btn_next_page = QPushButton("Trang sau ▶")
        self.btn_next_page.setObjectName("btnSecondary")
        self.btn_next_page.clicked.connect(self._next_page)

        self.btn_zoom_out = QPushButton("🔍 -")
        self.btn_zoom_out.setFixedWidth(40)
        self.btn_zoom_out.setObjectName("btnSecondary")
        self.btn_zoom_out.clicked.connect(self._zoom_out)

        self.lbl_zoom = QLabel("100%")
        self.lbl_zoom.setStyleSheet("color: #475569; font-size: 12px;")

        self.btn_zoom_in = QPushButton("🔍 +")
        self.btn_zoom_in.setFixedWidth(40)
        self.btn_zoom_in.setObjectName("btnSecondary")
        self.btn_zoom_in.clicked.connect(self._zoom_in)

        self.btn_auto_locate = QPushButton("🎯 Tự động tìm ô ký (NĐ 30)")
        self.btn_auto_locate.setStyleSheet(
            "background-color: #ecfdf5; color: #047857; font-weight: bold; border: 1px solid #a7f3d0; border-radius: 4px; padding: 4px 10px;"
        )
        self.btn_auto_locate.clicked.connect(self._auto_detect_signature_position)

        nav_layout.addWidget(self.btn_prev_page)
        nav_layout.addWidget(self.lbl_page_info)
        nav_layout.addWidget(self.btn_next_page)
        nav_layout.addSpacing(12)
        nav_layout.addWidget(self.btn_zoom_out)
        nav_layout.addWidget(self.lbl_zoom)
        nav_layout.addWidget(self.btn_zoom_in)
        nav_layout.addStretch(1)
        nav_layout.addWidget(self.btn_auto_locate)
        left_layout.addWidget(nav_toolbar)

        # Vùng cuộn chứa trang PDF tương tác
        self.scroll_area = QScrollArea()
        self.scroll_area.setObjectName("pdfViewer")
        self.scroll_area.setStyleSheet("background-color: #e2e8f0; border: 1px solid #cbd5e1; border-radius: 6px;")
        self.scroll_area.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.canvas = PdfInteractiveCanvas()
        self.canvas.position_changed.connect(self._on_canvas_position_changed)
        self.scroll_area.setWidget(self.canvas)
        self.scroll_area.setWidgetResizable(False)
        left_layout.addWidget(self.scroll_area, 1)

        # Chân trang thông tin vị trí ký
        self.lbl_coord_status = QLabel("Nhấp chuột vào tài liệu để chọn vị trí đóng chữ ký số / con dấu.")
        self.lbl_coord_status.setStyleSheet("color: #475569; font-size: 12px; font-weight: 500;")
        left_layout.addWidget(self.lbl_coord_status)

        # ====================== CỘT PHẢI: THIẾT LẬP KÝ SỐ ======================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(4, 0, 0, 0)
        right_layout.setSpacing(10)

        # 1. Trạng thái kết nối phần cứng / dịch vụ
        self.status_box = QFrame()
        self.status_box.setStyleSheet(
            "background-color: #f8fafc; border: 1px solid #cbd5e1; border-radius: 8px; padding: 6px 10px;"
        )
        status_layout = QHBoxLayout(self.status_box)
        status_layout.setContentsMargins(4, 4, 4, 4)

        self.lbl_status_icon = QLabel()
        self.lbl_status_icon.setPixmap(get_ui_pixmap("digital_sign", 20))
        self.lbl_status_text = QLabel("Đang kiểm tra thiết bị Token...")
        self.lbl_status_text.setStyleSheet("font-weight: 600; color: #1e293b; font-size: 12px;")

        self.btn_refresh = QPushButton("Quét lại")
        self.btn_refresh.setObjectName("btnSecondary")
        self.btn_refresh.setFixedWidth(75)
        self.btn_refresh.clicked.connect(self._refresh_environment)

        status_layout.addWidget(self.lbl_status_icon)
        status_layout.addWidget(self.lbl_status_text, 1)
        status_layout.addWidget(self.btn_refresh)
        right_layout.addWidget(self.status_box)

        # 2. Cơ chế ký số
        channel_group = QGroupBox("Cơ chế ký số")
        channel_layout = QHBoxLayout(channel_group)
        self.rb_native = QRadioButton("Ký trực tiếp bằng Token")
        self.rb_vgca = QRadioButton("Ký qua VGCA SignTool")
        self.rb_native.setChecked(True)
        self.rb_native.toggled.connect(self._on_channel_changed)

        channel_layout.addWidget(self.rb_native)
        channel_layout.addWidget(self.rb_vgca)
        right_layout.addWidget(channel_group)

        # 3. Lựa chọn Chứng thư số trong Token
        self.cert_group = QGroupBox("Chứng thư số trong USB Token")
        cert_layout = QVBoxLayout(self.cert_group)

        self.cb_certs = QComboBox()
        self.cb_certs.currentIndexChanged.connect(self._on_cert_selected)
        cert_layout.addWidget(self.cb_certs)

        self.lbl_cert_detail = QLabel("Chưa chọn chứng thư.")
        self.lbl_cert_detail.setStyleSheet("color: #64748b; font-size: 11px; line-height: 1.3;")
        self.lbl_cert_detail.setWordWrap(True)
        cert_layout.addWidget(self.lbl_cert_detail)
        right_layout.addWidget(self.cert_group)

        # 4. Thể thức ký theo Nghị định 30/2020/NĐ-CP
        type_group = QGroupBox("Thể thức và thẩm quyền ký (NĐ 30/2020)")
        type_layout = QVBoxLayout(type_group)

        row_type = QHBoxLayout()
        self.rb_personal = QRadioButton("Chữ ký cá nhân (Người ký)")
        self.rb_agency = QRadioButton("Con dấu số cơ quan")
        self.rb_combined = QRadioButton("Ký kết hợp (Dấu trùm 1/3 chữ ký)")
        self.rb_personal.setChecked(True)
        self.rb_personal.toggled.connect(self._on_sign_type_toggled)
        self.rb_agency.toggled.connect(self._on_sign_type_toggled)
        self.rb_combined.toggled.connect(self._on_sign_type_toggled)

        self.type_btn_group = QButtonGroup(self)
        self.type_btn_group.addButton(self.rb_personal, 0)
        self.type_btn_group.addButton(self.rb_agency, 1)
        self.type_btn_group.addButton(self.rb_combined, 2)

        row_type.addWidget(self.rb_personal)
        row_type.addWidget(self.rb_agency)
        row_type.addWidget(self.rb_combined)
        type_layout.addLayout(row_type)

        # Họ tên / Đơn vị
        row_name = QHBoxLayout()
        self.lbl_name_caption = QLabel("Họ tên người ký:")
        self.lbl_name_caption.setFixedWidth(105)
        self.txt_signer_name = QLineEdit()
        self.txt_signer_name.textChanged.connect(self._sync_canvas_preview)
        row_name.addWidget(self.lbl_name_caption)
        row_name.addWidget(self.txt_signer_name)
        type_layout.addLayout(row_name)

        # Chức vụ
        row_title = QHBoxLayout()
        self.lbl_title_caption = QLabel("Chức vụ:")
        self.lbl_title_caption.setFixedWidth(105)
        self.txt_signer_title = QLineEdit()
        self.txt_signer_title.textChanged.connect(self._sync_canvas_preview)
        row_title.addWidget(self.lbl_title_caption)
        row_title.addWidget(self.txt_signer_title)
        type_layout.addLayout(row_title)

        # Tùy chọn ảnh chữ ký / con dấu
        row_img = QHBoxLayout()
        self.lbl_img_caption = QLabel("Ảnh chữ ký/dấu:")
        self.lbl_img_caption.setFixedWidth(105)
        self.txt_img_path = QLineEdit()
        self.txt_img_path.setPlaceholderText("Tùy chọn: Chọn ảnh .png trong suốt")
        self.txt_img_path.textChanged.connect(self._sync_canvas_preview)
        self.btn_browse_img = QPushButton("Chọn...")
        self.btn_browse_img.setObjectName("btnSecondary")
        self.btn_browse_img.setFixedWidth(65)
        self.btn_browse_img.clicked.connect(self._browse_image)

        row_img.addWidget(self.lbl_img_caption)
        row_img.addWidget(self.txt_img_path)
        row_img.addWidget(self.btn_browse_img)
        type_layout.addLayout(row_img)

        right_layout.addWidget(type_group)

        # 5. Kích thước ô ký (Nghị định 30)
        dim_group = QGroupBox("Kích thước ô ký số")
        dim_layout = QHBoxLayout(dim_group)

        dim_layout.addWidget(QLabel("Rộng:"))
        self.sp_width = QSpinBox()
        self.sp_width.setRange(35, 110)
        self.sp_width.setValue(65)
        self.sp_width.setSuffix(" mm")
        self.sp_width.valueChanged.connect(self._on_stamp_size_changed)
        dim_layout.addWidget(self.sp_width)

        dim_layout.addWidget(QLabel("Cao:"))
        self.sp_height = QSpinBox()
        self.sp_height.setRange(20, 80)
        self.sp_height.setValue(32)
        self.sp_height.setSuffix(" mm")
        self.sp_height.valueChanged.connect(self._on_stamp_size_changed)
        dim_layout.addWidget(self.sp_height)

        right_layout.addWidget(dim_group)
        right_layout.addStretch(1)

        # 6. Các nút hành động
        btn_layout = QHBoxLayout()
        self.btn_sign = QPushButton("THỰC HIỆN KÝ SỐ VĂN BẢN")
        self.btn_sign.setIcon(get_ui_icon("digital_sign"))
        self.btn_sign.setIconSize(QSize(18, 18))
        self.btn_sign.setStyleSheet(
            "background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #7c3aed, stop:1 #6d28d9); "
            "color: white; font-weight: bold; font-size: 13px; min-height: 42px; border-radius: 6px;"
        )
        self.btn_sign.clicked.connect(self._execute_sign)

        self.btn_cancel = QPushButton("Đóng")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.setFixedWidth(80)
        self.btn_cancel.clicked.connect(self.reject)

        btn_layout.addWidget(self.btn_sign, 1)
        btn_layout.addWidget(self.btn_cancel)
        right_layout.addLayout(btn_layout)

        # Thiết lập splitter
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setStretchFactor(0, 65)
        splitter.setStretchFactor(1, 35)
        main_layout.addWidget(splitter)

    def _load_document(self):
        """Nạp tệp PDF vào Canvas và tự động định vị khối ký"""
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            self.lbl_coord_status.setText("⚠ Tệp PDF không tồn tại.")
            return

        ok = self.canvas.load_pdf(self.pdf_path)
        if not ok or self.canvas.page_count == 0:
            self.lbl_coord_status.setText("⚠ Không thể mở tệp PDF.")
            return

        self._update_page_nav_controls()
        self._auto_detect_signature_position()

    def _update_page_nav_controls(self):
        total = self.canvas.page_count
        cur = self.canvas.current_page + 1
        self.lbl_page_info.setText(f"Trang {cur} / {total}")
        self.btn_prev_page.setEnabled(self.canvas.current_page > 0)
        self.btn_next_page.setEnabled(self.canvas.current_page < total - 1)

    def _prev_page(self):
        if self.canvas.current_page > 0:
            new_p = self.canvas.current_page - 1
            self.canvas.set_current_page(new_p)
            self.canvas.set_stamp_position(new_p, self.canvas.norm_x, self.canvas.norm_y)
            self._update_page_nav_controls()
            self._on_canvas_position_changed(new_p, self.canvas.norm_x, self.canvas.norm_y)

    def _next_page(self):
        if self.canvas.current_page < self.canvas.page_count - 1:
            new_p = self.canvas.current_page + 1
            self.canvas.set_current_page(new_p)
            self.canvas.set_stamp_position(new_p, self.canvas.norm_x, self.canvas.norm_y)
            self._update_page_nav_controls()
            self._on_canvas_position_changed(new_p, self.canvas.norm_x, self.canvas.norm_y)

    def _zoom_in(self):
        new_zoom = min(2.0, self.canvas.zoom + 0.15)
        self.canvas.set_zoom(new_zoom)
        self.lbl_zoom.setText(f"{int(new_zoom * 100)}%")

    def _zoom_out(self):
        new_zoom = max(0.5, self.canvas.zoom - 0.15)
        self.canvas.set_zoom(new_zoom)
        self.lbl_zoom.setText(f"{int(new_zoom * 100)}%")

    def _auto_detect_signature_position(self):
        """Quét tự động khối ký số trong PDF và chuyển ngay đến trang chuẩn"""
        anchor = PdfSignatureLocator.find_signature_anchor(self.pdf_path)
        target_p = anchor.get("page_index", max(0, self.canvas.page_count - 1))
        norm_x = anchor.get("norm_x", 0.55)
        norm_y = anchor.get("norm_y", 0.72)

        # Chuyển canvas đến trang ký và cập nhật tọa độ
        self.canvas.set_current_page(target_p)
        self.canvas.set_stamp_position(target_p, norm_x, norm_y)
        self._update_page_nav_controls()

        # Tự động điền họ tên / chức danh nếu phát hiện được
        detected_signer = anchor.get("signer_name", "").strip()
        detected_auth = anchor.get("authority", "").strip()
        if detected_signer and not self.txt_signer_name.text():
            self.txt_signer_name.setText(detected_signer)
        if detected_auth and not self.txt_signer_title.text():
            self.txt_signer_title.setText(detected_auth)

        self._on_canvas_position_changed(target_p, norm_x, norm_y)

    def _on_canvas_position_changed(self, page_idx: int, norm_x: float, norm_y: float):
        msg = f"✔ Đang đặt chữ ký tại: Trang {page_idx + 1}/{self.canvas.page_count} (X={norm_x*100:.1f}%, Y={norm_y*100:.1f}%)"
        self.lbl_coord_status.setText(msg)
        self._update_page_nav_controls()

    def _on_stamp_size_changed(self):
        w = float(self.sp_width.value())
        h = float(self.sp_height.value())
        self.canvas.set_stamp_size(w, h)

    def _sync_canvas_preview(self):
        if self.rb_personal.isChecked():
            sign_type = "personal"
        elif self.rb_combined.isChecked():
            sign_type = "combined"
        else:
            sign_type = "agency"

        name = self.txt_signer_name.text().strip()
        title = self.txt_signer_title.text().strip()
        agency = name if sign_type == "agency" else (self.default_agency or "CƠ QUAN BAN HÀNH")
        img_path = self.txt_img_path.text().strip()
        if not os.path.exists(img_path):
            img_path = None
        self.canvas.set_signer_preview_data(sign_type, name, title, agency, img_path)

    def _refresh_environment(self):
        """Quét lại thiết bị Token và dịch vụ Ban Cơ yếu"""
        self.certs = WindowsCertStoreHelper.list_certificates()
        self.cb_certs.clear()

        vgca_ok, vgca_msg, port = VGCABridge.check_service_status()

        if self.certs:
            vgca_count = sum(1 for c in self.certs if c.is_vgca)
            for c in self.certs:
                self.cb_certs.addItem(c.display_name, c)

            if vgca_count > 0:
                self.status_box.setStyleSheet("background-color: #f0fdf4; border: 1px solid #86efac; border-radius: 8px; padding: 6px 10px;")
                self.lbl_status_text.setText(f"✔ Đã kết nối USB Token Ban Cơ yếu ({vgca_count} chứng thư)")
                self.lbl_status_text.setStyleSheet("font-weight: 600; color: #15803d; font-size: 12px;")
            else:
                self.status_box.setStyleSheet("background-color: #f0f9ff; border: 1px solid #7dd3fc; border-radius: 8px; padding: 6px 10px;")
                self.lbl_status_text.setText(f"✔ Đã nhận diện {len(self.certs)} chứng thư số khả dụng")
                self.lbl_status_text.setStyleSheet("font-weight: 600; color: #0369a1; font-size: 12px;")
            self.btn_sign.setEnabled(True)
        elif vgca_ok:
            self.status_box.setStyleSheet("background-color: #f5f3ff; border: 1px solid #c4b5fd; border-radius: 8px; padding: 6px 10px;")
            self.lbl_status_text.setText(f"✔ Dịch vụ Ban Cơ yếu ({vgca_msg})")
            self.lbl_status_text.setStyleSheet("font-weight: 600; color: #6d28d9; font-size: 12px;")
            self.rb_vgca.setChecked(True)
            self.btn_sign.setEnabled(True)
        else:
            self.status_box.setStyleSheet("background-color: #fffbeb; border: 1px solid #fde68a; border-radius: 8px; padding: 6px 10px;")
            self.lbl_status_text.setText("⚠ Chưa phát hiện USB Token. Hãy cắm Token vào máy.")
            self.lbl_status_text.setStyleSheet("font-weight: 600; color: #b45309; font-size: 12px;")
            self.lbl_cert_detail.setText("Không tìm thấy chứng thư số nào.")
            self.btn_sign.setEnabled(False)

        self._sync_canvas_preview()

    def _on_channel_changed(self):
        use_vgca = self.rb_vgca.isChecked()
        self.cert_group.setEnabled(not use_vgca)

    def _on_cert_selected(self, idx: int):
        if idx < 0 or idx >= len(self.certs):
            return
        cert = self.certs[idx]
        detail_text = (
            f"<b>Chủ thể:</b> {cert.subject}<br/>"
            f"<b>Tổ chức cấp:</b> {cert.issuer}"
        )
        self.lbl_cert_detail.setText(detail_text)

        if not self.txt_signer_name.text():
            self.txt_signer_name.setText(cert.common_name)
        if cert.title and not self.txt_signer_title.text():
            self.txt_signer_title.setText(cert.title)

        self._sync_canvas_preview()

    def _on_sign_type_toggled(self):
        if self.rb_personal.isChecked():
            self.lbl_name_caption.setText("Họ tên người ký:")
            self.lbl_title_caption.setText("Chức vụ:")
            self.lbl_img_caption.setText("Ảnh chữ ký tay:")
            self.sp_width.setValue(65)
            self.sp_height.setValue(32)
        elif self.rb_combined.isChecked():
            self.lbl_name_caption.setText("Họ tên người ký:")
            self.lbl_title_caption.setText("Chức vụ:")
            self.lbl_img_caption.setText("Ảnh dấu/chữ ký:")
            self.sp_width.setValue(75)
            self.sp_height.setValue(36)
        else:
            self.lbl_name_caption.setText("Tên cơ quan:")
            self.lbl_title_caption.setText("Thẩm quyền:")
            self.lbl_img_caption.setText("Ảnh con dấu đỏ:")
            self.sp_width.setValue(45)
            self.sp_height.setValue(45)
            if self.default_agency and not self.txt_signer_name.text():
                self.txt_signer_name.setText(self.default_agency)

        self._on_stamp_size_changed()
        self._sync_canvas_preview()

    def _browse_image(self):
        tip = "Chọn ảnh chữ ký tay (.png trong suốt)" if self.rb_personal.isChecked() else "Chọn ảnh con dấu đỏ (.png)"
        path, _ = QFileDialog.getOpenFileName(
            self, tip, "", "Hình ảnh (*.png *.jpg *.jpeg *.bmp)"
        )
        if path:
            self.txt_img_path.setText(path)
            self._sync_canvas_preview()

    def _execute_sign(self):
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            QMessageBox.warning(self, "Tệp không tồn tại", f"Không tìm thấy tệp PDF nguồn: {self.pdf_path}")
            return

        folder = os.path.dirname(os.path.abspath(self.pdf_path))
        base_name = os.path.basename(self.pdf_path)
        stem, ext = os.path.splitext(base_name)
        clean_stem = re.sub(r'(_ChuanHoa|_KySo|_signed)+$', '', stem, flags=re.IGNORECASE).strip()
        output_pdf = os.path.join(folder, f"{clean_stem}_ChuanHoa_KySo{ext}")

        # 1. Ký qua phần mềm Ban Cơ yếu (VGCA SignTool)
        if self.rb_vgca.isChecked():
            ok, msg = VGCABridge.open_with_vgca_tool(self.pdf_path)
            if ok:
                QMessageBox.information(
                    self, "Ký số Ban Cơ yếu",
                    "Đã kích hoạt tệp văn bản sang phần mềm Ban Cơ yếu.\n"
                    "Vui lòng hoàn tất việc nhập mã PIN và chọn mẫu con dấu trên cửa sổ Ban Cơ yếu."
                )
                self.signed_pdf_path = self.pdf_path
                self.accept()
            else:
                QMessageBox.warning(self, "Lỗi kết nối", msg)
            return

        # 2. Ký trực tiếp bằng Token (Native PAdES / CAPI)
        if not self.certs:
            QMessageBox.warning(self, "Thiếu Token", "Vui lòng cắm USB Token Ban Cơ yếu có chứa chứng thư số hợp lệ.")
            return

        selected_cert = self.cb_certs.currentData()
        if not selected_cert:
            selected_cert = self.certs[0]

        if self.rb_personal.isChecked():
            sign_type = "personal"
        elif self.rb_combined.isChecked():
            sign_type = "combined"
        else:
            sign_type = "agency"

        signer_name = self.txt_signer_name.text().strip()
        signer_title = self.txt_signer_title.text().strip()
        stamp_img = self.txt_img_path.text().strip() if os.path.exists(self.txt_img_path.text().strip()) else None

        self.btn_sign.setEnabled(False)
        self.btn_sign.setText("Đang gọi USB Token...")
        QApplication.processEvents()

        try:
            ok, msg = PdfDigitalSigner.sign_pdf_hybrid(
                input_pdf_path=self.pdf_path,
                output_pdf_path=output_pdf,
                cert_info=selected_cert,
                sign_type=sign_type,
                signer_name=signer_name,
                signer_title=signer_title,
                agency_name=signer_name if sign_type == "agency" else (self.default_agency or selected_cert.organization),
                stamp_image_path=stamp_img,
                target_page=self.canvas.target_page,
                norm_x=self.canvas.norm_x,
                norm_y=self.canvas.norm_y,
                stamp_width_mm=float(self.sp_width.value()),
                stamp_height_mm=float(self.sp_height.value())
            )

            if ok and os.path.exists(output_pdf):
                self.signed_pdf_path = output_pdf
                QMessageBox.information(
                    self, "Ký số Thành công",
                    f"Văn bản đã được ký số thành công theo đúng Nghị định 30/2020/NĐ-CP!\n\n"
                    f"Tệp kết quả: {os.path.basename(output_pdf)}\n"
                    f"Đã đóng dấu tại Trang {self.canvas.target_page + 1}."
                )
                self.accept()
            else:
                QMessageBox.warning(self, "Không thể ký số", msg)
        except Exception as e:
            QMessageBox.critical(self, "Lỗi ký số", f"Lỗi phát sinh: {str(e)}")
        finally:
            self.btn_sign.setEnabled(True)
            self.btn_sign.setText("THỰC HIỆN KÝ SỐ VĂN BẢN")
