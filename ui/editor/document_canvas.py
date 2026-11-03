"""
ui/editor/document_canvas.py: Khung vẽ Soạn thảo A4 Trực quan (WYSIWYG Document Canvas)
- Kết xuất phân trang thực tế (Discrete A4 Page Sheets) với đổ bóng chuẩn xác
- Không kéo dài trang vô tận: Phân tách rõ rệt từng trang A4 (Trang 1, Trang 2, Trang 3...)
- Hỗ trợ độc lập định hướng từng trang:
  + Trang 1: Luôn giữ khổ dọc (Portrait 210 x 297 mm) theo thể thức NĐ 30 / HD 36
  + Trang 2+: Tùy chọn độc lập xoay ngang (Landscape 297 x 210 mm) cho Phụ lục / Bảng biểu nhiều cột
              mà KHÔNG làm thay đổi khổ dọc của trang 1!
- Hiển thị ranh giới lề in NĐ 30: Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm
- Khối Thể thức Header (Quốc hiệu, Cơ quan) & Title ở trang đầu, Khối Footer (Nơi nhận, Ký tên) ở cuối văn bản
"""

import os
import re
from typing import Dict, Any, Optional, List
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTextEdit, QLabel,
    QFrame, QSizePolicy, QMessageBox, QFileDialog,
    QPushButton, QMenu
)
from PyQt6.QtGui import (
    QColor, QFont, QTextCursor, QTextCharFormat,
    QTextBlockFormat, QTextDocument, QPixmap, QTextTableFormat,
    QTextLength
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize

from ui.icon_factory import get_ui_icon


class PageSheetWidget(QWidget):
    """Một trang giấy A4 trực quan độc lập"""

    A4_WIDTH_PX = 794
    A4_HEIGHT_PX = 1123

    focus_received = pyqtSignal(int)
    orientation_changed = pyqtSignal(int, str)
    delete_requested = pyqtSignal(int)

    def __init__(self, page_index: int = 0, orientation: str = "portrait", parent=None):
        super().__init__(parent)
        self.page_index = page_index
        self.orientation = orientation
        self.zoom = 1.0

        self.setObjectName("page_sheet")
        self.setStyleSheet("""
            QWidget#page_sheet {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 2px;
            }
        """)

        self.sheet_layout = QVBoxLayout(self)
        self.sheet_layout.setSpacing(6)

        # Thanh điều khiển trang (Chỉ hiển thị từ trang 2 trở đi)
        self.top_page_bar = QWidget()
        bar_l = QHBoxLayout(self.top_page_bar)
        bar_l.setContentsMargins(0, 0, 0, 2)
        bar_l.setSpacing(8)

        self.lbl_page_title = QLabel(f"Trang {self.page_index + 1}")
        self.lbl_page_title.setStyleSheet("font-size: 11px; font-weight: 600; color: #475569; background: transparent;")
        bar_l.addWidget(self.lbl_page_title)
        bar_l.addStretch()

        self.btn_toggle_orient = QPushButton()
        self.btn_toggle_orient.setStyleSheet("""
            QPushButton {
                background-color: #f1f5f9; color: #1e293b; font-size: 11px;
                font-weight: 600; padding: 4px 10px; border: 1.5px solid #cbd5e1; border-radius: 4px;
            }
            QPushButton:hover { background-color: #eff6ff; border-color: #2563eb; color: #1d4ed8; }
        """)
        self.btn_toggle_orient.clicked.connect(self._on_toggle_orient_clicked)
        bar_l.addWidget(self.btn_toggle_orient)

        self.btn_delete_page = QPushButton("Xóa trang")
        self.btn_delete_page.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #dc2626; font-size: 11px;
                font-weight: 500; padding: 4px 8px; border: 1px solid #fecaca; border-radius: 4px;
            }
            QPushButton:hover { background-color: #fef2f2; border-color: #dc2626; }
        """)
        self.btn_delete_page.clicked.connect(lambda: self.delete_requested.emit(self.page_index))
        bar_l.addWidget(self.btn_delete_page)

        if self.page_index == 0:
            self.top_page_bar.setVisible(False)
        self.sheet_layout.addWidget(self.top_page_bar)

        # Số trang ở giữa lề trên (chỉ hiện từ trang 2 theo Nghị định 30)
        self.lbl_page_number = QLabel(f"- {self.page_index + 1} -")
        self.lbl_page_number.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_page_number.setStyleSheet("font-family: 'Times New Roman'; font-size: 13.5pt; color: #000000; background: transparent; margin-bottom: 2px;")
        if self.page_index == 0:
            self.lbl_page_number.setVisible(False)
        self.sheet_layout.addWidget(self.lbl_page_number)

        # Editor nội dung cho trang này
        self.editor = QTextEdit()
        self.editor.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: transparent;
                font-family: 'Times New Roman';
                font-size: 13pt;
                color: #0f172a;
                line-height: 1.25;
            }
        """)
        self.editor.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.editor.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.sheet_layout.addWidget(self.editor, 1)

        # Cài đặt kích thước ban đầu
        self.update_geometry()

    def _on_toggle_orient_clicked(self):
        new_orient = "portrait" if self.orientation == "landscape" else "landscape"
        self.orientation = new_orient
        self.update_geometry()
        self.orientation_changed.emit(self.page_index, self.orientation)

    def update_geometry(self):
        is_land = (self.orientation == "landscape")
        base_w = self.A4_HEIGHT_PX if is_land else self.A4_WIDTH_PX
        base_h = self.A4_WIDTH_PX if is_land else self.A4_HEIGHT_PX
        w = int(base_w * self.zoom)
        h = int(base_h * self.zoom)

        self.setFixedWidth(w)
        self.setMinimumHeight(h)

        # Lề NĐ 30: Trái 30mm (~113px), Phải 15mm (~57px), Trên 20mm (~76px), Dưới 20mm (~76px)
        top_m = 50 if self.page_index == 0 else 24
        self.sheet_layout.setContentsMargins(113, top_m, 57, 50)

        if is_land:
            self.btn_toggle_orient.setText("🔄 Đổi sang: Trang dọc (210×297 mm)")
            self.lbl_page_title.setText(f"Trang {self.page_index + 1} (Khổ ngang - Phụ lục / Bảng biểu)")
        else:
            self.btn_toggle_orient.setText("🔄 Đổi sang: Trang ngang (297×210 mm)")
            self.lbl_page_title.setText(f"Trang {self.page_index + 1} (Khổ dọc chuẩn)")

        self.lbl_page_number.setText(f"- {self.page_index + 1} -")


class DocumentCanvas(QWidget):
    """Bề mặt hiển thị các trang giấy A4 phân vùng soạn thảo trực quan"""
    document_modified = pyqtSignal()
    cursor_position_changed = pyqtSignal(int, int)  # line, col

    A4_WIDTH_PX = 794
    A4_HEIGHT_PX = 1123

    def __init__(self, parent=None):
        super().__init__(parent)
        self.zoom = 1.0
        self.lock_scaffolding = True
        self.orientation = "portrait"
        self._active_page_idx = 0

        self.metadata: Dict[str, Any] = {
            "standard_id": "ND30",
            "superior_agency": "ỦY BAN NHÂN DÂN HUYỆN",
            "issuing_agency": "ỦY BAN NHÂN DÂN XÃ",
            "doc_number": "Số: .../QĐ-UBND",
            "location": "Xã",
            "date_str": "..., ngày ... tháng ... năm 2026",
            "doc_type": "QUYẾT ĐỊNH",
            "doc_subject": "Về việc ban hành Kế hoạch phát triển kinh tế - xã hội năm 2026",
            "recipients": [
                "- Thường trực HĐND xã;",
                "- Chủ tịch, các PCT UBND xã;",
                "- Các ban ngành, đoàn thể xã;",
                "- Các thôn, tổ dân phố;",
                "- Lưu: VT."
            ],
            "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
            "signer_position": "CHỦ TỊCH",
            "signer_name": "Nguyễn Văn A"
        }

        # Bố cục chính căn giữa màn hình
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.main_layout.setContentsMargins(0, 16, 0, 40)
        self.main_layout.setSpacing(14)

        # 1. Thanh tiện ích phía trên trang giấy (Căn giữa)
        self.top_ctrl_widget = QWidget()
        self.top_ctrl_widget.setFixedWidth(self.A4_WIDTH_PX)
        top_ctrl_bar = QHBoxLayout(self.top_ctrl_widget)
        top_ctrl_bar.setContentsMargins(0, 0, 0, 0)
        top_ctrl_bar.setSpacing(8)

        lbl_page_info = QLabel("Khổ giấy: A4 (210 × 297 mm) • Chuẩn Thể thức Cấp Xã")
        lbl_page_info.setStyleSheet("color: #64748b; font-size: 11.5px; font-weight: 500;")
        top_ctrl_bar.addWidget(lbl_page_info)
        top_ctrl_bar.addStretch()

        self.btn_add_page_top = QPushButton("➕ Thêm trang mới (Ngắt trang)")
        self.btn_add_page_top.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #2563eb; font-weight: 600;
                border: 1.5px solid #bfdbfe; border-radius: 4px; padding: 4px 10px; font-size: 11.5px;
            }
            QPushButton:hover { background-color: #eff6ff; border-color: #2563eb; }
        """)
        self.btn_add_page_top.clicked.connect(lambda: self.add_page(orientation="portrait"))
        top_ctrl_bar.addWidget(self.btn_add_page_top)

        self.btn_add_landscape_top = QPushButton("📊 Thêm trang ngang (Landscape)")
        self.btn_add_landscape_top.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #0284c7; font-weight: 600;
                border: 1.5px solid #bae6fd; border-radius: 4px; padding: 4px 10px; font-size: 11.5px;
            }
            QPushButton:hover { background-color: #f0f9ff; border-color: #0284c7; }
        """)
        self.btn_add_landscape_top.clicked.connect(lambda: self.add_page(orientation="landscape"))
        top_ctrl_bar.addWidget(self.btn_add_landscape_top)

        self.btn_edit_meta = QPushButton("Chỉnh sửa thể thức (F2)")
        self.btn_edit_meta.setIcon(get_ui_icon("edit"))
        self.btn_edit_meta.setStyleSheet("""
            QPushButton {
                background-color: #ffffff; color: #1e293b; font-weight: 600;
                border: 1.5px solid #cbd5e1; border-radius: 4px;
                padding: 4px 12px; font-size: 11.5px;
            }
            QPushButton:hover { background-color: #f1f5f9; border-color: #94a3b8; }
        """)
        self.btn_edit_meta.clicked.connect(self.open_metadata_editor)
        top_ctrl_bar.addWidget(self.btn_edit_meta)

        self.main_layout.addWidget(self.top_ctrl_widget, alignment=Qt.AlignmentFlag.AlignHCenter)

        # 2. Khung chứa các trang giấy A4 rời rạc (Pages Layout)
        self.pages_container = QWidget()
        self.pages_layout = QVBoxLayout(self.pages_container)
        self.pages_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.pages_layout.setContentsMargins(0, 0, 0, 0)
        self.pages_layout.setSpacing(24)  # Khoảng cách 24px giữa các trang giấy A4 chân thực
        self.main_layout.addWidget(self.pages_container, alignment=Qt.AlignmentFlag.AlignHCenter)

        # Danh sách các trang A4
        self.pages: List[PageSheetWidget] = []

        # Khởi tạo khung Header và Title dùng cho Trang 1
        self._init_header_and_title_frames()

        # Khởi tạo khung Footer (Nơi nhận & Chữ ký)
        self._init_footer_frame()

        # Tạo trang 1 ban đầu
        self.add_page(orientation="portrait")

        # Nạp dữ liệu mẫu ban đầu
        self._refresh_scaffolding_display()
        self._load_sample_body()

    # ==================== QUẢN LÝ CÁC KHUNG THỂ THỨC PHÁP LÝ ====================
    def _init_header_and_title_frames(self):
        frame_hover_style = """
            QFrame {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 4px;
                padding: 2px;
            }
            QFrame:hover {
                background-color: #f8fafc;
                border: 1px dashed #93c5fd;
            }
        """

        # 1. Khung Header (Quốc hiệu & Cơ quan)
        self.header_frame = QFrame()
        self.header_frame.setCursor(Qt.CursorShape.PointingHandCursor)
        self.header_frame.setToolTip("Nhấp đúp chuột để chỉnh sửa Cơ quan, Quốc hiệu, Số hiệu, Địa danh (Phím F2)")
        self.header_frame.mouseDoubleClickEvent = lambda e: self.open_metadata_editor()
        self.header_frame.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.header_frame.customContextMenuRequested.connect(lambda p: self.open_metadata_editor())
        self.header_frame.setStyleSheet(frame_hover_style)
        h_layout = QHBoxLayout(self.header_frame)
        h_layout.setContentsMargins(4, 4, 4, 4)

        self.lbl_header_left = QLabel()
        self.lbl_header_left.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_header_left.setStyleSheet("color: #0f172a; font-family: 'Times New Roman'; background: transparent;")
        h_layout.addWidget(self.lbl_header_left, 1)

        self.lbl_header_right = QLabel()
        self.lbl_header_right.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_header_right.setStyleSheet("color: #0f172a; font-family: 'Times New Roman'; background: transparent;")
        h_layout.addWidget(self.lbl_header_right, 1)

        # 2. Khung Tiêu đề & Trích yếu
        self.title_frame = QFrame()
        self.title_frame.setCursor(Qt.CursorShape.PointingHandCursor)
        self.title_frame.setToolTip("Nhấp đúp chuột để chỉnh sửa Loại văn bản và Trích yếu (Phím F2)")
        self.title_frame.mouseDoubleClickEvent = lambda e: self.open_metadata_editor()
        self.title_frame.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.title_frame.customContextMenuRequested.connect(lambda p: self.open_metadata_editor())
        self.title_frame.setStyleSheet(frame_hover_style)
        t_layout = QVBoxLayout(self.title_frame)
        t_layout.setContentsMargins(0, 6, 0, 6)
        t_layout.setSpacing(2)

        self.lbl_doc_type = QLabel("QUYẾT ĐỊNH")
        self.lbl_doc_type.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_doc_type.setStyleSheet("font-family: 'Times New Roman'; font-size: 15pt; font-weight: bold; color: #0f172a; background: transparent;")
        t_layout.addWidget(self.lbl_doc_type)

        self.lbl_doc_subject = QLabel("Về việc phê duyệt kế hoạch chuyển đổi số năm 2026")
        self.lbl_doc_subject.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_doc_subject.setStyleSheet("font-family: 'Times New Roman'; font-size: 13.5pt; font-weight: bold; color: #1e293b; background: transparent;")
        self.lbl_doc_subject.setWordWrap(True)
        t_layout.addWidget(self.lbl_doc_subject)

    def _init_footer_frame(self):
        frame_hover_style = """
            QFrame {
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 4px;
                padding: 2px;
            }
            QFrame:hover {
                background-color: #f8fafc;
                border: 1px dashed #93c5fd;
            }
        """
        self.footer_frame = QFrame()
        self.footer_frame.setCursor(Qt.CursorShape.PointingHandCursor)
        self.footer_frame.setToolTip("Nhấp đúp chuột để chỉnh sửa Nơi nhận và Khối ký ban hành (Phím F2)")
        self.footer_frame.mouseDoubleClickEvent = lambda e: self.open_metadata_editor()
        self.footer_frame.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.footer_frame.customContextMenuRequested.connect(lambda p: self.open_metadata_editor())
        self.footer_frame.setStyleSheet(frame_hover_style)
        f_layout = QHBoxLayout(self.footer_frame)
        f_layout.setContentsMargins(4, 4, 4, 4)

        self.lbl_footer_left = QLabel()
        self.lbl_footer_left.setAlignment(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop)
        self.lbl_footer_left.setStyleSheet("color: #0f172a; font-family: 'Times New Roman'; background: transparent;")
        f_layout.addWidget(self.lbl_footer_left, 1)

        self.lbl_footer_right = QLabel()
        self.lbl_footer_right.setAlignment(Qt.AlignmentFlag.AlignCenter | Qt.AlignmentFlag.AlignTop)
        self.lbl_footer_right.setStyleSheet("color: #0f172a; font-family: 'Times New Roman'; background: transparent;")
        f_layout.addWidget(self.lbl_footer_right, 1)

    # ==================== QUẢN LÝ CÁC TRANG A4 VÀ XOAY NGANG ĐỘC LẬP ====================
    def add_page(self, orientation: str = "portrait") -> PageSheetWidget:
        """Thêm một trang A4 mới vào tài liệu với tùy chọn khổ giấy độc lập"""
        idx = len(self.pages)
        page = PageSheetWidget(page_index=idx, orientation=orientation, parent=self)
        page.zoom = self.zoom

        # Nếu là Trang 1: chèn Header và Title lên trên
        if idx == 0:
            page.sheet_layout.insertWidget(0, self.header_frame)
            page.sheet_layout.insertWidget(1, self.title_frame)

        self.pages.append(page)
        self.pages_layout.addWidget(page, alignment=Qt.AlignmentFlag.AlignHCenter)

        # Chuyển Footer sang gắn ở cuối trang mới nhất
        self._attach_footer_to_last_page()

        # Kết nối sự kiện tương tác
        page.editor.cursorPositionChanged.connect(self._on_cursor_changed)
        page.editor.textChanged.connect(self.document_modified.emit)
        page.editor.customContextMenuRequested.connect(self._show_body_context_menu)
        page.editor.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)

        page.delete_requested.connect(self.remove_page)
        page.orientation_changed.connect(self._on_page_orientation_changed)

        self._active_page_idx = idx
        page.update_geometry()
        page.editor.setFocus()
        self.document_modified.emit()
        return page

    def remove_page(self, page_index: int):
        """Xóa trang bổ sung (không cho phép xóa trang 1)"""
        if page_index <= 0 or page_index >= len(self.pages):
            return

        page_to_remove = self.pages[page_index]
        self.pages_layout.removeWidget(page_to_remove)
        self.pages.pop(page_index)
        page_to_remove.deleteLater()

        # Cập nhật lại chỉ số cho các trang còn lại
        for i, p in enumerate(self.pages):
            p.page_index = i
            p.update_geometry()

        self._attach_footer_to_last_page()
        self._active_page_idx = max(0, min(self._active_page_idx, len(self.pages) - 1))
        self.pages[self._active_page_idx].editor.setFocus()
        self.document_modified.emit()

    def toggle_page_orientation(self, page_index: int):
        """Xoay ngang / dọc độc lập cho trang được chọn mà KHÔNG ảnh hưởng các trang khác"""
        if 0 <= page_index < len(self.pages):
            page = self.pages[page_index]
            new_orient = "portrait" if page.orientation == "landscape" else "landscape"
            page.orientation = new_orient
            page.update_geometry()
            self.document_modified.emit()

    def _on_page_orientation_changed(self, page_idx: int, new_orient: str):
        self.document_modified.emit()

    def _attach_footer_to_last_page(self):
        """Luôn gắn Khối Nơi nhận & Chữ ký ở cuối trang cuối cùng của tài liệu"""
        if self.footer_frame.parentWidget():
            self.footer_frame.parentWidget().layout().removeWidget(self.footer_frame)

        if self.pages:
            last_page = self.pages[-1]
            last_page.sheet_layout.addWidget(self.footer_frame)

    @property
    def active_editor(self) -> QTextEdit:
        """Trả về editor của trang đang được thao tác hoặc trang 1"""
        for i, p in enumerate(self.pages):
            if p.editor.hasFocus():
                self._active_page_idx = i
                return p.editor
        if 0 <= self._active_page_idx < len(self.pages):
            return self.pages[self._active_page_idx].editor
        return self.pages[0].editor if self.pages else None

    @property
    def body_editor(self) -> QTextEdit:
        """Alias tương thích 100% với các mã nguồn và unit test trước đó (trả về trang 1)"""
        return self.pages[0].editor if self.pages else None

    @property
    def page_sheet(self) -> QWidget:
        """Alias tham chiếu trang 1 cho các tiện ích tương thích"""
        return self.pages[0] if self.pages else self

    def set_zoom(self, zoom: float):
        self.zoom = max(0.4, min(2.5, zoom))
        for p in self.pages:
            p.zoom = self.zoom
            p.update_geometry()
        if hasattr(self, 'top_ctrl_widget'):
            first_w = self.pages[0].width() if self.pages else self.A4_WIDTH_PX
            self.top_ctrl_widget.setFixedWidth(first_w)

    def set_page_orientation(self, orientation: str):
        """Thiết lập định hướng cho trang hiện tại đang thao tác"""
        if 0 <= self._active_page_idx < len(self.pages):
            page = self.pages[self._active_page_idx]
            page.orientation = orientation
            page.update_geometry()
            self.document_modified.emit()

    def set_lock_scaffolding(self, locked: bool):
        self.lock_scaffolding = locked

    def open_metadata_editor(self):
        """Mở hộp thoại chỉnh sửa thể thức văn bản"""
        from ui.editor.metadata_editor_dlg import MetadataEditorDialog
        dlg = MetadataEditorDialog(self, metadata=self.metadata)
        if dlg.exec():
            self.set_metadata(dlg.get_metadata())
            self.document_modified.emit()

    def set_metadata(self, meta: Dict[str, Any]):
        self.metadata.update(meta)
        self._refresh_scaffolding_display()

    def _refresh_scaffolding_display(self):
        """Cập nhật hiển thị Quốc hiệu, Số hiệu, Nơi nhận và Khối ký theo metadata"""
        std = self.metadata.get("standard_id", "ND30").upper()
        sup = self.metadata.get("superior_agency", "").strip().upper()
        iss = self.metadata.get("issuing_agency", "").strip().upper()
        num = self.metadata.get("doc_number", "Số: .../QĐ-UBND")
        date = self.metadata.get("date_str", "Hà Nội, ngày ... tháng ... năm 2026")

        div_sym = "———*———" if std == "HD36" else "—————"

        # Cột Trái Header
        html_l = f"<div style='text-align: center; font-size: 11pt;'>"
        if sup:
            html_l += f"<div style='font-size: 11pt;'>{sup}</div>"
        html_l += f"<div style='font-weight: bold; font-size: 11.5pt;'>{iss or 'ỦY BAN NHÂN DÂN XÃ'}</div>"
        html_l += f"<div style='font-weight: bold; font-size: 8pt; margin: 1px 0;'>{div_sym}</div>"
        html_l += f"<div style='font-size: 11.5pt; margin-top: 2px;'>{num}</div>"
        html_l += "</div>"
        self.lbl_header_left.setText(html_l)

        # Cột Phải Header
        html_r = f"<div style='text-align: center; font-size: 11.5pt;'>"
        if std == "HD36":
            html_r += "<div style='font-weight: bold; font-size: 12pt;'>ĐẢNG CỘNG SẢN VIỆT NAM</div>"
            html_r += "<div style='font-weight: bold; font-size: 9pt; margin: 1px 0;'>———————————————</div>"
        else:
            html_r += "<div style='font-weight: bold; font-size: 11.5pt;'>CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM</div>"
            html_r += "<div style='font-weight: bold; font-size: 12pt;'>Độc lập - Tự do - Hạnh phúc</div>"
            html_r += "<div style='font-weight: bold; font-size: 9pt; margin: 1px 0;'>———————————————</div>"
        html_r += f"<div style='font-style: italic; font-size: 12pt; margin-top: 3px;'>{date}</div>"
        html_r += "</div>"
        self.lbl_header_right.setText(html_r)

        # Tiêu đề & Trích yếu
        self.lbl_doc_type.setText(self.metadata.get("doc_type", "QUYẾT ĐỊNH").upper())
        self.lbl_doc_subject.setText(self.metadata.get("doc_subject", ""))

        # Cột Trái Footer: Nơi nhận
        recs = self.metadata.get("recipients", [])
        html_fl = "<div style='font-size: 10.5pt; line-height: 1.2;'>"
        html_fl += "<div style='font-weight: bold; font-style: italic;'>Nơi nhận:</div>"
        for r in recs:
            html_fl += f"<div>{r}</div>"
        html_fl += "</div>"
        self.lbl_footer_left.setText(html_fl)

        # Cột Phải Footer: Thẩm quyền & Chữ ký
        auth = self.metadata.get("authority_prefix", "TM. ỦY BAN NHÂN DÂN")
        pos = self.metadata.get("signer_position", "CHỦ TỊCH")
        name = self.metadata.get("signer_name", "Nguyễn Văn A")

        html_fr = "<div style='text-align: center; font-size: 12pt; line-height: 1.25;'>"
        if auth:
            html_fr += f"<div style='font-weight: bold;'>{auth}</div>"
        if pos:
            html_fr += f"<div style='font-weight: bold;'>{pos}</div>"
        html_fr += "<div style='height: 48px;'>&nbsp;<br/>&nbsp;</div>"
        html_fr += f"<div style='font-weight: bold; font-size: 12.5pt;'>{name}</div>"
        html_fr += "</div>"
        self.lbl_footer_right.setText(html_fr)

    def _load_sample_body(self):
        sample_html = """
        <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;</p>
        <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Theo đề nghị của Công chức Văn phòng - Thống kê xã,</p>
        <p style='text-align: center; font-weight: bold; margin: 12px 0;'>QUYẾT ĐỊNH:</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 1.</b> Ban hành kèm theo Quyết định này Kế hoạch chuyển đổi số và chuẩn hóa thể thức văn bản hành chính năm 2026 trên địa bàn xã.</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 2.</b> Quyết định này có hiệu lực thi hành kể từ ngày ký.</p>
        <p style='text-align: justify; text-indent: 36px;'><b>Điều 3.</b> Công chức Văn phòng - Thống kê, các ban ngành đoàn thể xã và các thôn, tổ dân phố chịu trách nhiệm thi hành Quyết định này./.</p>
        """
        if self.pages:
            self.pages[0].editor.setHtml(sample_html)

    def _on_cursor_changed(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            line = cursor.blockNumber() + 1
            col = cursor.positionInBlock() + 1
            self.cursor_position_changed.emit(line, col)

    # ==================== MENU CHUỘT PHẢI ====================
    def _show_body_context_menu(self, pos):
        ed = self.active_editor
        if not ed:
            return
        cursor = ed.textCursor()
        table = cursor.currentTable()
        menu = ed.createStandardContextMenu()

        if table:
            menu.addSeparator()
            tbl_menu = menu.addMenu(get_ui_icon("table"), "Bảng biểu")

            act_row_above = tbl_menu.addAction("Thêm hàng phía trên")
            act_row_above.triggered.connect(self.insert_table_row_above)

            act_row_below = tbl_menu.addAction("Thêm hàng phía dưới")
            act_row_below.triggered.connect(self.insert_table_row_below)

            act_del_row = tbl_menu.addAction("Xóa hàng hiện tại")
            act_del_row.triggered.connect(self.delete_table_row)

            tbl_menu.addSeparator()

            act_col_left = tbl_menu.addAction("Thêm cột bên trái")
            act_col_left.triggered.connect(self.insert_table_col_left)

            act_col_right = tbl_menu.addAction("Thêm cột bên phải")
            act_col_right.triggered.connect(self.insert_table_col_right)

            act_del_col = tbl_menu.addAction("Xóa cột hiện tại")
            act_del_col.triggered.connect(self.delete_table_col)

            tbl_menu.addSeparator()

            act_merge = tbl_menu.addAction("Trộn ô đang chọn")
            act_merge.triggered.connect(self.merge_table_cells)

            act_split = tbl_menu.addAction("Tách ô")
            act_split.triggered.connect(self.split_table_cell)

            act_del_tbl = tbl_menu.addAction("Xóa toàn bộ bảng")
            act_del_tbl.triggered.connect(self.delete_current_table)

        menu.addSeparator()
        act_add_p = menu.addAction(get_ui_icon("page_break"), "Thêm trang mới (Ngắt trang)...")
        act_add_p.triggered.connect(lambda: self.add_page(orientation="portrait"))

        act_add_l = menu.addAction(get_ui_icon("page_break"), "Thêm trang Phụ lục ngang (Landscape)...")
        act_add_l.triggered.connect(lambda: self.add_page(orientation="landscape"))

        menu.addSeparator()
        act_meta = menu.addAction(get_ui_icon("edit"), "Sửa thể thức (F2)...")
        act_meta.triggered.connect(self.open_metadata_editor)

        menu.exec(ed.mapToGlobal(pos))

    # ==================== ĐIỀU KHIỂN ĐỊNH DẠNG TỪ RIBBON ====================
    def set_font_bold(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            weight = QFont.Weight.Normal if fmt.fontWeight() == QFont.Weight.Bold else QFont.Weight.Bold
            fmt.setFontWeight(weight)
            cursor.mergeCharFormat(fmt)

    def set_font_italic(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            fmt.setFontItalic(not fmt.fontItalic())
            cursor.mergeCharFormat(fmt)

    def set_font_underline(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            fmt.setFontUnderline(not fmt.fontUnderline())
            cursor.mergeCharFormat(fmt)

    def set_font_family(self, font_name: str):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            fmt.setFontFamily(font_name)
            cursor.mergeCharFormat(fmt)

    def set_font_size(self, size_pt: float):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            fmt.setFontPointSize(size_pt)
            cursor.mergeCharFormat(fmt)

    def set_font_color(self, color: QColor):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            fmt = cursor.charFormat()
            fmt.setForeground(color)
            cursor.mergeCharFormat(fmt)

    def set_alignment(self, align: Qt.AlignmentFlag):
        ed = self.active_editor
        if ed:
            ed.setAlignment(align)

    def set_line_spacing(self, spacing_val: float):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            block_fmt = cursor.blockFormat()
            block_fmt.setLineHeight(int(spacing_val * 100), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
            cursor.setBlockFormat(block_fmt)

    def set_first_line_indent(self, mm: float = 12.7):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            block_fmt = cursor.blockFormat()
            px = int(mm * 3.7795)
            block_fmt.setTextIndent(px)
            cursor.setBlockFormat(block_fmt)

    def insert_symbol(self, sym: str):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            cursor.insertText(sym)

    def insert_bullet_list(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            cursor.insertText("• ")

    def insert_numbered_list(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            cursor.insertText("1. ")

    def insert_table(self, rows: int, cols: int, border_type: str = "grid"):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            tbl_fmt = QTextTableFormat()
            tbl_fmt.setAlignment(Qt.AlignmentFlag.AlignCenter)
            tbl_fmt.setCellPadding(4)
            tbl_fmt.setCellSpacing(0)
            tbl_fmt.setBorder(1 if border_type == "grid" else 0)
            tbl_fmt.setBorderStyle(QTextTableFormat.BorderStyle.BorderStyle_Solid)
            tbl_fmt.setBorderBrush(QColor("#000000"))
            tbl_fmt.setWidth(QTextLength(QTextLength.Type.PercentageLength, 100))
            cursor.insertTable(rows, cols, tbl_fmt)
            ed.setTextCursor(cursor)
            self.document_modified.emit()

    def insert_table_row_above(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cell = table.cellAt(cursor)
                table.insertRows(cell.row(), 1)
                self.document_modified.emit()

    def insert_table_row_below(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cell = table.cellAt(cursor)
                table.insertRows(cell.row() + 1, 1)
                self.document_modified.emit()

    def delete_table_row(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table and table.rows() > 1:
                cell = table.cellAt(cursor)
                target_row = max(0, min(cell.row(), table.rows() - 2))
                target_col = max(0, min(cell.column(), table.columns() - 1))
                table.removeRows(cell.row(), 1)
                new_cell = table.cellAt(target_row, target_col)
                if new_cell.isValid():
                    ed.setTextCursor(new_cell.firstCursorPosition())
                self.document_modified.emit()

    def insert_table_col_left(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cell = table.cellAt(cursor)
                table.insertColumns(cell.column(), 1)
                self.document_modified.emit()

    def insert_table_col_right(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cell = table.cellAt(cursor)
                table.insertColumns(cell.column() + 1, 1)
                self.document_modified.emit()

    def delete_table_col(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table and table.columns() > 1:
                cell = table.cellAt(cursor)
                target_row = max(0, min(cell.row(), table.rows() - 1))
                target_col = max(0, min(cell.column(), table.columns() - 2))
                table.removeColumns(cell.column(), 1)
                new_cell = table.cellAt(target_row, target_col)
                if new_cell.isValid():
                    ed.setTextCursor(new_cell.firstCursorPosition())
                self.document_modified.emit()

    def merge_table_cells(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                table.mergeCells(cursor)
                self.document_modified.emit()

    def split_table_cell(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cell = table.cellAt(cursor)
                table.splitCell(cell.row(), cell.column(), 1, 1)
                self.document_modified.emit()

    def delete_current_table(self):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            table = cursor.currentTable()
            if table:
                cursor.setPosition(table.firstPosition() - 1)
                cursor.setPosition(table.lastPosition() + 1, QTextCursor.MoveMode.KeepAnchor)
                cursor.removeSelectedText()
                self.document_modified.emit()

    def insert_image(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Chọn hình ảnh chèn vào văn bản", "", "Hình ảnh (*.png *.jpg *.jpeg *.bmp)"
        )
        if file_path:
            ed = self.active_editor
            if not ed:
                return
            cursor = ed.textCursor()
            img = QPixmap(file_path)
            # Khổ ảnh tối đa vừa vặn trang in
            is_land = (self.pages[self._active_page_idx].orientation == "landscape") if self.pages else False
            max_w = 700 if is_land else 460
            max_h = 320
            if img.width() > max_w or img.height() > max_h:
                img = img.scaled(max_w, max_h, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)

            cursor.insertBlock()
            block_fmt = cursor.blockFormat()
            block_fmt.setAlignment(Qt.AlignmentFlag.AlignCenter)
            block_fmt.setTextIndent(0)
            cursor.setBlockFormat(block_fmt)
            cursor.insertImage(img.toImage())

            cursor.insertBlock()
            norm_fmt = cursor.blockFormat()
            norm_fmt.setAlignment(Qt.AlignmentFlag.AlignJustify)
            norm_fmt.setTextIndent(int(12.7 * 3.7795))
            cursor.setBlockFormat(norm_fmt)

            ed.setTextCursor(cursor)
            self.document_modified.emit()

    def insert_page_break(self):
        """Hành động ngắt trang: Tạo trang A4 tiếp theo độc lập"""
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            cursor.insertHtml("<div style='margin: 14px 0; border-top: 1.5px dashed #cbd5e1; text-align: center;'><span style='background-color: #f1f5f9; color: #64748b; font-size: 9pt; padding: 2px 14px; border-radius: 4px; font-family: Segoe UI, sans-serif;'>[NGẮT TRANG / PAGE BREAK]</span></div><p style='text-indent: 36px;'><br/></p>")
            ed.setTextCursor(cursor)

        # Tạo trang mới
        self.add_page(orientation="portrait")

    def get_body_paragraphs(self) -> List[str]:
        """Trích xuất danh sách các đoạn văn bản trong toàn bộ các trang để chuyển giao sang OpenXmlBuilder"""
        paras = []
        for idx, p in enumerate(self.pages):
            if idx > 0:
                # Đánh dấu ngắt trang và định hướng xoay cho OpenXmlBuilder
                if p.orientation == "landscape":
                    paras.append("[PAGE_BREAK_LANDSCAPE]")
                else:
                    paras.append("[PAGE_BREAK]")
            raw_text = p.editor.toPlainText()
            for line in raw_text.splitlines():
                line_str = line.strip()
                if not line_str:
                    continue
                if any(k in line_str.upper() for k in ["[PAGE_BREAK]", "PAGE BREAK", "NGẮT TRANG"]):
                    continue
                paras.append(line_str)
        return paras

    def get_document_sections(self) -> List[Dict[str, Any]]:
        """Trả về cấu trúc phân đoạn và định hướng từng trang phục vụ xuất Word và PDF"""
        sections = []
        total = len(self.pages)
        for idx, p in enumerate(self.pages):
            sections.append({
                "page_index": idx,
                "orientation": p.orientation,
                "html": p.editor.toHtml(),
                "text": p.editor.toPlainText(),
                "has_header": (idx == 0),
                "has_footer": (idx == total - 1)
            })
        return sections
