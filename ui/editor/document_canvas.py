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
    QPushButton, QMenu, QGraphicsDropShadowEffect
)
from PyQt6.QtGui import (
    QColor, QFont, QTextCursor, QTextCharFormat,
    QTextBlockFormat, QTextDocument, QPixmap, QTextTableFormat,
    QTextLength, QTextListFormat
)
from PyQt6.QtCore import Qt, pyqtSignal, QSize


class AdministrativeTextEdit(QTextEdit):
    """Trình biên tập nội dung trang có hỗ trợ phím tắt Word (Ctrl+Enter ngắt trang, bắt tràn dòng)"""
    page_break_requested = pyqtSignal()
    overflow_occurred = pyqtSignal()
    backspace_at_start = pyqtSignal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("""
            QTextEdit {
                border: none;
                background-color: transparent;
                font-family: 'Times New Roman';
                font-size: 13pt;
                color: #0f172a;
                line-height: 1.25;
            }
        """)
        self.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.textChanged.connect(self._check_height_overflow)

    def _get_format_for_level(self, level: int) -> QTextListFormat:
        fmt = QTextListFormat()
        fmt.setIndent(level)
        if level == 1:
            fmt.setStyle(QTextListFormat.Style.ListUpperRoman)
            if hasattr(fmt, "setNumberSuffix"):
                fmt.setNumberSuffix(".")
        elif level == 2:
            fmt.setStyle(QTextListFormat.Style.ListDecimal)
            if hasattr(fmt, "setNumberSuffix"):
                fmt.setNumberSuffix(".")
        elif level == 3:
            fmt.setStyle(QTextListFormat.Style.ListLowerAlpha)
            if hasattr(fmt, "setNumberSuffix"):
                fmt.setNumberSuffix(")")
        else:
            fmt.setStyle(QTextListFormat.Style.ListDisc)
        return fmt

    def _demote_current_list_item(self, cursor=None):
        """Phím Tab trong danh sách: Tăng cấp thụt lề / Đổi cấp phân cấp như Microsoft Word"""
        if cursor is None:
            cursor = self.textCursor()
        cur_block = cursor.block()
        tl = cur_block.textList()
        cur_indent = tl.format().indent() if tl else 1
        new_indent = min(cur_indent + 1, 4)

        # Tìm danh sách trước đó có cùng new_indent để nối số
        target_list = None
        b = cur_block.previous()
        while b.isValid():
            btl = b.textList()
            if btl and btl.format().indent() == new_indent:
                target_list = btl
                break
            elif btl and btl.format().indent() < new_indent:
                break
            b = b.previous()

        if target_list:
            target_list.add(cur_block)
        else:
            fmt = self._get_format_for_level(new_indent)
            cursor.createList(fmt)

    def _promote_current_list_item(self, cursor=None):
        """Phím Shift+Tab trong danh sách: Giảm cấp phân cấp / Thoát danh sách nếu ở cấp 1"""
        if cursor is None:
            cursor = self.textCursor()
        cur_block = cursor.block()
        tl = cur_block.textList()
        cur_indent = tl.format().indent() if tl else 1
        if cur_indent <= 1:
            # Thoát danh sách
            if tl:
                tl.remove(cur_block)
            bf = cursor.blockFormat()
            bf.setObjectIndex(-1)
            bf.setIndent(0)
            cursor.setBlockFormat(bf)
        else:
            new_indent = cur_indent - 1
            target_list = None
            b = cur_block.previous()
            while b.isValid():
                btl = b.textList()
                if btl and btl.format().indent() == new_indent:
                    target_list = btl
                    break
                b = b.previous()

            if target_list:
                target_list.add(cur_block)
            else:
                fmt = self._get_format_for_level(new_indent)
                cursor.createList(fmt)

    def keyPressEvent(self, event):
        # 1. Ctrl + Enter / Ctrl + Return => Ngắt trang sang trang A4 tiếp theo chuẩn xác
        if (event.modifiers() & Qt.KeyboardModifier.ControlModifier) and event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter):
            self.page_break_requested.emit()
            event.accept()
            return

        # 2. Phím Tab trong danh sách => Tăng cấp phân cấp (Word demote)
        if event.key() == Qt.Key.Key_Tab and not (event.modifiers() & (Qt.KeyboardModifier.ControlModifier | Qt.KeyboardModifier.ShiftModifier)):
            cur = self.textCursor()
            if cur.currentList():
                self._demote_current_list_item(cur)
                event.accept()
                return

        # 3. Phím Shift + Tab (hoặc Key_Backtab) trong danh sách => Giảm cấp phân cấp (Word promote)
        if (event.key() == Qt.Key.Key_Backtab) or (event.key() == Qt.Key.Key_Tab and (event.modifiers() & Qt.KeyboardModifier.ShiftModifier)):
            cur = self.textCursor()
            if cur.currentList():
                self._promote_current_list_item(cur)
                event.accept()
                return

        # 4. Phím Enter trên một dòng danh sách trống => Thoát danh sách (như Microsoft Word)
        if event.key() in (Qt.Key.Key_Return, Qt.Key.Key_Enter) and not (event.modifiers() & Qt.KeyboardModifier.ControlModifier):
            cur = self.textCursor()
            lst = cur.currentList()
            if lst and len(cur.block().text().strip()) == 0:
                lst.remove(cur.block())
                bf = cur.blockFormat()
                bf.setObjectIndex(-1)
                bf.setIndent(0)
                cur.setBlockFormat(bf)
                event.accept()
                return

        # 5. Backspace ở đầu trang rỗng => Xóa trang và nhảy về trang trước
        if event.key() == Qt.Key.Key_Backspace:
            cur = self.textCursor()
            if cur.position() == 0 and not cur.hasSelection():
                self.backspace_at_start.emit()

        super().keyPressEvent(event)

    def _check_height_overflow(self):
        # Tự động phát hiện khi nội dung vượt quá chiều cao trang in A4 thực tế
        vh = self.viewport().height()
        if vh > 150:
            doc_h = self.document().size().height()
            if doc_h > vh + 12:
                self.overflow_occurred.emit()


class PageSheetWidget(QFrame):
    """Một trang giấy A4 trực quan độc lập chuẩn quy cách Microsoft Word"""

    A4_WIDTH_PX = 794
    A4_HEIGHT_PX = 1123

    focus_received = pyqtSignal(int)
    orientation_changed = pyqtSignal(int, str)
    delete_requested = pyqtSignal(int)
    page_break_shortcut = pyqtSignal(int)
    page_overflow = pyqtSignal(int)
    page_backspace_start = pyqtSignal(int)

    def __init__(self, page_index: int = 0, orientation: str = "portrait", parent=None):
        super().__init__(parent)
        self.page_index = page_index
        self.orientation = orientation
        self.zoom = 1.0

        self.setObjectName("page_sheet")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet("""
            QFrame#page_sheet {
                background-color: #ffffff !important;
                border: 1px solid #94a3b8;
                border-radius: 2px;
            }
        """)

        # Đổ bóng chân thực như trang giấy A4 trên bàn làm việc trong Microsoft Word
        self.shadow_effect = QGraphicsDropShadowEffect(self)
        self.shadow_effect.setBlurRadius(20)
        self.shadow_effect.setOffset(0, 4)
        self.shadow_effect.setColor(QColor(0, 0, 0, 75))
        self.setGraphicsEffect(self.shadow_effect)

        self.sheet_layout = QVBoxLayout(self)
        self.sheet_layout.setSpacing(6)

        # Thanh điều khiển trang (Chỉ hiển thị từ trang 2 trở đi)
        self.top_page_bar = QWidget()
        self.top_page_bar.setStyleSheet("background: transparent;")
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
        self.editor = AdministrativeTextEdit(self)
        self.editor.page_break_requested.connect(lambda: self.page_break_shortcut.emit(self.page_index))
        self.editor.overflow_occurred.connect(lambda: self.page_overflow.emit(self.page_index))
        self.editor.backspace_at_start.connect(lambda: self.page_backspace_start.emit(self.page_index))
        self.sheet_layout.addWidget(self.editor, 1)

        # Cài đặt kích thước ban đầu chuẩn A4
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
        self.setFixedHeight(h)  # CỐ ĐỊNH CHIỀU CAO TRANG A4 CHUẨN XÁC - KHÔNG KÉO DÃN TRANG

        # Lề NĐ 30: Trái 30mm (~113px), Phải 15mm (~57px), Trên 20mm (~76px), Dưới 20mm (~76px)
        top_m = 60 if self.page_index == 0 else 24
        self.sheet_layout.setContentsMargins(113, top_m, 57, 50)

        if is_land:
            self.btn_toggle_orient.setText("Đổi sang: Trang dọc (210×297 mm)")
            self.lbl_page_title.setText(f"Trang {self.page_index + 1} (Khổ ngang - Phụ lục / Bảng biểu)")
        else:
            self.btn_toggle_orient.setText("Đổi sang: Trang ngang (297×210 mm)")
            self.lbl_page_title.setText(f"Trang {self.page_index + 1} (Khổ dọc chuẩn)")

        self.lbl_page_number.setText(f"- {self.page_index + 1} -")


class DocumentCanvas(QFrame):
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

        self.setObjectName("document_canvas")
        self.setAttribute(Qt.WidgetAttribute.WA_StyledBackground, True)
        self.setStyleSheet("""
            QFrame#document_canvas {
                background-color: #718096;
            }
        """)

        # Bố cục chính căn giữa màn hình
        self.main_layout = QVBoxLayout(self)
        self.main_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.main_layout.setContentsMargins(0, 16, 0, 60)
        self.main_layout.setSpacing(16)

        # 1. Thanh tiện ích phía trên trang giấy (Căn giữa)
        self.top_ctrl_widget = QWidget()
        self.top_ctrl_widget.setObjectName("top_ctrl_widget")
        self.top_ctrl_widget.setFixedWidth(self.A4_WIDTH_PX)
        self.top_ctrl_widget.setStyleSheet("""
            QWidget#top_ctrl_widget {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
            }
        """)
        top_ctrl_bar = QHBoxLayout(self.top_ctrl_widget)
        top_ctrl_bar.setContentsMargins(12, 6, 12, 6)
        top_ctrl_bar.setSpacing(10)

        lbl_page_info = QLabel("Khổ giấy: A4 (210 × 297 mm) • Chuẩn Thể thức Cấp Xã")
        lbl_page_info.setStyleSheet("color: #475569; font-size: 11.5px; font-weight: 600; background: transparent;")
        top_ctrl_bar.addWidget(lbl_page_info)
        top_ctrl_bar.addStretch()

        self.btn_add_page_top = QPushButton("Thêm trang mới (Ngắt trang)")
        self.btn_add_page_top.setStyleSheet("""
            QPushButton {
                background-color: #eff6ff; color: #2563eb; font-weight: 600;
                border: 1.5px solid #bfdbfe; border-radius: 4px; padding: 4px 10px; font-size: 11.5px;
            }
            QPushButton:hover { background-color: #dbeafe; border-color: #2563eb; }
        """)
        self.btn_add_page_top.clicked.connect(self.insert_page_break)
        top_ctrl_bar.addWidget(self.btn_add_page_top)

        self.btn_add_landscape_top = QPushButton("Thêm trang ngang (Landscape)")
        self.btn_add_landscape_top.setStyleSheet("""
            QPushButton {
                background-color: #f0f9ff; color: #0284c7; font-weight: 600;
                border: 1.5px solid #bae6fd; border-radius: 4px; padding: 4px 10px; font-size: 11.5px;
            }
            QPushButton:hover { background-color: #e0f2fe; border-color: #0284c7; }
        """)
        self.btn_add_landscape_top.clicked.connect(lambda: self.add_page(orientation="landscape"))
        top_ctrl_bar.addWidget(self.btn_add_landscape_top)

        self.btn_edit_meta = QPushButton("Chỉnh sửa thể thức (F2)")
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
        self.pages_container.setStyleSheet("background: transparent;")
        self.pages_layout = QVBoxLayout(self.pages_container)
        self.pages_layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        self.pages_layout.setContentsMargins(0, 0, 0, 0)
        self.pages_layout.setSpacing(40)  # Khoảng cách 40px giữa các trang giấy A4 tạo ranh giới ngắt trang rõ ràng
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
        page.page_break_shortcut.connect(self.split_page_at_cursor)
        page.page_overflow.connect(self._on_page_overflow)
        page.page_backspace_start.connect(self._on_backspace_at_start)

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
        if iss:
            html_l += f"<div style='font-weight: bold; font-size: 11.5pt;'>{iss}</div>"
        elif not sup:
            html_l += f"<div style='font-weight: bold; font-size: 11.5pt;'>CƠ QUAN BAN HÀNH</div>"
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
        doc_type_text = self.metadata.get("doc_type", "").strip().upper() or "VĂN BẢN"
        self.lbl_doc_type.setText(doc_type_text)
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
        auth = self.metadata.get("authority_prefix", "").strip()
        pos = self.metadata.get("signer_position", "").strip()
        name = self.metadata.get("signer_name", "").strip()

        html_fr = "<div style='text-align: center; font-size: 12pt; line-height: 1.25;'>"
        if auth:
            html_fr += f"<div style='font-weight: bold;'>{auth}</div>"
        if pos:
            html_fr += f"<div style='font-weight: bold;'>{pos}</div>"
        html_fr += "<div style='height: 48px;'>&nbsp;<br/>&nbsp;</div>"
        if name:
            html_fr += f"<div style='font-weight: bold; font-size: 12.5pt;'>{name}</div>"
        html_fr += "</div>"
        self.lbl_footer_right.setText(html_fr)

    @classmethod
    def get_template_for_type(cls, doc_type: str) -> Dict[str, Any]:
        """Trả về tiêu đề, trích yếu và nội dung mẫu tương ứng với từng loại văn bản"""
        dt = (doc_type or "QUYẾT ĐỊNH").strip().upper()
        if "TỜ TRÌNH" in dt:
            return {
                "doc_type": "TỜ TRÌNH",
                "doc_subject": "Về việc phê duyệt Kế hoạch chuyển đổi số và chuẩn hóa văn bản năm 2026",
                "doc_number": "Số: 01/TTr-UBND",
                "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "recipients": [
                    "- Ủy ban nhân dân huyện (để báo cáo);",
                    "- Thường trực Đảng ủy, HĐND xã;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: left; text-indent: 0px;'><b>Kính gửi:</b> Ủy ban nhân dân huyện.</p>
                <p style='text-align: justify; text-indent: 36px;'>Thực hiện Kế hoạch phát triển kinh tế - xã hội và chuyển đổi số giai đoạn 2026 - 2030, Ủy ban nhân dân xã kính trình Ủy ban nhân dân huyện xem xét, phê duyệt Kế hoạch chuyển đổi số và chuẩn hóa văn bản hành chính năm 2026 với các nội dung chính sau:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>1. Sự cần thiết ban hành:</b></p>
                <p style='text-align: justify; text-indent: 36px;'>Nhằm nâng cao hiệu lực, hiệu quả quản lý nhà nước và tăng cường kỷ cương hành chính trên địa bàn...</p>
                <p style='text-align: justify; text-indent: 36px;'><b>2. Nội dung đề xuất:</b></p>
                <p style='text-align: justify; text-indent: 36px;'>Đầu tư nâng cấp hệ thống hạ tầng công nghệ thông tin, tập huấn nghiệp vụ cho đội ngũ cán bộ, công chức xã.</p>
                <p style='text-align: justify; text-indent: 36px;'>Ủy ban nhân dân xã kính trình Ủy ban nhân dân huyện xem xét, quyết định./.</p>
                """
            }
        elif "BÁO CÁO" in dt:
            return {
                "doc_type": "BÁO CÁO",
                "doc_subject": "Về việc rà soát, đánh giá công tác cải cách hành chính và chuyển đổi số",
                "doc_number": "Số: 45/BC-SNV",
                "authority_prefix": "TM. SỞ NỘI VỤ",
                "signer_position": "GIÁM ĐỐC",
                "signer_name": "Nguyễn Văn A",
                "recipients": [
                    "- Ủy ban nhân dân tỉnh (để báo cáo);",
                    "- Giám đốc, các Phó Giám đốc Sở;",
                    "- Các phòng chuyên môn;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px;'>Thực hiện ý kiến chỉ đạo của Ủy ban nhân dân tỉnh tại Công văn số 123/UBND-TH, Sở Nội vụ đã chủ trì, phối hợp với các cơ quan, đơn vị tiến hành kiểm tra tình hình thực tế. Qua rà soát, nhìn chung các đơn vị đã nghiêm túc triển khai nhiệm vụ được giao.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>I. KẾT QUẢ ĐẠT ĐƯỢC</b></p>
                <p style='text-align: justify; text-indent: 36px;'>1. Toàn thể cán bộ, công chức đã tích cực tham gia phong trào thi đua và hoàn thành 100% kế hoạch đề ra.</p>
                <p style='text-align: justify; text-indent: 36px;'>2. Công tác chuyển đổi số và giải quyết thủ tục hành chính qua cổng dịch vụ công đạt tỷ lệ đúng hạn trên 98%.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>II. TỒN TẠI, HẠN CHẾ VÀ ĐỀ XUẤT KIẾN NGHỊ</b></p>
                <p style='text-align: justify; text-indent: 36px;'>Một số đơn vị chưa chủ động trong việc ứng dụng công nghệ thông tin. Đề nghị UBND tỉnh tiếp tục quan tâm, chỉ đạo./.</p>
                """
            }
        elif "CÔNG VĂN" in dt:
            return {
                "doc_type": "CÔNG VĂN",
                "doc_subject": "V/v phối hợp rà soát, thống kê số liệu phục vụ chuyển đổi số",
                "doc_number": "Số: 12/UBND-VP",
                "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "recipients": [
                    "- Như trên;",
                    "- Thường trực Đảng ủy xã;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: left; text-indent: 0px;'><b>Kính gửi:</b></p>
                <p style='text-align: left; text-indent: 36px;'>- Các phòng chuyên môn, ban ngành đoàn thể xã;</p>
                <p style='text-align: left; text-indent: 36px;'>- Trưởng các thôn, tổ dân phố trên địa bàn.</p>
                <p style='text-align: justify; text-indent: 36px;'>Để phục vụ công tác xây dựng cơ sở dữ liệu quốc gia và nâng cao chất lượng phục vụ nhân dân, Ủy ban nhân dân xã đề nghị các cơ quan, đơn vị triển khai thực hiện các nội dung sau:</p>
                <p style='text-align: justify; text-indent: 36px;'>1. Khẩn trương rà soát, lập danh sách các trường hợp thuộc diện quản lý gửi về UBND xã trước ngày 15 hàng tháng.</p>
                <p style='text-align: justify; text-indent: 36px;'>2. Phối hợp với công chức chuyên môn giải quyết kịp thời các vướng mắc phát sinh.</p>
                <p style='text-align: justify; text-indent: 36px;'>Nhận được công văn này, đề nghị các đơn vị nghiêm túc triển khai thực hiện./.</p>
                """
            }
        elif "THÔNG BÁO" in dt:
            return {
                "doc_type": "THÔNG BÁO",
                "doc_subject": "Về việc tổ chức tiếp nhận hồ sơ thủ tục hành chính trực tuyến",
                "doc_number": "Số: 08/TB-UBND",
                "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "recipients": [
                    "- Toàn thể nhân dân trên địa bàn;",
                    "- Bộ phận Một cửa;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px;'>Ủy ban nhân dân xã thông báo đến toàn thể nhân dân, tổ chức và cá nhân trên địa bàn xã nội dung sau:</p>
                <p style='text-align: justify; text-indent: 36px;'>1. Kể từ ngày 01 tháng 01 năm 2026, Ủy ban nhân dân xã thực hiện tiếp nhận 100% hồ sơ thủ tục hành chính trực tuyến qua Cổng Dịch vụ công quốc gia.</p>
                <p style='text-align: justify; text-indent: 36px;'>2. Đề nghị các tổ chức, công dân chủ động đăng ký tài khoản định danh điện tử VNeID để thực hiện các giao dịch hành chính.</p>
                <p style='text-align: justify; text-indent: 36px;'>Ủy ban nhân dân xã thông báo để toàn thể nhân dân được biết và phối hợp thực hiện./.</p>
                """
            }
        elif "CHỈ THỊ" in dt:
            return {
                "standard_id": "HD36",
                "superior_agency": "HUYỆN ỦY",
                "issuing_agency": "BAN THƯỜNG VỤ",
                "doc_type": "CHỈ THỊ",
                "doc_subject": "Về tăng cường sự lãnh đạo của Đảng đối với công tác cải cách hành chính năm 2026",
                "doc_number": "Số 02-CT/ĐU",
                "authority_prefix": "T/M BAN THƯỜNG VỤ",
                "signer_position": "BÍ THƯ",
                "signer_name": "Nguyễn Văn An",
                "recipients": [
                    "- Thường trực Huyện ủy (để b/c);",
                    "- Các chi bộ trực thuộc;",
                    "- Lưu: VP."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px;'>Trong thời gian qua, các cấp ủy Đảng đã tích cực triển khai phong trào thi đua và đạt nhiều kết quả quan trọng trong công tác xây dựng Đảng và phát triển kinh tế - xã hội.</p>
                <p style='text-align: justify; text-indent: 36px;'>Để tiếp tục đẩy mạnh ứng dụng công nghệ thông tin và cải cách hành chính trong toàn hệ thống chính trị,</p>
                <p style='text-align: center; font-weight: bold; margin: 12px 0;'>BAN THƯỜNG VỤ CHỈ THỊ:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>1.</b> Cấp ủy các chi bộ trực thuộc quán triệt sâu sắc các chỉ thị của Đảng về chuyển đổi số và chuẩn hóa văn bản.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>2.</b> Giao Văn phòng Đảng ủy thường xuyên theo dõi, đôn đốc và báo cáo Ban Thường vụ Đảng ủy kết quả thực hiện./.</p>
                """
            }
        elif "NGHỊ QUYẾT ĐẢNG" in dt or ("NGHỊ QUYẾT" in dt and "HD36" in str(dt)):
            return {
                "standard_id": "HD36",
                "superior_agency": "HUYỆN ỦY",
                "issuing_agency": "ĐẢNG BỘ XÃ",
                "doc_type": "NGHỊ QUYẾT",
                "doc_subject": "Về nhiệm vụ công tác xây dựng Đảng và phát triển kinh tế năm 2026",
                "doc_number": "Số 01-NQ/ĐU",
                "authority_prefix": "T/M ĐẢNG ỦY",
                "signer_position": "BÍ THƯ",
                "signer_name": "Nguyễn Văn An",
                "recipients": [
                    "- Ban Thường vụ Huyện ủy (để b/c);",
                    "- Các chi bộ trực thuộc;",
                    "- Lưu: VP Đảng ủy."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Điều lệ Đảng Cộng sản Việt Nam;</p>
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Quy chế làm việc của Ban Chấp hành Đảng bộ xã;</p>
                <p style='text-align: center; font-weight: bold; margin: 12px 0;'>QUYẾT NGHỊ:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 1.</b> Thông qua Kế hoạch phát triển công tác Đảng và các mục tiêu kinh tế - xã hội năm 2026.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 2.</b> Giao Ban Thường vụ Đảng ủy và Ủy ban nhân dân xã phối hợp tổ chức triển khai thực hiện thắng lợi Nghị quyết này./.</p>
                """
            }
        elif "KẾ HOẠCH" in dt:
            return {
                "doc_type": "KẾ HOẠCH",
                "doc_subject": "Về triển khai công tác cải cách hành chính và chuyển đổi số năm 2026",
                "doc_number": "Số: 10/KH-UBND",
                "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn An",
                "recipients": [
                    "- UBND huyện (để b/c);",
                    "- Các ban ngành đoàn thể;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px;'>Thực hiện Kế hoạch của Ủy ban nhân dân huyện, Ủy ban nhân dân xã ban hành Kế hoạch triển khai công tác cải cách hành chính năm 2026 như sau:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>I. MỤC ĐÍCH, YÊU CẦU</b></p>
                <p style='text-align: justify; text-indent: 36px;'>1. Tăng cường hiệu lực, hiệu quả quản trị công vụ và giải quyết thủ tục hành chính cho nhân dân.</p>
                <p style='text-align: justify; text-indent: 36px;'>2. 100% cán bộ, công chức sử dụng thành thạo hệ thống văn bản điện tử và ký số.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>II. NỘI DUNG VÀ BIỆN PHÁP THỰC HIỆN</b></p>
                <p style='text-align: justify; text-indent: 36px;'>Tổ chức rà soát, cắt giảm thời gian xử lý thủ tục hành chính; kiểm tra định kỳ hàng quý./.</p>
                """
            }
        elif "NGHỊ QUYẾT" in dt:
            return {
                "doc_type": "NGHỊ QUYẾT",
                "doc_subject": "Về nhiệm vụ phát triển kinh tế - xã hội năm 2026",
                "doc_number": "Số: 01/NQ-HĐND",
                "authority_prefix": "TM. HỘI ĐỒNG NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn An",
                "recipients": [
                    "- Thường trực HĐND huyện;",
                    "- Đảng ủy, HĐND, UBND xã;",
                    "- Đại biểu HĐND xã;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;</p>
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Sau khi xem xét Tờ trình của Ủy ban nhân dân xã và Báo cáo thẩm tra của các Ban HĐND xã,</p>
                <p style='text-align: center; font-weight: bold; margin: 12px 0;'>QUYẾT NGHỊ:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 1.</b> Thông qua Kế hoạch phát triển kinh tế - xã hội, quốc phòng - an ninh năm 2026...</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 2.</b> Giao Ủy ban nhân dân xã tổ chức triển khai thực hiện thắng lợi Nghị quyết này.</p>
                <p style='text-align: justify; text-indent: 36px;'>Nghị quyết này đã được Hội đồng nhân dân xã khóa XX thông qua ngày 20 tháng 12 năm 2025./.</p>
                """
            }
        else: # Mặc định QUYẾT ĐỊNH
            return {
                "doc_type": "QUYẾT ĐỊNH",
                "doc_subject": "Về việc ban hành Kế hoạch phát triển kinh tế - xã hội năm 2026",
                "doc_number": "Số: 15/QĐ-UBND",
                "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "recipients": [
                    "- Thường trực HĐND xã;",
                    "- Chủ tịch, các PCT UBND xã;",
                    "- Các ban ngành, đoàn thể xã;",
                    "- Các thôn, tổ dân phố;",
                    "- Lưu: VT."
                ],
                "body_html": """
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;</p>
                <p style='text-align: justify; text-indent: 36px; font-style: italic;'>Theo đề nghị của Công chức Văn phòng - Thống kê xã,</p>
                <p style='text-align: center; font-weight: bold; margin: 12px 0;'>QUYẾT ĐỊNH:</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 1.</b> Ban hành kèm theo Quyết định này Kế hoạch chuyển đổi số và chuẩn hóa thể thức văn bản hành chính năm 2026 trên địa bàn xã.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 2.</b> Quyết định này có hiệu lực thi hành kể từ ngày ký.</p>
                <p style='text-align: justify; text-indent: 36px;'><b>Điều 3.</b> Công chức Văn phòng - Thống kê, các ban ngành đoàn thể xã và các thôn, tổ dân phố chịu trách nhiệm thi hành Quyết định này./.</p>
                """
            }

    def _load_sample_body(self):
        tpl = self.get_template_for_type(self.metadata.get("doc_type", "QUYẾT ĐỊNH"))
        if self.pages:
            self.pages[0].editor.setHtml(tpl["body_html"])

    def apply_nd30_standards_and_alignment(self) -> Dict[str, Any]:
        """
        Tự động căn chỉnh toàn bộ tài liệu về chuẩn Nghị định 30/2020/NĐ-CP:
        - Chuẩn hóa lề in chuẩn: Trái 30mm (~113px), Phải 15mm (~57px), Trên 20mm, Dưới 20mm
        - Dò quét từng đoạn văn bản:
          + Loại bỏ các khoảng trắng / tab gõ thủ công ở đầu dòng gây xô lệch lề hàng
          + Căn giữa tiêu đề chính (QUYẾT ĐỊNH:, QUYẾT NGHỊ:, BÁO CÁO, TỜ TRÌNH, ...)
          + Căn trái dòng Kính gửi, Nơi nhận
          + Căn đều 2 lề (AlignJustify) cho nội dung, căn cứ, điều khoản
          + Thụt lề đầu dòng chuẩn NĐ 30: 36.0pt (~1.0 - 1.27 cm)
          + Giãn dòng 1.25, giãn đoạn: 0 pt trên, 6 pt dưới
        - Làm tươi lại khung thể thức và phát tín hiệu sửa đổi
        """
        fixed_count = 0
        stripped_count = 0

        for page in self.pages:
            page.update_geometry()
            doc = page.editor.document()
            cursor = QTextCursor(doc)
            cursor.beginEditBlock()
            block = doc.begin()
            while block.isValid():
                text = block.text()
                stripped = text.lstrip()
                if len(stripped) < len(text):
                    diff = len(text) - len(stripped)
                    c = QTextCursor(block)
                    c.movePosition(QTextCursor.MoveOperation.StartOfBlock)
                    c.movePosition(QTextCursor.MoveOperation.Right, QTextCursor.MoveMode.KeepAnchor, diff)
                    c.removeSelectedText()
                    stripped_count += 1

                clean_text = stripped.strip()
                if clean_text:
                    # Tự động chuẩn hóa dấu câu cho các dòng Căn cứ pháp lý (phải kết thúc bằng ';')
                    if clean_text.startswith("Căn cứ "):
                        if not (clean_text.endswith(";") or clean_text.endswith(".")):
                            c_punc = QTextCursor(block)
                            raw_b = block.text()
                            r_strip = raw_b.rstrip()
                            if len(r_strip) < len(raw_b):
                                c_punc.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                                c_punc.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.KeepAnchor, len(raw_b) - len(r_strip))
                                c_punc.removeSelectedText()
                            clean_text = block.text().strip()
                            c_punc.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                            if clean_text.endswith(",") or clean_text.endswith(":"):
                                c_punc.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.KeepAnchor, 1)
                                c_punc.insertText(";")
                            else:
                                c_punc.insertText(";")
                            clean_text = block.text().strip()

                    # Tự động chuẩn hóa dấu câu cho dòng Đề nghị trước lệnh ban hành (phải kết thúc bằng ',')
                    elif clean_text.startswith("Theo đề nghị của"):
                        if not clean_text.endswith(","):
                            c_punc = QTextCursor(block)
                            raw_b = block.text()
                            r_strip = raw_b.rstrip()
                            if len(r_strip) < len(raw_b):
                                c_punc.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                                c_punc.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.KeepAnchor, len(raw_b) - len(r_strip))
                                c_punc.removeSelectedText()
                            clean_text = block.text().strip()
                            c_punc.movePosition(QTextCursor.MoveOperation.EndOfBlock)
                            if clean_text.endswith(";") or clean_text.endswith(".") or clean_text.endswith(":"):
                                c_punc.movePosition(QTextCursor.MoveOperation.Left, QTextCursor.MoveMode.KeepAnchor, 1)
                                c_punc.insertText(",")
                            else:
                                c_punc.insertText(",")
                            clean_text = block.text().strip()

                    bf = block.blockFormat()
                    upper_text = clean_text.upper()

                    # Căn giữa cho tiêu đề chính
                    if upper_text in [
                        "QUYẾT ĐỊNH:", "QUYẾT ĐỊNH", "QUYẾT NGHỊ:", "QUYẾT NGHỊ",
                        "TỜ TRÌNH", "BÁO CÁO", "THÔNG BÁO", "KẾ HOẠCH"
                    ] or (len(clean_text) <= 40 and (upper_text.startswith("CHƯƠNG ") or upper_text.startswith("MỤC "))):
                        bf.setAlignment(Qt.AlignmentFlag.AlignCenter)
                        bf.setTextIndent(0.0)
                    elif clean_text.lower().startswith("kính gửi:") or clean_text.lower().startswith("nơi nhận:"):
                        bf.setAlignment(Qt.AlignmentFlag.AlignLeft)
                        bf.setTextIndent(0.0)
                    elif clean_text.startswith("- ") or clean_text.startswith("+ ") or clean_text.startswith("• "):
                        bf.setAlignment(Qt.AlignmentFlag.AlignJustify)
                        bf.setTextIndent(18.0)
                    else:
                        bf.setAlignment(Qt.AlignmentFlag.AlignJustify)
                        bf.setTextIndent(36.0)  # ~1.0 - 1.27 cm chuẩn NĐ 30

                    bf.setLineHeight(125, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
                    bf.setTopMargin(0.0)
                    bf.setBottomMargin(6.0)

                    c = QTextCursor(block)
                    c.setBlockFormat(bf)
                    fixed_count += 1

                block = block.next()
            cursor.endEditBlock()

        self._refresh_scaffolding_display()
        self.document_modified.emit()

        return {
            "fixed_paragraphs": fixed_count,
            "stripped_spaces": stripped_count,
            "margins": "Trái 30mm, Phải 15mm, Trên 20mm, Dưới 20mm"
        }

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

        # Chuẩn hóa ngôn ngữ tiếng Việt chuyên nghiệp cho Menu ngữ cảnh (chuột phải)
        trans_map = {
            'Undo': 'Quay lại',
            'Redo': 'Làm lại',
            'Cut': 'Cắt',
            'Copy': 'Sao chép',
            'Paste': 'Dán',
            'Delete': 'Xóa',
            'Select All': 'Chọn tất cả',
        }
        for act in menu.actions():
            raw = act.text()
            if raw:
                parts = raw.split('\t')
                label = parts[0].replace('&', '')
                shortcut = ('\t' + parts[1]) if len(parts) > 1 else ''
                if label in trans_map:
                    act.setText(trans_map[label] + shortcut)

        if table:
            menu.addSeparator()
            tbl_menu = menu.addMenu("Bảng biểu")

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
        act_add_p = menu.addAction("Thêm trang mới (Ngắt trang)...")
        act_add_p.triggered.connect(lambda: self.add_page(orientation="portrait"))

        act_add_l = menu.addAction("Thêm trang Phụ lục ngang (Landscape)...")
        act_add_l.triggered.connect(lambda: self.add_page(orientation="landscape"))

        menu.addSeparator()
        act_meta = menu.addAction("Sửa thể thức (F2)...")
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

    def _modify_selected_blocks_format(self, modifier_func):
        """Áp dụng thay đổi định dạng đoạn văn (QTextBlockFormat) lên tất cả các đoạn đang được chọn hoặc đoạn hiện tại"""
        ed = self.active_editor
        if not ed:
            return
        cursor = ed.textCursor()
        cursor.beginEditBlock()
        try:
            doc = ed.document()
            start_pos = cursor.selectionStart()
            end_pos = cursor.selectionEnd()
            start_block = doc.findBlock(start_pos)
            end_block = doc.findBlock(max(start_pos, end_pos - 1 if end_pos > start_pos else end_pos))

            curr = start_block
            while curr.isValid():
                b_cursor = QTextCursor(curr)
                fmt = curr.blockFormat()
                modifier_func(fmt)
                b_cursor.setBlockFormat(fmt)
                if curr == end_block:
                    break
                curr = curr.next()
        finally:
            cursor.endEditBlock()
        self.document_modified.emit()

    def set_line_spacing(self, spacing_val: float):
        """Đặt khoảng cách giãn dòng (1.0, 1.15, 1.2, 1.25, 1.5, 2.0)"""
        self._modify_selected_blocks_format(
            lambda fmt: fmt.setLineHeight(int(spacing_val * 100), QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
        )

    def set_first_line_indent(self, mm: float = 10.0):
        """Thụt dòng đầu tiên của đoạn (1.0cm chuẩn NĐ 30, 1.27cm chuẩn Word 0.5', 0cm)"""
        px = int(mm * 3.7795)
        self._modify_selected_blocks_format(lambda fmt: fmt.setTextIndent(px))

    def increase_indent(self, step_mm: float = 10.0):
        """Tăng thụt lề toàn bộ đoạn (Increase Indent) hoặc thăng cấp phân cấp danh sách nếu đang ở trong List"""
        ed = self.active_editor
        if ed and ed.textCursor().currentList() and hasattr(ed, "_demote_current_list_item"):
            ed._demote_current_list_item()
            self.document_modified.emit()
            return
        step_px = int(step_mm * 3.7795)
        self._modify_selected_blocks_format(lambda fmt: fmt.setLeftMargin(fmt.leftMargin() + step_px))

    def decrease_indent(self, step_mm: float = 10.0):
        """Giảm thụt lề toàn bộ đoạn (Decrease Indent) hoặc hạ cấp/thoát danh sách nếu đang ở trong List"""
        ed = self.active_editor
        if ed and ed.textCursor().currentList() and hasattr(ed, "_promote_current_list_item"):
            ed._promote_current_list_item()
            self.document_modified.emit()
            return
        step_px = int(step_mm * 3.7795)
        self._modify_selected_blocks_format(lambda fmt: fmt.setLeftMargin(max(0.0, fmt.leftMargin() - step_px)))

    def remove_space_before(self):
        """Xóa khoảng cách thừa phía trên đoạn văn (Remove Space Before Paragraph - 0pt)"""
        self._modify_selected_blocks_format(lambda fmt: fmt.setTopMargin(0.0))

    def remove_space_after(self):
        """Xóa khoảng cách thừa phía dưới đoạn văn (Remove Space After Paragraph - 0pt)"""
        self._modify_selected_blocks_format(lambda fmt: fmt.setBottomMargin(0.0))

    def remove_all_spacing(self):
        """Xóa sạch mọi khoảng thừa trên và dưới (0pt / 0pt)"""
        def _clear(fmt):
            fmt.setTopMargin(0.0)
            fmt.setBottomMargin(0.0)
        self._modify_selected_blocks_format(_clear)

    def add_space_before(self, pt: float = 6.0):
        """Thêm khoảng cách phía trên đoạn văn (Add Space Before Paragraph)"""
        px = pt * 1.333
        self._modify_selected_blocks_format(lambda fmt: fmt.setTopMargin(px))

    def add_space_after(self, pt: float = 6.0):
        """Thêm khoảng cách phía dưới đoạn văn (Add Space After Paragraph)"""
        px = pt * 1.333
        self._modify_selected_blocks_format(lambda fmt: fmt.setBottomMargin(px))

    def apply_nd30_spacing(self):
        """Chuẩn hóa cấu hình giãn dòng & đoạn theo Nghị định 30 (Dòng 1.2, Trước 0pt, Sau 6pt, Thụt 1.0cm)"""
        def _nd30(fmt):
            fmt.setLineHeight(120, QTextBlockFormat.LineHeightTypes.ProportionalHeight.value)
            fmt.setTopMargin(0.0)
            fmt.setBottomMargin(6.0 * 1.333)
            fmt.setTextIndent(int(10.0 * 3.7795))
            fmt.setAlignment(Qt.AlignmentFlag.AlignJustify)
        self._modify_selected_blocks_format(_nd30)

    def insert_symbol(self, sym: str):
        ed = self.active_editor
        if ed:
            cursor = ed.textCursor()
            cursor.insertText(sym)

    def apply_list_format(self, config: dict = None):
        """Áp dụng định dạng danh sách số / ký hiệu / phân cấp đa cấp thông minh chuẩn Microsoft Word"""
        ed = self.active_editor
        if not ed:
            return
        if not config:
            config = {"style": "decimal", "suffix": ".", "restart": False}

        style_key = config.get("style", "decimal")
        suffix = config.get("suffix", ".")
        prefix = config.get("prefix", "")
        restart = config.get("restart", False)
        indent = config.get("indent", 1)

        style_map = {
            "decimal": QTextListFormat.Style.ListDecimal,
            "roman_upper": QTextListFormat.Style.ListUpperRoman,
            "roman_lower": QTextListFormat.Style.ListLowerRoman,
            "alpha_lower": QTextListFormat.Style.ListLowerAlpha,
            "alpha_upper": QTextListFormat.Style.ListUpperAlpha,
            "disc": QTextListFormat.Style.ListDisc,
            "circle": QTextListFormat.Style.ListCircle,
            "square": QTextListFormat.Style.ListSquare,
        }

        if style_key == "multilevel":
            qt_style = QTextListFormat.Style.ListUpperRoman
            suffix = "."
            prefix = ""
            indent = 1
        else:
            qt_style = style_map.get(style_key, QTextListFormat.Style.ListDecimal)

        fmt = QTextListFormat()
        fmt.setStyle(qt_style)
        fmt.setIndent(indent)
        if hasattr(fmt, "setNumberSuffix"):
            fmt.setNumberSuffix(suffix)
        if hasattr(fmt, "setNumberPrefix") and prefix:
            fmt.setNumberPrefix(prefix)

        cursor = ed.textCursor()

        if not cursor.hasSelection():
            cur_block = cursor.block()
            tl = cur_block.textList()
            # Nếu đoạn đã ở danh sách cùng kiểu và cùng hậu tố -> Bấm lần nữa là Hủy danh sách (Toggle OFF)
            if tl and tl.format().style() == qt_style and getattr(tl.format(), "numberSuffix", lambda: "")() == suffix and not restart:
                tl.remove(cur_block)
                bf = cursor.blockFormat()
                bf.setObjectIndex(-1)
                bf.setIndent(0)
                cursor.setBlockFormat(bf)
                self.document_modified.emit()
                return

            # Tìm danh sách trước đó phù hợp để nối tiếp số thứ tự (tránh lỗi bấm từng đoạn toàn ra số 1)
            target_list = None
            if not restart:
                b = cur_block.previous()
                while b.isValid():
                    btl = b.textList()
                    if btl:
                        if btl.format().style() == qt_style and btl.format().indent() == indent:
                            target_list = btl
                        break
                    elif b.text().strip():
                        # Gặp đoạn văn bản bình thường không rỗng -> Dừng nối
                        break
                    b = b.previous()

            if target_list:
                target_list.add(cur_block)
            else:
                cursor.createList(fmt)
        else:
            # Có vùng chọn nhiều đoạn văn bản
            start_pos = cursor.selectionStart()
            end_pos = cursor.selectionEnd()
            doc = ed.document()
            start_block = doc.findBlock(start_pos)
            end_block = doc.findBlock(end_pos)

            # Kiểm tra xem danh sách trước khối chọn có thể nối không
            target_list = None
            if not restart:
                b = start_block.previous()
                while b.isValid():
                    btl = b.textList()
                    if btl:
                        if btl.format().style() == qt_style and btl.format().indent() == indent:
                            target_list = btl
                        break
                    elif b.text().strip():
                        break
                    b = b.previous()

            if target_list:
                cur = start_block
                while cur.isValid() and cur.position() <= end_block.position():
                    target_list.add(cur)
                    if cur == end_block:
                        break
                    cur = cur.next()
            else:
                cursor.createList(fmt)

        ed.setTextCursor(cursor)
        self.document_modified.emit()

    def insert_bullet_list(self):
        """Tạo/bật tắt danh sách gạch đầu dòng tự động (- / •)"""
        self.apply_list_format({"style": "disc", "restart": False})

    def insert_numbered_list(self):
        """Tạo/bật tắt danh sách đánh số tự động chuẩn Word (1, 2, 3... - Tự nối số tiếp theo)"""
        self.apply_list_format({"style": "decimal", "suffix": ".", "restart": False})

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

            try:
                import base64
                with open(file_path, "rb") as f_img:
                    b64_str = base64.b64encode(f_img.read()).decode("ascii")
                ext = os.path.splitext(file_path)[1].lower().replace(".", "")
                mime = "image/jpeg" if ext in ["jpg", "jpeg"] else "image/png"
                data_uri = f"data:{mime};base64,{b64_str}"

                img_format = QTextImageFormat()
                img_format.setName(data_uri)
                img_format.setWidth(img.width())
                img_format.setHeight(img.height())

                cursor.insertBlock()
                block_fmt = cursor.blockFormat()
                block_fmt.setAlignment(Qt.AlignmentFlag.AlignCenter)
                block_fmt.setTextIndent(0)
                cursor.setBlockFormat(block_fmt)
                cursor.insertImage(img_format)
            except Exception:
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
        """Hành động ngắt trang: Tách nội dung sang trang mới chuẩn xác như Word"""
        self.split_page_at_cursor(self._active_page_idx)

    def split_page_at_cursor(self, page_idx: int):
        """Ngắt trang chuẩn xác như Microsoft Word:
        - Tách toàn bộ đoạn văn từ vị trí con trỏ hiện tại sang trang tiếp theo
        - Tạo trang A4 mới với khung viền và đổ bóng rõ ràng
        - Tuyệt đối không chèn chữ thô làm hỏng bố cục văn bản
        """
        if not (0 <= page_idx < len(self.pages)):
            self.add_page(orientation="portrait")
            return

        cur_page = self.pages[page_idx]
        cur = cur_page.editor.textCursor()

        # Chọn từ vị trí con trỏ đến hết văn bản của trang hiện tại
        cur.movePosition(QTextCursor.MoveOperation.End, QTextCursor.MoveMode.KeepAnchor)
        trailing_text = cur.selectedText()
        trailing_html = cur.selection().toHtml() if trailing_text.strip() else ""
        cur.removeSelectedText()

        # Thêm trang A4 mới
        new_page = self.add_page(orientation="portrait")

        # Nếu có đoạn văn phía sau con trỏ, chuyển sang trang mới
        if trailing_text.strip():
            new_page.editor.setHtml(trailing_html)
            c = new_page.editor.textCursor()
            c.movePosition(QTextCursor.MoveOperation.Start)
            new_page.editor.setTextCursor(c)
        else:
            new_page.editor.setFocus()

        self.document_modified.emit()

    def _on_page_overflow(self, page_idx: int):
        """Tự động nhảy sang trang mới khi nội dung hết trang (Auto-pagination)"""
        if not (0 <= page_idx < len(self.pages)):
            return
        cur_page = self.pages[page_idx]
        cur = cur_page.editor.textCursor()

        # Nếu con trỏ đang ở cuối trang thì tự động nhảy tiếp sang trang mới
        if cur.atEnd():
            if page_idx + 1 >= len(self.pages):
                new_page = self.add_page(orientation="portrait")
                new_page.editor.setFocus()
            else:
                self.pages[page_idx + 1].editor.setFocus()

    def _on_backspace_at_start(self, page_idx: int):
        """Khi bấm Backspace ở đầu trang rỗng: xóa trang thừa và nhảy về cuối trang trước"""
        if page_idx <= 0 or page_idx >= len(self.pages):
            return
        cur_page = self.pages[page_idx]
        prev_page = self.pages[page_idx - 1]

        # Nếu trang hiện tại hoàn toàn rỗng, xóa trang này
        if not cur_page.editor.toPlainText().strip():
            self.remove_page(page_idx)
            prev_cur = prev_page.editor.textCursor()
            prev_cur.movePosition(QTextCursor.MoveOperation.End)
            prev_page.editor.setTextCursor(prev_cur)
            prev_page.editor.setFocus()

    def get_body_paragraphs(self) -> List[str]:
        """Trích xuất danh sách các đoạn văn bản và hình ảnh trong toàn bộ các trang để chuyển giao sang OpenXmlBuilder"""
        paras = []
        for idx, p in enumerate(self.pages):
            if idx > 0:
                # Đánh dấu ngắt trang và định hướng xoay cho OpenXmlBuilder
                if p.orientation == "landscape":
                    paras.append("[PAGE_BREAK_LANDSCAPE]")
                else:
                    paras.append("[PAGE_BREAK]")

            doc = p.editor.document()
            for b_idx in range(doc.blockCount()):
                b = doc.findBlockByNumber(b_idx)

                # Kiểm tra hình ảnh trong block
                it = b.begin()
                img_name = None
                while not it.atEnd():
                    frag = it.fragment()
                    if frag.isValid() and frag.charFormat().isImageFormat():
                        img_name = frag.charFormat().toImageFormat().name()
                        break
                    it += 1

                txt = b.text().strip()
                if img_name:
                    paras.append(f"[IMAGE_DATA:{img_name}]")
                    clean_txt = txt.replace("\ufffc", "").strip()
                    if clean_txt:
                        paras.append(clean_txt)
                elif txt:
                    if any(k in txt.upper() for k in ["[PAGE_BREAK]", "PAGE BREAK", "NGẮT TRANG"]):
                        continue
                    paras.append(txt)
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
