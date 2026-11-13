"""
ui/editor/metadata_editor_dlg.py: Hộp thoại Chỉnh sửa Thể thức Văn bản
Giao diện đồ họa hiện đại chuẩn công vụ (Modern Administrative UI):
- Thanh chọn quy chuẩn Segmented Pill Switch (NĐ 30 vs HD 36) cố định đỉnh hộp thoại
- Danh mục Loại văn bản (QComboBox) dùng QListView nền trắng, sửa triệt để lỗi dropdown bị nền đen
- Bộ chọn Thời gian ban hành thông minh: Hỗ trợ Lịch (Calendar Popup), nút Hôm nay, chế độ Dự thảo chờ ký
- Tự động sinh chuỗi thời gian chuẩn Nghị định 30 (thêm số 0 cho ngày < 10, tháng 1, 2)
- Khắc phục lỗi nhập ngày nhảy số (1020), loại bỏ hoàn toàn thanh cuộn ngang
"""

import re
from typing import Dict, Any, List
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QLineEdit,
    QPlainTextEdit, QComboBox, QRadioButton, QButtonGroup,
    QPushButton, QFrame, QWidget, QScrollArea, QDateEdit, QListView
)
from PyQt6.QtCore import Qt, QDate, QEvent, QObject


class DatePickerClickFilter(QObject):
    """Bắt sự kiện click hoặc focus vào DatePicker để tự động chuyển sang chế độ Ngày cụ thể"""
    def __init__(self, callback, parent=None):
        super().__init__(parent)
        self.callback = callback

    def eventFilter(self, obj, ev):
        if ev.type() in (QEvent.Type.MouseButtonPress, QEvent.Type.FocusIn):
            self.callback()
        return False


class MetadataEditorDialog(QDialog):
    """Hộp thoại chỉnh sửa thông tin thể thức hành chính & Đảng"""

    def __init__(self, parent=None, metadata: Dict[str, Any] = None):
        super().__init__(parent)
        self.setWindowTitle("Chỉnh sửa Thể thức Văn bản")
        self.resize(880, 760)
        self.setMinimumSize(820, 640)
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
                color: #0f172a;
                font-family: 'Segoe UI', -apple-system, Arial, sans-serif;
            }
            QLabel {
                color: #334155;
                font-size: 12px;
                font-weight: 600;
                background: transparent;
            }
            QLineEdit, QPlainTextEdit, QDateEdit {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 6px 10px;
                font-size: 13px;
                min-height: 24px;
            }
            QLineEdit:focus, QPlainTextEdit:focus, QDateEdit:focus {
                border: 1.5px solid #2563eb;
                background-color: #ffffff;
            }
            QComboBox {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 5px 10px;
                font-size: 13px;
                min-height: 26px;
            }
            QComboBox:hover {
                border-color: #94a3b8;
            }
            QComboBox:focus {
                border: 1.5px solid #2563eb;
            }
            QComboBox::drop-down {
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 24px;
                border-left: none;
            }
            QComboBox QAbstractItemView {
                background-color: #ffffff;
                color: #0f172a;
                selection-background-color: #eff6ff;
                selection-color: #2563eb;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px;
                outline: none;
            }
            QFrame#stdSwitchBox {
                background-color: #f1f5f9;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
                padding: 3px;
            }
            QFrame#stdSwitchBox QRadioButton {
                color: #475569;
                font-size: 13px;
                font-weight: 600;
                padding: 7px 18px;
                background-color: transparent;
                border: 1px solid transparent;
                border-radius: 6px;
                text-align: center;
            }
            QFrame#stdSwitchBox QRadioButton::indicator {
                width: 0px;
                height: 0px;
            }
            QFrame#stdSwitchBox QRadioButton:hover {
                background-color: #e2e8f0;
                color: #0f172a;
            }
            QFrame#stdSwitchBox QRadioButton:checked {
                background-color: #ffffff;
                border: 1px solid #cbd5e1;
                color: #2563eb;
                font-weight: 700;
            }
            /* Radio button tròn chuẩn mực, rõ ràng bên trong phần chọn ngày tháng */
            QRadioButton#rbDateSpecific, QRadioButton#rbDateDraft {
                color: #1e293b;
                font-size: 13px;
                font-weight: 600;
                background-color: transparent;
                border: none;
                padding: 4px 6px;
                spacing: 8px;
            }
            QRadioButton#rbDateSpecific::indicator, QRadioButton#rbDateDraft::indicator {
                width: 16px;
                height: 16px;
                border-radius: 8px;
                border: 2px solid #94a3b8;
                background-color: #ffffff;
            }
            QRadioButton#rbDateSpecific::indicator:hover, QRadioButton#rbDateDraft::indicator:hover {
                border-color: #2563eb;
            }
            QRadioButton#rbDateSpecific::indicator:checked, QRadioButton#rbDateDraft::indicator:checked {
                border: 5px solid #2563eb;
                background-color: #ffffff;
            }
            QPushButton {
                background-color: #ffffff;
                color: #334155;
                font-weight: 600;
                font-size: 13px;
                padding: 8px 22px;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
                border-color: #94a3b8;
            }
            QPushButton#btnSmallAction {
                padding: 5px 12px;
                font-size: 12px;
                background-color: #eff6ff;
                color: #1e40af;
                border: 1px solid #bfdbfe;
                border-radius: 5px;
            }
            QPushButton#btnSmallAction:hover {
                background-color: #dbeafe;
                border-color: #3b82f6;
            }
            QPushButton#btnPrimary {
                background-color: #2563eb;
                color: #ffffff;
                font-weight: 600;
                border: 1px solid #1d4ed8;
            }
            QPushButton#btnPrimary:hover {
                background-color: #1d4ed8;
            }
            QScrollArea {
                border: none;
                background: transparent;
            }
            QScrollBar:vertical {
                width: 6px;
                background: transparent;
                margin: 0px;
            }
            QScrollBar::handle:vertical {
                background: #cbd5e1;
                border-radius: 3px;
                min-height: 24px;
            }
            QScrollBar::handle:vertical:hover {
                background: #94a3b8;
            }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
                height: 0px;
            }
            QScrollBar::add-page:vertical, QScrollBar::sub-page:vertical {
                background: none;
            }
            QCalendarWidget QWidget#qt_calendar_navigationbar {
                background-color: #f8fafc;
                border-bottom: 1px solid #e2e8f0;
            }
            QCalendarWidget QToolButton {
                color: #1e293b;
                font-weight: 600;
                font-size: 12px;
                border: none;
                border-radius: 4px;
                padding: 4px 8px;
            }
            QCalendarWidget QToolButton:hover {
                background-color: #e2e8f0;
            }
            QCalendarWidget QAbstractItemView:enabled {
                background-color: #ffffff;
                color: #0f172a;
                selection-background-color: #2563eb;
                selection-color: #ffffff;
                font-size: 12px;
            }
        """)

        self.meta = dict(metadata or {})

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(22, 16, 22, 16)
        main_layout.setSpacing(10)

        # 0. Tiêu đề hộp thoại tinh gọn
        header_box = QVBoxLayout()
        header_box.setSpacing(3)
        lbl_title = QLabel("Chỉnh sửa Thể thức Văn bản")
        lbl_title.setStyleSheet("font-size: 16px; font-weight: 700; color: #0f172a;")
        lbl_subtitle = QLabel("Cập nhật thông tin Quốc hiệu, Cơ quan ban hành, Số hiệu, Trích yếu và Khối chữ ký")
        lbl_subtitle.setStyleSheet("font-size: 12px; font-weight: 400; color: #64748b;")
        header_box.addWidget(lbl_title)
        header_box.addWidget(lbl_subtitle)
        main_layout.addLayout(header_box)

        # 1. Phân loại Quy chuẩn áp dụng cố định ở đỉnh (Segmented Pill Switch)
        switch_frame = QFrame()
        switch_frame.setObjectName("stdSwitchBox")
        row_std = QHBoxLayout(switch_frame)
        row_std.setContentsMargins(4, 4, 4, 4)
        row_std.setSpacing(6)

        self.rb_nd30 = QRadioButton("Hành chính Nhà nước (Nghị định 30/2020)")
        self.rb_hd36 = QRadioButton("Cơ quan Đảng (Hướng dẫn 36-HD/VPTW)")
        std_id = self.meta.get("standard_id", "ND30").upper()
        if std_id == "HD36":
            self.rb_hd36.setChecked(True)
        else:
            self.rb_nd30.setChecked(True)

        self.grp_std = QButtonGroup(self)
        self.grp_std.addButton(self.rb_nd30, 0)
        self.grp_std.addButton(self.rb_hd36, 1)
        row_std.addWidget(self.rb_nd30, 1)
        row_std.addWidget(self.rb_hd36, 1)
        main_layout.addWidget(switch_frame)

        # Dòng hướng dẫn quy chuẩn động (Bật setWordWrap để không tràn chiều ngang)
        self.lbl_std_hint = QLabel()
        self.lbl_std_hint.setWordWrap(True)
        main_layout.addWidget(self.lbl_std_hint)

        # 2. Vùng cuộn QScrollArea bọc 3 thẻ nhóm thông tin (TẮT HẲN CUỘN NGANG)
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)

        self.scroll_widget = QWidget()
        self.scroll_widget.setStyleSheet("background: transparent;")
        scroll_layout = QVBoxLayout(self.scroll_widget)
        scroll_layout.setContentsMargins(0, 4, 4, 4)
        scroll_layout.setSpacing(12)

        # ---------- THẺ 1: CƠ QUAN BAN HÀNH & SỐ HIỆU VĂN BẢN ----------
        card_agency = self._create_card("1. Cơ quan ban hành & Số hiệu văn bản")
        g_ag = QGridLayout()
        g_ag.setSpacing(10)

        col1 = QVBoxLayout()
        col1.setSpacing(3)
        col1.addWidget(QLabel("Cơ quan cấp trên:"))
        self.txt_sup = QLineEdit(self.meta.get("superior_agency", ""))
        col1.addWidget(self.txt_sup)
        g_ag.addLayout(col1, 0, 0)

        col2 = QVBoxLayout()
        col2.setSpacing(3)
        col2.addWidget(QLabel("Cơ quan ban hành:"))
        self.txt_iss = QLineEdit(self.meta.get("issuing_agency", ""))
        col2.addWidget(self.txt_iss)
        g_ag.addLayout(col2, 0, 1)

        col3 = QVBoxLayout()
        col3.setSpacing(3)
        col3.addWidget(QLabel("Số và ký hiệu:"))
        self.txt_num = QLineEdit(self.meta.get("doc_number", ""))
        col3.addWidget(self.txt_num)
        g_ag.addLayout(col3, 1, 0)

        col4 = QVBoxLayout()
        col4.setSpacing(3)
        col4.addWidget(QLabel("Loại văn bản:"))
        self.cb_type = QComboBox()

        # Dùng QListView chuyên dụng với nền trắng sáng, loại bỏ hoàn toàn lỗi menu màu đen của Windows
        cb_list_view = QListView(self.cb_type)
        cb_list_view.setStyleSheet("""
            QListView {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px;
                outline: none;
            }
            QListView::item {
                background-color: #ffffff;
                color: #0f172a;
                min-height: 28px;
                padding: 4px 10px;
                border-radius: 4px;
            }
            QListView::item:hover {
                background-color: #f1f5f9;
                color: #0f172a;
            }
            QListView::item:selected {
                background-color: #eff6ff;
                color: #2563eb;
                font-weight: 600;
            }
        """)
        self.cb_type.setView(cb_list_view)
        col4.addWidget(self.cb_type)
        g_ag.addLayout(col4, 1, 1)

        card_agency.layout().addLayout(g_ag)
        scroll_layout.addWidget(card_agency)

        # ---------- THẺ 2: ĐỊA DANH, THỜI GIAN & TRÍCH YẾU (BỘ CHỌN LỊCH THÔNG MINH) ----------
        card_content = self._create_card("2. Địa danh, Thời gian & Trích yếu nội dung")
        l_cnt = QVBoxLayout()
        l_cnt.setSpacing(10)

        row_loc_date = QHBoxLayout()
        row_loc_date.setSpacing(14)

        # Cột địa danh
        c_loc = QVBoxLayout()
        c_loc.setSpacing(4)
        c_loc.addWidget(QLabel("Địa danh ban hành:"))
        from core.administrative_sanitizer import AdministrativeSanitizer
        init_loc = self.meta.get("location", "").strip()
        if not init_loc or init_loc in ["Xã", "Huyện", "Tỉnh", "...", "…"]:
            init_loc = AdministrativeSanitizer.infer_location_from_agencies(
                self.meta.get("issuing_agency", ""),
                self.meta.get("superior_agency", "")
            )
        self.txt_loc = QLineEdit(init_loc)
        self.txt_loc.setPlaceholderText("Ví dụ: Thuận Hạnh, Quảng Ninh, Hà Nội...")
        c_loc.addWidget(self.txt_loc)
        row_loc_date.addLayout(c_loc, 1)

        # Cột thời gian với Lịch Calendar popup thông minh
        c_date = QVBoxLayout()
        c_date.setSpacing(4)
        c_date.addWidget(QLabel("Thời gian ban hành văn bản:"))

        date_ctrl_row = QHBoxLayout()
        date_ctrl_row.setSpacing(8)

        # Tùy chọn Ngày cụ thể
        self.rb_date_specific = QRadioButton("Ngày cụ thể:")
        self.rb_date_specific.setObjectName("rbDateSpecific")
        date_ctrl_row.addWidget(self.rb_date_specific)

        # Ô chọn ngày với Lịch Popup Calendar bọc trong container bắt click khi disabled
        self.date_picker = QDateEdit()
        self.date_picker.setCalendarPopup(True)
        self.date_picker.setDisplayFormat("dd/MM/yyyy")
        self.date_picker.setFixedWidth(125)

        self.date_picker_container = QWidget()
        l_dp = QHBoxLayout(self.date_picker_container)
        l_dp.setContentsMargins(0, 0, 0, 0)
        l_dp.addWidget(self.date_picker)
        self._date_click_filter = DatePickerClickFilter(self._on_date_picker_interacted, self)
        self.date_picker_container.installEventFilter(self._date_click_filter)
        self.date_picker.installEventFilter(self._date_click_filter)
        if self.date_picker.lineEdit():
            self.date_picker.lineEdit().installEventFilter(self._date_click_filter)
        date_ctrl_row.addWidget(self.date_picker_container)

        # Nút lấy nhanh ngày hôm nay
        self.btn_today = QPushButton("Hôm nay")
        self.btn_today.setObjectName("btnSmallAction")
        self.btn_today.setToolTip("Lấy nhanh ngày tháng năm hiện tại")
        self.btn_today.clicked.connect(self._on_today_clicked)
        date_ctrl_row.addWidget(self.btn_today)

        date_ctrl_row.addSpacing(14)

        # Tùy chọn Chờ ký (để trống ... chờ ký sau)
        self.rb_date_draft = QRadioButton("Chờ ký (..., ngày ...)")
        self.rb_date_draft.setObjectName("rbDateDraft")
        self.rb_date_draft.setToolTip("Dành cho văn bản dự thảo chưa có ngày ký chính thức")
        date_ctrl_row.addWidget(self.rb_date_draft)

        date_ctrl_row.addStretch()
        c_date.addLayout(date_ctrl_row)

        self.grp_date_mode = QButtonGroup(self)
        self.grp_date_mode.addButton(self.rb_date_specific, 0)
        self.grp_date_mode.addButton(self.rb_date_draft, 1)

        # Nhãn xem trước chuỗi thời gian thực tế sẽ hiển thị trên văn bản
        self.lbl_date_preview = QLabel()
        self.lbl_date_preview.setStyleSheet("font-size: 11.5px; color: #1e40af; font-weight: 500; margin-top: 2px;")
        c_date.addWidget(self.lbl_date_preview)

        row_loc_date.addLayout(c_date, 2)
        l_cnt.addLayout(row_loc_date)

        # Trích yếu nội dung
        c_subj = QVBoxLayout()
        c_subj.setSpacing(4)
        c_subj.addWidget(QLabel("Trích yếu nội dung (Về việc...):"))
        self.txt_subject = QLineEdit(self.meta.get("doc_subject", ""))
        self.txt_subject.setPlaceholderText("Ví dụ: Về việc ban hành Kế hoạch phát triển kinh tế - xã hội...")
        c_subj.addWidget(self.txt_subject)
        l_cnt.addLayout(c_subj)

        card_content.layout().addLayout(l_cnt)
        scroll_layout.addWidget(card_content)

        # Thuộc tính txt_date ngầm đồng bộ để tương thích với toàn bộ mã nguồn cũ
        self.txt_date = QLineEdit()
        self.txt_date.setVisible(False)

        # ---------- THẺ 3: NƠI NHẬN & KHỐI CHỮ KÝ BAN HÀNH ----------
        card_sign = self._create_card("3. Nơi nhận & Khối chữ ký ban hành")
        row_bottom = QHBoxLayout()
        row_bottom.setSpacing(14)

        # Cột trái: Nơi nhận
        col_rec = QVBoxLayout()
        col_rec.setSpacing(4)
        col_rec.addWidget(QLabel("Nơi nhận (mỗi dòng một nơi nhận):"))
        self.txt_recipients = QPlainTextEdit()
        recs = self.meta.get("recipients", [])
        if isinstance(recs, list):
            self.txt_recipients.setPlainText("\n".join(recs))
        else:
            self.txt_recipients.setPlainText(str(recs))
        self.txt_recipients.setMinimumHeight(80)
        col_rec.addWidget(self.txt_recipients)
        row_bottom.addLayout(col_rec, 1)

        # Cột phải: Khối chữ ký
        col_signer = QVBoxLayout()
        col_signer.setSpacing(6)

        c_auth = QVBoxLayout()
        c_auth.setSpacing(2)
        c_auth.addWidget(QLabel("Quyền hạn ký:"))
        self.txt_auth = QLineEdit(self.meta.get("authority_prefix", ""))
        c_auth.addWidget(self.txt_auth)
        col_signer.addLayout(c_auth)

        row_pn = QHBoxLayout()
        row_pn.setSpacing(10)
        c_pos = QVBoxLayout()
        c_pos.setSpacing(2)
        c_pos.addWidget(QLabel("Chức vụ người ký:"))
        self.txt_pos = QLineEdit(self.meta.get("signer_position", ""))
        c_pos.addWidget(self.txt_pos)
        row_pn.addLayout(c_pos, 1)

        c_name = QVBoxLayout()
        c_name.setSpacing(2)
        c_name.addWidget(QLabel("Họ tên người ký:"))
        self.txt_name = QLineEdit(self.meta.get("signer_name", ""))
        c_name.addWidget(self.txt_name)
        row_pn.addLayout(c_name, 1)
        col_signer.addLayout(row_pn)

        row_bottom.addLayout(col_signer, 1)
        card_sign.layout().addLayout(row_bottom)
        scroll_layout.addWidget(card_sign)

        self.scroll_area.setWidget(self.scroll_widget)
        main_layout.addWidget(self.scroll_area, 1)

        # Kết nối sự kiện thay đổi Quy chuẩn (NĐ 30 vs HD 36)
        self.rb_nd30.toggled.connect(self._on_standard_toggled)
        self.rb_hd36.toggled.connect(self._on_standard_toggled)
        # Kích hoạt cập nhật lần đầu
        self._on_standard_toggled()

        # Nạp lại giá trị doc_type ban đầu nếu có
        curr_type = self.meta.get("doc_type", "QUYẾT ĐỊNH").upper()
        if curr_type:
            idx = self.cb_type.findText(curr_type)
            if idx >= 0:
                self.cb_type.setCurrentIndex(idx)
            else:
                self.cb_type.addItem(curr_type)
                self.cb_type.setCurrentText(curr_type)

        # Khởi tạo ngày tháng và kết nối bộ bắt sự kiện cập nhật thời gian thực
        self._init_date_controls(self.meta.get("date_str", ""))
        self.txt_loc.textChanged.connect(self._update_date_preview)
        self.date_picker.dateChanged.connect(self._on_date_picker_interacted)
        self.date_picker.dateChanged.connect(self._update_date_preview)
        self.rb_date_specific.toggled.connect(self._on_date_mode_toggled)
        self.rb_date_draft.toggled.connect(self._on_date_mode_toggled)

        # Thanh nút bấm dưới cùng (luôn ghim cố định ở đáy hộp thoại)
        line_div = QFrame()
        line_div.setFrameShape(QFrame.Shape.HLine)
        line_div.setFrameShadow(QFrame.Shadow.Sunken)
        line_div.setStyleSheet("color: #e2e8f0; margin-top: 4px;")
        main_layout.addWidget(line_div)

        btn_box = QHBoxLayout()
        btn_box.addStretch()

        self.btn_cancel = QPushButton("Hủy bỏ")
        self.btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("Cập nhật thay đổi")
        self.btn_save.setObjectName("btnPrimary")
        self.btn_save.setDefault(True)
        self.btn_save.clicked.connect(self._on_save_clicked)
        btn_box.addWidget(self.btn_save)

        main_layout.addLayout(btn_box)

    def _create_card(self, title: str) -> QFrame:
        """Tạo khung chứa dạng thẻ phẳng hiện đại"""
        card = QFrame()
        card.setObjectName("metaCard")
        card.setStyleSheet("""
            QFrame#metaCard {
                background-color: #ffffff;
                border: 1px solid #e2e8f0;
                border-radius: 8px;
            }
        """)
        l = QVBoxLayout(card)
        l.setContentsMargins(16, 12, 16, 14)
        l.setSpacing(8)

        lbl = QLabel(title)
        lbl.setStyleSheet("font-size: 13px; font-weight: 700; color: #1e293b; border: none; background: transparent;")
        l.addWidget(lbl)
        return card

    def _update_combo_types(self, type_list: List[str], default: str = ""):
        cur = self.cb_type.currentText().strip().upper()
        self.cb_type.clear()
        self.cb_type.addItems(type_list)
        if cur and cur in type_list:
            self.cb_type.setCurrentText(cur)
        elif default and default in type_list:
            self.cb_type.setCurrentText(default)

    def _init_date_controls(self, date_str: str):
        """Khởi tạo giá trị cho DatePicker từ chuỗi date_str hiện có"""
        today = QDate.currentDate()
        is_draft = False
        parsed_date = today

        if date_str:
            if "..." in date_str or ".." in date_str:
                is_draft = True
                m_y = re.search(r'năm\s+(\d{4})', date_str)
                if m_y:
                    try:
                        parsed_date = QDate(int(m_y.group(1)), today.month(), today.day())
                    except Exception:
                        pass
            else:
                m = re.search(r'ngày\s+(\d{1,2})\s+tháng\s+(\d{1,2})\s+năm\s+(\d{4})', date_str, re.IGNORECASE)
                if m:
                    try:
                        parsed_date = QDate(int(m.group(3)), int(m.group(2)), int(m.group(1)))
                    except Exception:
                        pass
                else:
                    m2 = re.search(r'(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})', date_str)
                    if m2:
                        try:
                            parsed_date = QDate(int(m2.group(3)), int(m2.group(2)), int(m2.group(1)))
                        except Exception:
                            pass

        if not parsed_date.isValid():
            parsed_date = today

        self.date_picker.setDate(parsed_date)
        if is_draft:
            self.rb_date_draft.setChecked(True)
        else:
            self.rb_date_specific.setChecked(True)

        self._on_date_mode_toggled()

    def _on_date_picker_interacted(self):
        """Tự động chuyển về Ngày cụ thể khi người dùng click, focus hoặc đổi ngày trên bộ chọn"""
        if not self.rb_date_specific.isChecked():
            self.rb_date_specific.setChecked(True)

    def _on_today_clicked(self):
        """Lấy ngày hôm nay và tự động kích hoạt Ngày cụ thể"""
        self.rb_date_specific.setChecked(True)
        self.date_picker.setDate(QDate.currentDate())
        self._on_date_mode_toggled()

    def _on_date_mode_toggled(self):
        """Chuyển đổi giữa chế độ ngày cụ thể và chờ ký"""
        is_specific = self.rb_date_specific.isChecked()
        self.date_picker.setEnabled(is_specific)
        if is_specific:
            self.date_picker.setStyleSheet("background-color: #ffffff; color: #0f172a; border: 1px solid #cbd5e1;")
        else:
            self.date_picker.setStyleSheet("background-color: #f8fafc; color: #94a3b8; border: 1px solid #e2e8f0;")
        # Nút Hôm nay luôn sẵn sàng để người dùng có thể quay lại bất cứ lúc nào
        self.btn_today.setEnabled(True)
        self._update_date_preview()

    def _get_current_formatted_date(self) -> str:
        """Tạo chuỗi thời gian chuẩn thể thức NĐ 30 & HD 36"""
        loc = self.txt_loc.text().strip() or "Hà Nội"
        if self.rb_date_draft.isChecked():
            year = self.date_picker.date().year()
            return f"{loc}, ngày ... tháng ... năm {year}"
        else:
            d = self.date_picker.date()
            # Theo NĐ 30: ngày dưới 10 và tháng 1, 2 phải ghi thêm số 0 ở trước
            return f"{loc}, ngày {d.day():02d} tháng {d.month():02d} năm {d.year()}"

    def _update_date_preview(self):
        """Cập nhật dòng xem trước ngày tháng hiển thị trên văn bản"""
        date_str = self._get_current_formatted_date()
        self.lbl_date_preview.setText(f"Hiển thị trên văn bản: <b>{date_str}</b>")
        self.txt_date.setText(date_str)

    def _on_standard_toggled(self):
        """Thay đổi linh hoạt giao diện và dữ liệu mẫu khi người dùng bấm chuyển NĐ 30 / HD 36"""
        is_party = self.rb_hd36.isChecked()
        if is_party:
            self.lbl_std_hint.setText("[Quy chuẩn Đảng - Hướng dẫn 36-HD/VPTW]: Tiêu đề 'ĐẢNG CỘNG SẢN VIỆT NAM', số hiệu dạng 'Số 01-QĐ/ĐU' (không có dấu hai chấm, dùng gạch nối '-').")
            self.lbl_std_hint.setStyleSheet("""
                font-size: 12px;
                font-weight: 500;
                line-height: 1.4;
                color: #92400e;
                background-color: #fef3c7;
                border: 1px solid #fde68a;
                border-radius: 6px;
                padding: 8px 14px;
            """)

            self.txt_sup.setPlaceholderText("Ví dụ: TỈNH ỦY... hoặc HUYỆN ỦY...")
            self.txt_iss.setPlaceholderText("Ví dụ: ĐẢNG BỘ XÃ... hoặc BAN THƯỜNG VỤ...")
            self.txt_num.setPlaceholderText("Ví dụ: Số 01-QĐ/ĐU hoặc Số 12-NQ/TU")
            self.txt_auth.setPlaceholderText("T/M BAN THƯỜNG VỤ hoặc T/M ĐẢNG ỦY")
            self.txt_pos.setPlaceholderText("BÍ THƯ hoặc PHÓ BÍ THƯ")
            self.txt_recipients.setPlaceholderText("- Ban Thường vụ Huyện ủy (để b/c);\n- Các chi bộ trực thuộc;\n- Lưu: VP Đảng ủy.")

            party_types = ["NGHỊ QUYẾT ĐẢNG", "QUYẾT ĐỊNH ĐẢNG", "CHỈ THỊ ĐẢNG", "KẾT LUẬN", "KẾ HOẠCH ĐẢNG", "THÔNG BÁO ĐẢNG", "TỜ TRÌNH", "BÁO CÁO", "BIÊN BẢN"]
            self._update_combo_types(party_types, default="QUYẾT ĐỊNH ĐẢNG")

            # Tự động gợi ý chuyển đổi nếu trường đang mang mẫu mặc định của Nhà nước
            sup_cur = self.txt_sup.text().strip().upper()
            if sup_cur in ["ỦY BAN NHÂN DÂN HUYỆN", "UBND HUYỆN", ""]:
                self.txt_sup.setText("HUYỆN ỦY")
            iss_cur = self.txt_iss.text().strip().upper()
            if iss_cur in ["ỦY BAN NHÂN DÂN XÃ", "UBND XÃ", ""]:
                self.txt_iss.setText("ĐẢNG ỦY XÃ")
            auth_cur = self.txt_auth.text().strip().upper()
            if auth_cur in ["TM. ỦY BAN NHÂN DÂN", "TM. UBND", ""]:
                self.txt_auth.setText("T/M ĐẢNG ỦY")
            pos_cur = self.txt_pos.text().strip().upper()
            if pos_cur in ["CHỦ TỊCH", ""]:
                self.txt_pos.setText("BÍ THƯ")
            num_cur = self.txt_num.text().strip()
            if "QĐ-UBND" in num_cur or "/QĐ-" in num_cur or not num_cur:
                self.txt_num.setText("Số 01-QĐ/ĐU")
        else:
            self.lbl_std_hint.setText("[Quy chuẩn Nhà nước - Nghị định 30/2020/NĐ-CP]: Tiêu ngữ 'Độc lập - Tự do - Hạnh phúc', số hiệu dạng 'Số: 15/QĐ-UBND' (có dấu hai chấm, gạch chéo '/').")
            self.lbl_std_hint.setStyleSheet("""
                font-size: 12px;
                font-weight: 500;
                line-height: 1.4;
                color: #1e40af;
                background-color: #eff6ff;
                border: 1px solid #bfdbfe;
                border-radius: 6px;
                padding: 8px 14px;
            """)

            self.txt_sup.setPlaceholderText("Ví dụ: ỦY BAN NHÂN DÂN HUYỆN... hoặc TỈNH ỦY...")
            self.txt_iss.setPlaceholderText("Ví dụ: ỦY BAN NHÂN DÂN XÃ... hoặc SỞ NỘI VỤ...")
            self.txt_num.setPlaceholderText("Ví dụ: Số: 15/QĐ-UBND hoặc Số: 02/BC-SNV")
            self.txt_auth.setPlaceholderText("TM. ỦY BAN NHÂN DÂN hoặc KT. CHỦ TỊCH")
            self.txt_pos.setPlaceholderText("CHỦ TỊCH hoặc PHÓ CHỦ TỊCH")
            self.txt_recipients.setPlaceholderText("- Thường trực HĐND xã;\n- Chủ tịch, các PCT UBND xã;\n- Lưu: VT.")

            state_types = ["QUYẾT ĐỊNH", "CÔNG VĂN", "TỜ TRÌNH", "THÔNG BÁO", "KẾ HOẠCH", "BÁO CÁO", "CHỈ THỊ", "BIÊN BẢN"]
            self._update_combo_types(state_types, default="QUYẾT ĐỊNH")

            # Tự động gợi ý chuyển đổi nếu trường đang mang mẫu mặc định của Đảng
            sup_cur = self.txt_sup.text().strip().upper()
            if sup_cur in ["HUYỆN ỦY", "TỈNH ỦY", ""]:
                self.txt_sup.setText("ỦY BAN NHÂN DÂN HUYỆN")
            iss_cur = self.txt_iss.text().strip().upper()
            if iss_cur in ["ĐẢNG ỦY XÃ", "BAN THƯỜNG VỤ", ""]:
                self.txt_iss.setText("ỦY BAN NHÂN DÂN XÃ")
            auth_cur = self.txt_auth.text().strip().upper()
            if auth_cur in ["T/M ĐẢNG ỦY", "T/M BAN THƯỜNG VỤ", ""]:
                self.txt_auth.setText("TM. ỦY BAN NHÂN DÂN")
            pos_cur = self.txt_pos.text().strip().upper()
            if pos_cur in ["BÍ THƯ", ""]:
                self.txt_pos.setText("CHỦ TỊCH")
            num_cur = self.txt_num.text().strip()
            if "-QĐ/ĐU" in num_cur or "-NQ/" in num_cur or not num_cur:
                self.txt_num.setText("Số: 01/QĐ-UBND")

    def _on_save_clicked(self):
        loc = self.txt_loc.text().strip() or "Hà Nội"
        date_str = self._get_current_formatted_date()

        raw_rec = self.txt_recipients.toPlainText()
        rec_lines = [r.strip() for r in raw_rec.splitlines() if r.strip()]
        if not rec_lines:
            rec_lines = ["- Thường trực HĐND;", "- Lãnh đạo UBND;", "- Lưu: VT."]

        sup_in = self.txt_sup.text().strip().upper()
        iss_in = self.txt_iss.text().strip().upper()
        from core.administrative_sanitizer import AdministrativeSanitizer
        sup_norm, iss_norm = AdministrativeSanitizer.split_and_normalize_agency_header(sup_in, iss_in)

        self.meta.update({
            "standard_id": "HD36" if self.rb_hd36.isChecked() else "ND30",
            "superior_agency": sup_norm,
            "issuing_agency": iss_norm,
            "doc_number": self.txt_num.text().strip(),
            "location": loc,
            "date_str": date_str,
            "doc_type": self.cb_type.currentText().strip().upper(),
            "doc_subject": self.txt_subject.text().strip(),
            "recipients": rec_lines,
            "authority_prefix": self.txt_auth.text().strip().upper(),
            "signer_position": self.txt_pos.text().strip().upper(),
            "signer_name": self.txt_name.text().strip()
        })
        self.accept()

    def get_metadata(self) -> Dict[str, Any]:
        return self.meta
