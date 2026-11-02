"""
ui/editor/metadata_editor_dlg.py: Hộp thoại Chỉnh sửa Toàn diện Khung Thể thức Văn bản
Cho phép cán bộ biên tập nhanh các thông tin pháp lý bắt buộc:
- Cơ quan cấp trên, Cơ quan ban hành, Số ký hiệu
- Địa danh ban hành, Ngày tháng năm
- Thể loại văn bản, Trích yếu nội dung
- Danh sách Nơi nhận (nhiều dòng)
- Khối Thẩm quyền ký, Chức vụ và Họ tên người ký
"""

from typing import Dict, Any, List
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QComboBox, QRadioButton, QButtonGroup,
    QGroupBox, QPushButton, QFrame, QScrollArea, QWidget
)
from PyQt6.QtCore import Qt
from ui.icon_factory import get_ui_icon


class MetadataEditorDialog(QDialog):
    """Hộp thoại chỉnh sửa thông tin thể thức hành chính & Đảng"""

    def __init__(self, parent=None, metadata: Dict[str, Any] = None):
        super().__init__(parent)
        self.setWindowTitle("Chỉnh sửa Thông tin Thể thức Văn bản (NĐ 30 & HD 36)")
        self.resize(620, 680)
        self.setMinimumSize(540, 540)

        self.meta = dict(metadata or {})

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # Header hướng dẫn
        hdr = QFrame()
        hdr.setStyleSheet("background-color: #f0f9ff; border: 1px solid #bae6fd; border-radius: 6px; padding: 8px 12px;")
        hdr_layout = QVBoxLayout(hdr)
        lbl_h = QLabel("CHỈNH SỬA THÔNG TIN THỂ THỨC VĂN BẢN")
        lbl_h.setStyleSheet("color: #0369a1; font-weight: bold; font-size: 13px;")
        lbl_sub = QLabel("Thay đổi Địa danh, Cơ quan, Số hiệu, Nơi nhận hoặc Khối ký. Hệ thống sẽ tự động cập nhật lại khung thể thức trên trang soạn thảo.")
        lbl_sub.setStyleSheet("color: #0284c7; font-size: 11.5px;")
        lbl_sub.setWordWrap(True)
        hdr_layout.addWidget(lbl_h)
        hdr_layout.addWidget(lbl_sub)
        main_layout.addWidget(hdr)

        # Scroll Area chứa nội dung
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setFrameShape(QFrame.Shape.NoFrame)
        scroll.setStyleSheet("background: transparent;")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(2, 4, 8, 4)
        layout.setSpacing(12)

        # 1. Cơ quan & Số hiệu
        grp_agency = QGroupBox("1. Cơ quan ban hành & Số hiệu")
        l_ag = QVBoxLayout(grp_agency)

        row_std = QHBoxLayout()
        self.rb_nd30 = QRadioButton("Khối Hành chính Nhà nước (Nghị định 30)")
        self.rb_hd36 = QRadioButton("Khối Cơ quan Đảng (Hướng dẫn 36)")
        std_id = self.meta.get("standard_id", "ND30").upper()
        if std_id == "HD36":
            self.rb_hd36.setChecked(True)
        else:
            self.rb_nd30.setChecked(True)

        self.grp_std = QButtonGroup(self)
        self.grp_std.addButton(self.rb_nd30, 0)
        self.grp_std.addButton(self.rb_hd36, 1)
        row_std.addWidget(self.rb_nd30)
        row_std.addWidget(self.rb_hd36)
        l_ag.addLayout(row_std)

        row_sup = QHBoxLayout()
        row_sup.addWidget(QLabel("Cơ quan cấp trên:"))
        self.txt_sup = QLineEdit(self.meta.get("superior_agency", ""))
        self.txt_sup.setPlaceholderText("Ví dụ: TỈNH ỦY ĐẮK NÔNG hoặc ỦY BAN NHÂN DÂN...")
        self.txt_sup.setMinimumHeight(30)
        row_sup.addWidget(self.txt_sup, 1)
        l_ag.addLayout(row_sup)

        row_iss = QHBoxLayout()
        row_iss.addWidget(QLabel("Cơ quan ban hành:"))
        self.txt_iss = QLineEdit(self.meta.get("issuing_agency", ""))
        self.txt_iss.setPlaceholderText("Ví dụ: ĐẢNG BỘ XÃ THUẬN HẠNH hoặc THÀNH PHỐ HÀ NỘI...")
        self.txt_iss.setMinimumHeight(30)
        row_iss.addWidget(self.txt_iss, 1)
        l_ag.addLayout(row_iss)

        row_num = QHBoxLayout()
        row_num.addWidget(QLabel("Số và ký hiệu:"))
        self.txt_num = QLineEdit(self.meta.get("doc_number", ""))
        self.txt_num.setPlaceholderText("Ví dụ: Số 16-QĐ/ĐU hoặc Số: 125/QĐ-UBND")
        self.txt_num.setMinimumHeight(30)
        row_num.addWidget(self.txt_num, 1)
        l_ag.addLayout(row_num)

        layout.addWidget(grp_agency)

        # 2. Địa danh, Thời gian & Trích yếu
        grp_content = QGroupBox("2. Địa danh, Thời gian & Trích yếu")
        l_cnt = QVBoxLayout(grp_content)

        row_loc_date = QHBoxLayout()
        row_loc_date.addWidget(QLabel("Địa danh:"))
        self.txt_loc = QLineEdit(self.meta.get("location", "Hà Nội"))
        self.txt_loc.setMinimumHeight(30)
        row_loc_date.addWidget(self.txt_loc, 1)

        row_loc_date.addWidget(QLabel("Ngày tháng năm:"))
        self.txt_date = QLineEdit(self.meta.get("date_str", ""))
        self.txt_date.setPlaceholderText("Ví dụ: Thuận Hạnh, ngày 02 tháng 11 năm 2026")
        self.txt_date.setMinimumHeight(30)
        row_loc_date.addWidget(self.txt_date, 2)
        l_cnt.addLayout(row_loc_date)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Loại văn bản:"))
        self.cb_type = QComboBox()
        self.cb_type.setMinimumHeight(30)
        types = [
            "QUYẾT ĐỊNH", "CÔNG VĂN", "TỜ TRÌNH", "THÔNG BÁO",
            "NGHỊ QUYẾT", "CHỈ THỊ", "KẾ HOẠCH", "BÁO CÁO", "QUY ĐỊNH"
        ]
        self.cb_type.addItems(types)
        curr_type = self.meta.get("doc_type", "QUYẾT ĐỊNH").upper()
        if curr_type in types:
            self.cb_type.setCurrentText(curr_type)
        else:
            self.cb_type.addItem(curr_type)
            self.cb_type.setCurrentText(curr_type)
        row_type.addWidget(self.cb_type, 1)
        l_cnt.addLayout(row_type)

        l_cnt.addWidget(QLabel("Trích yếu nội dung (Về việc...):"))
        self.txt_subject = QLineEdit(self.meta.get("doc_subject", ""))
        self.txt_subject.setPlaceholderText("Ví dụ: Về việc phân công nhiệm vụ cán bộ phụ trách chi bộ...")
        self.txt_subject.setMinimumHeight(30)
        l_cnt.addWidget(self.txt_subject)

        layout.addWidget(grp_content)

        # 3. Nơi nhận
        grp_rec = QGroupBox("3. Nơi nhận văn bản")
        l_rec = QVBoxLayout(grp_rec)
        lbl_rec_hint = QLabel("Nhập danh sách nơi nhận (mỗi dòng một đơn vị, ví dụ: - Như Điều 3;):")
        lbl_rec_hint.setStyleSheet("color: #64748b; font-size: 11.5px;")
        l_rec.addWidget(lbl_rec_hint)

        self.txt_recipients = QPlainTextEdit()
        recs = self.meta.get("recipients", [])
        if isinstance(recs, list):
            self.txt_recipients.setPlainText("\n".join(recs))
        else:
            self.txt_recipients.setPlainText(str(recs))
        self.txt_recipients.setMinimumHeight(90)
        l_rec.addWidget(self.txt_recipients)
        layout.addWidget(grp_rec)

        # 4. Khối Ký ban hành
        grp_signer = QGroupBox("4. Thẩm quyền ký & Người ký")
        l_sign = QVBoxLayout(grp_signer)

        row_auth = QHBoxLayout()
        row_auth.addWidget(QLabel("Quyền hạn ký:"))
        self.txt_auth = QLineEdit(self.meta.get("authority_prefix", ""))
        self.txt_auth.setPlaceholderText("Ví dụ: T/M BAN THƯỜNG VỤ hoặc TM. ỦY BAN NHÂN DÂN")
        self.txt_auth.setMinimumHeight(30)
        row_auth.addWidget(self.txt_auth, 1)
        l_sign.addLayout(row_auth)

        row_pos_name = QHBoxLayout()
        row_pos_name.addWidget(QLabel("Chức vụ:"))
        self.txt_pos = QLineEdit(self.meta.get("signer_position", ""))
        self.txt_pos.setPlaceholderText("Ví dụ: BÍ THƯ hoặc CHỦ TỊCH")
        self.txt_pos.setMinimumHeight(30)
        row_pos_name.addWidget(self.txt_pos, 1)

        row_pos_name.addWidget(QLabel("Họ và tên:"))
        self.txt_name = QLineEdit(self.meta.get("signer_name", ""))
        self.txt_name.setPlaceholderText("Ví dụ: K'Thanh hoặc Lê Văn Cường")
        self.txt_name.setMinimumHeight(30)
        row_pos_name.addWidget(self.txt_name, 1)
        l_sign.addLayout(row_pos_name)

        layout.addWidget(grp_signer)

        scroll.setWidget(container)
        main_layout.addWidget(scroll, 1)

        # Nút điều khiển
        btn_box = QHBoxLayout()
        self.btn_cancel = QPushButton("Hủy bỏ")
        self.btn_cancel.clicked.connect(self.reject)
        btn_box.addWidget(self.btn_cancel)

        self.btn_save = QPushButton("✔ CẬP NHẬT THỂ THỨC VĂN BẢN")
        self.btn_save.setIcon(get_ui_icon("save"))
        self.btn_save.setStyleSheet("""
            QPushButton {
                background-color: #059669; color: white; font-weight: bold;
                padding: 9px 20px; border-radius: 6px; font-size: 12.5px;
            }
            QPushButton:hover { background-color: #047857; }
        """)
        self.btn_save.clicked.connect(self._on_save_clicked)
        btn_box.addWidget(self.btn_save)

        main_layout.addLayout(btn_box)

    def _on_save_clicked(self):
        # Đồng bộ ngày tháng nếu trống
        loc = self.txt_loc.text().strip() or "Hà Nội"
        date_str = self.txt_date.text().strip()
        if not date_str:
            date_str = f"{loc}, ngày ... tháng ... năm 2026"

        # Tách danh sách nơi nhận
        raw_rec = self.txt_recipients.toPlainText()
        rec_lines = [r.strip() for r in raw_rec.splitlines() if r.strip()]
        if not rec_lines:
            rec_lines = ["- Như Điều 3;", "- Lưu: VT."]

        self.meta.update({
            "standard_id": "HD36" if self.rb_hd36.isChecked() else "ND30",
            "superior_agency": self.txt_sup.text().strip().upper(),
            "issuing_agency": self.txt_iss.text().strip().upper(),
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
