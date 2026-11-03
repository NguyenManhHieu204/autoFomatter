"""
ui/editor/metadata_editor_dlg.py: Hộp thoại Chỉnh sửa Thể thức Văn bản (Chuyên dụng Cấp Xã)
Cho phép cán bộ biên tập các trường thông tin thể thức:
- Cơ quan cấp trên, Cơ quan ban hành, Số ký hiệu
- Địa danh, Ngày tháng năm, Loại văn bản, Trích yếu
- Nơi nhận và Khối ký ban hành
"""

from typing import Dict, Any, List
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QPlainTextEdit, QComboBox, QRadioButton, QButtonGroup,
    QGroupBox, QPushButton, QScrollArea, QWidget
)
from PyQt6.QtCore import Qt


class MetadataEditorDialog(QDialog):
    """Hộp thoại chỉnh sửa thông tin thể thức hành chính & Đảng cấp xã"""

    def __init__(self, parent=None, metadata: Dict[str, Any] = None):
        super().__init__(parent)
        self.setWindowTitle("Chỉnh sửa thể thức văn bản cấp xã")
        self.resize(640, 680)
        self.setMinimumSize(560, 560)
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
            }
            QLabel {
                color: #1e293b;
                font-weight: 500;
                font-size: 12px;
            }
        """)

        self.meta = dict(metadata or {})

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(12)

        # Scroll Area chứa nội dung
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        container = QWidget()
        layout = QVBoxLayout(container)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(12)

        # 1. Cơ quan & Số hiệu
        grp_agency = QGroupBox("1. Cơ quan ban hành & Số ký hiệu (Cấp xã)")
        l_ag = QVBoxLayout(grp_agency)
        l_ag.setSpacing(8)

        row_std = QHBoxLayout()
        self.rb_nd30 = QRadioButton("Hành chính Nhà nước (Nghị định 30)")
        self.rb_hd36 = QRadioButton("Cơ quan Đảng (Hướng dẫn 36)")
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
        lbl_sup = QLabel("Cơ quan cấp trên:")
        lbl_sup.setMinimumWidth(110)
        row_sup.addWidget(lbl_sup)
        self.txt_sup = QLineEdit(self.meta.get("superior_agency", ""))
        self.txt_sup.setPlaceholderText("Ví dụ: ỦY BAN NHÂN DÂN HUYỆN... hoặc HUYỆN ỦY...")
        row_sup.addWidget(self.txt_sup, 1)
        l_ag.addLayout(row_sup)

        row_iss = QHBoxLayout()
        lbl_iss = QLabel("Cơ quan ban hành:")
        lbl_iss.setMinimumWidth(110)
        row_iss.addWidget(lbl_iss)
        self.txt_iss = QLineEdit(self.meta.get("issuing_agency", ""))
        self.txt_iss.setPlaceholderText("Ví dụ: ỦY BAN NHÂN DÂN XÃ... hoặc ĐẢNG ỦY XÃ...")
        row_iss.addWidget(self.txt_iss, 1)
        l_ag.addLayout(row_iss)

        row_num = QHBoxLayout()
        lbl_num = QLabel("Số và ký hiệu:")
        lbl_num.setMinimumWidth(110)
        row_num.addWidget(lbl_num)
        self.txt_num = QLineEdit(self.meta.get("doc_number", ""))
        self.txt_num.setPlaceholderText("Ví dụ: Số: .../QĐ-UBND hoặc Số ...-QĐ/ĐU")
        row_num.addWidget(self.txt_num, 1)
        l_ag.addLayout(row_num)

        layout.addWidget(grp_agency)

        # 2. Địa danh, Thời gian & Trích yếu
        grp_content = QGroupBox("2. Địa danh, Thời gian & Trích yếu")
        l_cnt = QVBoxLayout(grp_content)
        l_cnt.setSpacing(8)

        row_loc_date = QHBoxLayout()
        row_loc_date.addWidget(QLabel("Địa danh:"))
        self.txt_loc = QLineEdit(self.meta.get("location", "Xã"))
        self.txt_loc.setPlaceholderText("Ví dụ: Xã Thuận Hạnh")
        row_loc_date.addWidget(self.txt_loc, 1)

        row_loc_date.addWidget(QLabel("Ngày tháng:"))
        self.txt_date = QLineEdit(self.meta.get("date_str", ""))
        self.txt_date.setPlaceholderText("Thuận Hạnh, ngày ... tháng ... năm 2026")
        row_loc_date.addWidget(self.txt_date, 2)
        l_cnt.addLayout(row_loc_date)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Loại văn bản:"))
        self.cb_type = QComboBox()
        types = [
            "QUYẾT ĐỊNH", "CÔNG VĂN", "TỜ TRÌNH", "THÔNG BÁO",
            "KẾ HOẠCH", "BÁO CÁO", "BIÊN BẢN", "NGHỊ QUYẾT ĐẢNG", "CHỈ THỊ ĐẢNG"
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
        self.txt_subject.setPlaceholderText("Ví dụ: Về việc ban hành Kế hoạch phát triển kinh tế - xã hội")
        l_cnt.addWidget(self.txt_subject)

        layout.addWidget(grp_content)

        # 3. Nơi nhận
        grp_rec = QGroupBox("3. Nơi nhận (Mỗi dòng một nơi nhận)")
        l_rec = QVBoxLayout(grp_rec)
        l_rec.setSpacing(6)

        self.txt_recipients = QPlainTextEdit()
        recs = self.meta.get("recipients", [])
        if isinstance(recs, list):
            self.txt_recipients.setPlainText("\n".join(recs))
        else:
            self.txt_recipients.setPlainText(str(recs))
        self.txt_recipients.setMinimumHeight(75)
        l_rec.addWidget(self.txt_recipients)
        layout.addWidget(grp_rec)

        # 4. Khối Ký ban hành
        grp_signer = QGroupBox("4. Thẩm quyền ký & Người ký")
        l_sign = QVBoxLayout(grp_signer)
        l_sign.setSpacing(8)

        row_auth = QHBoxLayout()
        lbl_a = QLabel("Quyền hạn ký:")
        lbl_a.setMinimumWidth(90)
        row_auth.addWidget(lbl_a)
        self.txt_auth = QLineEdit(self.meta.get("authority_prefix", ""))
        self.txt_auth.setPlaceholderText("TM. ỦY BAN NHÂN DÂN hoặc T/M ĐẢNG ỦY")
        row_auth.addWidget(self.txt_auth, 1)
        l_sign.addLayout(row_auth)

        row_pos_name = QHBoxLayout()
        row_pos_name.addWidget(QLabel("Chức vụ:"))
        self.txt_pos = QLineEdit(self.meta.get("signer_position", ""))
        self.txt_pos.setPlaceholderText("CHỦ TỊCH hoặc BÍ THƯ")
        row_pos_name.addWidget(self.txt_pos, 1)

        row_pos_name.addWidget(QLabel("Họ tên:"))
        self.txt_name = QLineEdit(self.meta.get("signer_name", ""))
        self.txt_name.setPlaceholderText("Họ và tên người ký")
        row_pos_name.addWidget(self.txt_name, 1)
        l_sign.addLayout(row_pos_name)

        layout.addWidget(grp_signer)

        scroll.setWidget(container)
        main_layout.addWidget(scroll, 1)

        # Nút điều khiển
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

    def _on_save_clicked(self):
        loc = self.txt_loc.text().strip() or "Xã"
        date_str = self.txt_date.text().strip()
        if not date_str:
            date_str = f"{loc}, ngày ... tháng ... năm 2026"

        raw_rec = self.txt_recipients.toPlainText()
        rec_lines = [r.strip() for r in raw_rec.splitlines() if r.strip()]
        if not rec_lines:
            rec_lines = ["- Thường trực HĐND xã;", "- Chủ tịch, các PCT UBND xã;", "- Lưu: VT."]

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
