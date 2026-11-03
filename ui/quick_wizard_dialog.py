"""
ui/quick_wizard_dialog.py: Hộp thoại Khởi tạo Văn bản Mới (Template Wizard)
Cho phép lựa chọn quy chuẩn thể thức (NĐ 30 / HD 36), cơ quan ban hành,
loại văn bản và trích yếu nội dung để nạp vào trình soạn thảo.
"""

from typing import Dict, Any, Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QPushButton, QGroupBox, QRadioButton,
    QButtonGroup, QMessageBox, QScrollArea, QWidget
)
from PyQt6.QtCore import Qt

from core.template_wizard import TemplateWizardController


class QuickWizardDialog(QDialog):
    """Hộp thoại khởi tạo văn bản chuẩn thể thức"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Khởi tạo văn bản mới")
        self.resize(580, 530)
        self.setMinimumSize(520, 480)
        self.wizard_controller = TemplateWizardController()
        self.result_metadata: Optional[Dict[str, Any]] = None

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(18, 16, 18, 16)
        main_layout.setSpacing(12)

        # Scroll Area chứa các trường thông tin
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        scroll_area.setStyleSheet("QScrollArea { border: none; background: transparent; }")

        scroll_content = QWidget()
        layout = QVBoxLayout(scroll_content)
        layout.setContentsMargins(0, 0, 4, 0)
        layout.setSpacing(12)

        # 1. Cơ quan ban hành
        prof_group = QGroupBox("1. Cơ quan ban hành")
        prof_layout = QVBoxLayout(prof_group)
        prof_layout.setSpacing(8)

        row_std = QHBoxLayout()
        self.rb_nd30 = QRadioButton("Hành chính Nhà nước (Nghị định 30)")
        self.rb_hd36 = QRadioButton("Cơ quan Đảng (Hướng dẫn 36)")
        self.rb_nd30.setChecked(True)
        self.rb_nd30.toggled.connect(self._on_standard_changed)

        self.grp_std = QButtonGroup(self)
        self.grp_std.addButton(self.rb_nd30, 0)
        self.grp_std.addButton(self.rb_hd36, 1)

        row_std.addWidget(self.rb_nd30)
        row_std.addWidget(self.rb_hd36)
        prof_layout.addLayout(row_std)

        row_prof = QHBoxLayout()
        lbl_prof = QLabel("Hồ sơ mẫu:")
        lbl_prof.setMinimumWidth(80)
        row_prof.addWidget(lbl_prof)
        self.cb_profiles = QComboBox()
        self.cb_profiles.currentIndexChanged.connect(self._on_profile_selected)
        row_prof.addWidget(self.cb_profiles, 1)
        prof_layout.addLayout(row_prof)

        layout.addWidget(prof_group)

        # 2. Loại văn bản và Trích yếu
        doc_group = QGroupBox("2. Thể loại & Trích yếu")
        doc_layout = QVBoxLayout(doc_group)
        doc_layout.setSpacing(8)

        row_type = QHBoxLayout()
        lbl_t = QLabel("Loại văn bản:")
        lbl_t.setMinimumWidth(80)
        row_type.addWidget(lbl_t)
        self.cb_doc_type = QComboBox()
        row_type.addWidget(self.cb_doc_type, 1)
        doc_layout.addLayout(row_type)

        doc_layout.addWidget(QLabel("Trích yếu nội dung (Về việc...):"))
        self.txt_subject = QLineEdit()
        self.txt_subject.setPlaceholderText("Ví dụ: Phê duyệt kế hoạch chuyển đổi số năm 2026")
        doc_layout.addWidget(self.txt_subject)

        layout.addWidget(doc_group)

        # 3. Người ký và địa danh
        sign_group = QGroupBox("3. Thẩm quyền ký & Địa danh")
        sign_layout = QVBoxLayout(sign_group)
        sign_layout.setSpacing(8)

        row_signer = QHBoxLayout()
        row_signer.addWidget(QLabel("Chức vụ:"))
        self.txt_position = QLineEdit()
        row_signer.addWidget(self.txt_position, 1)

        row_signer.addWidget(QLabel("Họ tên:"))
        self.txt_signer_name = QLineEdit()
        row_signer.addWidget(self.txt_signer_name, 1)
        sign_layout.addLayout(row_signer)

        row_loc = QHBoxLayout()
        row_loc.addWidget(QLabel("Địa danh:"))
        self.txt_location = QLineEdit()
        self.txt_location.setText("Xã")
        row_loc.addWidget(self.txt_location, 1)
        sign_layout.addLayout(row_loc)

        layout.addWidget(sign_group)

        scroll_area.setWidget(scroll_content)
        main_layout.addWidget(scroll_area, 1)

        # Nút điều khiển
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        self.btn_cancel = QPushButton("Hủy")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_create = QPushButton("Bắt đầu soạn thảo")
        self.btn_create.setObjectName("btnPrimary")
        self.btn_create.setDefault(True)
        self.btn_create.clicked.connect(self._create_and_open_editor)
        btn_layout.addWidget(self.btn_create)

        main_layout.addLayout(btn_layout)

        # Nạp danh sách ban đầu
        self._refresh_profiles_list()

    def _on_standard_changed(self):
        self._refresh_profiles_list()
        self._refresh_doc_types()

    def _refresh_profiles_list(self):
        std = "HD36" if self.rb_hd36.isChecked() else "ND30"
        self.cb_profiles.clear()
        profs = self.wizard_controller.get_profiles()
        matched = [p for p in profs if p.get("standard_id", "ND30").upper() == std]

        for p in matched:
            self.cb_profiles.addItem(p.get("name", "Cơ quan"), p)

        if not matched:
            self.cb_profiles.addItem("Cơ quan cấp xã mặc định", {
                "standard_id": std,
                "superior_agency": "HUYỆN ỦY" if std == "HD36" else "ỦY BAN NHÂN DÂN HUYỆN",
                "issuing_agency": "ĐẢNG ỦY XÃ" if std == "HD36" else "ỦY BAN NHÂN DÂN XÃ",
                "authority_prefix": "T/M ĐẢNG ỦY" if std == "HD36" else "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "BÍ THƯ" if std == "HD36" else "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "location": "Xã"
            })

        self._refresh_doc_types()

    def _refresh_doc_types(self):
        std = "HD36" if self.rb_hd36.isChecked() else "ND30"
        self.cb_doc_type.clear()
        types = self.wizard_controller.get_available_document_types(std)
        self.cb_doc_type.addItems(types)

    def _on_profile_selected(self, idx: int):
        data = self.cb_profiles.currentData()
        if data:
            self.txt_position.setText(data.get("signer_position", ""))
            self.txt_signer_name.setText(data.get("signer_name", ""))
            self.txt_location.setText(data.get("location", "Xã"))

    def _create_and_open_editor(self):
        subj = self.txt_subject.text().strip()
        if not subj:
            QMessageBox.warning(self, "Thiếu thông tin", "Vui lòng nhập trích yếu nội dung văn bản.")
            self.txt_subject.setFocus()
            return

        std = "HD36" if self.rb_hd36.isChecked() else "ND30"
        profile_data = self.cb_profiles.currentData() or {}
        doc_type_disp = self.cb_doc_type.currentText()

        # Tạo metadata
        self.result_metadata = {
            "standard_id": std,
            "superior_agency": profile_data.get("superior_agency", ""),
            "issuing_agency": profile_data.get("issuing_agency", ""),
            "doc_number": profile_data.get("doc_number_symbol", "QĐ-UBND"),
            "location": self.txt_location.text().strip() or "Xã",
            "doc_type": doc_type_disp.split("(")[0].strip().upper(),
            "doc_subject": f"Về việc {subj}" if not subj.lower().startswith("về việc") and not subj.lower().startswith("v/v") else subj,
            "authority_prefix": profile_data.get("authority_prefix", ""),
            "signer_position": self.txt_position.text().strip() or profile_data.get("signer_position", "CHỦ TỊCH"),
            "signer_name": self.txt_signer_name.text().strip() or profile_data.get("signer_name", "Nguyễn Văn A"),
            "recipients": profile_data.get("recipients", [])
        }
        self.accept()
