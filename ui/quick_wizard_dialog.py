"""
ui/quick_wizard_dialog.py: Hộp thoại Wizard Tạo Nhanh Văn bản Hành chính & Đảng
- Hỗ trợ cán bộ chọn nhanh Cơ quan ban hành từ danh mục Hồ sơ (Profiles)
- Lựa chọn loại văn bản theo chuẩn NĐ 30/2020/NĐ-CP hoặc HĐ 36-HD/VPTW
- Nhập trích yếu nội dung văn bản
- Tự động sinh khung thể thức và mở ngay vào Trình Soạn thảo Tích hợp (WordEditorWindow)
"""

import os
from typing import Dict, Any, Optional

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QComboBox,
    QLineEdit, QTextEdit, QPushButton, QGroupBox, QRadioButton,
    QButtonGroup, QFrame, QMessageBox, QFileDialog
)
from PyQt6.QtGui import QFont, QColor
from PyQt6.QtCore import Qt

from core.template_wizard import TemplateWizardController
from ui.editor.word_editor_window import WordEditorWindow
from ui.icon_factory import get_ui_icon


class QuickWizardDialog(QDialog):
    """Hộp thoại Wizard Tạo Nhanh Văn bản"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Tạo nhanh Văn bản Chuẩn Thể thức (Template Wizard)")
        self.setFixedSize(540, 560)
        self.wizard_controller = TemplateWizardController()
        self.created_editor_window: Optional[WordEditorWindow] = None

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 18, 20, 18)
        layout.setSpacing(14)

        # Header giới thiệu
        header_frame = QFrame()
        header_frame.setStyleSheet("background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 6px; padding: 10px;")
        h_layout = QVBoxLayout(header_frame)
        lbl_title = QLabel("KHỞI TẠO VĂN BẢN CHUẨN THỂ THỨC")
        lbl_title.setStyleSheet("color: #166534; font-weight: bold; font-size: 13px;")
        lbl_desc = QLabel("Hệ thống tự động dựng khung Quốc hiệu, Tiêu ngữ, Số hiệu, Nơi nhận và Khối ký.\nCán bộ chỉ cần nhập trích yếu và bắt đầu soạn thảo nội dung.")
        lbl_desc.setStyleSheet("color: #15803d; font-size: 11.5px; line-height: 1.3;")
        h_layout.addWidget(lbl_title)
        h_layout.addWidget(lbl_desc)
        layout.addWidget(header_frame)

        # 1. Chọn Quy chuẩn & Hồ sơ cơ quan
        prof_group = QGroupBox("1. Thông tin Cơ quan ban hành")
        prof_layout = QVBoxLayout(prof_group)

        row_std = QHBoxLayout()
        self.rb_nd30 = QRadioButton("Khối Hành chính Nhà nước (Nghị định 30)")
        self.rb_hd36 = QRadioButton("Khối Cơ quan Đảng (Hướng dẫn 36)")
        self.rb_nd30.setChecked(True)
        self.rb_nd30.toggled.connect(self._on_standard_changed)

        self.grp_std = QButtonGroup(self)
        self.grp_std.addButton(self.rb_nd30, 0)
        self.grp_std.addButton(self.rb_hd36, 1)

        row_std.addWidget(self.rb_nd30)
        row_std.addWidget(self.rb_hd36)
        prof_layout.addLayout(row_std)

        row_prof = QHBoxLayout()
        row_prof.addWidget(QLabel("Hồ sơ mẫu:"))
        self.cb_profiles = QComboBox()
        self.cb_profiles.currentIndexChanged.connect(self._on_profile_selected)
        row_prof.addWidget(self.cb_profiles, 1)
        prof_layout.addLayout(row_prof)

        layout.addWidget(prof_group)

        # 2. Loại văn bản và Trích yếu
        doc_group = QGroupBox("2. Thể loại & Trích yếu nội dung")
        doc_layout = QVBoxLayout(doc_group)

        row_type = QHBoxLayout()
        row_type.addWidget(QLabel("Loại văn bản:"))
        self.cb_doc_type = QComboBox()
        row_type.addWidget(self.cb_doc_type, 1)
        doc_layout.addLayout(row_type)

        doc_layout.addWidget(QLabel("Trích yếu nội dung (Về việc...):"))
        self.txt_subject = QLineEdit()
        self.txt_subject.setPlaceholderText("Ví dụ: Về việc phê duyệt kế hoạch chuyển đổi số năm 2026")
        doc_layout.addWidget(self.txt_subject)

        layout.addWidget(doc_group)

        # 3. Thông tin người ký và địa danh
        sign_group = QGroupBox("3. Thẩm quyền ký & Địa danh")
        sign_layout = QVBoxLayout(sign_group)

        row_signer = QHBoxLayout()
        row_signer.addWidget(QLabel("Chức vụ:"))
        self.txt_position = QLineEdit()
        row_signer.addWidget(self.txt_position, 1)

        row_signer.addWidget(QLabel("Họ tên:"))
        self.txt_signer_name = QLineEdit()
        row_signer.addWidget(self.txt_signer_name, 1)
        sign_layout.addLayout(row_signer)

        row_loc = QHBoxLayout()
        row_loc.addWidget(QLabel("Địa danh ban hành:"))
        self.txt_location = QLineEdit()
        self.txt_location.setText("Hà Nội")
        row_loc.addWidget(self.txt_location, 1)
        sign_layout.addLayout(row_loc)

        layout.addWidget(sign_group)

        # Nút Tạo văn bản
        btn_layout = QHBoxLayout()
        self.btn_cancel = QPushButton("Hủy bỏ")
        self.btn_cancel.clicked.connect(self.reject)
        btn_layout.addWidget(self.btn_cancel)

        self.btn_create = QPushButton("KHỞI TẠO & MỞ SOẠN THẢO NGAY")
        self.btn_create.setIcon(get_ui_icon("file_word"))
        self.btn_create.setStyleSheet("""
            QPushButton {
                background-color: #16a34a; color: white; font-weight: bold;
                padding: 9px 18px; border-radius: 5px; font-size: 12.5px;
            }
            QPushButton:hover { background-color: #15803d; }
        """)
        self.btn_create.clicked.connect(self._create_and_open_editor)
        btn_layout.addWidget(self.btn_create)

        layout.addLayout(btn_layout)

        # Nạp dữ liệu ban đầu
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
            self.cb_profiles.addItem("Cơ quan mặc định", {
                "standard_id": std,
                "superior_agency": "TỈNH ỦY" if std == "HD36" else "ỦY BAN NHÂN DÂN",
                "issuing_agency": "BAN THƯỜNG VỤ" if std == "HD36" else "TỈNH HÀ NỘI",
                "authority_prefix": "T/M BAN THƯỜNG VỤ" if std == "HD36" else "TM. ỦY BAN NHÂN DÂN",
                "signer_position": "BÍ THƯ" if std == "HD36" else "CHỦ TỊCH",
                "signer_name": "Nguyễn Văn A",
                "location": "Hà Nội"
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
            self.txt_location.setText(data.get("location", "Hà Nội"))

    def _create_and_open_editor(self):
        subj = self.txt_subject.text().strip()
        if not subj:
            QMessageBox.warning(self, "Thiếu trích yếu", "Vui lòng nhập trích yếu nội dung văn bản (Về việc...).")
            self.txt_subject.setFocus()
            return

        std = "HD36" if self.rb_hd36.isChecked() else "ND30"
        profile_data = self.cb_profiles.currentData() or {}
        doc_type_disp = self.cb_doc_type.currentText()

        # Tạo metadata
        meta = {
            "standard_id": std,
            "superior_agency": profile_data.get("superior_agency", ""),
            "issuing_agency": profile_data.get("issuing_agency", ""),
            "doc_number": profile_data.get("doc_number_symbol", "QĐ-UBND"),
            "location": self.txt_location.text().strip() or "Hà Nội",
            "doc_type": doc_type_disp.split("(")[0].strip().upper(),
            "doc_subject": f"Về việc {subj}" if not subj.lower().startswith("về việc") and not subj.lower().startswith("v/v") else subj,
            "authority_prefix": profile_data.get("authority_prefix", ""),
            "signer_position": self.txt_position.text().strip() or profile_data.get("signer_position", "CHỦ TỊCH"),
            "signer_name": self.txt_signer_name.text().strip() or profile_data.get("signer_name", "Nguyễn Văn A"),
            "recipients": profile_data.get("recipients", [])
        }

        # Khởi chạy cửa sổ soạn thảo tích hợp
        self.created_editor_window = WordEditorWindow(parent=self.parent(), initial_metadata=meta)
        self.created_editor_window.show()
        self.accept()
