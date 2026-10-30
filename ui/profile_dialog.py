"""
ui/profile_dialog.py: Hộp thoại Quản lý Hồ sơ Đơn vị / Cơ quan ban hành
Cho phép lưu trữ, tùy biến thông tin cơ quan và mẫu ký để cán bộ không phải nhập lại mỗi lần dùng.
"""

import json
import os
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QLineEdit,
    QComboBox, QTextEdit, QPushButton, QListWidget, QMessageBox,
    QFormLayout, QGroupBox
)
from PyQt6.QtCore import Qt, QSize
from PyQt6.QtGui import QIcon

from core.path_utils import get_config_dir, get_app_icon_path
from ui.icon_factory import get_ui_icon


class ProfileDialog(QDialog):
    def __init__(self, parent=None, config_dir=None):
        super().__init__(parent)
        self.setWindowTitle("Quản lý Hồ sơ Cơ quan & Đơn vị")
        self.resize(750, 520)

        app_icon = get_app_icon_path()
        if app_icon and os.path.exists(app_icon):
            self.setWindowIcon(QIcon(app_icon))

        if config_dir is None:
            config_dir = get_config_dir()
        self.config_dir = config_dir
        self.profiles_path = os.path.join(self.config_dir, "agency_profiles.json")

        self.profiles_data = self._load_profiles()
        self._init_ui()
        self._populate_list()

    def _load_profiles(self):
        if os.path.exists(self.profiles_path):
            try:
                with open(self.profiles_path, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"active_profile_id": "state_ubnd", "profiles": []}

    def _save_profiles(self):
        try:
            with open(self.profiles_path, "w", encoding="utf-8") as f:
                json.dump(self.profiles_data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            QMessageBox.critical(self, "Lỗi", f"Không thể lưu hồ sơ: {str(e)}")
            return False

    def _init_ui(self):
        main_layout = QHBoxLayout(self)

        # Cột trái: Danh sách hồ sơ & nút thao tác
        left_layout = QVBoxLayout()
        lbl_list = QLabel("<b>Danh mục cơ quan đã lưu:</b>")
        self.list_profiles = QListWidget()
        self.list_profiles.currentRowChanged.connect(self._on_profile_selected)

        btn_layout = QHBoxLayout()
        self.btn_new = QPushButton("Thêm mới")
        self.btn_new.setObjectName("btnSecondary")
        self.btn_new.setIcon(get_ui_icon("file_add"))
        self.btn_new.setIconSize(QSize(16, 16))
        self.btn_new.clicked.connect(self._on_new_profile)

        self.btn_delete = QPushButton("Xóa")
        self.btn_delete.setObjectName("btnSecondary")
        self.btn_delete.setIcon(get_ui_icon("trash"))
        self.btn_delete.setIconSize(QSize(16, 16))
        self.btn_delete.clicked.connect(self._on_delete_profile)

        btn_layout.addWidget(self.btn_new)
        btn_layout.addWidget(self.btn_delete)

        left_layout.addWidget(lbl_list)
        left_layout.addWidget(self.list_profiles)
        left_layout.addLayout(btn_layout)

        # Cột phải: Form chi tiết
        right_group = QGroupBox("Thông tin chi tiết hồ sơ cơ quan")
        form_layout = QFormLayout(right_group)

        self.txt_id = QLineEdit()
        self.txt_id.setPlaceholderText("Mã định danh (ví dụ: ubnd_tinh, danguy_khoi)")

        self.txt_name = QLineEdit()
        self.txt_name.setPlaceholderText("Tên gọi hồ sơ hiển thị")

        self.cb_standard = QComboBox()
        self.cb_standard.addItem("Văn bản Hành chính (NĐ 30/2020/NĐ-CP)", "ND30")
        self.cb_standard.addItem("Văn bản Đảng (HD 36-HD/VPTW)", "HD36")

        self.txt_superior = QLineEdit()
        self.txt_superior.setPlaceholderText("Cơ quan chủ quản / cấp ủy cấp trên")

        self.txt_issuing = QLineEdit()
        self.txt_issuing.setPlaceholderText("Cơ quan ban hành văn bản")

        self.txt_location = QLineEdit()
        self.txt_location.setPlaceholderText("Địa danh ban hành (ví dụ: Hà Nội)")

        self.txt_symbol = QLineEdit()
        self.txt_symbol.setPlaceholderText("Mã ký hiệu số (ví dụ: QĐ-UBND, CV/TU)")

        self.txt_auth = QLineEdit()
        self.txt_auth.setPlaceholderText("Quyền hạn ký (ví dụ: TM. ỦY BAN NHÂN DÂN hoặc T/M BAN THƯỜNG VỤ)")

        self.txt_pos = QLineEdit()
        self.txt_pos.setPlaceholderText("Chức vụ người ký (ví dụ: CHỦ TỊCH, BÍ THƯ)")

        self.txt_signer = QLineEdit()
        self.txt_signer.setPlaceholderText("Họ và tên người ký")

        self.txt_recipients = QTextEdit()
        self.txt_recipients.setPlaceholderText("Danh sách nơi nhận (mỗi dòng một đơn vị)")
        self.txt_recipients.setMaximumHeight(80)

        form_layout.addRow("Mã hồ sơ:", self.txt_id)
        form_layout.addRow("Tên hiển thị:", self.txt_name)
        form_layout.addRow("Quy chuẩn:", self.cb_standard)
        form_layout.addRow("Cơ quan cấp trên:", self.txt_superior)
        form_layout.addRow("Cơ quan ban hành:", self.txt_issuing)
        form_layout.addRow("Địa danh:", self.txt_location)
        form_layout.addRow("Ký hiệu văn bản:", self.txt_symbol)
        form_layout.addRow("Quyền hạn ký:", self.txt_auth)
        form_layout.addRow("Chức vụ:", self.txt_pos)
        form_layout.addRow("Họ tên người ký:", self.txt_signer)
        form_layout.addRow("Nơi nhận:", self.txt_recipients)

        right_layout = QVBoxLayout()
        right_layout.addWidget(right_group)

        action_layout = QHBoxLayout()
        action_layout.addStretch()
        self.btn_save = QPushButton("Lưu hồ sơ này")
        self.btn_save.setIcon(get_ui_icon("status_success"))
        self.btn_save.setIconSize(QSize(16, 16))
        self.btn_save.clicked.connect(self._on_save_current_profile)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setObjectName("btnSecondary")
        self.btn_close.clicked.connect(self.accept)

        action_layout.addWidget(self.btn_save)
        action_layout.addWidget(self.btn_close)
        right_layout.addLayout(action_layout)

        main_layout.addLayout(left_layout, 1)
        main_layout.addLayout(right_layout, 2)

    def _populate_list(self):
        self.list_profiles.clear()
        profiles = self.profiles_data.get("profiles", [])
        for p in profiles:
            self.list_profiles.addItem(p.get("name", p.get("id", "")))
        if profiles:
            self.list_profiles.setCurrentRow(0)

    def _on_profile_selected(self, row: int):
        profiles = self.profiles_data.get("profiles", [])
        if 0 <= row < len(profiles):
            p = profiles[row]
            self.txt_id.setText(p.get("id", ""))
            self.txt_name.setText(p.get("name", ""))

            std = p.get("standard_id", "ND30")
            idx = 0 if std == "ND30" else 1
            self.cb_standard.setCurrentIndex(idx)

            self.txt_superior.setText(p.get("superior_agency", ""))
            self.txt_issuing.setText(p.get("issuing_agency", ""))
            self.txt_location.setText(p.get("location", ""))
            self.txt_symbol.setText(p.get("doc_number_symbol", ""))
            self.txt_auth.setText(p.get("authority_prefix", ""))
            self.txt_pos.setText(p.get("signer_position", ""))
            self.txt_signer.setText(p.get("signer_name", ""))
            self.txt_recipients.setPlainText("\n".join(p.get("recipients", [])))

    def _on_new_profile(self):
        self.txt_id.setText("agency_new")
        self.txt_name.setText("Hồ sơ cơ quan mới")
        self.cb_standard.setCurrentIndex(0)
        self.txt_superior.clear()
        self.txt_issuing.clear()
        self.txt_location.setText("Hà Nội")
        self.txt_symbol.clear()
        self.txt_auth.clear()
        self.txt_pos.clear()
        self.txt_signer.clear()
        self.txt_recipients.setPlainText("- Lãnh đạo đơn vị;\n- Lưu: VT.")
        self.txt_name.setFocus()

    def _on_save_current_profile(self):
        pid = self.txt_id.text().strip()
        name = self.txt_name.text().strip()
        if not pid or not name:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng nhập Mã và Tên hiển thị hồ sơ!")
            return

        recipients = [line.strip() for line in self.txt_recipients.toPlainText().split("\n") if line.strip()]
        new_prof = {
            "id": pid,
            "name": name,
            "standard_id": self.cb_standard.currentData(),
            "superior_agency": self.txt_superior.text().strip(),
            "issuing_agency": self.txt_issuing.text().strip(),
            "location": self.txt_location.text().strip(),
            "doc_number_symbol": self.txt_symbol.text().strip(),
            "authority_prefix": self.txt_auth.text().strip(),
            "signer_position": self.txt_pos.text().strip(),
            "signer_name": self.txt_signer.text().strip(),
            "recipients": recipients
        }

        profiles = self.profiles_data.setdefault("profiles", [])
        found_idx = -1
        for i, p in enumerate(profiles):
            if p.get("id") == pid:
                found_idx = i
                break

        if found_idx >= 0:
            profiles[found_idx] = new_prof
        else:
            profiles.append(new_prof)

        if self._save_profiles():
            QMessageBox.information(self, "Thành công", f"Đã lưu hồ sơ '{name}'!")
            self._populate_list()

    def _on_delete_profile(self):
        row = self.list_profiles.currentRow()
        profiles = self.profiles_data.get("profiles", [])
        if 0 <= row < len(profiles):
            target = profiles[row]
            reply = QMessageBox.question(
                self, "Xác nhận",
                f"Bạn có chắc muốn xóa hồ sơ '{target.get('name')}'?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No
            )
            if reply == QMessageBox.StandardButton.Yes:
                profiles.pop(row)
                self._save_profiles()
                self._populate_list()
