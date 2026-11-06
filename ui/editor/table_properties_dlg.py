"""
ui/editor/table_properties_dlg.py: Hộp thoại Chèn và Cấu hình Bảng biểu Nâng cao
- Thiết lập số hàng, số cột (Rows & Columns)
- Kiểu viền: Bảng lưới văn thư chuẩn (Grid) hoặc Bảng ẩn viền (Invisible)
- Tùy chọn bảo vệ bảng nhiều trang (Lặp tiêu đề tblHeader, chống gãy dòng cantSplit)
- Tự động chia đều bề rộng cột theo khổ in 165mm
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QGridLayout, QLabel, QSpinBox,
    QPushButton, QGroupBox, QRadioButton, QButtonGroup, QCheckBox, QDialogButtonBox
)
from PyQt6.QtCore import Qt


class TablePropertiesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chèn Bảng biểu Hành chính")
        self.setMinimumSize(480, 420)
        self.resize(500, 440)
        self.setStyleSheet("""
            QDialog {
                background-color: #f8fafc;
                color: #0f172a;
            }
            QGroupBox {
                background-color: #ffffff;
                border: 1.5px solid #cbd5e1;
                border-radius: 8px;
                font-weight: 600;
                font-size: 13px;
                color: #0f172a;
                margin-top: 14px;
                padding: 16px 14px 12px 14px;
            }
            QGroupBox::title {
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 14px;
                padding: 0 8px;
                background-color: #ffffff;
                color: #1e293b;
            }
            QLabel {
                color: #0f172a;
                font-weight: 500;
                font-size: 13px;
                background: transparent;
            }
            QSpinBox {
                background-color: #ffffff;
                color: #0f172a;
                border: 1.5px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px 10px;
                font-size: 13px;
                min-height: 30px;
            }
            QSpinBox:focus {
                border-color: #2563eb;
            }
            QRadioButton, QCheckBox {
                color: #0f172a;
                font-size: 13px;
                font-weight: 500;
                spacing: 10px;
                background: transparent;
                min-height: 28px;
                padding: 2px 4px;
            }
            QRadioButton::indicator, QCheckBox::indicator {
                width: 18px;
                height: 18px;
            }
            QPushButton {
                background-color: #ffffff;
                color: #0f172a;
                font-weight: 500;
                font-size: 13px;
                padding: 7px 20px;
                min-height: 28px;
                border: 1.5px solid #cbd5e1;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
                border-color: #94a3b8;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 16, 20, 18)
        layout.setSpacing(14)

        # 1. Kích thước bảng (bố cục 2 cột gọn gàng, rộng rãi)
        dim_group = QGroupBox("Kích thước bảng")
        dim_grid = QGridLayout(dim_group)
        dim_grid.setContentsMargins(12, 12, 12, 10)
        dim_grid.setHorizontalSpacing(16)
        dim_grid.setVerticalSpacing(10)

        lbl_r = QLabel("Số hàng:")
        lbl_r.setMinimumWidth(65)
        dim_grid.addWidget(lbl_r, 0, 0)
        self.sp_rows = QSpinBox()
        self.sp_rows.setRange(1, 100)
        self.sp_rows.setValue(4)
        self.sp_rows.setMinimumWidth(85)
        dim_grid.addWidget(self.sp_rows, 0, 1)

        lbl_c = QLabel("Số cột:")
        lbl_c.setMinimumWidth(65)
        dim_grid.addWidget(lbl_c, 0, 2)
        self.sp_cols = QSpinBox()
        self.sp_cols.setRange(1, 15)
        self.sp_cols.setValue(3)
        self.sp_cols.setMinimumWidth(85)
        dim_grid.addWidget(self.sp_cols, 0, 3)

        layout.addWidget(dim_group)

        # 2. Kiểu viền khung
        border_group = QGroupBox("Đường viền thể thức")
        border_layout = QVBoxLayout(border_group)
        border_layout.setContentsMargins(12, 12, 12, 10)
        border_layout.setSpacing(8)

        self.rb_grid = QRadioButton("Viền khung lưới chuẩn văn thư (1/2 pt)")
        self.rb_grid.setChecked(True)
        self.rb_none = QRadioButton("Bảng ẩn viền (Dùng định vị Quốc hiệu / Chữ ký)")

        self.border_group = QButtonGroup(self)
        self.border_group.addButton(self.rb_grid, 0)
        self.border_group.addButton(self.rb_none, 1)

        border_layout.addWidget(self.rb_grid)
        border_layout.addWidget(self.rb_none)
        layout.addWidget(border_group)

        # 3. Tùy chọn bảo vệ nhiều trang
        protect_group = QGroupBox("Bảo vệ trang in")
        protect_layout = QVBoxLayout(protect_group)
        protect_layout.setContentsMargins(12, 12, 12, 10)
        protect_layout.setSpacing(8)

        self.chk_header = QCheckBox("Lặp lại hàng tiêu đề trên từng trang (tblHeader)")
        self.chk_header.setChecked(True)
        self.chk_cant_split = QCheckBox("Chống bẻ đôi hàng qua 2 trang (cantSplit)")
        self.chk_cant_split.setChecked(True)

        protect_layout.addWidget(self.chk_header)
        protect_layout.addWidget(self.chk_cant_split)
        layout.addWidget(protect_group)

        # Buttons
        btns = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        btns.accepted.connect(self.accept)
        btns.rejected.connect(self.reject)
        layout.addWidget(btns)

    def get_table_config(self):
        return {
            "rows": self.sp_rows.value(),
            "cols": self.sp_cols.value(),
            "border_type": "grid" if self.rb_grid.isChecked() else "none",
            "repeat_header": self.chk_header.isChecked(),
            "cant_split": self.chk_cant_split.isChecked()
        }
