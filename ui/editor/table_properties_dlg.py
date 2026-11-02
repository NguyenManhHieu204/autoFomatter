"""
ui/editor/table_properties_dlg.py: Hộp thoại Chèn và Cấu hình Bảng biểu Nâng cao
- Thiết lập số hàng, số cột (Rows & Columns)
- Kiểu viền: Bảng lưới văn thư chuẩn (Grid) hoặc Bảng ẩn viền (Invisible)
- Tùy chọn bảo vệ bảng nhiều trang (Lặp tiêu đề tblHeader, chống gãy dòng cantSplit)
- Tự động chia đều bề rộng cột theo khổ in 165mm
"""

from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QSpinBox,
    QPushButton, QGroupBox, QRadioButton, QButtonGroup, QCheckBox, QDialogButtonBox
)
from PyQt6.QtCore import Qt


class TablePropertiesDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Chèn Bảng biểu Hành chính")
        self.resize(450, 390)
        self.setMinimumSize(400, 360)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(12)

        # 1. Kích thước bảng
        dim_group = QGroupBox("Kích thước bảng")
        dim_layout = QVBoxLayout(dim_group)
        dim_layout.setSpacing(8)

        row_layout = QHBoxLayout()
        lbl_r = QLabel("Số hàng:")
        lbl_r.setMinimumWidth(80)
        row_layout.addWidget(lbl_r)
        self.sp_rows = QSpinBox()
        self.sp_rows.setRange(1, 100)
        self.sp_rows.setValue(4)
        self.sp_rows.setMinimumHeight(28)
        self.sp_rows.setMinimumWidth(90)
        row_layout.addWidget(self.sp_rows)
        row_layout.addStretch()

        col_layout = QHBoxLayout()
        lbl_c = QLabel("Số cột:")
        lbl_c.setMinimumWidth(80)
        col_layout.addWidget(lbl_c)
        self.sp_cols = QSpinBox()
        self.sp_cols.setRange(1, 15)
        self.sp_cols.setValue(3)
        self.sp_cols.setMinimumHeight(28)
        self.sp_cols.setMinimumWidth(90)
        col_layout.addWidget(self.sp_cols)
        col_layout.addStretch()

        dim_layout.addLayout(row_layout)
        dim_layout.addLayout(col_layout)
        layout.addWidget(dim_group)

        # 2. Kiểu viền khung
        border_group = QGroupBox("Đường viền thể thức")
        border_layout = QVBoxLayout(border_group)
        border_layout.setSpacing(6)

        self.rb_grid = QRadioButton("Viền khung lưới chuẩn văn thư (1/2 pt)")
        self.rb_grid.setMinimumHeight(26)
        self.rb_grid.setChecked(True)
        self.rb_none = QRadioButton("Bảng ẩn viền (Dùng định vị Quốc hiệu / Chữ ký)")
        self.rb_none.setMinimumHeight(26)

        self.border_group = QButtonGroup(self)
        self.border_group.addButton(self.rb_grid, 0)
        self.border_group.addButton(self.rb_none, 1)

        border_layout.addWidget(self.rb_grid)
        border_layout.addWidget(self.rb_none)
        layout.addWidget(border_group)

        # 3. Tùy chọn bảo vệ nhiều trang
        protect_group = QGroupBox("Bảo vệ trang in")
        protect_layout = QVBoxLayout(protect_group)
        protect_layout.setSpacing(6)

        self.chk_header = QCheckBox("Lặp lại hàng tiêu đề trên từng trang (tblHeader)")
        self.chk_header.setMinimumHeight(26)
        self.chk_header.setChecked(True)
        self.chk_cant_split = QCheckBox("Chống bẻ đôi hàng qua 2 trang (cantSplit)")
        self.chk_cant_split.setMinimumHeight(26)
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
