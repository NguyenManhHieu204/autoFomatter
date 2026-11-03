"""
ui/editor/administrative_ribbon.py: Dải băng điều khiển Hành chính (Administrative Ribbon Bar)
Thiết kế tối giản, hiện đại, tái lập trải nghiệm văn phòng quen thuộc:
- Thẻ 1: Văn bản (Home) - Phông chữ, Cỡ chữ, Định dạng, Căn lề, Giãn dòng, Danh sách
- Thẻ 2: Chèn (Insert) - Bảng biểu, Thao tác Hàng/Cột/Trộn ô, Chèn Ảnh, Ngắt trang
- Thẻ 3: Bố cục (Page Layout) - Hướng giấy, Lề in NĐ 30, Bật/Tắt Thước đo
- Thẻ 4: Thể thức (Admin Tools) - Sửa Thể thức (F2), Khóa khung, Rà soát tuân thủ
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QComboBox, QToolButton, QLabel, QFrame, QButtonGroup,
    QColorDialog
)
from PyQt6.QtGui import QFont, QColor
from PyQt6.QtCore import Qt, pyqtSignal, QSize

from ui.icon_factory import get_ui_icon


class AdministrativeRibbonBar(QWidget):
    # Signals
    save_document_requested = pyqtSignal()
    export_pdf_requested = pyqtSignal()
    sign_document_requested = pyqtSignal()

    format_bold_triggered = pyqtSignal()
    format_italic_triggered = pyqtSignal()
    format_underline_triggered = pyqtSignal()
    font_size_changed = pyqtSignal(float)
    font_family_changed = pyqtSignal(str)
    font_color_changed = pyqtSignal(QColor)

    align_left_triggered = pyqtSignal()
    align_center_triggered = pyqtSignal()
    align_right_triggered = pyqtSignal()
    align_justify_triggered = pyqtSignal()

    line_spacing_changed = pyqtSignal(float)
    indent_first_line_triggered = pyqtSignal(float)
    bullet_list_triggered = pyqtSignal()
    numbered_list_triggered = pyqtSignal()

    insert_table_requested = pyqtSignal()
    insert_image_requested = pyqtSignal()
    page_break_requested = pyqtSignal()
    insert_symbol_requested = pyqtSignal(str)

    table_insert_row_above_triggered = pyqtSignal()
    table_insert_row_below_triggered = pyqtSignal()
    table_delete_row_triggered = pyqtSignal()
    table_insert_col_left_triggered = pyqtSignal()
    table_insert_col_right_triggered = pyqtSignal()
    table_delete_col_triggered = pyqtSignal()
    table_merge_cells_triggered = pyqtSignal()
    table_split_cell_triggered = pyqtSignal()

    page_orientation_changed = pyqtSignal(str)
    margins_preset_triggered = pyqtSignal(str)
    toggle_ruler_triggered = pyqtSignal(bool)

    edit_metadata_requested = pyqtSignal()
    switch_doc_type_requested = pyqtSignal(str)
    check_compliance_requested = pyqtSignal()
    lock_scaffolding_toggled = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(105)
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QTabWidget::pane {
                border-top: 1px solid #e2e8f0;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #ffffff;
                color: #64748b;
                padding: 6px 16px;
                border: none;
                border-bottom: 2px solid transparent;
                font-size: 12px;
                font-weight: 500;
            }
            QTabBar::tab:selected {
                color: #2563eb;
                border-bottom: 2px solid #2563eb;
                font-weight: 600;
            }
            QTabBar::tab:hover {
                color: #0f172a;
            }
            QToolButton {
                border: 1px solid transparent;
                border-radius: 4px;
                padding: 4px 6px;
                background: transparent;
                font-size: 11.5px;
                color: #334155;
            }
            QToolButton:hover {
                background-color: #f1f5f9;
                border-color: #cbd5e1;
            }
            QToolButton:checked {
                background-color: #eff6ff;
                border-color: #bfdbfe;
                color: #1d4ed8;
            }
            QComboBox {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 2px 20px 2px 6px;
                font-size: 11.5px;
                background: white;
                min-height: 24px;
            }
        """)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        self.tabs = QTabWidget()
        layout.addWidget(self.tabs)

        self._setup_home_tab()
        self._setup_insert_tab()
        self._setup_layout_tab()
        self._setup_admin_tab()

    def _create_separator(self) -> QFrame:
        line = QFrame()
        line.setFrameShape(QFrame.Shape.VLine)
        line.setFrameShadow(QFrame.Shadow.Sunken)
        line.setStyleSheet("color: #e2e8f0; margin: 4px 6px;")
        return line

    # 1. THẺ VĂN BẢN (HOME)
    def _setup_home_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(10, 4, 10, 4)
        layout.setSpacing(6)

        # Nhóm Tệp & Xuất bản trực tiếp
        file_box = QVBoxLayout()
        file_box.setSpacing(2)

        row_f1 = QHBoxLayout()
        row_f1.setSpacing(4)
        self.btn_ribbon_save = QToolButton()
        self.btn_ribbon_save.setText("Lưu (.docx)")
        self.btn_ribbon_save.setIcon(get_ui_icon("file_word"))
        self.btn_ribbon_save.setIconSize(QSize(16, 16))
        self.btn_ribbon_save.setStyleSheet("font-weight: 600; color: #2563eb; background: #eff6ff; border: 1px solid #bfdbfe;")
        self.btn_ribbon_save.clicked.connect(self.save_document_requested.emit)
        row_f1.addWidget(self.btn_ribbon_save)

        self.btn_ribbon_pdf = QToolButton()
        self.btn_ribbon_pdf.setText("Xuất PDF")
        self.btn_ribbon_pdf.setIcon(get_ui_icon("file_pdf"))
        self.btn_ribbon_pdf.setIconSize(QSize(16, 16))
        self.btn_ribbon_pdf.setStyleSheet("font-weight: 600; color: #dc2626; background: #fef2f2; border: 1px solid #fecaca;")
        self.btn_ribbon_pdf.clicked.connect(self.export_pdf_requested.emit)
        row_f1.addWidget(self.btn_ribbon_pdf)
        file_box.addLayout(row_f1)

        self.btn_ribbon_sign = QToolButton()
        self.btn_ribbon_sign.setText("Ký số điện tử (NĐ 30)")
        self.btn_ribbon_sign.setIcon(get_ui_icon("digital_sign"))
        self.btn_ribbon_sign.setIconSize(QSize(16, 16))
        self.btn_ribbon_sign.setStyleSheet("font-weight: 600; color: #7c3aed; background: #f5f3ff; border: 1px solid #ddd6fe;")
        self.btn_ribbon_sign.clicked.connect(self.sign_document_requested.emit)
        file_box.addWidget(self.btn_ribbon_sign)

        layout.addLayout(file_box)
        layout.addWidget(self._create_separator())

        # Phông chữ
        font_box = QVBoxLayout()
        font_box.setSpacing(3)

        row_pickers = QHBoxLayout()
        self.cb_font_family = QComboBox()
        self.cb_font_family.addItems(["Times New Roman", "Arial", "Calibri"])
        self.cb_font_family.setFixedWidth(140)
        self.cb_font_family.currentTextChanged.connect(self.font_family_changed.emit)
        row_pickers.addWidget(self.cb_font_family)

        self.cb_font_size = QComboBox()
        self.cb_font_size.setFixedWidth(54)
        for s in [10, 11, 12, 13, 14, 15, 16, 18, 20]:
            self.cb_font_size.addItem(str(s))
        self.cb_font_size.setCurrentText("13")
        self.cb_font_size.currentTextChanged.connect(self._on_font_size_changed)
        row_pickers.addWidget(self.cb_font_size)
        font_box.addLayout(row_pickers)

        row_styles = QHBoxLayout()
        self.btn_bold = QToolButton()
        self.btn_bold.setText("B")
        self.btn_bold.setCheckable(True)
        self.btn_bold.setFont(QFont("Times New Roman", 10, QFont.Weight.Bold))
        self.btn_bold.clicked.connect(self.format_bold_triggered.emit)
        row_styles.addWidget(self.btn_bold)

        self.btn_italic = QToolButton()
        self.btn_italic.setText("I")
        self.btn_italic.setCheckable(True)
        font_i = QFont("Times New Roman", 10)
        font_i.setItalic(True)
        self.btn_italic.setFont(font_i)
        self.btn_italic.clicked.connect(self.format_italic_triggered.emit)
        row_styles.addWidget(self.btn_italic)

        self.btn_underline = QToolButton()
        self.btn_underline.setText("U")
        self.btn_underline.setCheckable(True)
        font_u = QFont("Times New Roman", 10)
        font_u.setUnderline(True)
        self.btn_underline.setFont(font_u)
        self.btn_underline.clicked.connect(self.format_underline_triggered.emit)
        row_styles.addWidget(self.btn_underline)

        self.btn_color = QToolButton()
        self.btn_color.setText("A")
        self.btn_color.setToolTip("Màu chữ")
        self.btn_color.setStyleSheet("color: #2563eb; font-weight: bold;")
        self.btn_color.clicked.connect(self._choose_text_color)
        row_styles.addWidget(self.btn_color)

        font_box.addLayout(row_styles)
        layout.addLayout(font_box)
        layout.addWidget(self._create_separator())

        # Căn lề & Giãn dòng
        para_box = QVBoxLayout()
        para_box.setSpacing(3)

        row_align = QHBoxLayout()
        self.btn_left = QToolButton()
        self.btn_left.setIcon(get_ui_icon("align_left"))
        self.btn_left.setIconSize(QSize(16, 16))
        self.btn_left.setToolTip("Căn trái")
        self.btn_left.clicked.connect(self.align_left_triggered.emit)
        row_align.addWidget(self.btn_left)

        self.btn_center = QToolButton()
        self.btn_center.setIcon(get_ui_icon("align_center"))
        self.btn_center.setIconSize(QSize(16, 16))
        self.btn_center.setToolTip("Căn giữa")
        self.btn_center.clicked.connect(self.align_center_triggered.emit)
        row_align.addWidget(self.btn_center)

        self.btn_right = QToolButton()
        self.btn_right.setIcon(get_ui_icon("align_right"))
        self.btn_right.setIconSize(QSize(16, 16))
        self.btn_right.setToolTip("Căn phải")
        self.btn_right.clicked.connect(self.align_right_triggered.emit)
        row_align.addWidget(self.btn_right)

        self.btn_justify = QToolButton()
        self.btn_justify.setIcon(get_ui_icon("align_justify"))
        self.btn_justify.setIconSize(QSize(16, 16))
        self.btn_justify.setToolTip("Căn đều hai bên")
        self.btn_justify.setChecked(True)
        self.btn_justify.clicked.connect(self.align_justify_triggered.emit)
        row_align.addWidget(self.btn_justify)
        para_box.addLayout(row_align)

        row_spacing = QHBoxLayout()
        lbl_sp = QLabel("Giãn dòng:")
        lbl_sp.setStyleSheet("font-size: 11px; color: #64748b;")
        row_spacing.addWidget(lbl_sp)

        self.cb_line_spacing = QComboBox()
        self.cb_line_spacing.setFixedWidth(56)
        self.cb_line_spacing.addItems(["1.0", "1.15", "1.2", "1.25", "1.5"])
        self.cb_line_spacing.setCurrentText("1.2")
        self.cb_line_spacing.currentTextChanged.connect(self._on_line_spacing_changed)
        row_spacing.addWidget(self.cb_line_spacing)

        btn_indent = QToolButton()
        btn_indent.setText("Thụt lề 1.27cm")
        btn_indent.clicked.connect(lambda: self.indent_first_line_triggered.emit(12.7))
        row_spacing.addWidget(btn_indent)

        para_box.addLayout(row_spacing)
        layout.addLayout(para_box)
        layout.addWidget(self._create_separator())

        # Danh sách
        list_box = QVBoxLayout()
        list_box.setSpacing(3)
        btn_bullet = QToolButton()
        btn_bullet.setText("Gạch đầu dòng")
        btn_bullet.clicked.connect(self.bullet_list_triggered.emit)
        list_box.addWidget(btn_bullet)

        btn_num = QToolButton()
        btn_num.setText("Đánh số tự động")
        btn_num.clicked.connect(self.numbered_list_triggered.emit)
        list_box.addWidget(btn_num)
        layout.addLayout(list_box)

        layout.addStretch()
        self.tabs.addTab(tab, "Văn bản")

    # 2. THẺ CHÈN (INSERT)
    def _setup_insert_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(10, 4, 10, 4)
        layout.setSpacing(6)

        btn_table = QToolButton()
        btn_table.setIcon(get_ui_icon("table"))
        btn_table.setIconSize(QSize(18, 18))
        btn_table.setText("Chèn bảng")
        btn_table.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_table.clicked.connect(self.insert_table_requested.emit)
        layout.addWidget(btn_table)

        layout.addWidget(self._create_separator())

        # Thao tác bảng
        tbl_ops = QVBoxLayout()
        tbl_ops.setSpacing(2)

        row_1 = QHBoxLayout()
        row_1.setSpacing(3)
        b_ra = QToolButton()
        b_ra.setText("+ Hàng trên")
        b_ra.clicked.connect(self.table_insert_row_above_triggered.emit)
        row_1.addWidget(b_ra)

        b_rb = QToolButton()
        b_rb.setText("+ Hàng dưới")
        b_rb.clicked.connect(self.table_insert_row_below_triggered.emit)
        row_1.addWidget(b_rb)

        b_dr = QToolButton()
        b_dr.setText("Xóa hàng")
        b_dr.clicked.connect(self.table_delete_row_triggered.emit)
        row_1.addWidget(b_dr)
        tbl_ops.addLayout(row_1)

        row_2 = QHBoxLayout()
        row_2.setSpacing(3)
        b_cl = QToolButton()
        b_cl.setText("+ Cột trái")
        b_cl.clicked.connect(self.table_insert_col_left_triggered.emit)
        row_2.addWidget(b_cl)

        b_cr = QToolButton()
        b_cr.setText("+ Cột phải")
        b_cr.clicked.connect(self.table_insert_col_right_triggered.emit)
        row_2.addWidget(b_cr)

        b_dc = QToolButton()
        b_dc.setText("Xóa cột")
        b_dc.clicked.connect(self.table_delete_col_triggered.emit)
        row_2.addWidget(b_dc)

        b_mg = QToolButton()
        b_mg.setText("Trộn ô")
        b_mg.clicked.connect(self.table_merge_cells_triggered.emit)
        row_2.addWidget(b_mg)

        b_sp = QToolButton()
        b_sp.setText("Tách ô")
        b_sp.clicked.connect(self.table_split_cell_triggered.emit)
        row_2.addWidget(b_sp)

        tbl_ops.addLayout(row_2)
        layout.addLayout(tbl_ops)

        layout.addWidget(self._create_separator())

        btn_img = QToolButton()
        btn_img.setIcon(get_ui_icon("image"))
        btn_img.setIconSize(QSize(18, 18))
        btn_img.setText("Hình ảnh")
        btn_img.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_img.clicked.connect(self.insert_image_requested.emit)
        layout.addWidget(btn_img)

        btn_break = QToolButton()
        btn_break.setIcon(get_ui_icon("page_break"))
        btn_break.setIconSize(QSize(18, 18))
        btn_break.setText("Ngắt trang")
        btn_break.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_break.clicked.connect(self.page_break_requested.emit)
        layout.addWidget(btn_break)

        layout.addWidget(self._create_separator())

        # Ký hiệu
        sym_box = QHBoxLayout()
        for sym in ["§", "°", "©", "®", "→"]:
            b_s = QToolButton()
            b_s.setText(sym)
            b_s.setFixedSize(24, 24)
            b_s.clicked.connect(lambda checked, s=sym: self.insert_symbol_requested.emit(s))
            sym_box.addWidget(b_s)
        layout.addLayout(sym_box)

        layout.addStretch()
        self.tabs.addTab(tab, "Chèn")

    # 3. THẺ BỐ CỤC (LAYOUT)
    def _setup_layout_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(8)

        btn_p = QToolButton()
        btn_p.setText("Trang dọc")
        btn_p.setCheckable(True)
        btn_p.setChecked(True)
        btn_p.clicked.connect(lambda: self.page_orientation_changed.emit("portrait"))
        layout.addWidget(btn_p)

        btn_l = QToolButton()
        btn_l.setText("Trang ngang")
        btn_l.setCheckable(True)
        btn_l.clicked.connect(lambda: self.page_orientation_changed.emit("landscape"))
        layout.addWidget(btn_l)

        grp = QButtonGroup(self)
        grp.addButton(btn_p)
        grp.addButton(btn_l)

        layout.addWidget(self._create_separator())

        btn_m = QToolButton()
        btn_m.setText("Lề chuẩn NĐ 30")
        btn_m.clicked.connect(lambda: self.margins_preset_triggered.emit("nd30"))
        layout.addWidget(btn_m)

        layout.addWidget(self._create_separator())

        btn_ruler = QToolButton()
        btn_ruler.setText("Thước đo (Ruler)")
        btn_ruler.setCheckable(True)
        btn_ruler.setChecked(True)
        btn_ruler.toggled.connect(self.toggle_ruler_triggered.emit)
        layout.addWidget(btn_ruler)

        layout.addStretch()
        self.tabs.addTab(tab, "Bố cục")

    # 4. THẺ THỂ THỨC (ADMIN)
    def _setup_admin_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(8)

        btn_meta = QToolButton()
        btn_meta.setIcon(get_ui_icon("edit"))
        btn_meta.setIconSize(QSize(16, 16))
        btn_meta.setText("Sửa thể thức (F2)")
        btn_meta.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_meta.clicked.connect(self.edit_metadata_requested.emit)
        layout.addWidget(btn_meta)

        layout.addWidget(self._create_separator())

        lbl_t = QLabel("Chuyển mẫu:")
        lbl_t.setStyleSheet("font-size: 11px; color: #64748b;")
        layout.addWidget(lbl_t)

        self.cb_type = QComboBox()
        self.cb_type.setFixedWidth(140)
        self.cb_type.addItems([
            "QUYẾT ĐỊNH", "CÔNG VĂN", "TỜ TRÌNH", "THÔNG BÁO",
            "NGHỊ QUYẾT ĐẢNG", "CHỈ THỊ ĐẢNG"
        ])
        self.cb_type.currentTextChanged.connect(self.switch_doc_type_requested.emit)
        layout.addWidget(self.cb_type)

        layout.addWidget(self._create_separator())

        self.btn_lock = QToolButton()
        self.btn_lock.setText("Khóa thể thức")
        self.btn_lock.setCheckable(True)
        self.btn_lock.setChecked(True)
        self.btn_lock.toggled.connect(self._on_lock_toggled)
        layout.addWidget(self.btn_lock)

        layout.addWidget(self._create_separator())

        btn_audit = QToolButton()
        btn_audit.setText("Kiểm tra chuẩn thể thức")
        btn_audit.clicked.connect(self.check_compliance_requested.emit)
        layout.addWidget(btn_audit)

        layout.addStretch()
        self.tabs.addTab(tab, "Thể thức")

    def _on_font_size_changed(self, text: str):
        try:
            self.font_size_changed.emit(float(text))
        except ValueError:
            pass

    def _choose_text_color(self):
        color = QColorDialog.getColor(QColor(0, 0, 0), self, "Màu chữ")
        if color.isValid():
            self.font_color_changed.emit(color)

    def _on_line_spacing_changed(self, text: str):
        try:
            self.line_spacing_changed.emit(float(text))
        except ValueError:
            pass

    def _on_lock_toggled(self, checked: bool):
        self.btn_lock.setText("Khóa thể thức" if checked else "Mở khóa thể thức")
        self.lock_scaffolding_toggled.emit(checked)
