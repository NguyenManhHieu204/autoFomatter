"""
ui/editor/administrative_ribbon.py: Dải băng điều khiển Hành chính (Administrative Ribbon Bar)
Thiết kế tối giản, hiện đại, tái lập trải nghiệm văn phòng quen thuộc:
- Thẻ 1: Văn bản (Home) - Phông chữ, Cỡ chữ, Định dạng, Căn lề, Giãn dòng, Danh sách
- Thẻ 2: Chèn (Insert) - Bảng biểu, Thao tác Hàng/Cột/Trộn ô, Chèn Ảnh, Ngắt trang
- Thẻ 3: Bố cục (Page Layout) - Hướng giấy, Lề in NĐ 30, Bật/Tắt Thước đo
- Thẻ 4: Thể thức (Admin Tools) - Sửa Thể thức (F2), Khóa khung, Rà soát tuân thủ
"""

import os
from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget,
    QComboBox, QToolButton, QLabel, QFrame, QButtonGroup,
    QColorDialog, QMenu
)
from PyQt6.QtGui import QFont, QColor, QIcon
from PyQt6.QtCore import Qt, pyqtSignal, QSize


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
    increase_indent_triggered = pyqtSignal()
    decrease_indent_triggered = pyqtSignal()
    remove_space_before_triggered = pyqtSignal()
    remove_space_after_triggered = pyqtSignal()
    remove_all_spacing_triggered = pyqtSignal()
    add_space_before_triggered = pyqtSignal(float)
    add_space_after_triggered = pyqtSignal(float)
    apply_nd30_spacing_triggered = pyqtSignal()

    bullet_list_triggered = pyqtSignal()
    numbered_list_triggered = pyqtSignal()
    apply_list_format_triggered = pyqtSignal(dict)

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
        self.setFixedHeight(112)
        self.icons_dir = os.path.join(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))), "assets", "icons")
        self.setStyleSheet("""
            QWidget {
                background-color: #ffffff;
                color: #0f172a;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QTabWidget::pane {
                border-top: 2px solid #cbd5e1;
                background-color: #ffffff;
            }
            QTabBar {
                background-color: #f1f5f9;
            }
            QTabBar::tab {
                background-color: #f1f5f9;
                color: #1e293b;
                padding: 7px 18px;
                border: 1px solid #cbd5e1;
                border-bottom: none;
                margin-right: 3px;
                border-top-left-radius: 5px;
                border-top-right-radius: 5px;
                font-size: 12.5px;
                font-weight: 600;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #2563eb;
                border-top: 3px solid #2563eb;
                font-weight: 700;
            }
            QTabBar::tab:hover {
                background-color: #ffffff;
                color: #1d4ed8;
            }
            QToolButton {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 4px 6px;
                background-color: #ffffff;
                font-size: 12px;
                font-weight: 600;
                color: #0f172a;
            }
            QToolButton:hover {
                background-color: #eff6ff;
                border-color: #2563eb;
                color: #1d4ed8;
            }
            QToolButton:checked {
                background-color: #dbeafe;
                border-color: #1d4ed8;
                color: #1e40af;
                font-weight: 700;
            }
            QLabel {
                color: #0f172a;
                font-size: 12px;
                font-weight: 600;
                background: transparent;
            }
            QComboBox {
                border: 1.5px solid #cbd5e1;
                border-radius: 4px;
                padding: 3px 22px 3px 8px;
                font-size: 12px;
                font-weight: 600;
                color: #0f172a;
                background-color: #ffffff;
                min-height: 26px;
            }
            QComboBox:focus {
                border-color: #2563eb;
            }
            QMenu {
                background-color: #ffffff;
                color: #0f172a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 20px 6px 10px;
                border-radius: 4px;
                font-size: 12px;
            }
            QMenu::item:selected {
                background-color: #eff6ff;
                color: #2563eb;
            }
            QMenu::separator {
                height: 1px;
                background-color: #e2e8f0;
                margin: 4px 6px;
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

    def _icon(self, name: str) -> QIcon:
        p = os.path.join(self.icons_dir, name)
        return QIcon(p) if os.path.exists(p) else QIcon()

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
        layout.setSpacing(8)

        # 1. Nhóm Phông chữ & Cỡ chữ (Typography)
        font_box = QVBoxLayout()
        font_box.setSpacing(4)

        row_pickers = QHBoxLayout()
        row_pickers.setSpacing(6)
        self.cb_font_family = QComboBox()
        self.cb_font_family.addItems(["Times New Roman", "Arial", "Calibri"])
        self.cb_font_family.setFixedWidth(160)
        self.cb_font_family.currentTextChanged.connect(self.font_family_changed.emit)
        row_pickers.addWidget(self.cb_font_family)

        self.cb_font_size = QComboBox()
        self.cb_font_size.setFixedWidth(60)
        for s in [10, 11, 12, 13, 14, 15, 16, 18, 20]:
            self.cb_font_size.addItem(str(s))
        self.cb_font_size.setCurrentText("13")
        self.cb_font_size.currentTextChanged.connect(self._on_font_size_changed)
        row_pickers.addWidget(self.cb_font_size)
        font_box.addLayout(row_pickers)

        row_styles = QHBoxLayout()
        row_styles.setSpacing(3)
        self.btn_bold = QToolButton()
        self.btn_bold.setText("B")
        self.btn_bold.setFixedSize(28, 26)
        self.btn_bold.setCheckable(True)
        self.btn_bold.setFont(QFont("Times New Roman", 10, QFont.Weight.Bold))
        self.btn_bold.setToolTip("In đậm (Ctrl+B)")
        self.btn_bold.clicked.connect(self.format_bold_triggered.emit)
        row_styles.addWidget(self.btn_bold)

        self.btn_italic = QToolButton()
        self.btn_italic.setText("I")
        self.btn_italic.setFixedSize(28, 26)
        self.btn_italic.setCheckable(True)
        font_i = QFont("Times New Roman", 10)
        font_i.setItalic(True)
        self.btn_italic.setFont(font_i)
        self.btn_italic.setToolTip("In nghiêng (Ctrl+I)")
        self.btn_italic.clicked.connect(self.format_italic_triggered.emit)
        row_styles.addWidget(self.btn_italic)

        self.btn_underline = QToolButton()
        self.btn_underline.setText("U")
        self.btn_underline.setFixedSize(28, 26)
        self.btn_underline.setCheckable(True)
        font_u = QFont("Times New Roman", 10)
        font_u.setUnderline(True)
        self.btn_underline.setFont(font_u)
        self.btn_underline.setToolTip("Gạch chân (Ctrl+U)")
        self.btn_underline.clicked.connect(self.format_underline_triggered.emit)
        row_styles.addWidget(self.btn_underline)

        self.btn_color = QToolButton()
        self.btn_color.setText("A")
        self.btn_color.setFixedSize(28, 26)
        self.btn_color.setToolTip("Màu chữ")
        self.btn_color.setStyleSheet("color: #2563eb; font-weight: bold;")
        self.btn_color.clicked.connect(self._choose_text_color)
        row_styles.addWidget(self.btn_color)

        font_box.addLayout(row_styles)
        layout.addLayout(font_box)
        layout.addWidget(self._create_separator())

        # 2. Nhóm Căn lề, Thụt lề & Giãn dòng/đoạn (Paragraph)
        para_box = QVBoxLayout()
        para_box.setSpacing(4)

        # Dòng 1: Căn lề, Thụt lề, Danh sách biểu tượng chuẩn Word
        row_align = QHBoxLayout()
        row_align.setSpacing(3)

        self.btn_left = QToolButton()
        self.btn_left.setIcon(self._icon("align_left.svg"))
        self.btn_left.setIconSize(QSize(18, 18))
        self.btn_left.setFixedSize(28, 26)
        self.btn_left.setToolTip("Căn lề trái (Ctrl+L)")
        self.btn_left.clicked.connect(self.align_left_triggered.emit)
        row_align.addWidget(self.btn_left)

        self.btn_center = QToolButton()
        self.btn_center.setIcon(self._icon("align_center.svg"))
        self.btn_center.setIconSize(QSize(18, 18))
        self.btn_center.setFixedSize(28, 26)
        self.btn_center.setToolTip("Căn giữa (Ctrl+E)")
        self.btn_center.clicked.connect(self.align_center_triggered.emit)
        row_align.addWidget(self.btn_center)

        self.btn_right = QToolButton()
        self.btn_right.setIcon(self._icon("align_right.svg"))
        self.btn_right.setIconSize(QSize(18, 18))
        self.btn_right.setFixedSize(28, 26)
        self.btn_right.setToolTip("Căn lề phải (Ctrl+R)")
        self.btn_right.clicked.connect(self.align_right_triggered.emit)
        row_align.addWidget(self.btn_right)

        self.btn_justify = QToolButton()
        self.btn_justify.setIcon(self._icon("align_justify.svg"))
        self.btn_justify.setIconSize(QSize(18, 18))
        self.btn_justify.setFixedSize(28, 26)
        self.btn_justify.setToolTip("Căn đều hai bên (Ctrl+J)")
        self.btn_justify.setChecked(True)
        self.btn_justify.clicked.connect(self.align_justify_triggered.emit)
        row_align.addWidget(self.btn_justify)

        # Nút giảm / tăng thụt lề chuẩn Word
        self.btn_dec_indent = QToolButton()
        self.btn_dec_indent.setIcon(self._icon("indent_decrease.svg"))
        self.btn_dec_indent.setIconSize(QSize(18, 18))
        self.btn_dec_indent.setFixedSize(28, 26)
        self.btn_dec_indent.setToolTip("Giảm thụt lề cả đoạn")
        self.btn_dec_indent.clicked.connect(self.decrease_indent_triggered.emit)
        row_align.addWidget(self.btn_dec_indent)

        self.btn_inc_indent = QToolButton()
        self.btn_inc_indent.setIcon(self._icon("indent_increase.svg"))
        self.btn_inc_indent.setIconSize(QSize(18, 18))
        self.btn_inc_indent.setFixedSize(28, 26)
        self.btn_inc_indent.setToolTip("Tăng thụt lề cả đoạn")
        self.btn_inc_indent.clicked.connect(self.increase_indent_triggered.emit)
        row_align.addWidget(self.btn_inc_indent)

        # Nút danh sách Bullet (Kèm menu chọn kiểu ký hiệu)
        self.btn_bullet = QToolButton()
        self.btn_bullet.setIcon(self._icon("list_bullet.svg"))
        self.btn_bullet.setIconSize(QSize(18, 18))
        self.btn_bullet.setFixedSize(38, 26)
        self.btn_bullet.setToolTip("Danh sách ký hiệu đầu dòng (Nhấp để bật/tắt, bấm mũi tên để chọn kiểu)")
        self.btn_bullet.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        self.btn_bullet.clicked.connect(self.bullet_list_triggered.emit)

        menu_bullet = QMenu(self.btn_bullet)
        menu_bullet.setStyleSheet("QMenu { background-color: #ffffff; color: #0f172a; border: 1px solid #cbd5e1; font-size: 12px; } QMenu::item:selected { background-color: #eff6ff; color: #1d4ed8; }")
        act_b_disc = menu_bullet.addAction("•   Dấu chấm tròn đặc (Disc)")
        act_b_disc.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "disc"}))
        act_b_circle = menu_bullet.addAction("○   Dấu tròn rỗng (Circle)")
        act_b_circle.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "circle"}))
        act_b_square = menu_bullet.addAction("■   Dấu ô vuông (Square)")
        act_b_square.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "square"}))
        self.btn_bullet.setMenu(menu_bullet)
        row_align.addWidget(self.btn_bullet)

        # Nút Đánh số tự động (Kèm menu phân cấp, La Mã, chữ cái, điều khoản)
        self.btn_num = QToolButton()
        self.btn_num.setIcon(self._icon("list_numbers.svg"))
        self.btn_num.setIconSize(QSize(18, 18))
        self.btn_num.setFixedSize(38, 26)
        self.btn_num.setToolTip("Đánh số tự động (Nhấp để bật/tắt, bấm mũi tên để chọn kiểu)")
        self.btn_num.setPopupMode(QToolButton.ToolButtonPopupMode.MenuButtonPopup)
        self.btn_num.clicked.connect(self.numbered_list_triggered.emit)

        menu_num = QMenu(self.btn_num)
        menu_num.setStyleSheet("QMenu { background-color: #ffffff; color: #0f172a; border: 1px solid #cbd5e1; font-size: 12px; } QMenu::item:selected { background-color: #eff6ff; color: #1d4ed8; }")
        
        act_num_dec = menu_num.addAction("1.  2.  3.  (Số Ả Rập chuẩn)")
        act_num_dec.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "decimal", "suffix": "."}))

        act_num_paren = menu_num.addAction("1)  2)  3)  (Số đóng ngoặc đơn)")
        act_num_paren.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "decimal", "suffix": ")"}))

        act_num_roman_u = menu_num.addAction("I.  II.  III.  (Số La Mã in hoa)")
        act_num_roman_u.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "roman_upper", "suffix": "."}))

        act_num_roman_l = menu_num.addAction("i.  ii.  iii.  (Số La Mã chữ thường)")
        act_num_roman_l.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "roman_lower", "suffix": "."}))

        act_num_alpha_l = menu_num.addAction("a)  b)  c)  (Điểm quy định — NĐ 30)")
        act_num_alpha_l.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "alpha_lower", "suffix": ")"}))

        act_num_alpha_u = menu_num.addAction("A.  B.  C.  (Chữ cái in hoa)")
        act_num_alpha_u.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "alpha_upper", "suffix": "."}))

        act_num_clause = menu_num.addAction("Điều 1.  Điều 2.  (Điều khoản văn bản)")
        act_num_clause.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "decimal", "prefix": "Điều ", "suffix": "."}))

        act_num_multi = menu_num.addAction("I. -> 1. -> a) -> -  (Phân cấp đa cấp)")
        act_num_multi.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "multilevel"}))

        menu_num.addSeparator()

        act_num_cont = menu_num.addAction("Tiếp tục đánh số từ danh sách trước")
        act_num_cont.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "decimal", "suffix": ".", "restart": False}))

        act_num_restart = menu_num.addAction("Bắt đầu lại từ 1 (Restart numbering)")
        act_num_restart.triggered.connect(lambda: self.apply_list_format_triggered.emit({"style": "decimal", "suffix": ".", "restart": True}))

        self.btn_num.setMenu(menu_num)
        row_align.addWidget(self.btn_num)

        para_box.addLayout(row_align)

        # Dòng 2: Giãn dòng & Giãn đoạn chuẩn Microsoft Word
        row_spacing = QHBoxLayout()
        row_spacing.setSpacing(6)
        lbl_sp = QLabel("Giãn dòng:")
        lbl_sp.setStyleSheet("font-size: 11.5px; color: #475569;")
        row_spacing.addWidget(lbl_sp)

        self.cb_line_spacing = QComboBox()
        self.cb_line_spacing.setFixedWidth(78)
        self.cb_line_spacing.addItems(["1.0", "1.15", "1.2", "1.25", "1.5", "2.0"])
        self.cb_line_spacing.setCurrentText("1.25")
        self.cb_line_spacing.currentTextChanged.connect(self._on_line_spacing_changed)
        row_spacing.addWidget(self.cb_line_spacing)

        # Menu Thụt đầu dòng (First Line Indent)
        self.btn_first_line = QToolButton()
        self.btn_first_line.setText("Thụt dòng đầu")
        self.btn_first_line.setToolTip("Quy chuẩn thụt dòng đầu đoạn văn")
        self.btn_first_line.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        menu_fl = QMenu(self.btn_first_line)
        act_fl_10 = menu_fl.addAction("Thụt dòng đầu 1.0 cm (Nghị định 30)")
        act_fl_10.triggered.connect(lambda: self.indent_first_line_triggered.emit(10.0))
        act_fl_127 = menu_fl.addAction("Thụt dòng đầu 1.27 cm (Chuẩn Word 0.5\")")
        act_fl_127.triggered.connect(lambda: self.indent_first_line_triggered.emit(12.7))
        act_fl_none = menu_fl.addAction("Không thụt dòng (0 cm)")
        act_fl_none.triggered.connect(lambda: self.indent_first_line_triggered.emit(0.0))
        self.btn_first_line.setMenu(menu_fl)
        row_spacing.addWidget(self.btn_first_line)

        # Nút menu Giãn cách đoạn văn
        self.btn_paragraph_spacing = QToolButton()
        self.btn_paragraph_spacing.setText("Khoảng cách đoạn")
        self.btn_paragraph_spacing.setToolTip("Điều chỉnh khoảng cách phía trên và phía dưới đoạn văn")
        self.btn_paragraph_spacing.setPopupMode(QToolButton.ToolButtonPopupMode.InstantPopup)
        self.btn_paragraph_spacing.setStyleSheet("font-weight: 500;")

        menu_sp = QMenu(self.btn_paragraph_spacing)
        act_rm_before = menu_sp.addAction("Xóa khoảng cách trên đoạn (0 pt)")
        act_rm_before.triggered.connect(self.remove_space_before_triggered.emit)

        act_rm_after = menu_sp.addAction("Xóa khoảng cách dưới đoạn (0 pt)")
        act_rm_after.triggered.connect(self.remove_space_after_triggered.emit)

        act_rm_all = menu_sp.addAction("Xóa toàn bộ khoảng cách trên & dưới (0 pt / 0 pt)")
        act_rm_all.triggered.connect(self.remove_all_spacing_triggered.emit)

        menu_sp.addSeparator()

        act_add_before = menu_sp.addAction("Thêm 6 pt khoảng cách trước đoạn")
        act_add_before.triggered.connect(lambda: self.add_space_before_triggered.emit(6.0))

        act_add_after = menu_sp.addAction("Thêm 6 pt khoảng cách sau đoạn")
        act_add_after.triggered.connect(lambda: self.add_space_after_triggered.emit(6.0))

        menu_sp.addSeparator()

        act_nd30 = menu_sp.addAction("Chuẩn hóa Nghị định 30 (Dòng: 1.25, Trước: 0 pt, Sau: 6 pt)")
        act_nd30.triggered.connect(self.apply_nd30_spacing_triggered.emit)

        self.btn_paragraph_spacing.setMenu(menu_sp)
        row_spacing.addWidget(self.btn_paragraph_spacing)

        para_box.addLayout(row_spacing)
        layout.addLayout(para_box)
        layout.addWidget(self._create_separator())

        # 3. Nút căn chuẩn NĐ 30 trực tiếp ngay trên thanh Văn bản
        quick_box = QVBoxLayout()
        quick_box.setSpacing(4)
        btn_quick_align = QToolButton()
        btn_quick_align.setIcon(self._icon("margins.svg"))
        btn_quick_align.setIconSize(QSize(20, 20))
        btn_quick_align.setText("Căn chuẩn NĐ 30")
        btn_quick_align.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_quick_align.setStyleSheet("""
            QToolButton {
                font-weight: 700;
                color: #1e40af;
                background-color: #eff6ff;
                border: 1.5px solid #3b82f6;
                padding: 6px 10px;
                border-radius: 4px;
            }
            QToolButton:hover {
                background-color: #dbeafe;
                border-color: #1d4ed8;
            }
        """)
        btn_quick_align.setToolTip("Tự động căn lại lề in NĐ 30 (30-15-20-20 mm), xóa bỏ thụt lề thủ công, căn đều hai bên (Justify)")
        btn_quick_align.clicked.connect(lambda: self.margins_preset_triggered.emit("nd30"))
        quick_box.addWidget(btn_quick_align)
        layout.addLayout(quick_box)

        layout.addStretch()
        self.tabs.addTab(tab, "Văn bản")

    # 2. THẺ CHÈN (INSERT)
    def _setup_insert_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 4, 12, 4)
        layout.setSpacing(10)

        # 1. Nhóm Bảng biểu
        btn_table = QToolButton()
        btn_table.setIcon(self._icon("table.svg"))
        btn_table.setIconSize(QSize(22, 22))
        btn_table.setText("Bảng biểu")
        btn_table.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_table.setFixedSize(76, 56)
        btn_table.setToolTip("Chèn bảng mới vào vị trí con trỏ")
        btn_table.clicked.connect(self.insert_table_requested.emit)
        layout.addWidget(btn_table)

        # Thao tác cấu trúc bảng: 2 hàng cân đối (mỗi hàng 4 nút kích thước 28x26)
        tbl_ops_widget = QWidget()
        tbl_ops = QVBoxLayout(tbl_ops_widget)
        tbl_ops.setContentsMargins(0, 0, 0, 0)
        tbl_ops.setSpacing(3)

        # Hàng 1: Thêm hàng và cột (4 nút)
        row_1 = QHBoxLayout()
        row_1.setContentsMargins(0, 0, 0, 0)
        row_1.setSpacing(3)

        b_ra = QToolButton()
        b_ra.setIcon(self._icon("table_row_add_above.svg"))
        b_ra.setIconSize(QSize(18, 18))
        b_ra.setFixedSize(28, 26)
        b_ra.setToolTip("Thêm hàng phía trên")
        b_ra.clicked.connect(self.table_insert_row_above_triggered.emit)
        row_1.addWidget(b_ra)

        b_rb = QToolButton()
        b_rb.setIcon(self._icon("table_row_add_below.svg"))
        b_rb.setIconSize(QSize(18, 18))
        b_rb.setFixedSize(28, 26)
        b_rb.setToolTip("Thêm hàng phía dưới")
        b_rb.clicked.connect(self.table_insert_row_below_triggered.emit)
        row_1.addWidget(b_rb)

        b_cl = QToolButton()
        b_cl.setIcon(self._icon("table_col_add_left.svg"))
        b_cl.setIconSize(QSize(18, 18))
        b_cl.setFixedSize(28, 26)
        b_cl.setToolTip("Thêm cột bên trái")
        b_cl.clicked.connect(self.table_insert_col_left_triggered.emit)
        row_1.addWidget(b_cl)

        b_cr = QToolButton()
        b_cr.setIcon(self._icon("table_col_add_right.svg"))
        b_cr.setIconSize(QSize(18, 18))
        b_cr.setFixedSize(28, 26)
        b_cr.setToolTip("Thêm cột bên phải")
        b_cr.clicked.connect(self.table_insert_col_right_triggered.emit)
        row_1.addWidget(b_cr)
        tbl_ops.addLayout(row_1)

        # Hàng 2: Xóa hàng, cột và Trộn/Tách ô (4 nút đối xứng)
        row_2 = QHBoxLayout()
        row_2.setContentsMargins(0, 0, 0, 0)
        row_2.setSpacing(3)

        b_dr = QToolButton()
        b_dr.setIcon(self._icon("table_row_delete.svg"))
        b_dr.setIconSize(QSize(18, 18))
        b_dr.setFixedSize(28, 26)
        b_dr.setToolTip("Xóa hàng đang chọn")
        b_dr.clicked.connect(self.table_delete_row_triggered.emit)
        row_2.addWidget(b_dr)

        b_dc = QToolButton()
        b_dc.setIcon(self._icon("table_col_delete.svg"))
        b_dc.setIconSize(QSize(18, 18))
        b_dc.setFixedSize(28, 26)
        b_dc.setToolTip("Xóa cột đang chọn")
        b_dc.clicked.connect(self.table_delete_col_triggered.emit)
        row_2.addWidget(b_dc)

        b_mg = QToolButton()
        b_mg.setIcon(self._icon("table_merge.svg"))
        b_mg.setIconSize(QSize(18, 18))
        b_mg.setFixedSize(28, 26)
        b_mg.setToolTip("Trộn các ô đang chọn (Merge Cells)")
        b_mg.clicked.connect(self.table_merge_cells_triggered.emit)
        row_2.addWidget(b_mg)

        b_sp = QToolButton()
        b_sp.setIcon(self._icon("table_split.svg"))
        b_sp.setIconSize(QSize(18, 18))
        b_sp.setFixedSize(28, 26)
        b_sp.setToolTip("Tách ô (Split Cell)")
        b_sp.clicked.connect(self.table_split_cell_triggered.emit)
        row_2.addWidget(b_sp)
        tbl_ops.addLayout(row_2)

        layout.addWidget(tbl_ops_widget)

        layout.addWidget(self._create_separator())

        # 2. Nhóm Minh họa & Trang
        btn_img = QToolButton()
        btn_img.setIcon(self._icon("image.svg"))
        btn_img.setIconSize(QSize(22, 22))
        btn_img.setText("Hình ảnh")
        btn_img.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_img.setFixedSize(76, 56)
        btn_img.setToolTip("Chèn hình ảnh từ máy tính")
        btn_img.clicked.connect(self.insert_image_requested.emit)
        layout.addWidget(btn_img)

        btn_break = QToolButton()
        btn_break.setIcon(self._icon("page_break.svg"))
        btn_break.setIconSize(QSize(22, 22))
        btn_break.setText("Ngắt trang")
        btn_break.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_break.setFixedSize(78, 56)
        btn_break.setToolTip("Ngắt trang sang trang A4 mới (Ctrl+Enter)")
        btn_break.clicked.connect(self.page_break_requested.emit)
        layout.addWidget(btn_break)

        layout.addWidget(self._create_separator())

        # 3. Ký hiệu văn bản hành chính
        sym_box = QHBoxLayout()
        sym_box.setSpacing(3)
        for sym in ["§", "°", "©", "®", "→", "•", "✓"]:
            b_s = QToolButton()
            b_s.setText(sym)
            b_s.setFixedSize(26, 26)
            b_s.setToolTip(f"Chèn ký hiệu '{sym}'")
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
        layout.setSpacing(10)

        btn_p = QToolButton()
        btn_p.setIcon(self._icon("orientation_portrait.svg"))
        btn_p.setIconSize(QSize(18, 18))
        btn_p.setText("Trang dọc")
        btn_p.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_p.setCheckable(True)
        btn_p.setChecked(True)
        btn_p.clicked.connect(lambda: self.page_orientation_changed.emit("portrait"))
        layout.addWidget(btn_p)

        btn_l = QToolButton()
        btn_l.setIcon(self._icon("orientation_landscape.svg"))
        btn_l.setIconSize(QSize(18, 18))
        btn_l.setText("Trang ngang")
        btn_l.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_l.setCheckable(True)
        btn_l.clicked.connect(lambda: self.page_orientation_changed.emit("landscape"))
        layout.addWidget(btn_l)

        grp = QButtonGroup(self)
        grp.addButton(btn_p)
        grp.addButton(btn_l)

        layout.addWidget(self._create_separator())

        btn_m = QToolButton()
        btn_m.setIcon(self._icon("margins.svg"))
        btn_m.setIconSize(QSize(20, 20))
        btn_m.setText("Lề chuẩn NĐ 30")
        btn_m.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_m.setStyleSheet("""
            QToolButton {
                font-weight: 700;
                color: #1e40af;
                background-color: #eff6ff;
                border: 1.5px solid #3b82f6;
                padding: 6px 14px;
                border-radius: 4px;
            }
            QToolButton:hover {
                background-color: #dbeafe;
                border-color: #1d4ed8;
            }
        """)
        btn_m.setToolTip("Tự động căn lề (30-15-20-20 mm), xóa bỏ thụt lề thủ công, căn đều hai bên (Justify)")
        btn_m.clicked.connect(lambda: self.margins_preset_triggered.emit("nd30"))
        layout.addWidget(btn_m)

        layout.addWidget(self._create_separator())

        btn_ruler = QToolButton()
        btn_ruler.setIcon(self._icon("ruler.svg"))
        btn_ruler.setIconSize(QSize(18, 18))
        btn_ruler.setText("Thước đo (Ruler)")
        btn_ruler.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
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
        layout.setSpacing(10)

        btn_meta = QToolButton()
        btn_meta.setIcon(self._icon("edit.svg"))
        btn_meta.setIconSize(QSize(18, 18))
        btn_meta.setText("Sửa thể thức (F2)")
        btn_meta.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_meta.setStyleSheet("font-weight: 600; padding: 6px 12px;")
        btn_meta.clicked.connect(self.edit_metadata_requested.emit)
        layout.addWidget(btn_meta)

        layout.addWidget(self._create_separator())

        lbl_t = QLabel("Chuyển mẫu:")
        lbl_t.setStyleSheet("font-size: 11px; color: #64748b;")
        layout.addWidget(lbl_t)

        self.cb_type = QComboBox()
        self.cb_type.setFixedWidth(145)
        self.cb_type.addItems([
            "QUYẾT ĐỊNH", "CÔNG VĂN", "TỜ TRÌNH", "THÔNG BÁO",
            "BÁO CÁO", "NGHỊ QUYẾT ĐẢNG", "CHỈ THỊ ĐẢNG"
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
        btn_audit.setIcon(self._icon("compliance_check.svg"))
        btn_audit.setIconSize(QSize(20, 20))
        btn_audit.setText("Kiểm tra chuẩn thể thức")
        btn_audit.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextBesideIcon)
        btn_audit.setStyleSheet("""
            QToolButton {
                font-weight: 700;
                color: #065f46;
                background-color: #ecfdf5;
                border: 1.5px solid #10b981;
                padding: 6px 14px;
                border-radius: 4px;
            }
            QToolButton:hover {
                background-color: #d1fae5;
                border-color: #059669;
            }
        """)
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
