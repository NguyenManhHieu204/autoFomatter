"""
ui/editor/administrative_ribbon.py: Dải băng điều khiển Hành chính (Administrative Ribbon Interface)
Tái lập toàn diện trải nghiệm của Microsoft Word theo 4 thẻ chuyên biệt:
- Thẻ 1: Văn bản (Home) - Clipboard, Font, Đoạn văn, Căn lề, Giãn dòng
- Thẻ 2: Chèn (Insert) - Bảng biểu, Hình ảnh, Ngắt trang, Ký hiệu đặc biệt
- Thẻ 3: Bố cục (Page Layout) - Khổ A4, Hướng giấy, Lề in NĐ 30, Bật/Tắt Thước đo
- Thẻ 4: Thể thức (Admin Tools) - Khóa khung, Đổi loại văn bản, Rà soát tuân thủ
"""

from PyQt6.QtWidgets import (
    QWidget, QVBoxLayout, QHBoxLayout, QTabWidget, QToolBar,
    QComboBox, QSpinBox, QToolButton, QLabel, QFrame, QButtonGroup,
    QColorDialog, QInputDialog, QMessageBox, QMenu
)
from PyQt6.QtGui import QIcon, QFont, QColor, QAction
from PyQt6.QtCore import Qt, pyqtSignal, QSize

from ui.icon_factory import get_ui_icon


class AdministrativeRibbonBar(QWidget):
    # Signals phát ra cho canvas nhận lệnh
    format_bold_triggered = pyqtSignal()
    format_italic_triggered = pyqtSignal()
    format_underline_triggered = pyqtSignal()
    font_size_changed = pyqtSignal(float)
    font_color_changed = pyqtSignal(QColor)
    highlight_color_changed = pyqtSignal(QColor)

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

    page_orientation_changed = pyqtSignal(str)  # 'portrait' hoặc 'landscape'
    margins_preset_triggered = pyqtSignal(str)   # 'nd30' hoặc 'custom'
    toggle_ruler_triggered = pyqtSignal(bool)

    switch_doc_type_requested = pyqtSignal(str)
    check_compliance_requested = pyqtSignal()
    lock_scaffolding_toggled = pyqtSignal(bool)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(115)
        self.setStyleSheet("""
            QWidget {
                background-color: #f8fafc;
                font-family: 'Segoe UI', Arial, sans-serif;
            }
            QTabWidget::pane {
                border-top: 1px solid #cbd5e1;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #f1f5f9;
                color: #475569;
                padding: 6px 16px;
                border: 1px solid #e2e8f0;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                font-size: 12px;
                font-weight: 500;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #1e3a8a;
                border-color: #cbd5e1;
                font-weight: 600;
            }
            QToolButton {
                border: 1px solid transparent;
                border-radius: 4px;
                padding: 4px;
                background: transparent;
                font-size: 11px;
                color: #334155;
            }
            QToolButton:hover {
                background-color: #f1f5f9;
                border-color: #cbd5e1;
            }
            QToolButton:checked {
                background-color: #e0e7ff;
                border-color: #a5b4fc;
                color: #3730a3;
            }
            QComboBox {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                padding: 2px 6px;
                font-size: 11px;
                background: white;
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

    # ==================== 1. THẺ VĂN BẢN (HOME) ====================
    def _setup_home_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(8, 4, 8, 4)
        layout.setSpacing(4)

        # Khối Phông chữ (Font)
        font_box = QVBoxLayout()
        font_box.setSpacing(3)

        row_font_pickers = QHBoxLayout()
        self.cb_font_family = QComboBox()
        self.cb_font_family.addItem("Times New Roman (Chuẩn NĐ 30)")
        self.cb_font_family.setEnabled(False)  # Khóa phông chuẩn theo quy định
        self.cb_font_family.setFixedWidth(160)
        row_font_pickers.addWidget(self.cb_font_family)

        self.cb_font_size = QComboBox()
        self.cb_font_size.setFixedWidth(55)
        for s in [10, 11, 12, 12.5, 13, 13.5, 14, 15, 16, 18, 20, 24]:
            self.cb_font_size.addItem(str(s))
        self.cb_font_size.setCurrentText("13")
        self.cb_font_size.currentTextChanged.connect(self._on_font_size_changed)
        row_font_pickers.addWidget(self.cb_font_size)
        font_box.addLayout(row_font_pickers)

        row_font_styles = QHBoxLayout()
        self.btn_bold = QToolButton()
        self.btn_bold.setText("B")
        self.btn_bold.setCheckable(True)
        self.btn_bold.setFont(QFont("Times New Roman", 11, QFont.Weight.Bold))
        self.btn_bold.clicked.connect(self.format_bold_triggered.emit)
        row_font_styles.addWidget(self.btn_bold)

        self.btn_italic = QToolButton()
        self.btn_italic.setText("I")
        self.btn_italic.setCheckable(True)
        font_i = QFont("Times New Roman", 11)
        font_i.setItalic(True)
        self.btn_italic.setFont(font_i)
        self.btn_italic.clicked.connect(self.format_italic_triggered.emit)
        row_font_styles.addWidget(self.btn_italic)

        self.btn_underline = QToolButton()
        self.btn_underline.setText("U")
        self.btn_underline.setCheckable(True)
        font_u = QFont("Times New Roman", 11)
        font_u.setUnderline(True)
        self.btn_underline.setFont(font_u)
        self.btn_underline.clicked.connect(self.format_underline_triggered.emit)
        row_font_styles.addWidget(self.btn_underline)

        self.btn_text_color = QToolButton()
        self.btn_text_color.setText("A")
        self.btn_text_color.setToolTip("Màu chữ")
        self.btn_text_color.setStyleSheet("color: #1e3a8a; font-weight: bold;")
        self.btn_text_color.clicked.connect(self._choose_text_color)
        row_font_styles.addWidget(self.btn_text_color)

        font_box.addLayout(row_font_styles)
        layout.addLayout(font_box)
        layout.addWidget(self._create_separator())

        # Khối Đoạn văn & Căn lề (Paragraph)
        para_box = QVBoxLayout()
        para_box.setSpacing(3)

        row_align = QHBoxLayout()
        self.btn_left = QToolButton()
        self.btn_left.setIcon(get_ui_icon("align_left"))
        self.btn_left.setToolTip("Căn lề trái")
        self.btn_left.clicked.connect(self.align_left_triggered.emit)
        row_align.addWidget(self.btn_left)

        self.btn_center = QToolButton()
        self.btn_center.setIcon(get_ui_icon("align_center"))
        self.btn_center.setToolTip("Căn giữa")
        self.btn_center.clicked.connect(self.align_center_triggered.emit)
        row_align.addWidget(self.btn_center)

        self.btn_right = QToolButton()
        self.btn_right.setIcon(get_ui_icon("align_right"))
        self.btn_right.setToolTip("Căn lề phải")
        self.btn_right.clicked.connect(self.align_right_triggered.emit)
        row_align.addWidget(self.btn_right)

        self.btn_justify = QToolButton()
        self.btn_justify.setIcon(get_ui_icon("align_justify"))
        self.btn_justify.setToolTip("Căn đều hai lề (Chuẩn NĐ 30)")
        self.btn_justify.setChecked(True)
        self.btn_justify.clicked.connect(self.align_justify_triggered.emit)
        row_align.addWidget(self.btn_justify)

        para_box.addLayout(row_align)

        row_spacing = QHBoxLayout()
        lbl_spacing = QLabel("Giãn dòng:")
        lbl_spacing.setStyleSheet("font-size: 11px; color: #64748b;")
        row_spacing.addWidget(lbl_spacing)

        self.cb_line_spacing = QComboBox()
        self.cb_line_spacing.setFixedWidth(65)
        self.cb_line_spacing.addItems(["1.0", "1.15", "1.2", "1.25", "1.5"])
        self.cb_line_spacing.setCurrentText("1.2")
        self.cb_line_spacing.currentTextChanged.connect(self._on_line_spacing_changed)
        row_spacing.addWidget(self.cb_line_spacing)

        btn_first_line = QToolButton()
        btn_first_line.setText("Thụt lề 1.27cm")
        btn_first_line.setToolTip("Thụt đầu dòng 1.27cm chuẩn thể thức")
        btn_first_line.clicked.connect(lambda: self.indent_first_line_triggered.emit(12.7))
        row_spacing.addWidget(btn_first_line)

        para_box.addLayout(row_spacing)
        layout.addLayout(para_box)
        layout.addWidget(self._create_separator())

        # Khối Danh sách (Lists)
        list_box = QVBoxLayout()
        list_box.setSpacing(4)
        btn_bullet = QToolButton()
        btn_bullet.setText("• Gạch đầu dòng (-)")
        btn_bullet.clicked.connect(self.bullet_list_triggered.emit)
        list_box.addWidget(btn_bullet)

        btn_num = QToolButton()
        btn_num.setText("1. Số thứ tự (1, 2, 3)")
        btn_num.clicked.connect(self.numbered_list_triggered.emit)
        list_box.addWidget(btn_num)
        layout.addLayout(list_box)

        layout.addStretch()
        self.tabs.addTab(tab, "Văn bản (Home)")

    # ==================== 2. THẺ CHÈN (INSERT) ====================
    def _setup_insert_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(10)

        # Chèn bảng
        btn_table = QToolButton()
        btn_table.setIcon(get_ui_icon("table"))
        btn_table.setIconSize(QSize(22, 22))
        btn_table.setText("Chèn Bảng biểu")
        btn_table.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_table.clicked.connect(self.insert_table_requested.emit)
        layout.addWidget(btn_table)

        # Chèn ảnh
        btn_image = QToolButton()
        btn_image.setIcon(get_ui_icon("image"))
        btn_image.setIconSize(QSize(22, 22))
        btn_image.setText("Chèn Hình ảnh")
        btn_image.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_image.clicked.connect(self.insert_image_requested.emit)
        layout.addWidget(btn_image)

        # Ngắt trang
        btn_break = QToolButton()
        btn_break.setIcon(get_ui_icon("page_break"))
        btn_break.setIconSize(QSize(22, 22))
        btn_break.setText("Ngắt trang mới")
        btn_break.setToolButtonStyle(Qt.ToolButtonStyle.ToolButtonTextUnderIcon)
        btn_break.clicked.connect(self.page_break_requested.emit)
        layout.addWidget(btn_break)

        layout.addWidget(self._create_separator())

        # Ký hiệu đặc biệt
        lbl_sym = QLabel("Ký tự văn thư:")
        lbl_sym.setStyleSheet("font-size: 11px; color: #64748b;")
        layout.addWidget(lbl_sym)

        sym_layout = QHBoxLayout()
        for sym in ["§", "°", "©", "®", "™", "→", "✔"]:
            b_sym = QToolButton()
            b_sym.setText(sym)
            b_sym.setFixedSize(26, 26)
            b_sym.clicked.connect(lambda checked, s=sym: self.insert_symbol_requested.emit(s))
            sym_layout.addWidget(b_sym)
        layout.addLayout(sym_layout)

        layout.addStretch()
        self.tabs.addTab(tab, "Chèn (Insert)")

    # ==================== 3. THẺ BỐ CỤC TRANG (PAGE LAYOUT) ====================
    def _setup_layout_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(10)

        # Hướng giấy
        btn_portrait = QToolButton()
        btn_portrait.setText("Trang Dọc (A4)")
        btn_portrait.setCheckable(True)
        btn_portrait.setChecked(True)
        btn_portrait.clicked.connect(lambda: self.page_orientation_changed.emit("portrait"))
        layout.addWidget(btn_portrait)

        btn_landscape = QToolButton()
        btn_landscape.setText("Trang Ngang (A4)")
        btn_landscape.setCheckable(True)
        btn_landscape.clicked.connect(lambda: self.page_orientation_changed.emit("landscape"))
        layout.addWidget(btn_landscape)

        grp_orient = QButtonGroup(self)
        grp_orient.addButton(btn_portrait)
        grp_orient.addButton(btn_landscape)

        layout.addWidget(self._create_separator())

        # Căn lề NĐ 30
        btn_margin_nd30 = QToolButton()
        btn_margin_nd30.setText("Lề chuẩn NĐ 30\n(20-20-30-15 mm)")
        btn_margin_nd30.clicked.connect(lambda: self.margins_preset_triggered.emit("nd30"))
        layout.addWidget(btn_margin_nd30)

        layout.addWidget(self._create_separator())

        # Bật/Tắt thước đo
        btn_ruler = QToolButton()
        btn_ruler.setText("Thước căn lề (Ruler)")
        btn_ruler.setCheckable(True)
        btn_ruler.setChecked(True)
        btn_ruler.toggled.connect(self.toggle_ruler_triggered.emit)
        layout.addWidget(btn_ruler)

        layout.addStretch()
        self.tabs.addTab(tab, "Bố cục trang (Page Layout)")

    # ==================== 4. THẺ THỂ THỨC (ADMIN TOOLS) ====================
    def _setup_admin_tab(self):
        tab = QWidget()
        layout = QHBoxLayout(tab)
        layout.setContentsMargins(12, 6, 12, 6)
        layout.setSpacing(10)

        # Đổi loại văn bản
        lbl_type = QLabel("Chuyển mẫu:")
        lbl_type.setStyleSheet("font-size: 11px; color: #64748b;")
        layout.addWidget(lbl_type)

        self.cb_switch_type = QComboBox()
        self.cb_switch_type.setFixedWidth(150)
        self.cb_switch_type.addItems([
            "QUYẾT ĐỊNH",
            "CÔNG VĂN",
            "TỜ TRÌNH",
            "THÔNG BÁO",
            "NGHỊ QUYẾT ĐẢNG",
            "CHỈ THỊ ĐẢNG"
        ])
        self.cb_switch_type.currentTextChanged.connect(self.switch_doc_type_requested.emit)
        layout.addWidget(self.cb_switch_type)

        layout.addWidget(self._create_separator())

        # Khóa / Mở vùng thể thức
        self.btn_lock = QToolButton()
        self.btn_lock.setText("Khóa Khung Thể thức")
        self.btn_lock.setCheckable(True)
        self.btn_lock.setChecked(True)
        self.btn_lock.setStyleSheet("color: #15803d; font-weight: bold;")
        self.btn_lock.toggled.connect(self._on_lock_toggled)
        layout.addWidget(self.btn_lock)

        layout.addWidget(self._create_separator())

        # Rà soát tuân thủ
        btn_check = QToolButton()
        btn_check.setText("Rà soát Chuẩn thể thức\n(NĐ 30 / HD 36)")
        btn_check.setStyleSheet("background-color: #f0fdf4; border: 1px solid #86efac; color: #166534; font-weight: 600;")
        btn_check.clicked.connect(self.check_compliance_requested.emit)
        layout.addWidget(btn_check)

        layout.addStretch()
        self.tabs.addTab(tab, "Công cụ Thể thức (Admin Tools)")

    # ==================== TRÌNH ĐIỀU KHIỂN SỰ KIỆN ====================
    def _on_font_size_changed(self, text: str):
        try:
            sz = float(text)
            self.font_size_changed.emit(sz)
        except ValueError:
            pass

    def _choose_text_color(self):
        color = QColorDialog.getColor(QColor(0, 0, 0), self, "Chọn màu chữ")
        if color.isValid():
            self.font_color_changed.emit(color)

    def _on_line_spacing_changed(self, text: str):
        try:
            sp = float(text)
            self.line_spacing_changed.emit(sp)
        except ValueError:
            pass

    def _on_lock_toggled(self, checked: bool):
        if checked:
            self.btn_lock.setText("Khóa Khung Thể thức")
            self.btn_lock.setStyleSheet("color: #15803d; font-weight: bold;")
        else:
            self.btn_lock.setText("Mở Khóa Khung")
            self.btn_lock.setStyleSheet("color: #b45309; font-weight: bold;")
        self.lock_scaffolding_toggled.emit(checked)
