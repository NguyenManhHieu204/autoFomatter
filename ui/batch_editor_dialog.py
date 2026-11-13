"""
ui/batch_editor_dialog.py: Giao diện Xử lý Hàng loạt Văn bản (Batch Processor: Content & Rename)
- Hỗ trợ đổi địa danh hàng loạt (Thuận Hạnh, Đức An, Nam Bình, Kiến Đức...)
- Hỗ trợ chỉnh sửa số hiệu văn bản hàng loạt:
  Chuyển dấu chấm thành khoảng trống dự thảo chuẩn (Số: ............../abc -> Số:               /abc)
  Chuyển khoảng trống thành dấu chấm, đổi ký hiệu viết tắt, gán số hiệu cụ thể
- Hỗ trợ xóa ngày/đổi ngày hàng loạt (chuyển sang ngày dự thảo hoặc đặt ngày cố định)
- Tự động chuẩn hóa ngắt dòng cơ quan ban hành theo chuẩn NĐ 30:
  Dòng 1: ỦY BAN NHÂN DÂN
  Dòng 2: XÃ THUẬN HẠNH / XÃ ĐỨC AN...
- Tìm kiếm & Thay thế văn bản hàng loạt
- Hỗ trợ xuất sang thư mục tùy chọn hoặc ghi đè an toàn
- ĐỔI TÊN TỆP HÀNG LOẠT:
  Xóa nhanh hậu tố '_chuanhoa', '_ChuanHoa', '_chuẩn hóa' sau khi định dạng
  Tìm & thay thế trong tên tệp, thêm tiền tố/hậu tố, đổi kiểu chữ, tự động xử lý trùng tên
  Bảng xem trước trực quan tức thì
- XỬ LÝ ĐỒNG THỜI: Vừa chỉnh sửa nội dung vừa đổi tên tệp trong 1 lần chạy duy nhất.
"""

import os
import datetime
from typing import List, Optional, Dict, Any, Set
from PyQt6.QtWidgets import (
    QDialog, QVBoxLayout, QHBoxLayout, QLabel, QPushButton, QLineEdit,
    QCheckBox, QComboBox, QRadioButton, QButtonGroup, QGroupBox, QListWidget,
    QProgressBar, QMessageBox, QFileDialog, QFrame, QSplitter, QScrollArea,
    QWidget, QSizePolicy, QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt6.QtGui import QFont, QColor

from core.batch_editor import (
    BatchEditor, BatchEditOptions, BatchEditItemResult,
    BatchRenameOptions, BatchRenameItemResult
)
from core.path_utils import get_resource_path


class BatchCombinedWorkerThread(QThread):
    progress_updated = pyqtSignal(int, int, str)
    finished_processing = pyqtSignal(list)

    def __init__(
        self,
        file_paths: List[str],
        edit_options: Optional[BatchEditOptions],
        rename_options: Optional[BatchRenameOptions]
    ):
        super().__init__()
        self.file_paths = file_paths
        self.edit_options = edit_options
        self.rename_options = rename_options

    def run(self):
        def on_prog(cur, total, msg):
            self.progress_updated.emit(cur, total, msg)

        results = BatchEditor.process_combined_batch(
            self.file_paths,
            self.edit_options,
            self.rename_options,
            progress_callback=on_prog
        )
        self.finished_processing.emit(results)


class BatchEditorDialog(QDialog):
    """Hộp thoại Xử lý Hàng loạt Văn bản: Chỉnh sửa nội dung & Đổi tên tệp đồng thời"""

    def __init__(self, parent=None, initial_files: Optional[List[str]] = None, initial_tab: int = 0):
        super().__init__(parent)
        self.setWindowTitle("Xử lý Hàng loạt Văn bản (Nội dung & Đổi tên tệp)")
        self.resize(1120, 720)
        self.setMinimumSize(960, 600)
        self.file_paths: List[str] = list(initial_files) if initial_files else []
        self._file_paths_set: Set[str] = set(self.file_paths)
        self._initial_tab = initial_tab
        self.worker_thread: Optional[BatchCombinedWorkerThread] = None

        # Timer debounce cho Live Preview để chống giật/lag khi có nhiều tệp
        self._preview_timer = QTimer(self)
        self._preview_timer.setSingleShot(True)
        self._preview_timer.setInterval(120)
        self._preview_timer.timeout.connect(self._run_rename_preview_calculation)

        self._init_ui()

    def _init_ui(self):
        chk_icon_path = get_resource_path("assets/icons/check_white.svg").replace("\\", "/")

        self.setStyleSheet(f"""
            QDialog {{
                background-color: #f8fafc;
            }}
            QGroupBox {{
                font-weight: 700;
                font-size: 12px;
                color: #1e3a8a;
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                margin-top: 8px;
                padding: 8px 8px 6px 8px;
                background-color: #ffffff;
            }}
            QGroupBox::title {{
                subcontrol-origin: margin;
                subcontrol-position: top left;
                left: 8px;
                padding: 0 4px;
                background-color: #ffffff;
                color: #1e3a8a;
                font-weight: bold;
            }}
            QLineEdit {{
                border: 1px solid #cbd5e1;
                border-radius: 5px;
                padding: 4px 8px;
                min-height: 24px;
                font-size: 12px;
                background-color: #ffffff;
                color: #0f172a;
            }}
            QLineEdit:focus {{
                border: 1.5px solid #2563eb;
                background-color: #f8faff;
            }}
            QComboBox {{
                border: 1px solid #cbd5e1;
                border-radius: 5px;
                padding: 4px 24px 4px 8px;
                min-height: 24px;
                font-size: 12px;
                background-color: #ffffff;
                color: #0f172a;
            }}
            QComboBox:focus {{
                border: 1.5px solid #2563eb;
            }}
            QComboBox::drop-down {{
                subcontrol-origin: padding;
                subcontrol-position: top right;
                width: 20px;
                border-left: 1px solid #cbd5e1;
                border-top-right-radius: 5px;
                border-bottom-right-radius: 5px;
                background-color: #f8fafc;
            }}
            QComboBox QAbstractItemView {{
                background-color: #ffffff;
                color: #0f172a;
                border: 1.5px solid #94a3b8;
                border-radius: 5px;
                selection-background-color: #eff6ff;
                selection-color: #1d4ed8;
                padding: 4px;
                outline: none;
            }}
            QComboBox QAbstractItemView::item {{
                min-height: 26px;
                padding: 4px 8px;
                color: #0f172a;
                background-color: #ffffff;
            }}
            QComboBox QAbstractItemView::item:hover,
            QComboBox QAbstractItemView::item:selected {{
                background-color: #eff6ff;
                color: #1d4ed8;
            }}
            QCheckBox, QRadioButton {{
                font-size: 12px;
                color: #0f172a;
                spacing: 6px;
                background-color: transparent;
            }}
            QCheckBox::indicator {{
                width: 16px;
                height: 16px;
                border: 1.5px solid #64748b;
                border-radius: 3px;
                background-color: #ffffff;
            }}
            QCheckBox::indicator:hover {{
                border-color: #2563eb;
                background-color: #eff6ff;
            }}
            QCheckBox::indicator:checked {{
                border-color: #2563eb;
                background-color: #2563eb;
                image: url({chk_icon_path});
            }}
            QRadioButton::indicator {{
                width: 16px;
                height: 16px;
                border: 1.5px solid #64748b;
                border-radius: 8px;
                background-color: #ffffff;
            }}
            QRadioButton::indicator:hover {{
                border-color: #2563eb;
                background-color: #eff6ff;
            }}
            QRadioButton::indicator:checked {{
                border: 4.5px solid #2563eb;
                background-color: #ffffff;
            }}
            QTabWidget::pane {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                background-color: #ffffff;
                top: -1px;
            }}
            QTabBar::tab {{
                background-color: #e2e8f0;
                color: #475569;
                font-weight: 600;
                font-size: 12px;
                padding: 7px 18px;
                border: 1px solid #cbd5e1;
                border-bottom: none;
                border-top-left-radius: 6px;
                border-top-right-radius: 6px;
                margin-right: 4px;
            }}
            QTabBar::tab:selected {{
                background-color: #ffffff;
                color: #1e3a8a;
                font-weight: 700;
                border-bottom: 2px solid #2563eb;
            }}
            QTabBar::tab:hover:!selected {{
                background-color: #f1f5f9;
                color: #1e293b;
            }}
            QTableWidget {{
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                background-color: #ffffff;
                color: #0f172a;
                font-size: 12px;
                gridline-color: #f1f5f9;
            }}
            QHeaderView::section {{
                background-color: #f8fafc;
                color: #1e293b;
                font-weight: 700;
                font-size: 11.5px;
                border: none;
                border-bottom: 1.5px solid #cbd5e1;
                border-right: 1px solid #e2e8f0;
                padding: 5px 8px;
            }}
            QTableWidget::item {{
                padding: 4px 6px;
                color: #0f172a;
                background-color: #ffffff;
            }}
            QTableWidget::item:selected {{
                background-color: #eff6ff;
                color: #1d4ed8;
            }}
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setContentsMargins(12, 12, 12, 12)
        main_layout.setSpacing(8)

        # 1. Splitter: Cột Trái (Danh sách tệp) & Cột Phải (Tabs cấu hình)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # --- Cột Trái: Danh sách tệp ---
        left_widget = QFrame()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 0, 0)
        left_layout.setSpacing(6)

        self.lbl_files_hdr = QLabel(f"Danh sách tệp áp dụng ({len(self.file_paths)} tệp):")
        self.lbl_files_hdr.setStyleSheet("font-weight: 700; color: #1e293b; font-size: 12px;")
        left_layout.addWidget(self.lbl_files_hdr)

        self.list_files = QListWidget()
        self.list_files.setStyleSheet("""
            QListWidget {
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                background-color: #ffffff;
                font-size: 12px;
                padding: 4px;
            }
            QListWidget::item {
                padding: 4px 6px;
                border-bottom: 1px solid #f1f5f9;
            }
            QListWidget::item:selected {
                background-color: #eff6ff;
                color: #1d4ed8;
                font-weight: 600;
            }
        """)
        # Nạp tệp ban đầu siêu tốc
        self.list_files.setUpdatesEnabled(False)
        self.list_files.addItems([os.path.basename(p) for p in self.file_paths])
        self.list_files.setUpdatesEnabled(True)
        left_layout.addWidget(self.list_files, 1)

        btns_file_row = QHBoxLayout()
        btns_file_row.setSpacing(6)
        btn_add_f = QPushButton("Thêm tệp...")
        btn_add_f.clicked.connect(self._add_files)
        btn_add_dir = QPushButton("Thêm thư mục...")
        btn_add_dir.clicked.connect(self._add_directory)
        btn_rem = QPushButton("Xóa")
        btn_rem.clicked.connect(self._remove_selected_file)
        btn_clr = QPushButton("Xóa hết")
        btn_clr.clicked.connect(self._clear_files)

        for b in [btn_add_f, btn_add_dir, btn_rem, btn_clr]:
            b.setStyleSheet("""
                QPushButton {
                    border: 1px solid #cbd5e1;
                    border-radius: 4px;
                    background-color: #ffffff;
                    padding: 5px 8px;
                    font-size: 11.5px;
                    color: #334155;
                }
                QPushButton:hover {
                    background-color: #f1f5f9;
                    border-color: #94a3b8;
                }
            """)
            btns_file_row.addWidget(b)
        left_layout.addLayout(btns_file_row)
        splitter.addWidget(left_widget)

        # --- Cột Phải: QTabWidget chứa Tab 1 (Chỉnh sửa nội dung) và Tab 2 (Đổi tên tệp) ---
        self.tab_widget = QTabWidget()

        # Tab 1: Chỉnh sửa Nội dung
        tab_content_widget = QWidget()
        l_tab_c = QVBoxLayout(tab_content_widget)
        l_tab_c.setContentsMargins(6, 6, 6, 6)
        l_tab_c.setSpacing(6)

        # Checkbox công tắc chính cho Tab 1
        self.chk_enable_content_tab = QCheckBox("Bật chỉnh sửa nội dung văn bản (chuẩn Nghị định 30)")
        self.chk_enable_content_tab.setChecked(True)
        self.chk_enable_content_tab.setStyleSheet("font-weight: 700; color: #1e3a8a; font-size: 12.5px;")
        l_tab_c.addWidget(self.chk_enable_content_tab)

        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.Shape.NoFrame)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.scroll_area.setStyleSheet("QScrollArea { background-color: transparent; border: none; }")

        scroll_content = QWidget()
        scroll_content.setStyleSheet("background-color: transparent;")
        right_layout = QVBoxLayout(scroll_content)
        right_layout.setContentsMargins(0, 0, 8, 4)
        right_layout.setSpacing(6)

        # Nhóm 1: Địa danh
        grp_loc = QGroupBox("1. Thay đổi Địa danh hàng loạt")
        l_loc = QVBoxLayout(grp_loc)
        l_loc.setSpacing(5)

        row_loc = QHBoxLayout()
        self.chk_enable_loc = QCheckBox("Đổi sang địa danh mới:")
        self.chk_enable_loc.setChecked(False)
        self.txt_new_loc = QLineEdit()
        self.txt_new_loc.setPlaceholderText("Ví dụ: Thuận Hạnh, Đức An, Nam Bình, Kiến Đức...")
        row_loc.addWidget(self.chk_enable_loc)
        row_loc.addWidget(self.txt_new_loc, 1)
        l_loc.addLayout(row_loc)

        self.chk_sync_agency = QCheckBox("Tự động đồng bộ tên cơ quan ban hành (vd: XÃ THUẬN HẠNH / XÃ ĐỨC AN)")
        self.chk_sync_agency.setChecked(True)
        self.chk_sync_agency.setStyleSheet("color: #475569; font-size: 11.5px; margin-left: 20px;")
        l_loc.addWidget(self.chk_sync_agency)
        right_layout.addWidget(grp_loc)

        # Nhóm 2: Số, ký hiệu văn bản hàng loạt
        grp_num = QGroupBox("2. Chỉnh sửa Số, ký hiệu văn bản hàng loạt")
        l_num = QVBoxLayout(grp_num)
        l_num.setSpacing(5)

        self.chk_enable_doc_num = QCheckBox("Bật chỉnh sửa Số, ký hiệu văn bản hàng loạt:")
        self.chk_enable_doc_num.setChecked(False)
        l_num.addWidget(self.chk_enable_doc_num)

        row_mode_num = QHBoxLayout()
        lbl_mode_num = QLabel("Chế độ định dạng số:")
        lbl_mode_num.setStyleSheet("font-size: 11.5px; color: #475569; margin-left: 20px;")
        self.cb_doc_num_mode = QComboBox()
        self.cb_doc_num_mode.addItem("Chuyển dấu chấm thành khoảng trống dự thảo: 'Số:               /abc' [NĐ 30]", "dots_to_spaces")
        self.cb_doc_num_mode.addItem("Chuyển khoảng trống thành dấu chấm: 'Số: ............../abc'", "spaces_to_dots")
        self.cb_doc_num_mode.addItem("Nhập số hiệu cụ thể tùy chỉnh...", "custom")
        self.cb_doc_num_mode.currentIndexChanged.connect(self._on_doc_num_mode_changed)
        row_mode_num.addWidget(lbl_mode_num)
        row_mode_num.addWidget(self.cb_doc_num_mode, 1)
        l_num.addLayout(row_mode_num)

        row_sym = QHBoxLayout()
        lbl_sym = QLabel("Đổi ký hiệu viết tắt (sau dấu /):")
        lbl_sym.setStyleSheet("font-size: 11.5px; color: #475569; margin-left: 20px;")
        self.txt_new_symbol = QLineEdit()
        self.txt_new_symbol.setPlaceholderText("Để trống sẽ giữ nguyên ký hiệu gốc của từng tệp (vd: TB-ĐKĐĐ, QĐ-UBND)")
        row_sym.addWidget(lbl_sym)
        row_sym.addWidget(self.txt_new_symbol, 1)
        l_num.addLayout(row_sym)

        self.txt_custom_doc_num = QLineEdit()
        self.txt_custom_doc_num.setPlaceholderText("Ví dụ: Số:    /TB-ĐKĐĐ hoặc Số: 125/QĐ-UBND")
        self.txt_custom_doc_num.setStyleSheet("margin-left: 20px;")
        self.txt_custom_doc_num.setVisible(False)
        l_num.addWidget(self.txt_custom_doc_num)
        right_layout.addWidget(grp_num)

        # Nhóm 3: Ngày tháng
        grp_date = QGroupBox("3. Xóa ngày hoặc Thay đổi Ngày tháng hàng loạt")
        l_date = QVBoxLayout(grp_date)
        l_date.setSpacing(5)

        self.chk_enable_date = QCheckBox("Áp dụng thay đổi ngày tháng:")
        self.chk_enable_date.setChecked(True)
        l_date.addWidget(self.chk_enable_date)

        self.cb_date_mode = QComboBox()
        self.cb_date_mode.addItem("Dự thảo có khoảng trống (ngày    tháng    năm 2026) [Khuyên dùng]", "draft_spaces")
        self.cb_date_mode.addItem("Dự thảo dấu chấm (ngày ... tháng ... năm 2026)", "draft_dots")
        self.cb_date_mode.addItem("Xóa toàn bộ (ngày ... tháng ... năm ...)", "clear")
        self.cb_date_mode.addItem("Đặt ngày hôm nay (ngày hiện tại)", "today")
        self.cb_date_mode.addItem("Nhập ngày cụ thể tùy chỉnh...", "custom")
        self.cb_date_mode.currentIndexChanged.connect(self._on_date_mode_changed)
        l_date.addWidget(self.cb_date_mode)

        self.txt_custom_date = QLineEdit()
        self.txt_custom_date.setPlaceholderText("Ví dụ: ngày 15 tháng 11 năm 2026")
        self.txt_custom_date.setVisible(False)
        l_date.addWidget(self.txt_custom_date)
        right_layout.addWidget(grp_date)

        # Nhóm 4: Cơ quan ban hành (Nghị định 30)
        grp_agency = QGroupBox("4. Chuẩn hóa Cơ quan ban hành (Nghị định 30)")
        l_ag = QVBoxLayout(grp_agency)
        l_ag.setSpacing(5)

        self.chk_auto_agency = QCheckBox("Tự động sửa ngắt dòng cơ quan theo chuẩn NĐ 30:")
        self.chk_auto_agency.setChecked(True)
        l_ag.addWidget(self.chk_auto_agency)

        lbl_ag_hint = QLabel("Dòng 1: ỦY BAN NHÂN DÂN (in hoa, đứng) | Dòng 2: XÃ THUẬN HẠNH / XÃ ĐỨC AN... (in hoa, đậm)\nKhắc phục triệt để lỗi 'ỦY BAN NHÂN DÂN XÃ' bị dính trên dòng 1.")
        lbl_ag_hint.setStyleSheet("font-size: 11px; color: #15803d; background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 4px; padding: 4px 8px; margin-left: 20px;")
        lbl_ag_hint.setWordWrap(True)
        l_ag.addWidget(lbl_ag_hint)

        self.chk_custom_agency = QCheckBox("Gán cơ quan cố định tùy chỉnh:")
        self.chk_custom_agency.setChecked(False)
        l_ag.addWidget(self.chk_custom_agency)

        row_c_ag = QHBoxLayout()
        self.txt_c_sup = QLineEdit()
        self.txt_c_sup.setPlaceholderText("Dòng 1: ỦY BAN NHÂN DÂN")
        self.txt_c_iss = QLineEdit()
        self.txt_c_iss.setPlaceholderText("Dòng 2: XÃ THUẬN HẠNH")
        row_c_ag.addWidget(self.txt_c_sup)
        row_c_ag.addWidget(self.txt_c_iss)
        l_ag.addLayout(row_c_ag)
        right_layout.addWidget(grp_agency)

        # Nhóm 5: Tìm & Thay thế văn bản
        grp_fr = QGroupBox("5. Tìm kiếm & Thay thế văn bản (Toàn bộ tài liệu)")
        l_fr = QVBoxLayout(grp_fr)
        l_fr.setSpacing(5)

        self.chk_enable_fr = QCheckBox("Bật tìm kiếm & thay thế:")
        l_fr.addWidget(self.chk_enable_fr)

        row_fr = QHBoxLayout()
        self.txt_find = QLineEdit()
        self.txt_find.setPlaceholderText("Từ khóa cần tìm...")
        self.txt_replace = QLineEdit()
        self.txt_replace.setPlaceholderText("Thay thế bằng...")
        row_fr.addWidget(self.txt_find)
        row_fr.addWidget(self.txt_replace)
        l_fr.addLayout(row_fr)

        self.chk_match_case = QCheckBox("Phân biệt chữ hoa / chữ thường")
        l_fr.addWidget(self.chk_match_case)
        right_layout.addWidget(grp_fr)

        # Nhóm 6: Nơi lưu tệp
        grp_save = QGroupBox("6. Nơi lưu tệp kết quả")
        l_save = QVBoxLayout(grp_save)
        l_save.setSpacing(5)

        self.rb_overwrite = QRadioButton("Cập nhật đè trực tiếp lên tệp đã chọn")
        self.rb_overwrite.setChecked(True)
        self.rb_custom_dir = QRadioButton("Xuất sang thư mục riêng:")
        l_save.addWidget(self.rb_overwrite)
        l_save.addWidget(self.rb_custom_dir)

        row_out = QHBoxLayout()
        self.txt_out_folder = QLineEdit()
        self.txt_out_folder.setPlaceholderText("Chọn thư mục xuất tệp...")
        self.txt_out_folder.setReadOnly(True)
        self.btn_browse_out = QPushButton("Chọn thư mục...")
        self.btn_browse_out.setStyleSheet("""
            QPushButton {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                background-color: #ffffff;
                padding: 4px 10px;
                font-size: 11.5px;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
            }
        """)
        self.btn_browse_out.clicked.connect(self._browse_output_dir)
        row_out.addWidget(self.txt_out_folder, 1)
        row_out.addWidget(self.btn_browse_out)
        l_save.addLayout(row_out)
        right_layout.addWidget(grp_save)

        right_layout.addStretch()
        self.scroll_area.setWidget(scroll_content)
        l_tab_c.addWidget(self.scroll_area, 1)

        # Tab 2: Đổi tên tệp (giao diện tối ưu, bảng xem trước chiếm trọn không gian)
        tab_rename_widget = self._create_optimized_rename_tab()

        self.tab_widget.addTab(tab_content_widget, "Chỉnh sửa nội dung")
        self.tab_widget.addTab(tab_rename_widget, "Đổi tên tệp")
        if self._initial_tab in (0, 1):
            self.tab_widget.setCurrentIndex(self._initial_tab)

        splitter.addWidget(self.tab_widget)
        splitter.setSizes([380, 720])
        main_layout.addWidget(splitter, 1)

        # 3. Thanh tiến trình
        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        self.progress_bar.setFixedHeight(18)
        self.progress_bar.setStyleSheet("""
            QProgressBar {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                text-align: center;
                height: 18px;
                font-size: 11px;
                background-color: #f1f5f9;
            }
            QProgressBar::chunk {
                background-color: #2563eb;
                border-radius: 3px;
            }
        """)
        main_layout.addWidget(self.progress_bar)

        # 4. Hàng Nút Hành động & Trạng thái ở đáy hộp thoại
        act_layout = QHBoxLayout()
        act_layout.setContentsMargins(0, 4, 0, 0)
        act_layout.setSpacing(10)

        self.lbl_status = QLabel("")
        self.lbl_status.setStyleSheet("color: #475569; font-size: 12px; font-weight: 500;")
        act_layout.addWidget(self.lbl_status, 1)

        self.btn_execute = QPushButton("BẮT ĐẦU XỬ LÝ HÀNG LOẠT")
        self.btn_execute.setStyleSheet("""
            QPushButton {
                background-color: #2563eb;
                color: #ffffff;
                font-weight: bold;
                font-size: 13px;
                padding: 9px 26px;
                border-radius: 6px;
                border: 1px solid #1d4ed8;
            }
            QPushButton:hover {
                background-color: #1d4ed8;
            }
            QPushButton:disabled {
                background-color: #94a3b8;
                border-color: #cbd5e1;
            }
        """)
        self.btn_execute.setToolTip("Thực hiện đồng thời chỉnh sửa nội dung văn bản và đổi tên tệp đã chọn")
        self.btn_execute.clicked.connect(self._start_combined_batch)

        self.btn_close = QPushButton("Đóng")
        self.btn_close.setStyleSheet("""
            QPushButton {
                border: 1px solid #cbd5e1;
                border-radius: 6px;
                background-color: #ffffff;
                padding: 8px 18px;
                font-size: 12px;
                color: #334155;
            }
            QPushButton:hover {
                background-color: #f1f5f9;
            }
        """)
        self.btn_close.clicked.connect(self.accept)

        act_layout.addWidget(self.btn_execute)
        act_layout.addWidget(self.btn_close)
        main_layout.addLayout(act_layout)

        # Popup styling an toàn chống đen nền dark mode
        def _style_combo_popup(cb: QComboBox):
            v = cb.view()
            pal = v.palette()
            pal.setColor(pal.ColorRole.Base, QColor('#ffffff'))
            pal.setColor(pal.ColorRole.Window, QColor('#ffffff'))
            pal.setColor(pal.ColorRole.Text, QColor('#0f172a'))
            pal.setColor(pal.ColorRole.Highlight, QColor('#2563eb'))
            pal.setColor(pal.ColorRole.HighlightedText, QColor('#ffffff'))
            v.setPalette(pal)
            pw = v.parentWidget()
            if pw:
                pw.setStyleSheet("background-color: #ffffff; border: 1.5px solid #94a3b8; border-radius: 6px;")
            v.setStyleSheet("""
                QListView {
                    background-color: #ffffff;
                    color: #0f172a;
                    border: 1.5px solid #94a3b8;
                    border-radius: 6px;
                    padding: 4px;
                    outline: none;
                }
                QListView::item {
                    min-height: 28px;
                    padding: 4px 8px;
                    color: #0f172a;
                    background-color: #ffffff;
                }
                QListView::item:hover {
                    background-color: #eff6ff;
                    color: #1d4ed8;
                }
                QListView::item:selected {
                    background-color: #2563eb;
                    color: #ffffff;
                }
            """)

        _style_combo_popup(self.cb_doc_num_mode)
        _style_combo_popup(self.cb_date_mode)
        _style_combo_popup(self.cb_name_case)
        _style_combo_popup(self.cb_collision_policy)

    def _create_optimized_rename_tab(self) -> QWidget:
        """Tạo giao diện Tab 2: Tối ưu hóa không gian, bảng xem trước chiếm trọn diện tích"""
        widget = QWidget()
        layout = QVBoxLayout(widget)
        layout.setContentsMargins(6, 6, 6, 6)
        layout.setSpacing(6)

        # Checkbox công tắc chính cho Tab 2
        self.chk_enable_rename_tab = QCheckBox("Bật đổi tên tệp hàng loạt")
        self.chk_enable_rename_tab.setChecked(True)
        self.chk_enable_rename_tab.setStyleSheet("font-weight: 700; color: #1e3a8a; font-size: 12.5px;")
        layout.addWidget(self.chk_enable_rename_tab)

        # Nhóm tùy chọn đổi tên gọn gàng (chỉ chiếm ~110px)
        grp_opts = QGroupBox("Tùy chọn quy tắc đổi tên tệp")
        l_opts = QVBoxLayout(grp_opts)
        l_opts.setContentsMargins(8, 8, 8, 8)
        l_opts.setSpacing(5)

        # Hàng 1: Xóa hậu tố chuẩn hóa & xóa chuỗi tùy chỉnh
        row1 = QHBoxLayout()
        row1.setSpacing(10)
        self.chk_remove_chuanhoa = QCheckBox("Xóa hậu tố '_chuanhoa', '_ChuanHoa', '_chuẩn hóa'")
        self.chk_remove_chuanhoa.setChecked(True)
        self.chk_remove_chuanhoa.setStyleSheet("font-weight: 600; color: #1e293b;")
        
        lbl_rem_c = QLabel("Xóa thêm cụm từ:")
        lbl_rem_c.setStyleSheet("color: #475569; font-size: 11.5px;")
        self.txt_remove_custom = QLineEdit()
        self.txt_remove_custom.setPlaceholderText("Ví dụ: - copy, _fixed, 2026...")
        row1.addWidget(self.chk_remove_chuanhoa)
        row1.addWidget(lbl_rem_c)
        row1.addWidget(self.txt_remove_custom, 1)
        l_opts.addLayout(row1)

        # Hàng 2: Tìm & Thay thế trong tên tệp
        row2 = QHBoxLayout()
        row2.setSpacing(8)
        self.chk_enable_name_fr = QCheckBox("Tìm & Thay thế:")
        self.chk_enable_name_fr.setChecked(False)
        self.txt_name_find = QLineEdit()
        self.txt_name_find.setPlaceholderText("Cụm từ cần tìm trong tên tệp (vd: Thuận Hạnh)...")
        self.txt_name_replace = QLineEdit()
        self.txt_name_replace.setPlaceholderText("Thay thế bằng (vd: Đức An)...")
        self.chk_name_match_case = QCheckBox("Phân biệt hoa/thường")
        row2.addWidget(self.chk_enable_name_fr)
        row2.addWidget(self.txt_name_find, 1)
        row2.addWidget(self.txt_name_replace, 1)
        row2.addWidget(self.chk_name_match_case)
        l_opts.addLayout(row2)

        # Hàng 3: Tiền tố, Hậu tố & Định dạng
        row3 = QHBoxLayout()
        row3.setSpacing(8)
        self.txt_name_prefix = QLineEdit()
        self.txt_name_prefix.setPlaceholderText("Tiền tố (vd: VB_)")
        self.txt_name_suffix = QLineEdit()
        self.txt_name_suffix.setPlaceholderText("Hậu tố (vd: _final)")

        self.cb_name_case = QComboBox()
        self.cb_name_case.addItem("Giữ nguyên kiểu chữ", "keep")
        self.cb_name_case.addItem("Chữ thường (abc)", "lower")
        self.cb_name_case.addItem("CHỮ HOA (ABC)", "upper")
        self.cb_name_case.addItem("Viết Hoa Đầu Từ (Abc)", "title")

        self.cb_collision_policy = QComboBox()
        self.cb_collision_policy.addItem("Tự thêm số (1), (2)... nếu trùng", "auto_index")
        self.cb_collision_policy.addItem("Bỏ qua nếu trùng tên", "skip")
        self.cb_collision_policy.addItem("Ghi đè", "overwrite")

        self.chk_name_numbering = QCheckBox("Đánh số TT")

        row3.addWidget(self.txt_name_prefix, 1)
        row3.addWidget(self.txt_name_suffix, 1)
        row3.addWidget(self.cb_name_case)
        row3.addWidget(self.cb_collision_policy)
        row3.addWidget(self.chk_name_numbering)
        l_opts.addLayout(row3)
        layout.addWidget(grp_opts)

        # Header Bảng xem trước
        row_prev_hdr = QHBoxLayout()
        self.lbl_preview_summary = QLabel("Xem trước danh sách đổi tên:")
        self.lbl_preview_summary.setStyleSheet("font-weight: 700; color: #1e293b; font-size: 12px;")
        btn_refresh = QPushButton("Làm mới bảng")
        btn_refresh.setStyleSheet("""
            QPushButton {
                border: 1px solid #cbd5e1;
                border-radius: 4px;
                background-color: #ffffff;
                padding: 3px 8px;
                font-size: 11px;
                color: #475569;
            }
            QPushButton:hover { background-color: #f1f5f9; }
        """)
        btn_refresh.clicked.connect(self._run_rename_preview_calculation)
        row_prev_hdr.addWidget(self.lbl_preview_summary, 1)
        row_prev_hdr.addWidget(btn_refresh)
        layout.addLayout(row_prev_hdr)

        # Bảng xem trước rộng rãi chiếm toàn bộ chiều cao còn lại
        self.table_preview = QTableWidget()
        self.table_preview.setColumnCount(4)
        self.table_preview.setHorizontalHeaderLabels(["STT", "Tên tệp gốc (Hiện tại)", "Tên tệp mới (Dự kiến)", "Trạng thái"])
        self.table_preview.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        self.table_preview.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        self.table_preview.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        self.table_preview.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeMode.ResizeToContents)
        self.table_preview.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_preview.setEditTriggers(QTableWidget.EditTrigger.NoEditTriggers)
        self.table_preview.setAlternatingRowColors(True)
        layout.addWidget(self.table_preview, 1)

        # Lắng nghe sự kiện để kích hoạt debounce preview
        self.chk_enable_rename_tab.stateChanged.connect(self._trigger_preview_update)
        self.chk_remove_chuanhoa.stateChanged.connect(self._trigger_preview_update)
        self.txt_remove_custom.textChanged.connect(self._trigger_preview_update)
        self.chk_enable_name_fr.stateChanged.connect(self._trigger_preview_update)
        self.txt_name_find.textChanged.connect(self._trigger_preview_update)
        self.txt_name_replace.textChanged.connect(self._trigger_preview_update)
        self.chk_name_match_case.stateChanged.connect(self._trigger_preview_update)
        self.txt_name_prefix.textChanged.connect(self._trigger_preview_update)
        self.txt_name_suffix.textChanged.connect(self._trigger_preview_update)
        self.cb_name_case.currentIndexChanged.connect(self._trigger_preview_update)
        self.cb_collision_policy.currentIndexChanged.connect(self._trigger_preview_update)
        self.chk_name_numbering.stateChanged.connect(self._trigger_preview_update)

        return widget

    def showEvent(self, event):
        super().showEvent(event)
        QTimer.singleShot(0, lambda: self.scroll_area.verticalScrollBar().setValue(0))
        self._trigger_preview_update()

    def _trigger_preview_update(self):
        """Kích hoạt debounce cập nhật xem trước để không bị nghẽn GUI khi người dùng gõ phím"""
        self._preview_timer.start()

    def _get_rename_options(self) -> Optional[BatchRenameOptions]:
        if not self.chk_enable_rename_tab.isChecked():
            return None
        return BatchRenameOptions(
            remove_chuanhoa=self.chk_remove_chuanhoa.isChecked(),
            remove_custom_text=self.txt_remove_custom.text().strip(),
            enable_find_replace=self.chk_enable_name_fr.isChecked(),
            find_text=self.txt_name_find.text().strip(),
            replace_text=self.txt_name_replace.text(),
            match_case=self.chk_name_match_case.isChecked(),
            prefix=self.txt_name_prefix.text(),
            suffix=self.txt_name_suffix.text(),
            case_transform=self.cb_name_case.currentData() or "keep",
            enable_numbering=self.chk_name_numbering.isChecked(),
            collision_policy=self.cb_collision_policy.currentData() or "auto_index"
        )

    def _run_rename_preview_calculation(self):
        """Cập nhật bảng xem trước siêu tốc với setUpdatesEnabled(False)"""
        if not hasattr(self, "table_preview"):
            return

        opts = self._get_rename_options()
        if not opts:
            self.table_preview.setRowCount(0)
            self.lbl_preview_summary.setText("Chức năng đổi tên tệp đang tắt (chỉ chỉnh sửa nội dung).")
            return

        previews = BatchEditor.preview_rename(self.file_paths, opts)

        self.table_preview.setUpdatesEnabled(False)
        self.table_preview.setRowCount(len(previews))
        changed_count = 0

        for row, item in enumerate(previews):
            # STT
            stt_item = QTableWidgetItem(str(row + 1))
            stt_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            stt_item.setForeground(QColor("#64748b"))
            self.table_preview.setItem(row, 0, stt_item)

            # Tên cũ
            old_item = QTableWidgetItem(item["old_name"])
            old_item.setForeground(QColor("#1e293b"))
            self.table_preview.setItem(row, 1, old_item)

            # Tên mới
            new_item = QTableWidgetItem(item["new_name"])
            font = new_item.font()
            if item["changed"]:
                changed_count += 1
                font.setBold(True)
                new_item.setFont(font)
                new_item.setForeground(QColor("#2563eb"))
            else:
                new_item.setForeground(QColor("#64748b"))
            self.table_preview.setItem(row, 2, new_item)

            # Trạng thái
            st_item = QTableWidgetItem(item["status"])
            st_font = st_item.font()
            st_font.setBold(True)
            st_item.setFont(st_font)
            st_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)

            if item["status_type"] == "rename":
                st_item.setForeground(QColor("#16a34a"))
            elif item["status_type"] == "collision":
                st_item.setForeground(QColor("#d97706"))
            elif item["status_type"] == "error":
                st_item.setForeground(QColor("#dc2626"))
            else:
                st_item.setForeground(QColor("#94a3b8"))
            self.table_preview.setItem(row, 3, st_item)

        self.table_preview.setUpdatesEnabled(True)
        self.lbl_preview_summary.setText(
            f"Xem trước danh sách đổi tên: Có {changed_count}/{len(previews)} tệp sẽ được đổi tên"
        )

    def _start_combined_batch(self):
        """Thực hiện xử lý hàng loạt: chỉnh sửa nội dung và đổi tên đồng thời"""
        if not self.file_paths:
            QMessageBox.warning(self, "Cảnh báo", "Vui lòng chọn ít nhất một tệp để xử lý.")
            return

        enable_content = self.chk_enable_content_tab.isChecked()
        enable_rename = self.chk_enable_rename_tab.isChecked()

        if not enable_content and not enable_rename:
            QMessageBox.warning(
                self,
                "Cảnh báo",
                "Vui lòng tích chọn ít nhất một chức năng:\n"
                "• 'Bật chỉnh sửa nội dung văn bản'\n"
                "• 'Bật đổi tên tệp hàng loạt'"
            )
            return

        out_dir = self.txt_out_folder.text().strip() if self.rb_custom_dir.isChecked() else None
        if self.rb_custom_dir.isChecked() and (not out_dir or not os.path.exists(out_dir)):
            QMessageBox.warning(self, "Cảnh báo", "Thư mục xuất tệp riêng không hợp lệ hoặc chưa được chọn.")
            return

        edit_options = None
        if enable_content:
            edit_options = BatchEditOptions(
                enable_location=self.chk_enable_loc.isChecked(),
                new_location=self.txt_new_loc.text().strip(),
                sync_agency_location=self.chk_sync_agency.isChecked(),
                enable_doc_number=self.chk_enable_doc_num.isChecked(),
                doc_number_mode=self.cb_doc_num_mode.currentData(),
                new_symbol=self.txt_new_symbol.text().strip(),
                custom_doc_number=self.txt_custom_doc_num.text().strip(),
                enable_date=self.chk_enable_date.isChecked(),
                date_mode=self.cb_date_mode.currentData(),
                custom_date_text=self.txt_custom_date.text().strip(),
                enable_agency=self.chk_custom_agency.isChecked(),
                new_superior_agency=self.txt_c_sup.text().strip(),
                new_issuing_agency=self.txt_c_iss.text().strip(),
                auto_fix_agency_split=self.chk_auto_agency.isChecked(),
                enable_find_replace=self.chk_enable_fr.isChecked(),
                find_text=self.txt_find.text().strip(),
                replace_text=self.txt_replace.text(),
                match_case=self.chk_match_case.isChecked(),
                output_dir=out_dir
            )

        rename_options = self._get_rename_options() if enable_rename else None

        # Hỏi xác nhận ngắn gọn
        action_desc = []
        if enable_content:
            action_desc.append("Chỉnh sửa nội dung chuẩn NĐ 30")
        if enable_rename:
            action_desc.append("Đổi tên tệp hàng loạt")

        confirm_msg = f"Bắt đầu xử lý {len(self.file_paths)} tệp đã chọn?\nTác vụ: {' & '.join(action_desc)}"
        if self.rb_overwrite.isChecked() and not out_dir:
            confirm_msg += "\n(Lưu ý: Tệp sẽ được cập nhật trực tiếp trên ổ đĩa)."

        confirm = QMessageBox.question(
            self,
            "Xác nhận xử lý hàng loạt",
            confirm_msg,
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.Yes
        )
        if confirm != QMessageBox.StandardButton.Yes:
            return

        self.btn_execute.setEnabled(False)
        self.progress_bar.setVisible(True)
        self.progress_bar.setRange(0, len(self.file_paths))
        self.progress_bar.setValue(0)
        self.lbl_status.setText("Đang khởi chạy xử lý hàng loạt...")

        self.worker_thread = BatchCombinedWorkerThread(self.file_paths, edit_options, rename_options)
        self.worker_thread.progress_updated.connect(self._on_worker_progress)
        self.worker_thread.finished_processing.connect(self._on_worker_finished)
        self.worker_thread.start()

    def _on_worker_progress(self, current: int, total: int, msg: str):
        self.progress_bar.setValue(current)
        self.lbl_status.setText(msg)

    def _on_worker_finished(self, results: List[BatchEditItemResult]):
        self.btn_execute.setEnabled(True)
        self.progress_bar.setValue(len(self.file_paths))

        success_count = sum(1 for r in results if r.success)
        fail_count = len(results) - success_count

        # Cập nhật lại danh sách đường dẫn tệp sau khi xử lý & đổi tên
        parent = self.parent()
        updated_paths = []
        path_mapping = {}

        for res in results:
            if res.success and res.output_path:
                updated_paths.append(res.output_path)
                path_mapping[res.file_path] = res.output_path
            else:
                updated_paths.append(res.file_path)

        self.file_paths = updated_paths
        self._file_paths_set = set(self.file_paths)

        # Cập nhật hiển thị danh sách cột trái siêu tốc
        self.list_files.setUpdatesEnabled(False)
        self.list_files.clear()
        self.list_files.addItems([os.path.basename(p) for p in self.file_paths])
        self.list_files.setUpdatesEnabled(True)

        # Cập nhật danh sách ở màn hình chính nếu có
        if parent and hasattr(parent, "files_queue") and hasattr(parent, "table_files"):
            parent.table_files.setUpdatesEnabled(False)
            for idx, old_p in enumerate(parent.files_queue):
                if old_p in path_mapping:
                    new_p = path_mapping[old_p]
                    parent.files_queue[idx] = new_p
                    item = parent.table_files.item(idx, 0)
                    if item:
                        item.setText(os.path.basename(new_p))
            parent.table_files.setUpdatesEnabled(True)

        self._trigger_preview_update()

        msg = f"Đã hoàn thành xử lý hàng loạt!\n- Thành công: {success_count}/{len(results)} tệp"
        if fail_count > 0:
            msg += f"\n- Thất bại: {fail_count} tệp"

        self.lbl_status.setText(f"Hoàn thành: {success_count} tệp thành công.")
        QMessageBox.information(self, "Kết quả xử lý hàng loạt", msg)

    def _on_doc_num_mode_changed(self, idx: int):
        data = self.cb_doc_num_mode.currentData()
        self.txt_custom_doc_num.setVisible(data == "custom")

    def _on_date_mode_changed(self, idx: int):
        data = self.cb_date_mode.currentData()
        self.txt_custom_date.setVisible(data == "custom")

    def _browse_output_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục lưu tệp kết quả")
        if folder:
            self.txt_out_folder.setText(folder)
            self.rb_custom_dir.setChecked(True)

    def _add_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Chọn tệp cần xử lý", "", "Tệp văn bản (*.docx *.pdf);;Mọi tệp (*.*)")
        if files:
            new_items = []
            for f in files:
                if f not in self._file_paths_set:
                    self._file_paths_set.add(f)
                    self.file_paths.append(f)
                    new_items.append(os.path.basename(f))
            if new_items:
                self.list_files.setUpdatesEnabled(False)
                self.list_files.addItems(new_items)
                self.list_files.setUpdatesEnabled(True)
                self.lbl_files_hdr.setText(f"Danh sách tệp áp dụng ({len(self.file_paths)} tệp):")
                self._trigger_preview_update()

    def _add_directory(self):
        folder = QFileDialog.getExistingDirectory(self, "Chọn thư mục chứa tệp")
        if folder:
            new_items = []
            for root, _, files in os.walk(folder):
                for f in files:
                    if not f.startswith("~$") and f.lower().endswith((".docx", ".pdf")):
                        full_p = os.path.join(root, f)
                        if full_p not in self._file_paths_set:
                            self._file_paths_set.add(full_p)
                            self.file_paths.append(full_p)
                            new_items.append(f)
            if new_items:
                self.list_files.setUpdatesEnabled(False)
                self.list_files.addItems(new_items)
                self.list_files.setUpdatesEnabled(True)
                self.lbl_files_hdr.setText(f"Danh sách tệp áp dụng ({len(self.file_paths)} tệp):")
                self._trigger_preview_update()

    def _remove_selected_file(self):
        row = self.list_files.currentRow()
        if 0 <= row < len(self.file_paths):
            p = self.file_paths.pop(row)
            self._file_paths_set.discard(p)
            self.list_files.takeItem(row)
            self.lbl_files_hdr.setText(f"Danh sách tệp áp dụng ({len(self.file_paths)} tệp):")
            self._trigger_preview_update()

    def _clear_files(self):
        self.file_paths.clear()
        self._file_paths_set.clear()
        self.list_files.clear()
        self.lbl_files_hdr.setText("Danh sách tệp áp dụng (0 tệp):")
        self._trigger_preview_update()
