"""
ui/main_window.py: Giao diện chính của ứng dụng Auto-Formatter (PyQt6)
Cung cấp trải nghiệm kéo-thả trực quan, xử lý bất đồng bộ (QThread),
thống kê báo cáo theo thẻ và bảng cảnh báo chi tiết.
"""

import os
import sys
import json
from typing import List, Dict, Any, Optional

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QPushButton, QFileDialog, QTableWidget, QTableWidgetItem,
    QHeaderView, QRadioButton, QButtonGroup, QComboBox, QCheckBox,
    QProgressBar, QGroupBox, QFrame, QMessageBox, QSplitter,
    QAbstractScrollArea, QSizePolicy
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal, QMimeData, QSize
from PyQt6.QtGui import QDragEnterEvent, QDropEvent, QIcon, QFont, QColor, QPixmap

from core.formatter_engine import FormatterEngine, FormattingReport
from core.win32_helper import Win32WordHelper
from core.path_utils import get_config_dir, get_app_icon_path
from ui.icon_factory import get_ui_icon, get_ui_pixmap
from ui.styles import MAIN_STYLE
from ui.quick_wizard_dialog import QuickWizardDialog
from ui.editor.word_editor_window import WordEditorWindow


class ProcessingWorker(QThread):
    progress_updated = pyqtSignal(int, str)
    file_started = pyqtSignal(int, str)
    file_finished = pyqtSignal(int, FormattingReport, dict)
    all_finished = pyqtSignal(bool)

    def __init__(
        self,
        files: List[str],
        standard_id: str,
        profile_data: Optional[Dict[str, Any]],
        export_pdf: bool,
        config_dir: str,
        auto_ocr_scans: bool = True
    ):
        super().__init__()
        self.files = files
        self.standard_id = standard_id
        self.profile_data = profile_data
        self.export_pdf = export_pdf
        self.config_dir = config_dir
        self.auto_ocr_scans = auto_ocr_scans
        self.engine = FormatterEngine(config_dir=config_dir)
        self.win32_helper = Win32WordHelper()
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        total = len(self.files)
        # Sử dụng word_session để khởi tạo 1 tiến trình Word duy nhất cho toàn bộ batch
        with self.win32_helper.word_session() as word_app:
            for idx, file_path in enumerate(self.files):
                if self._is_cancelled:
                    break

                pct = int((idx / total) * 100)
                base_name = os.path.basename(file_path)
                self.file_started.emit(idx, base_name)

                # Nếu là file PDF, tự động nhận dạng OCR để bóc tách thành văn bản
                if file_path.lower().endswith(".pdf"):
                    self.progress_updated.emit(pct, f"Đang xử lý tệp PDF: {base_name} ({idx + 1}/{total})...")
                    report = self.engine.format_scanned_document_with_ocr(
                        input_file=file_path,
                        standard_id=self.standard_id,
                        profile_override=self.profile_data,
                        progress_callback=lambda p, msg: self.progress_updated.emit(pct + int(p * 0.5 / total), msg)
                    )
                else:
                    self.progress_updated.emit(pct, f"Đang chuẩn hóa: {base_name} ({idx + 1}/{total})...")
                    # 1. Chạy engine định dạng chính (có Safe Mode bảo vệ ảnh scan)
                    report = self.engine.format_document(
                        input_docx_path=file_path,
                        standard_id=self.standard_id,
                        profile_override=self.profile_data
                    )

                    # Tự động chuyển giao sang OCR nếu phát hiện ảnh scan và người dùng bật auto_ocr_scans
                    if report.is_scanned_image and self.auto_ocr_scans and not self._is_cancelled:
                        self.progress_updated.emit(pct, f"Tự động chạy OCR bóc tách chữ: {base_name}...")
                        report = self.engine.format_scanned_document_with_ocr(
                            input_file=file_path,
                            standard_id=self.standard_id,
                            profile_override=self.profile_data,
                            progress_callback=lambda p, msg: self.progress_updated.emit(pct + int(p * 0.5 / total), msg)
                        )

                if self._is_cancelled:
                    break

                # 2. Xử lý qua MS Word COM nếu thành công và có tùy chọn
                com_res = {"success": False, "pdf_exported": False, "pdf_path": ""}
                if report.success and not file_path.lower().endswith(".pdf"):
                    com_res = self.win32_helper.process_document_via_word(
                        docx_path=report.output_file,
                        export_pdf_also=self.export_pdf,
                        standard_id=self.standard_id,
                        word_app=word_app
                    )

                self.file_finished.emit(idx, report, com_res)

        if self._is_cancelled:
            self.progress_updated.emit(0, "Đã dừng xử lý theo yêu cầu.")
        else:
            self.progress_updated.emit(100, "Hoàn tất xử lý tất cả các tài liệu!")
        self.all_finished.emit(self._is_cancelled)


class OcrWorker(QThread):
    progress_updated = pyqtSignal(int, str)
    ocr_finished = pyqtSignal(FormattingReport, dict)

    def __init__(
        self,
        file_path: str,
        standard_id: str,
        profile_data: Optional[Dict[str, Any]],
        export_pdf: bool,
        config_dir: str
    ):
        super().__init__()
        self.file_path = file_path
        self.standard_id = standard_id
        self.profile_data = profile_data
        self.export_pdf = export_pdf
        self.config_dir = config_dir
        self.engine = FormatterEngine(config_dir=config_dir)
        self.win32_helper = Win32WordHelper()
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        base_name = os.path.basename(self.file_path)
        self.progress_updated.emit(10, f"Đang bóc tách ảnh và chạy OCR ngoại tuyến: {base_name}...")

        report = self.engine.format_scanned_document_with_ocr(
            input_file=self.file_path,
            standard_id=self.standard_id,
            profile_override=self.profile_data,
            progress_callback=lambda pct, msg: self.progress_updated.emit(pct, msg)
        )

        com_res = {"success": False, "pdf_exported": False, "pdf_path": ""}
        if report.success and not self._is_cancelled:
            com_res = self.win32_helper.process_document_via_word(
                docx_path=report.output_file,
                export_pdf_also=self.export_pdf,
                standard_id=self.standard_id
            )

        self.progress_updated.emit(100, "Hoàn tất nhận dạng OCR và chuẩn hóa!")
        self.ocr_finished.emit(report, com_res)


class DropZoneWidget(QFrame):
    files_dropped = pyqtSignal(list)

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("dropZone")
        self.setAcceptDrops(True)

        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)

        self.lbl_icon = QLabel()
        self.lbl_icon.setPixmap(get_ui_pixmap("cloud_upload", 46))
        self.lbl_icon.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.lbl_text = QLabel("KÉO VÀ THẢ TỆP DỰ THẢO HOẶC THƯ MỤC VÀO ĐÂY")
        self.lbl_text.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_text.setStyleSheet("color: #1e293b; font-weight: 700; font-size: 13px;")

        self.lbl_sub = QLabel("Hỗ trợ tệp văn bản .docx và tệp scan .pdf (Tự động quét các tệp bên trong thư mục)")
        self.lbl_sub.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.lbl_sub.setStyleSheet("color: #64748b; font-size: 11px;")

        layout.addWidget(self.lbl_icon)
        layout.addWidget(self.lbl_text)
        layout.addWidget(self.lbl_sub)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        file_paths = []
        for url in event.mimeData().urls():
            p = url.toLocalFile()
            if os.path.isdir(p):
                # Tối ưu: Kéo thả cả thư mục -> tự động quét đệ quy các tệp bên trong
                for root, _, files in os.walk(p):
                    for f in files:
                        if f.lower().endswith((".docx", ".pdf")):
                            file_paths.append(os.path.join(root, f))
            elif p.lower().endswith((".docx", ".pdf")):
                file_paths.append(p)

        if file_paths:
            self.files_dropped.emit(file_paths)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Auto-Formatter 1.0 - Chuẩn hóa Thể thức Văn bản Đảng & Hành chính")
        self.resize(1180, 780)
        self.setMinimumSize(960, 600)
        self.setStyleSheet(MAIN_STYLE)

        icon_path = get_app_icon_path()
        if icon_path and os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.config_dir = get_config_dir()
        self.files_queue: List[str] = []
        self.reports_map: Dict[int, FormattingReport] = {}
        self.com_results_map: Dict[int, dict] = {}
        self.profiles_data: Dict[str, Any] = {}
        self.current_report: Optional[FormattingReport] = None
        self.active_issue_filter: str = "all"
        self.stat_cards: List[QFrame] = []

        self._load_profiles_data()
        self._init_ui()
        self._refresh_profiles_dropdown()

    def _load_profiles_data(self):
        p_path = os.path.join(self.config_dir, "agency_profiles.json")
        if os.path.exists(p_path):
            try:
                with open(p_path, "r", encoding="utf-8") as f:
                    self.profiles_data = json.load(f)
            except Exception:
                self.profiles_data = {"profiles": []}
        else:
            self.profiles_data = {"profiles": []}

    def _init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)
        main_layout.setContentsMargins(14, 14, 14, 14)
        main_layout.setSpacing(10)

        # Splitter chia đôi màn hình: Trái (Đầu vào & Cấu hình) - Phải (Kết quả & Cảnh báo)
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # ====================== CỘT TRÁI ======================
        left_widget = QWidget()
        left_layout = QVBoxLayout(left_widget)
        left_layout.setContentsMargins(0, 0, 8, 0)
        left_layout.setSpacing(10)

        # Drop Zone
        self.drop_zone = DropZoneWidget()
        self.drop_zone.files_dropped.connect(self._add_files)
        left_layout.addWidget(self.drop_zone)

        # Nút Tạo nhanh & Soạn thảo văn bản mới
        self.btn_quick_create = QPushButton("Soạn thảo văn bản mới...")
        self.btn_quick_create.setIcon(get_ui_icon("file_word"))
        self.btn_quick_create.setIconSize(QSize(16, 16))
        self.btn_quick_create.setStyleSheet("""
            QPushButton {
                background-color: #059669; color: white; font-weight: bold;
                padding: 7px 14px; border-radius: 5px; border: none; font-size: 12px;
            }
            QPushButton:hover { background-color: #047857; }
        """)
        self.btn_quick_create.clicked.connect(self._open_quick_wizard)
        left_layout.addWidget(self.btn_quick_create)

        # Nút quản lý danh sách file
        file_btns_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("Thêm tệp...")
        self.btn_add_files.setObjectName("btnSecondary")
        self.btn_add_files.setIcon(get_ui_icon("file_add"))
        self.btn_add_files.setIconSize(QSize(16, 16))
        self.btn_add_files.clicked.connect(self._choose_files_dialog)

        self.btn_add_folder = QPushButton("Thêm thư mục...")
        self.btn_add_folder.setObjectName("btnSecondary")
        self.btn_add_folder.setIcon(get_ui_icon("folder_add"))
        self.btn_add_folder.setIconSize(QSize(16, 16))
        self.btn_add_folder.clicked.connect(self._choose_folder_dialog)

        self.btn_remove_file = QPushButton("Xóa tệp")
        self.btn_remove_file.setObjectName("btnSecondary")
        self.btn_remove_file.setIcon(get_ui_icon("trash"))
        self.btn_remove_file.setIconSize(QSize(16, 16))
        self.btn_remove_file.clicked.connect(self._remove_selected_file)

        self.btn_clear_all = QPushButton("Xóa danh sách")
        self.btn_clear_all.setObjectName("btnSecondary")
        self.btn_clear_all.setIcon(get_ui_icon("clear_all"))
        self.btn_clear_all.setIconSize(QSize(16, 16))
        self.btn_clear_all.clicked.connect(self._clear_all_files)

        file_btns_layout.addWidget(self.btn_add_files)
        file_btns_layout.addWidget(self.btn_add_folder)
        file_btns_layout.addWidget(self.btn_remove_file)
        file_btns_layout.addWidget(self.btn_clear_all)
        left_layout.addLayout(file_btns_layout)

        # Bảng danh sách file
        self.table_files = QTableWidget(0, 3)
        self.table_files.setHorizontalHeaderLabels(["Tên tệp dự thảo", "Kích thước", "Trạng thái"])
        self.table_files.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.table_files.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table_files.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        self.table_files.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        self.table_files.horizontalHeader().setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        self.table_files.setColumnWidth(1, 85)
        self.table_files.setColumnWidth(2, 115)
        self.table_files.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self.table_files.cellClicked.connect(self._on_file_selected)
        left_layout.addWidget(self.table_files, 1)

        # Khung Cấu hình
        config_group = QGroupBox("Cấu hình Thể thức & Cơ quan Ban hành")
        cfg_layout = QVBoxLayout(config_group)

        # Lựa chọn quy chuẩn
        rule_layout = QHBoxLayout()
        lbl_rule = QLabel("<b>Quy chuẩn:</b>")
        self.rb_nd30 = QRadioButton("Hành chính Nhà nước (NĐ 30/2020/NĐ-CP)")
        self.rb_hd36 = QRadioButton("Cơ quan Đảng (HD 36-HD/VPTW)")
        self.rb_nd30.setChecked(True)
        self.rb_nd30.toggled.connect(self._on_standard_changed)

        rule_layout.addWidget(lbl_rule)
        rule_layout.addWidget(self.rb_nd30)
        rule_layout.addWidget(self.rb_hd36)
        rule_layout.addStretch()
        cfg_layout.addLayout(rule_layout)

        # Hồ sơ đơn vị
        prof_layout = QHBoxLayout()
        lbl_prof = QLabel("Hồ sơ cơ quan:")
        self.cb_profiles = QComboBox()
        self.cb_profiles.currentIndexChanged.connect(self._on_profile_changed)

        self.btn_manage_prof = QPushButton("Quản lý...")
        self.btn_manage_prof.setObjectName("btnSecondary")
        self.btn_manage_prof.setIcon(get_ui_icon("settings"))
        self.btn_manage_prof.setIconSize(QSize(15, 15))
        self.btn_manage_prof.clicked.connect(self._open_profile_dialog)

        prof_layout.addWidget(lbl_prof)
        prof_layout.addWidget(self.cb_profiles, 1)
        prof_layout.addWidget(self.btn_manage_prof)
        cfg_layout.addLayout(prof_layout)

        # Tùy chọn nâng cao
        opts_layout = QVBoxLayout()
        opts_row1 = QHBoxLayout()
        self.chk_export_pdf = QCheckBox("Tự động xuất tệp PDF khi xử lý")
        self.chk_export_pdf.setChecked(False)

        self.chk_open_folder = QCheckBox("Mở thư mục khi hoàn tất")
        self.chk_open_folder.setChecked(True)

        opts_row1.addWidget(self.chk_export_pdf)
        opts_row1.addWidget(self.chk_open_folder)
        opts_row1.addStretch()

        opts_row2 = QHBoxLayout()
        self.chk_auto_ocr = QCheckBox("Tự động nhận dạng OCR khi phát hiện tệp ảnh scan")
        self.chk_auto_ocr.setChecked(True)
        self.chk_auto_ocr.setToolTip("Khi phát hiện tệp scan dạng ảnh, tự động bóc tách thành văn bản chữ chuẩn hóa")
        opts_row2.addWidget(self.chk_auto_ocr)
        opts_row2.addStretch()

        opts_layout.addLayout(opts_row1)
        opts_layout.addLayout(opts_row2)
        cfg_layout.addLayout(opts_layout)

        left_layout.addWidget(config_group)

        # Nút Chạy và Nút Hủy
        run_btns_layout = QHBoxLayout()
        self.btn_start = QPushButton("BẮT ĐẦU CHUẨN HÓA THỂ THỨC")
        self.btn_start.setObjectName("btnRun")
        self.btn_start.setIcon(get_ui_icon("play"))
        self.btn_start.setIconSize(QSize(18, 18))
        self.btn_start.clicked.connect(self._start_processing)

        self.btn_cancel = QPushButton("Dừng lại")
        self.btn_cancel.setObjectName("btnSecondary")
        self.btn_cancel.setIcon(get_ui_icon("stop"))
        self.btn_cancel.setIconSize(QSize(16, 16))
        self.btn_cancel.setStyleSheet("background-color: #fee2e2; color: #dc2626; font-weight: bold; border: 1px solid #f87171; min-height: 38px; border-radius: 6px;")
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self._cancel_processing)

        run_btns_layout.addWidget(self.btn_start, 3)
        run_btns_layout.addWidget(self.btn_cancel, 1)
        left_layout.addLayout(run_btns_layout)

        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setTextVisible(True)
        left_layout.addWidget(self.progress_bar)

        self.lbl_status = QLabel("Sẵn sàng xử lý tài liệu.")
        self.lbl_status.setStyleSheet("color: #64748b; font-size: 12px;")
        left_layout.addWidget(self.lbl_status)

        # ====================== CỘT PHẢI ======================
        right_widget = QWidget()
        right_layout = QVBoxLayout(right_widget)
        right_layout.setContentsMargins(8, 0, 0, 0)
        right_layout.setSpacing(10)

        # 4 Thẻ thống kê với biểu tượng nhận diện chuyên biệt và hỗ trợ lọc danh mục
        stats_layout = QHBoxLayout()

        self.card_paragraphs = self._create_stat_card("Đoạn văn chuẩn hóa", "0", "stat_paragraph", "all", "#eff6ff", "#bfdbfe")
        self.card_spaces = self._create_stat_card("Khoảng trắng dọn", "0", "stat_whitespace", "whitespace", "#ecfeff", "#a5f3fc")
        self.card_punct = self._create_stat_card("Chính tả & Dấu câu", "0", "stat_punctuation", "spelling_punct", "#f0fdf4", "#bbf7d0")
        self.card_warnings = self._create_stat_card("Cảnh báo viết hoa", "0", "stat_warning", "capitalization", "#fffbeb", "#fde68a")

        self.stat_cards = [self.card_paragraphs, self.card_spaces, self.card_punct, self.card_warnings]

        stats_layout.addWidget(self.card_paragraphs)
        stats_layout.addWidget(self.card_spaces)
        stats_layout.addWidget(self.card_punct)
        stats_layout.addWidget(self.card_warnings)
        right_layout.addLayout(stats_layout)

        # Bảng cảnh báo & Linter
        self.report_group = QGroupBox("Chi tiết lỗi")
        rep_layout = QVBoxLayout(self.report_group)

        self.table_issues = QTableWidget(0, 4)
        self.table_issues.setHorizontalHeaderLabels(["Loại lỗi", "Nội dung gốc", "Chuẩn hóa / Gợi ý", "Quy định & Giải thích"])
        self.table_issues.setSizeAdjustPolicy(QAbstractScrollArea.SizeAdjustPolicy.AdjustIgnored)
        self.table_issues.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        self.table_issues.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table_issues.setVerticalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAsNeeded)
        self.table_issues.setWordWrap(True)

        header = self.table_issues.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(1, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(2, QHeaderView.ResizeMode.Interactive)
        header.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)

        self.table_issues.setColumnWidth(0, 130)
        self.table_issues.setColumnWidth(1, 140)
        self.table_issues.setColumnWidth(2, 140)
        self.table_issues.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        rep_layout.addWidget(self.table_issues)

        right_layout.addWidget(self.report_group, 1)

        # Thanh nút thao tác trên file đã hoàn thành
        actions_layout = QHBoxLayout()
        self.btn_run_ocr = QPushButton("Nhận dạng OCR & Chuẩn hóa")
        self.btn_run_ocr.setIcon(get_ui_icon("ocr_scan"))
        self.btn_run_ocr.setIconSize(QSize(17, 17))
        self.btn_run_ocr.setStyleSheet("background-color: #d97706; color: white; font-weight: bold;")
        self.btn_run_ocr.setEnabled(False)
        self.btn_run_ocr.clicked.connect(self._run_ocr_on_selected_file)

        self.btn_open_word = QPushButton("Mở file Word")
        self.btn_open_word.setIcon(get_ui_icon("file_word"))
        self.btn_open_word.setIconSize(QSize(17, 17))
        self.btn_open_word.setStyleSheet("background-color: #2563eb; color: white; font-weight: 600;")
        self.btn_open_word.setEnabled(False)
        self.btn_open_word.clicked.connect(self._open_output_word)

        self.btn_open_pdf = QPushButton("Mở file PDF")
        self.btn_open_pdf.setIcon(get_ui_icon("file_pdf"))
        self.btn_open_pdf.setIconSize(QSize(17, 17))
        self.btn_open_pdf.setStyleSheet("background-color: #dc2626; color: white; font-weight: 600;")
        self.btn_open_pdf.setEnabled(False)
        self.btn_open_pdf.clicked.connect(self._open_output_pdf)

        self.btn_sign_pdf = QPushButton("Ký số (NĐ 30)")
        self.btn_sign_pdf.setIcon(get_ui_icon("digital_sign"))
        self.btn_sign_pdf.setIconSize(QSize(17, 17))
        self.btn_sign_pdf.setStyleSheet("background-color: #7c3aed; color: white; font-weight: 600;")
        self.btn_sign_pdf.setToolTip("Ký số điện tử bằng USB Token Ban Cơ yếu hoặc VGCA SignTool theo Nghị định 30/2020/NĐ-CP")
        self.btn_sign_pdf.setEnabled(False)
        self.btn_sign_pdf.clicked.connect(self._open_signature_dialog)

        self.btn_embedded_edit = QPushButton("Soạn thảo tích hợp")
        self.btn_embedded_edit.setIcon(get_ui_icon("file_word"))
        self.btn_embedded_edit.setIconSize(QSize(17, 17))
        self.btn_embedded_edit.setStyleSheet("background-color: #059669; color: white; font-weight: 600;")
        self.btn_embedded_edit.setEnabled(False)
        self.btn_embedded_edit.setToolTip("Mở tài liệu trong Trình Soạn thảo Hành chính WYSIWYG tích hợp")
        self.btn_embedded_edit.clicked.connect(self._open_in_embedded_editor)

        self.btn_open_dir = QPushButton("Mở thư mục")
        self.btn_open_dir.setIcon(get_ui_icon("folder_open"))
        self.btn_open_dir.setIconSize(QSize(17, 17))
        self.btn_open_dir.setEnabled(False)
        self.btn_open_dir.setObjectName("btnSecondary")
        self.btn_open_dir.clicked.connect(self._open_output_folder)

        actions_layout.addWidget(self.btn_run_ocr)
        actions_layout.addWidget(self.btn_embedded_edit)
        actions_layout.addWidget(self.btn_open_word)
        actions_layout.addWidget(self.btn_open_pdf)
        actions_layout.addWidget(self.btn_sign_pdf)
        actions_layout.addWidget(self.btn_open_dir)
        right_layout.addLayout(actions_layout)

        # Thêm 2 cột vào Splitter
        left_widget.setMinimumWidth(320)
        right_widget.setMinimumWidth(340)
        splitter.addWidget(left_widget)
        splitter.addWidget(right_widget)
        splitter.setChildrenCollapsible(False)
        splitter.setStretchFactor(0, 5)
        splitter.setStretchFactor(1, 6)

        main_layout.addWidget(splitter)

    def _create_stat_card(self, title: str, initial_value: str, icon_name: str, category: str, bg_tint: str = "#eff6ff", border_tint: str = "#bfdbfe") -> QFrame:
        card = QFrame()
        card.setObjectName("statCard")
        card.setSizePolicy(QSizePolicy.Policy.Preferred, QSizePolicy.Policy.Fixed)
        card.setMinimumWidth(80)
        card.setCursor(Qt.CursorShape.PointingHandCursor)
        card.setProperty("active", "false")
        card.category = category

        if category == "all":
            card.setToolTip("Bấm để hiển thị toàn bộ các lỗi & đoạn văn")
        elif category == "whitespace":
            card.setToolTip("Bấm để lọc và chỉ hiển thị lỗi khoảng trắng thừa")
        elif category == "spelling_punct":
            card.setToolTip("Bấm để lọc và chỉ hiển thị lỗi chính tả & dấu câu")
        elif category == "capitalization":
            card.setToolTip("Bấm để lọc và chỉ hiển thị cảnh báo quy tắc viết hoa")

        lay = QHBoxLayout(card)
        lay.setContentsMargins(10, 8, 10, 8)
        lay.setSpacing(8)

        lbl_badge = QLabel()
        lbl_badge.setFixedSize(36, 36)
        lbl_badge.setStyleSheet(f"background-color: {bg_tint}; border: 1px solid {border_tint}; border-radius: 18px;")
        lbl_badge.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lbl_badge.setPixmap(get_ui_pixmap(icon_name, 18))
        lay.addWidget(lbl_badge)

        info_lay = QVBoxLayout()
        info_lay.setSpacing(1)
        lbl_v = QLabel(initial_value)
        lbl_v.setObjectName("statValue")
        lbl_t = QLabel(title)
        lbl_t.setObjectName("statTitle")
        info_lay.addWidget(lbl_v)
        info_lay.addWidget(lbl_t)
        lay.addLayout(info_lay, 1)

        card.value_label = lbl_v  # tham chiếu nhanh
        card.mousePressEvent = lambda e, cat=category: self._on_stat_card_clicked(cat)
        return card

    def _on_stat_card_clicked(self, category: str):
        if self.active_issue_filter == category and category != "all":
            self.active_issue_filter = "all"
        else:
            self.active_issue_filter = category

        for c in self.stat_cards:
            is_act = (c.category == self.active_issue_filter and self.active_issue_filter != "all")
            c.setProperty("active", "true" if is_act else "false")
            c.style().unpolish(c)
            c.style().polish(c)

        self._render_current_issues_table()

    def _refresh_profiles_dropdown(self):
        self.cb_profiles.clear()
        self.cb_profiles.addItem("Tự động nhận diện từ tài liệu gốc (Khuyên dùng)", None)
        profiles = self.profiles_data.get("profiles", [])
        current_std = "HD36" if self.rb_hd36.isChecked() else "ND30"

        for p in profiles:
            if p.get("standard_id", "ND30") == current_std:
                self.cb_profiles.addItem(f"Áp dụng mẫu: {p.get('name', p.get('id'))}", p)

        self.cb_profiles.setCurrentIndex(0)

    def _on_standard_changed(self):
        self._refresh_profiles_dropdown()

    def _on_profile_changed(self, idx: int):
        pass

    def _open_profile_dialog(self):
        from ui.profile_dialog import ProfileDialog
        dlg = ProfileDialog(self, config_dir=self.config_dir)
        if dlg.exec():
            self._load_profiles_data()
            self._refresh_profiles_dropdown()

    def _choose_files_dialog(self):
        files, _ = QFileDialog.getOpenFileNames(
            self,
            "Chọn các tệp Word hoặc PDF cần chuẩn hóa thể thức",
            "",
            "Văn bản & Tài liệu scan (*.docx *.pdf);;Word Documents (*.docx);;PDF Documents (*.pdf)"
        )
        if files:
            self._add_files(files)

    def _choose_folder_dialog(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Chọn thư mục chứa các tệp văn bản dự thảo (.docx, .pdf)"
        )
        if folder:
            collected = []
            for root, _, files in os.walk(folder):
                for f in files:
                    if f.lower().endswith((".docx", ".pdf")):
                        collected.append(os.path.join(root, f))
            if collected:
                self._add_files(collected)
            else:
                QMessageBox.information(
                    self,
                    "Không tìm thấy tệp",
                    "Không tìm thấy tệp .docx hoặc .pdf nào trong thư mục đã chọn."
                )

    def _add_files(self, paths: List[str]):
        from core.security_guard import SecurityGuard
        for path in paths:
            ok, msg = SecurityGuard.validate_input_file(path)
            if not ok:
                QMessageBox.warning(self, "Cảnh báo An toàn Tệp", f"Từ chối tệp '{os.path.basename(path)}':\n{msg}")
                continue

            if path not in self.files_queue:
                self.files_queue.append(path)
                row = self.table_files.rowCount()
                self.table_files.insertRow(row)

                file_name = os.path.basename(path)
                size_kb = f"{os.path.getsize(path) / 1024:.1f} KB" if os.path.exists(path) else "-"

                self.table_files.setItem(row, 0, QTableWidgetItem(file_name))
                self.table_files.setItem(row, 1, QTableWidgetItem(size_kb))

                status_item = QTableWidgetItem("Chờ xử lý")
                status_item.setIcon(get_ui_icon("status_pending"))
                status_item.setForeground(QColor("#64748b"))
                self.table_files.setItem(row, 2, status_item)

        self.lbl_status.setText(f"Đã nạp {len(self.files_queue)} tài liệu trong danh sách.")

    def _remove_selected_file(self):
        row = self.table_files.currentRow()
        if 0 <= row < len(self.files_queue):
            self.files_queue.pop(row)
            self.table_files.removeRow(row)
            self.reports_map.pop(row, None)
            self.com_results_map.pop(row, None)

    def _clear_all_files(self):
        self.files_queue.clear()
        self.table_files.setRowCount(0)
        self.reports_map.clear()
        self.com_results_map.clear()
        self._clear_reports_view()

    def _clear_reports_view(self):
        self.current_report = None
        self.active_issue_filter = "all"
        for c in getattr(self, "stat_cards", []):
            c.setProperty("active", "false")
            c.style().unpolish(c)
            c.style().polish(c)
        self.card_paragraphs.value_label.setText("0")
        self.card_spaces.value_label.setText("0")
        self.card_punct.value_label.setText("0")
        self.card_warnings.value_label.setText("0")
        self.table_issues.setRowCount(0)
        if hasattr(self, "report_group"):
            self.report_group.setTitle("Chi tiết lỗi")
        self.btn_open_word.setEnabled(False)
        self.btn_embedded_edit.setEnabled(False)
        self.btn_open_pdf.setEnabled(False)
        self.btn_sign_pdf.setEnabled(False)
        self.btn_open_dir.setEnabled(False)
        self.btn_run_ocr.setEnabled(False)

    def _cancel_processing(self):
        if hasattr(self, 'worker') and self.worker.isRunning():
            self.worker.cancel()
            self.lbl_status.setText("Đang dừng tiến trình...")
            self.btn_cancel.setEnabled(False)
        if hasattr(self, 'ocr_worker') and self.ocr_worker.isRunning():
            self.ocr_worker.cancel()
            self.lbl_status.setText("Đang dừng nhận dạng OCR...")
            self.btn_cancel.setEnabled(False)

    def _start_processing(self):
        if not self.files_queue:
            QMessageBox.warning(self, "Chưa chọn tệp", "Vui lòng kéo-thả hoặc nhấn 'Thêm tệp...' để chọn tài liệu cần chuẩn hóa!")
            return

        self.btn_start.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.btn_run_ocr.setEnabled(False)
        self.progress_bar.setValue(0)
        self._clear_reports_view()

        standard_id = "HD36" if self.rb_hd36.isChecked() else "ND30"
        profile_data = self.cb_profiles.currentData()
        export_pdf = self.chk_export_pdf.isChecked()
        auto_ocr_scans = self.chk_auto_ocr.isChecked()

        # Tạo worker chạy ngầm
        self.worker = ProcessingWorker(
            files=self.files_queue,
            standard_id=standard_id,
            profile_data=profile_data,
            export_pdf=export_pdf,
            config_dir=self.config_dir,
            auto_ocr_scans=auto_ocr_scans
        )
        self.worker.progress_updated.connect(self._on_worker_progress)
        self.worker.file_started.connect(self._on_file_started)
        self.worker.file_finished.connect(self._on_file_finished)
        self.worker.all_finished.connect(self._on_all_finished)
        self.worker.start()

    def _on_file_started(self, row: int, base_name: str):
        if 0 <= row < self.table_files.rowCount():
            status_item = self.table_files.item(row, 2)
            if status_item:
                status_item.setText("Đang xử lý...")
                status_item.setIcon(get_ui_icon("status_running"))
                status_item.setForeground(QColor("#2563eb"))
            self.table_files.selectRow(row)

    def _on_worker_progress(self, pct: int, msg: str):
        self.progress_bar.setValue(pct)
        self.lbl_status.setText(msg)

    def _on_file_finished(self, row: int, report: FormattingReport, com_res: dict):
        self.reports_map[row] = report
        self.com_results_map[row] = com_res

        status_item = self.table_files.item(row, 2)
        if report.success:
            if report.is_scanned_image and not any(i.issue_type == "ocr_processed" for i in report.all_issues):
                status_item.setText("Tệp ảnh scan (Bảo vệ)")
                status_item.setIcon(get_ui_icon("status_scan"))
                status_item.setForeground(QColor("#d97706"))
            elif any(i.issue_type == "ocr_processed" for i in report.all_issues):
                status_item.setText("Đã OCR & Chuẩn hóa")
                status_item.setIcon(get_ui_icon("status_ocr_success"))
                status_item.setForeground(QColor("#16a34a"))
            else:
                status_item.setText("Đã chuẩn hóa")
                status_item.setIcon(get_ui_icon("status_success"))
                status_item.setForeground(QColor("#16a34a"))
        else:
            status_item.setText("Lỗi")
            status_item.setIcon(get_ui_icon("status_error"))
            status_item.setForeground(QColor("#dc2626"))

        # Nếu là dòng được chọn hiện tại hoặc dòng đầu tiên, hiển thị báo cáo
        if self.table_files.currentRow() == row or self.table_files.currentRow() == -1:
            self._display_report(report, com_res)

    def _on_all_finished(self, was_cancelled: bool = False):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_run_ocr.setEnabled(True)

        if was_cancelled:
            QMessageBox.information(
                self,
                "Đã Dừng",
                "Đã dừng tiến trình chuẩn hóa theo yêu cầu của bạn."
            )
            return

        QMessageBox.information(
            self,
            "Hoàn tất Xử lý",
            f"Đã xử lý xong {len(self.files_queue)} tài liệu!\n"
            "Các tệp văn bản đã được căn chỉnh thể thức chuẩn xác theo quy định."
        )

        if self.chk_open_folder.isChecked() and self.files_queue:
            first_rep = self.reports_map.get(0)
            if first_rep and os.path.exists(first_rep.output_file):
                folder = os.path.dirname(first_rep.output_file)
                os.startfile(folder)

    def _on_file_selected(self, row: int, col: int):
        report = self.reports_map.get(row)
        com_res = self.com_results_map.get(row, {})
        if report:
            self._display_report(report, com_res)
        else:
            self._clear_reports_view()

    def _display_report(self, report: FormattingReport, com_res: dict):
        self.current_report = report
        # Cập nhật số liệu trên thẻ
        self.card_paragraphs.value_label.setText(str(report.paragraphs_formatted))
        self.card_spaces.value_label.setText(str(report.spaces_fixed))
        total_spelling_punct = report.punctuation_fixed + report.spelling_fixed + report.dates_fixed + report.quotes_fixed
        self.card_punct.value_label.setText(str(total_spelling_punct))
        self.card_warnings.value_label.setText(str(report.cap_warnings_count))

        # Hiển thị bảng chi tiết lỗi theo bộ lọc đang chọn
        self._render_current_issues_table()

        # Kích hoạt các nút mở file
        self.btn_open_word.setEnabled(os.path.exists(report.output_file))
        self.btn_embedded_edit.setEnabled(os.path.exists(report.output_file))
        # Nút PDF luôn kích hoạt nếu đã có tệp Word kết quả (để tạo PDF tức thì theo yêu cầu)
        self.btn_open_pdf.setEnabled(os.path.exists(report.output_file))
        self.btn_sign_pdf.setEnabled(os.path.exists(report.output_file))
        self.btn_open_dir.setEnabled(os.path.exists(report.output_file))

        # Bật nút chạy OCR nếu là tệp ảnh scan hoặc khuyến nghị OCR
        self.btn_run_ocr.setEnabled(True)
        if report.is_scanned_image or report.ocr_recommended:
            self.btn_run_ocr.setText("Nhận dạng OCR & Chuẩn hóa")
            self.btn_run_ocr.setToolTip("Khuyên dùng: Bóc tách hình ảnh scan thành văn bản có thể chỉnh sửa")
        else:
            self.btn_run_ocr.setText("Nhận dạng lại bằng OCR")
            self.btn_run_ocr.setToolTip("Bóc tách lại nội dung tài liệu bằng công nghệ OCR ngoại tuyến")

    def _render_current_issues_table(self):
        if not self.current_report:
            self.table_issues.setRowCount(0)
            self.report_group.setTitle("Chi tiết lỗi")
            return

        all_issues = self.current_report.all_issues
        f = self.active_issue_filter

        if f == "whitespace":
            filtered = [i for i in all_issues if i.issue_type == "whitespace"]
            self.report_group.setTitle(f"Chi tiết lỗi — Khoảng trắng thừa ({len(filtered)})")
        elif f == "spelling_punct":
            filtered = [i for i in all_issues if i.issue_type in ["spelling", "punctuation", "quote", "date"]]
            self.report_group.setTitle(f"Chi tiết lỗi — Chính tả & Dấu câu ({len(filtered)})")
        elif f == "capitalization":
            filtered = [i for i in all_issues if i.issue_type == "capitalization"]
            self.report_group.setTitle(f"Chi tiết lỗi — Cảnh báo viết hoa ({len(filtered)})")
        else:
            filtered = all_issues
            self.report_group.setTitle(f"Chi tiết lỗi ({len(filtered)})")

        self.table_issues.setRowCount(0)
        for issue in filtered:
            r = self.table_issues.rowCount()
            self.table_issues.insertRow(r)

            type_name = {
                "spelling": "Chính tả & Dấu chữ",
                "punctuation": "Dấu câu hành chính",
                "whitespace": "Khoảng trắng thừa",
                "quote": "Ngoặc kép chuẩn",
                "date": "Thể thức ngày tháng",
                "capitalization": "Quy tắc Viết hoa",
                "scanned_image": "Ảnh scan",
                "ocr_processed": "Xử lý OCR"
            }.get(issue.issue_type, issue.issue_type)

            type_icons = {
                "spelling": "stat_punctuation",
                "punctuation": "stat_punctuation",
                "whitespace": "stat_whitespace",
                "quote": "stat_punctuation",
                "date": "stat_punctuation",
                "capitalization": "stat_warning",
                "scanned_image": "status_scan",
                "ocr_processed": "ocr_scan",
            }
            icon_name = type_icons.get(issue.issue_type, "stat_paragraph")
            orig_text = issue.original.strip()
            repl_text = issue.replacement.strip()
            orig_disp = orig_text if len(orig_text) <= 50 else orig_text[:47] + "..."
            repl_disp = repl_text if len(repl_text) <= 50 else repl_text[:47] + "..."

            it_type = QTableWidgetItem(get_ui_icon(icon_name), type_name)
            it_orig = QTableWidgetItem(orig_disp)
            it_orig.setToolTip(orig_text)
            it_repl = QTableWidgetItem(repl_disp)
            it_repl.setToolTip(repl_text)
            it_msg = QTableWidgetItem(issue.message)
            it_msg.setToolTip(issue.message)

            if issue.severity == "warning":
                it_type.setForeground(QColor("#d97706"))
                it_msg.setForeground(QColor("#b45309"))
            else:
                it_type.setForeground(QColor("#16a34a"))
                it_repl.setForeground(QColor("#16a34a"))

            self.table_issues.setItem(r, 0, it_type)
            self.table_issues.setItem(r, 1, it_orig)
            self.table_issues.setItem(r, 2, it_repl)
            self.table_issues.setItem(r, 3, it_msg)

    def _run_ocr_on_selected_file(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0

        if 0 <= row < len(self.files_queue):
            file_path = self.files_queue[row]
            standard_id = "HD36" if self.rb_hd36.isChecked() else "ND30"
            profile_data = self.cb_profiles.currentData()
            export_pdf = self.chk_export_pdf.isChecked()

            self.btn_run_ocr.setEnabled(False)
            self.btn_start.setEnabled(False)
            self.btn_cancel.setEnabled(True)
            self.progress_bar.setValue(5)
            self.lbl_status.setText(f"Đang bóc tách ảnh scan và chạy OCR: {os.path.basename(file_path)}...")

            self.ocr_worker = OcrWorker(
                file_path=file_path,
                standard_id=standard_id,
                profile_data=profile_data,
                export_pdf=export_pdf,
                config_dir=self.config_dir
            )
            self.ocr_worker.progress_updated.connect(self._on_worker_progress)
            self.ocr_worker.ocr_finished.connect(lambda rep, com: self._on_single_ocr_finished(row, rep, com))
            self.ocr_worker.start()

    def _on_single_ocr_finished(self, row: int, report: FormattingReport, com_res: dict):
        self.btn_start.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.btn_run_ocr.setEnabled(True)
        self.reports_map[row] = report
        self.com_results_map[row] = com_res

        status_item = self.table_files.item(row, 2)
        if report.success:
            status_item.setText("Đã OCR & Chuẩn hóa")
            status_item.setIcon(get_ui_icon("status_ocr_success"))
            status_item.setForeground(QColor("#16a34a"))
            self._display_report(report, com_res)
            QMessageBox.information(
                self,
                "Nhận dạng OCR Thành công",
                f"Đã nhận dạng thành công các dòng chữ từ ảnh scan và tự động chuẩn hóa thể thức văn bản!\n"
                f"Tệp kết quả: {os.path.basename(report.output_file)}"
            )
        else:
            status_item.setText("Lỗi OCR")
            status_item.setIcon(get_ui_icon("status_error"))
            status_item.setForeground(QColor("#dc2626"))
            QMessageBox.critical(self, "Lỗi OCR", f"Không thể nhận dạng OCR: {report.error_message}")

    def _open_quick_wizard(self):
        """Mở Wizard tạo nhanh văn bản chuẩn thể thức NĐ 30 / HD 36"""
        wiz = QuickWizardDialog(self)
        wiz.exec()

    def _open_in_embedded_editor(self):
        """Mở văn bản hiện tại trong Trình Soạn thảo Hành chính WYSIWYG tích hợp"""
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        target = report.output_file if (report and os.path.exists(report.output_file)) else (self.files_queue[row] if row < len(self.files_queue) else "")
        if target and os.path.exists(target):
            std = self._get_selected_standard()
            editor = WordEditorWindow(self, initial_metadata={"standard_id": std}, current_file_path=target)
            editor.show()
        else:
            # Nếu chưa có tệp, mở Wizard để tạo mới
            self._open_quick_wizard()

    def _open_output_word(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        if report and os.path.exists(report.output_file):
            target = os.path.abspath(report.output_file)
            try:
                os.startfile(target)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", target], shell=False)
                except Exception as e:
                    QMessageBox.warning(self, "Mở Word", f"Không thể mở tệp Word: {str(e)}")

    def _set_override_wait_cursor(self, busy: bool):
        try:
            if busy:
                QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
            else:
                QApplication.restoreOverrideCursor()
        except Exception:
            pass

    def _open_output_pdf(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        if not report or not os.path.exists(report.output_file):
            return

        base, _ = os.path.splitext(report.output_file)
        pdf_path = os.path.abspath(f"{base}.pdf")

        # Nếu file PDF chưa tồn tại -> Tự động chuyển đổi từ file Word kết quả ngay lập tức
        if not os.path.exists(pdf_path):
            self.lbl_status.setText("Đang tự động chuyển đổi sang tệp PDF bằng MS Word...")
            self._set_override_wait_cursor(True)
            try:
                from core.com_helper import ComHelper
                com = ComHelper()
                ok, msg = com.convert_docx_to_pdf(report.output_file, pdf_path)
                if not ok or not os.path.exists(pdf_path):
                    self._set_override_wait_cursor(False)
                    QMessageBox.warning(self, "Xuất PDF", f"Không thể xuất file PDF (yêu cầu máy có cài MS Word): {msg}")
                    return
                self.com_results_map.setdefault(row, {})["pdf_path"] = pdf_path
            except Exception as e:
                self._set_override_wait_cursor(False)
                QMessageBox.warning(self, "Lỗi xuất PDF", f"Lỗi khi chuyển đổi PDF: {str(e)}")
                return
            finally:
                self._set_override_wait_cursor(False)
                self.lbl_status.setText("Đã xuất tệp PDF thành công.")

        if os.path.exists(pdf_path):
            try:
                os.startfile(pdf_path)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", pdf_path], shell=False)
                except Exception:
                    QMessageBox.information(
                        self, "Mở PDF",
                        f"Tệp PDF đã sẵn sàng tại:\n{pdf_path}"
                    )

    def _open_output_folder(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0
        report = self.reports_map.get(row)
        if report and os.path.exists(report.output_file):
            folder = os.path.abspath(os.path.dirname(report.output_file))
            try:
                os.startfile(folder)
            except Exception:
                try:
                    import subprocess
                    subprocess.Popen(["explorer.exe", folder], shell=False)
                except Exception:
                    pass

    def _open_signature_dialog(self):
        row = self.table_files.currentRow()
        if row < 0 and self.files_queue:
            row = 0

        pdf_path = ""
        report = self.reports_map.get(row)

        if report and report.output_file and os.path.exists(report.output_file):
            base, ext = os.path.splitext(report.output_file)
            if ext.lower() == ".pdf":
                pdf_path = report.output_file
            else:
                pdf_path = f"{base}.pdf"
                # Nếu tệp PDF chưa tồn tại -> Tự động chuyển đổi từ file Word kết quả ngay
                if not os.path.exists(pdf_path):
                    self.lbl_status.setText("Đang tự động xuất tệp PDF để chuẩn bị ký số...")
                    self._set_override_wait_cursor(True)
                    try:
                        from core.com_helper import ComHelper
                        com = ComHelper()
                        ok, msg = com.convert_docx_to_pdf(report.output_file, pdf_path)
                        if not ok or not os.path.exists(pdf_path):
                            self._set_override_wait_cursor(False)
                            QMessageBox.warning(self, "Xuất PDF", f"Không thể xuất file PDF để ký số: {msg}")
                            return
                    except Exception as e:
                        self._set_override_wait_cursor(False)
                        QMessageBox.warning(self, "Lỗi xuất PDF", f"Lỗi chuyển đổi PDF: {str(e)}")
                        return
                    finally:
                        self._set_override_wait_cursor(False)
        elif 0 <= row < len(self.files_queue):
            orig = self.files_queue[row]
            if orig.lower().endswith(".pdf") and os.path.exists(orig):
                pdf_path = orig

        if not pdf_path or not os.path.exists(pdf_path):
            QMessageBox.information(
                self, "Chọn tài liệu",
                "Vui lòng chọn một tài liệu hoặc thực hiện chuẩn hóa để xuất PDF trước khi ký số."
            )
            return

        agency_name = ""
        prof = self.cb_profiles.currentData()
        if prof and isinstance(prof, dict):
            agency_name = prof.get("name", prof.get("id", ""))

        from ui.signature_dialog import SignatureDialog
        dlg = SignatureDialog(self, pdf_path=pdf_path, default_agency=agency_name)
        if dlg.exec() and dlg.signed_pdf_path and os.path.exists(dlg.signed_pdf_path):
            self.lbl_status.setText(f"Đã hoàn tất ký số văn bản: {os.path.basename(dlg.signed_pdf_path)}")
            try:
                os.startfile(dlg.signed_pdf_path)
            except Exception:
                pass
