"""
ui/styles.py: Bộ giao diện giao diện phẳng hiện đại (Clean Minimalist Modern UI)
Thiết kế theo chuẩn mực ứng dụng văn phòng cao cấp (Windows 11 Fluent / macOS Modern)
- Bảng màu trung tính tinh tế, độ tương phản sắc nét, màu form sáng rõ
- Loại bỏ hiệu ứng đổ bóng cồng kềnh, gradient sặc sỡ và phong cách AI-slop
- Tối ưu hóa hiển thị trên màn hình độ phân giải cao (High-DPI), không cắt chữ
"""

MAIN_STYLE = """
/* Reset & Base Typography */
QWidget {
    font-family: 'Segoe UI', -apple-system, BlinkMacSystemFont, Arial, sans-serif;
    font-size: 13px;
    color: #0f172a;
    background-color: #f8fafc;
}

QLabel {
    background-color: transparent;
    color: #0f172a;
}

QMainWindow {
    background-color: #f1f5f9;
}

QDialog {
    background-color: #ffffff;
    color: #0f172a;
}

/* Menu Bar Hiện Đại & Tối Giản */
QMenuBar {
    background-color: #ffffff;
    color: #1e293b;
    font-size: 12px;
    padding: 3px 8px;
    border-bottom: 1px solid #e2e8f0;
}

QMenuBar::item {
    background-color: transparent;
    padding: 5px 12px;
    border-radius: 4px;
}

QMenuBar::item:selected {
    background-color: #f1f5f9;
    color: #0f172a;
}

QMenu {
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 4px;
}

QMenu::item {
    padding: 6px 24px 6px 12px;
    border-radius: 4px;
}

QMenu::item:selected {
    background-color: #f1f5f9;
    color: #2563eb;
}

QMenu::separator {
    height: 1px;
    background-color: #e2e8f0;
    margin: 4px 6px;
}

/* Khung nhóm GroupBox Tinh Gọn - Tương phản cao, dễ nhìn */
QGroupBox {
    font-weight: 600;
    font-size: 12px;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    margin-top: 14px;
    background-color: #ffffff;
    padding: 14px 12px 10px 12px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 12px;
    padding: 0 6px;
    background-color: #ffffff;
    color: #1e293b;
}

/* Vùng Kéo & Thả (Drop Zone) */
#dropZone {
    border: 1.5px dashed #cbd5e1;
    border-radius: 8px;
    background-color: #ffffff;
    min-height: 105px;
}

#dropZone:hover {
    border-color: #3b82f6;
    background-color: #f8fafc;
}

/* Nút bấm Chuẩn mực (Buttons) */
QPushButton {
    background-color: #ffffff;
    color: #0f172a;
    font-weight: 500;
    font-size: 12px;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 6px 14px;
    min-height: 22px;
}

QPushButton:hover {
    background-color: #f8fafc;
    border-color: #94a3b8;
}

QPushButton:pressed {
    background-color: #f1f5f9;
}

QPushButton:disabled {
    background-color: #f1f5f9;
    color: #94a3b8;
    border-color: #e2e8f0;
}

/* Nút chính (Primary Action) */
QPushButton#btnPrimary, QPushButton.btn-primary {
    background-color: #2563eb;
    color: #ffffff;
    border: 1px solid #1d4ed8;
    font-weight: 600;
}

QPushButton#btnPrimary:hover, QPushButton.btn-primary:hover {
    background-color: #1d4ed8;
}

QPushButton#btnPrimary:pressed, QPushButton.btn-primary:pressed {
    background-color: #1e40af;
}

/* Nút Đỏ / Xuất PDF */
QPushButton#btnDanger {
    background-color: #dc2626;
    color: #ffffff;
    border: 1px solid #b91c1c;
    font-weight: 600;
}

QPushButton#btnDanger:hover {
    background-color: #b91c1c;
}

/* Nút Tím / Ký số */
QPushButton#btnPurple {
    background-color: #7c3aed;
    color: #ffffff;
    border: 1px solid #6d28d9;
    font-weight: 600;
}

QPushButton#btnPurple:hover {
    background-color: #6d28d9;
}

/* Nút Chạy Chuẩn hóa Lớn */
QPushButton#btnRun {
    background-color: #16a34a;
    color: #ffffff;
    font-size: 13px;
    font-weight: 600;
    padding: 9px 20px;
    border: 1px solid #15803d;
    border-radius: 6px;
}

QPushButton#btnRun:hover {
    background-color: #15803d;
}

QPushButton#btnRun:pressed {
    background-color: #166534;
}

/* Nút Dừng */
QPushButton#btnCancel {
    background-color: #ffffff;
    color: #dc2626;
    border: 1px solid #fecaca;
}

QPushButton#btnCancel:hover {
    background-color: #fef2f2;
    border-color: #fca5a5;
}

/* Ô nhập liệu & Dropdown & SpinBox */
QLineEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 4px 8px;
    min-height: 28px;
    color: #0f172a;
    font-size: 12.5px;
}

QLineEdit:focus {
    border-color: #2563eb;
}

QPlainTextEdit {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 6px 8px;
    color: #0f172a;
    font-size: 12.5px;
}

QPlainTextEdit:focus {
    border-color: #2563eb;
}

QComboBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 4px 28px 4px 8px;
    min-height: 28px;
    color: #0f172a;
    font-size: 12.5px;
}

QComboBox:focus {
    border-color: #2563eb;
}

QComboBox::drop-down {
    subcontrol-origin: padding;
    subcontrol-position: top right;
    width: 24px;
    border-left: 1px solid #cbd5e1;
    border-top-right-radius: 4px;
    border-bottom-right-radius: 4px;
}

QComboBox QAbstractItemView {
    background-color: #ffffff;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    selection-background-color: #f1f5f9;
    selection-color: #2563eb;
    padding: 4px;
}

QSpinBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 4px;
    padding: 3px 8px;
    min-height: 28px;
    color: #0f172a;
    font-size: 12.5px;
}

QSpinBox:focus {
    border-color: #2563eb;
}

QSpinBox::up-button, QSpinBox::down-button {
    width: 20px;
    border: none;
    background-color: #f1f5f9;
}

QSpinBox::up-button:hover, QSpinBox::down-button:hover {
    background-color: #e2e8f0;
}

/* Radio Button & Checkbox - Tương phản cao, rõ ràng */
QRadioButton, QCheckBox {
    spacing: 8px;
    font-size: 12.5px;
    color: #0f172a;
    background-color: transparent;
    padding: 3px 0;
    min-height: 22px;
}

QRadioButton::indicator, QCheckBox::indicator {
    width: 16px;
    height: 16px;
}

/* Bảng dữ liệu */
QTableWidget {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    gridline-color: #f1f5f9;
    selection-background-color: #eff6ff;
    selection-color: #1e3a8a;
}

QHeaderView::section {
    background-color: #f8fafc;
    color: #475569;
    font-weight: 600;
    font-size: 12px;
    padding: 7px 10px;
    border: none;
    border-bottom: 1px solid #e2e8f0;
}

/* Thanh cuộn hiện đại (Modern Scrollbar) */
QScrollBar:vertical {
    border: none;
    background: #f1f5f9;
    width: 8px;
    margin: 0px;
}

QScrollBar::handle:vertical {
    background: #cbd5e1;
    min-height: 25px;
    border-radius: 4px;
}

QScrollBar::handle:vertical:hover {
    background: #94a3b8;
}

QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical {
    height: 0px;
}

QScrollBar:horizontal {
    border: none;
    background: #f1f5f9;
    height: 8px;
    margin: 0px;
}

QScrollBar::handle:horizontal {
    background: #cbd5e1;
    min-width: 25px;
    border-radius: 4px;
}

QScrollBar::handle:horizontal:hover {
    background: #94a3b8;
}

QScrollBar::add-line:horizontal, QScrollBar::sub-line:horizontal {
    width: 0px;
}

/* Thanh tiến trình */
QProgressBar {
    border: 1px solid #e2e8f0;
    border-radius: 4px;
    text-align: center;
    background-color: #f1f5f9;
    height: 14px;
    color: #0f172a;
    font-size: 11px;
    font-weight: 600;
}

QProgressBar::chunk {
    background-color: #2563eb;
    border-radius: 3px;
}

/* Thẻ thống kê Tinh giản (Stat Cards) */
QFrame#statCard {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 6px;
    padding: 6px 12px;
}

QFrame#statCard:hover {
    border-color: #cbd5e1;
    background-color: #f8fafc;
}

QFrame#statCard[active="true"] {
    border: 1.5px solid #2563eb;
    background-color: #eff6ff;
}

QLabel#statValue {
    font-size: 20px;
    font-weight: 700;
    color: #0f172a;
}

QLabel#statTitle {
    font-size: 11px;
    color: #64748b;
    font-weight: 500;
}
"""
