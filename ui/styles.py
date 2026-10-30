"""
ui/styles.py: Bộ giao diện QSS chuyên nghiệp cho ứng dụng Auto-Formatter
Phong cách: Administrative Modern - Sang trọng, chuẩn mực cơ quan Nhà nước và Đảng
"""

MAIN_STYLE = """
QWidget {
    font-family: 'Segoe UI', 'Helvetica Neue', Arial, sans-serif;
    font-size: 13px;
    color: #1e293b;
    background-color: #f8fafc;
}

QMainWindow {
    background-color: #f1f5f9;
}

/* Header Banner */
#headerBanner {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #0f172a, stop:0.5 #1e3a8a, stop:1 #b91c1c);
    border-radius: 8px;
    padding: 16px;
    color: white;
}

#headerTitle {
    font-size: 18px;
    font-weight: bold;
    color: #ffffff;
}

#headerSubtitle {
    font-size: 12px;
    color: #cbd5e1;
}

/* GroupBox Card */
QGroupBox {
    font-weight: bold;
    font-size: 13px;
    color: #0f172a;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    margin-top: 14px;
    background-color: #ffffff;
    padding-top: 14px;
}

QGroupBox::title {
    subcontrol-origin: margin;
    subcontrol-position: top left;
    left: 14px;
    padding: 0 6px;
    background-color: #ffffff;
}

/* Drag & Drop Zone */
#dropZone {
    border: 2px dashed #94a3b8;
    border-radius: 10px;
    background-color: #f8fafc;
    min-height: 110px;
}

#dropZone:hover {
    border-color: #2563eb;
    background-color: #eff6ff;
}

/* Buttons */
QPushButton {
    background-color: #2563eb;
    color: white;
    font-weight: 600;
    border: none;
    border-radius: 6px;
    padding: 8px 16px;
    min-height: 20px;
}

QPushButton:hover {
    background-color: #1d4ed8;
}

QPushButton:pressed {
    background-color: #1e40af;
}

QPushButton:disabled {
    background-color: #94a3b8;
    color: #e2e8f0;
}

/* Secondary Button */
QPushButton#btnSecondary {
    background-color: #ffffff;
    color: #334155;
    border: 1px solid #cbd5e1;
}

QPushButton#btnSecondary:hover {
    background-color: #f1f5f9;
    border-color: #94a3b8;
}

/* Big Run Button */
QPushButton#btnRun {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #16a34a, stop:1 #15803d);
    color: white;
    font-size: 15px;
    font-weight: bold;
    padding: 12px 24px;
    border-radius: 8px;
}

QPushButton#btnRun:hover {
    background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 #15803d, stop:1 #166534);
}

/* Input Fields */
QLineEdit, QComboBox {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    padding: 6px 10px;
    color: #0f172a;
}

QLineEdit:focus, QComboBox:focus {
    border-color: #2563eb;
}

/* Table Widget */
QTableWidget {
    background-color: #ffffff;
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    gridline-color: #f1f5f9;
    selection-background-color: #dbeafe;
    selection-color: #1e3a8a;
}

QHeaderView::section {
    background-color: #f8fafc;
    color: #475569;
    font-weight: bold;
    font-size: 12px;
    padding: 8px;
    border: none;
    border-bottom: 1px solid #cbd5e1;
}

/* Progress Bar */
QProgressBar {
    border: 1px solid #cbd5e1;
    border-radius: 6px;
    text-align: center;
    background-color: #e2e8f0;
    height: 18px;
    color: #0f172a;
    font-weight: bold;
}

QProgressBar::chunk {
    background-color: #2563eb;
    border-radius: 5px;
}

/* Stat Cards Hiện Đại & Tương Tác */
QFrame#statCard {
    background-color: #ffffff;
    border: 1px solid #e2e8f0;
    border-radius: 10px;
    padding: 8px 12px;
}

QFrame#statCard:hover {
    border-color: #3b82f6;
    background-color: #f0f7ff;
}

QFrame#statCard[active="true"] {
    border: 2px solid #2563eb;
    background-color: #eff6ff;
}

QLabel#statValue {
    font-size: 22px;
    font-weight: 700;
    color: #0f172a;
}

QLabel#statTitle {
    font-size: 11px;
    color: #64748b;
    font-weight: 600;
}
"""
