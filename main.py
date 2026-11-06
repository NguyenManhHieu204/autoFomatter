"""
main.py: Điểm khởi chạy của ứng dụng Auto-Formatter
Hỗ trợ:
1. Chế độ Giao diện đồ họa (GUI Desktop App - PyQt6) khi chạy bình thường hoặc mở file .exe
2. Chế độ Dòng lệnh (CLI) khi cán bộ cần chuẩn hóa tự động hàng loạt qua script
"""

import sys
import os
import argparse

# Thiết lập UTF-8 cho console Windows
if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass


def run_cli(args):
    from core.formatter_engine import FormatterEngine
    from core.win32_helper import Win32WordHelper

    print("=" * 65)
    print("  AUTO-FORMATTER 2.0 - CHUẨN HÓA THỂ THỨC VĂN BẢN ĐẢNG & NHÀ NƯỚC")
    print("=" * 65)

    input_path = args.file
    if not os.path.exists(input_path):
        print(f"[!] Lỗi: Tệp hoặc thư mục đầu vào không tồn tại: {input_path}")
        sys.exit(1)

    std = args.standard.upper()
    print(f"[*] Quy chuẩn áp dụng: {std}")

    # Thu thập danh sách tệp nếu là thư mục
    if os.path.isdir(input_path):
        target_files = []
        for root, _, files in os.walk(input_path):
            for f in files:
                if f.lower().endswith((".docx", ".pdf")) and not f.endswith("_ChuanHoa.docx") and not f.endswith(".bak"):
                    target_files.append(os.path.join(root, f))
        if not target_files:
            print(f"[!] Không tìm thấy tệp .docx hoặc .pdf nào trong thư mục: {input_path}")
            sys.exit(1)
        print(f"[*] Đã phát hiện {len(target_files)} tệp cần xử lý.")
    else:
        target_files = [input_path]

    engine = FormatterEngine()
    helper = Win32WordHelper()

    # Sử dụng phiên Word COM chung cho toàn bộ batch
    with helper.word_session() as word_app:
        for idx, current_file in enumerate(target_files, 1):
            print(f"\n--- [{idx}/{len(target_files)}] Đang xử lý: {os.path.basename(current_file)} ---")
            output_file = args.output if (len(target_files) == 1 and args.output) else None

            if args.ocr or current_file.lower().endswith(".pdf"):
                print("    [*] Đang chạy OCR nhận dạng chữ từ ảnh scan ngoại tuyến...")
                report = engine.format_scanned_document_with_ocr(
                    input_file=current_file,
                    output_docx_path=output_file,
                    standard_id=std,
                    progress_callback=lambda p, msg: print(f"        [{p}%] {msg}")
                )
            else:
                report = engine.format_document(
                    input_docx_path=current_file,
                    output_docx_path=output_file,
                    standard_id=std
                )

                # Tự động chuyển giao OCR nếu là ảnh scan
                if report.is_scanned_image and args.ocr:
                    print("    [*] Phát hiện tài liệu scan dạng ảnh, tự động bóc tách chữ bằng OCR...")
                    report = engine.format_scanned_document_with_ocr(
                        input_file=current_file,
                        output_docx_path=output_file,
                        standard_id=std,
                        progress_callback=lambda p, msg: print(f"        [{p}%] {msg}")
                    )

            if not report.success:
                print(f"    [!] Chuẩn hóa thất bại: {report.error_message}")
                continue

            print(f"    [✔] Đã tạo tệp chuẩn hóa: {report.output_file}")
            if report.safe_mode_activated and not any(i.issue_type == "ocr_processed" for i in report.all_issues):
                print(f"        [!] CHÚ Ý: Phát hiện tệp dạng ảnh scan ({report.image_count} hình ảnh)!")
                print("        [!] Đã kích hoạt Chế độ Bảo vệ Hình ảnh (Safe Mode) để không xóa hay làm lệch ảnh.")
                print("        [!] Bạn có thể thêm cờ --ocr để tự động bóc tách thành văn bản chữ: python main.py --cli -f <file> --ocr")
            else:
                print(f"        - Số đoạn văn đã định dạng : {report.paragraphs_formatted}")
                print(f"        - Khoảng trắng thừa đã dọn : {report.spaces_fixed}")
                print(f"        - Dấu câu đã chuẩn hóa     : {report.punctuation_fixed}")
                print(f"        - Ngày tháng đã sửa        : {report.dates_fixed}")
                print(f"        - Cảnh báo viết hoa        : {report.cap_warnings_count}")

            if args.pdf:
                print("    [*] Đang xuất tệp PDF qua MS Word COM...")
                res = helper.process_document_via_word(
                    docx_path=report.output_file,
                    export_pdf_also=True,
                    standard_id=std,
                    word_app=word_app
                )
                if res["pdf_exported"]:
                    print(f"    [✔] Đã xuất PDF thành công: {res['pdf_path']}")
                else:
                    print(f"    [!] Không thể xuất PDF: {res['message']}")


def run_gui():
    # Cấu hình High DPI cho màn hình độ phân giải cao
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"

    # Đăng ký AppUserModelID để Windows Taskbar hiển thị icon chuẩn của ứng dụng
    if sys.platform.startswith('win'):
        try:
            import ctypes
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("DangHanhChinh.AutoFormatter.2.0")
        except Exception:
            pass

    from PyQt6.QtWidgets import QApplication, QMessageBox
    from PyQt6.QtGui import QIcon
    from ui.main_window import MainWindow
    from core.path_utils import get_app_icon_path

    # Bắt ngoại lệ toàn cục tránh Qt6 abort / qFatal đột ngột
    def handle_exception(exc_type, exc_value, exc_traceback):
        import traceback
        if issubclass(exc_type, KeyboardInterrupt):
            sys.__excepthook__(exc_type, exc_value, exc_traceback)
            return
        err_msg = "".join(traceback.format_exception(exc_type, exc_value, exc_traceback))
        try:
            print(f"[ERROR] Ngoại lệ chưa bắt:\n{err_msg}", file=sys.stderr)
        except Exception:
            pass
        try:
            QMessageBox.critical(
                None,
                "Lỗi ứng dụng",
                f"Đã phát sinh sự cố:\n{str(exc_value)}\n\nỨng dụng vẫn đang tiếp tục hoạt động."
            )
        except Exception:
            pass

    sys.excepthook = handle_exception

    app = QApplication(sys.argv)
    app.setApplicationName("Auto-Formatter")
    app.setApplicationDisplayName("")

    # Thiết lập icon ứng dụng cho khay hệ thống và thanh tác vụ
    icon_path = get_app_icon_path()
    if icon_path and os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    window = MainWindow()
    if icon_path and os.path.exists(icon_path):
        window.setWindowIcon(QIcon(icon_path))

    window.show()

    sys.exit(app.exec())


def main():
    parser = argparse.ArgumentParser(description="Auto-Formatter: Công cụ Chuẩn hóa Thể thức Văn bản Đảng & Hành chính")
    parser.add_argument("--cli", action="store_true", help="Chạy ở chế độ dòng lệnh (CLI)")
    parser.add_argument("-f", "--file", type=str, help="Đường dẫn tệp .docx cần chuẩn hóa")
    parser.add_argument("-o", "--output", type=str, help="Đường dẫn tệp xuất ra (mặc định: [tên]_ChuanHoa.docx)")
    parser.add_argument("-s", "--standard", type=str, default="ND30", choices=["ND30", "HD36"], help="Quy chuẩn: ND30 hoặc HD36")
    parser.add_argument("--pdf", action="store_true", help="Tự động xuất thêm file PDF qua MS Word")
    parser.add_argument("--ocr", action="store_true", help="Kích hoạt nhận dạng chữ OCR cho tệp ảnh scan hoặc PDF")

    args, unknown = parser.parse_known_args()

    if args.cli or args.file:
        run_cli(args)
    else:
        run_gui()


if __name__ == "__main__":
    main()
