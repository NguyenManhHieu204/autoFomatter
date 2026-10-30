# CÔNG CỤ TỰ ĐỘNG CHUẨN HÓA THỂ THỨC VĂN BẢN ĐẢNG & HÀNH CHÍNH (AUTO-FORMATTER)
**Phiên bản:** 1.0  
**Môi trường triển khai:** Desktop App chạy Ngoại tuyến hoàn toàn (Offline 100% / Air-gapped) trên Windows.

---

## 1. Mục đích & Phạm vi Quy chuẩn Áp dụng
Hệ thống giải quyết triệt để bài toán tự động hóa quy trình căn chỉnh thể thức, kỹ thuật trình bày và rà soát lỗi chính tả hành chính đối với các dự thảo văn bản (`.docx`) trước khi trình ký, bảo đảm tính thống nhất và tuân thủ tuyệt đối quy định hiện hành:

1. **Khối Đảng:** [Hướng dẫn số 36-HD/VPTW](file:///d:/AutoFormatter/config/hd36_rules.json) ngày 03/4/2018 của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng.
2. **Khối Hành chính Nhà nước:** [Nghị định số 30/2020/NĐ-CP](file:///d:/AutoFormatter/config/nd30_rules.json) ngày 05/3/2020 của Chính phủ về công tác văn thư.

---

## 2. Kiến trúc 4 Module Cốt lõi

- **Module 1 - Bộ máy Định dạng Trang & Đoạn văn (`core/formatter_engine.py`):**
  - Khổ giấy: Chuẩn A4 (210 mm x 297 mm).
  - Định lượng lề: Lề trên 20 mm, lề dưới 20 mm, lề trái 30 mm, lề phải 15 mm.
  - Phông chữ: Times New Roman (TCVN 6909:2001), cỡ chữ nội dung 14 pt (đồng nhất).
  - Giãn dòng 1.2 lines; khoảng cách đoạn: tối thiểu 6 pt; lùi đầu dòng chuẩn 1.27 cm (1 Tab).
  - Tự động đánh số trang ở giữa lề trên (ẩn ở trang đầu).
- **Module 2 - Tự động Dựng Khung Đầu Văn bản (`core/header_footer_builder.py`):**
  - Sử dụng cấu trúc Bảng ẩn viền (2 cột) để định vị tuyệt đối Quốc hiệu / Tiêu đề Đảng và Tên cơ quan ban hành, khắc phục hoàn toàn lỗi dùng phím Space hoặc Tab gây lệch dòng khi in.
  - Xử lý đường kẻ chân tiêu ngữ và ngày tháng chuẩn xác.
- **Module 3 - Khởi tạo Khung Cuối Văn bản (`core/header_footer_builder.py`):**
  - Tự động dựng bảng 2 cột không viền cho Nơi nhận (trái) và Quyền hạn, Chức danh, Chữ ký (phải).
  - Phân định rõ ràng: Khối Đảng có dấu gạch chéo (`T/M`, `K/T`, `T/L`); Khối Nhà nước có dấu chấm (`TM.`, `KT.`, `TL.`, `TUQ.`).
  - Tự động chừa 30–35 mm để ký tay và đóng dấu.
- **Module 4 - Bộ máy Rà soát Lỗi Kỹ thuật & Chính tả Hành chính (`core/spell_linter.py`):**
  - Tự động làm sạch khoảng trắng thừa (2+ spaces -> 1 space).
  - Chuẩn hóa dấu câu (xóa khoảng trắng trước dấu câu `, . ; : ! ?`, thêm khoảng trắng sau dấu câu).
  - Chuẩn hóa ngoặc kép thẳng `""` thành ngoặc kép cong `“ ”`.
  - Chuẩn hóa thể thức ngày tháng: số 0 trước ngày 1–9; tháng 1, 2 có số 0; tháng 3–12 không có số 0.
  - Rà soát lỗi viết hoa tùy tiện danh từ chung và nhắc nhở viết hoa tên cơ quan, tổ chức theo Phụ lục II Nghị định 30/2020/NĐ-CP.
- **Module Mở rộng - Bộ máy OCR Bóc tách Ảnh Scan Ngoại tuyến (`core/ocr_engine.py`):**
  - Tự động phát hiện tài liệu dạng ảnh scan hoặc tệp Word chỉ chứa ảnh chụp trang (chưa có lớp chữ OCR).
  - Tự động kích hoạt **Chế độ Bảo vệ Hình ảnh (Safe Mode)**: Tuyệt đối không xóa, không đè bảng làm mất hoặc lệch ảnh scan.
  - Tích hợp động cơ **RapidOCR (ONNX Runtime)** chạy offline 100% không cần Internet để trích xuất chữ từ ảnh scan/PDF thành văn bản có thể chỉnh sửa và tự động đưa vào quy trình chuẩn hóa thể thức.

---

## 3. Cấu trúc Thư mục Dự án

```
AutoFormatter/
├── main.py                      # Điểm khởi chạy (GUI Desktop / CLI)
├── config/
│   ├── hd36_rules.json          # Quy chuẩn Hướng dẫn 36-HD/VPTW
│   ├── nd30_rules.json          # Quy chuẩn Nghị định 30/2020/NĐ-CP
│   ├── agency_profiles.json     # Danh mục hồ sơ cơ quan đã lưu
│   ├── capitalization_dict.txt  # Từ điển danh từ/tổ chức hành chính
│   └── capitalization_rules.json# Quy tắc viết hoa tên cơ quan & chức danh
├── core/
│   ├── doc_parser.py            # Bóc tách cấu trúc file & phát hiện ảnh scan
│   ├── formatter_engine.py      # Module 1: Định dạng trang, đoạn văn, bảng
│   ├── header_footer_builder.py # Module 2 & 3: Bảng ẩn tiêu đề và chữ ký
│   ├── spell_linter.py          # Module 4: Regex dọn rác và linter viết hoa
│   ├── ocr_engine.py            # Module OCR: Nhận dạng ảnh scan ngoại tuyến
│   ├── win32_helper.py          # Điều khiển MS Word COM & xuất PDF
│   └── path_utils.py            # Xử lý đường dẫn tài nguyên cho PyInstaller
├── templates/
│   ├── Mau_Van_Ban_Dang_HD36.docx
│   └── Mau_Van_Ban_Hanh_Chinh_ND30.docx
├── ui/
│   ├── main_window.py           # Giao diện chính PyQt6 với Drag & Drop và nút OCR
│   ├── profile_dialog.py        # Hộp thoại quản lý hồ sơ cơ quan
│   └── styles.py                # Giao diện hiện đại Administrative Modern
├── tests/                       # Bộ kiểm thử tự động
├── build_exe.spec               # File cấu hình PyInstaller đóng gói 1 tệp .exe
├── installer.iss                # Script Inno Setup tạo bộ cài đặt Windows
└── requirements.txt             # Danh sách thư viện Python
```

---

## 4. Hướng dẫn Sử dụng

### 4.1. Khởi chạy Giao diện Đồ họa (GUI)
Nhấp đúp vào file `AutoFormatter.exe` (hoặc chạy lệnh):
```bash
python main.py
```
1. **Kéo - thả** một hoặc nhiều file `.docx` cần chuẩn hóa vào vùng nhận tệp.
2. Chọn **Quy chuẩn** (Hành chính NĐ 30 hoặc Khối Đảng HD 36) và **Hồ sơ cơ quan**.
3. Nhấn **▶ BẮT ĐẦU CHUẨN HÓA THỂ THỨC**.
4. Xem báo cáo chi tiết các lỗi đã xử lý và nhấp nút mở file Word kết quả (`_ChuanHoa.docx`) hoặc tệp PDF.

### 4.2. Chế độ Dòng lệnh (CLI)
Chuẩn hóa nhanh file đơn lẻ hoặc theo kịch bản tự động:
```bash
# Chuẩn hóa theo Nghị định 30 và xuất kèm PDF
python main.py --cli -f "du_thao.docx" -s ND30 --pdf

# Chuẩn hóa theo Hướng dẫn 36 (Khối Đảng)
python main.py --cli -f "nghi_quyet.docx" -s HD36
```

---

## 5. Hướng dẫn Đóng gói Thành file Thực thi (.exe)
Ứng dụng có sẵn cấu hình PyInstaller để đóng gói thành 1 file duy nhất không phụ thuộc Python:
```bash
pyinstaller build_exe.spec --noconfirm
```
File thực thi độc lập sẽ xuất hiện tại thư mục `dist/AutoFormatter.exe`.
Có thể copy sang máy tính bất kỳ chạy Windows mà không cần cài đặt thêm phần mềm phụ trợ.
