# CÔNG CỤ TỰ ĐỘNG CHUẨN HÓA THỂ THỨC VĂN BẢN ĐẢNG & HÀNH CHÍNH (AUTO-FORMATTER)
**Phiên bản:** 2.0 (Bản Thương Mại / Commercial Edition)  
**Môi trường triển khai:** Ứng dụng Desktop chạy Ngoại tuyến hoàn toàn (Offline 100% / Air-gapped) trên hệ điều hành Windows.  
**Cơ chế vận hành:** Tải về là chạy ngay (Portable / Zero-Install), không yêu cầu cài đặt phần mềm phụ trợ, không cần kết nối Internet.

---

## 1. MỤC ĐÍCH & PHẠM VI QUY CHUẨN PHÁP LÝ

Hệ thống giải quyết triệt để bài toán chuẩn hóa thể thức, kỹ thuật trình bày, rà soát lỗi chính tả hành chính và ký số điện tử đối với các văn bản dự thảo trước khi trình ký và phát hành:

1. **Khối Hành chính Nhà nước:** Tuân thủ tuyệt đối [Nghị định số 30/2020/NĐ-CP](file:///d:/AutoFormatter/config/nd30_rules.json) ngày 05/3/2020 của Chính phủ về công tác văn thư (áp dụng cho các cơ quan hành chính nhà nước, đơn vị sự nghiệp công lập, doanh nghiệp nhà nước).
2. **Khối Cơ quan Đảng:** Tuân thủ nghiêm ngặt [Hướng dẫn số 36-HD/VPTW](file:///d:/AutoFormatter/config/hd36_rules.json) ngày 03/4/2018 của Văn phòng Trung ương Đảng về thể thức và kỹ thuật trình bày văn bản của Đảng (áp dụng cho các cơ quan, tổ chức Đảng từ Trung ương đến cấp Xã/Phường, Chi bộ).

---

## 2. YÊU CẦU HỆ THỐNG MÁY TÍNH (SYSTEM REQUIREMENTS)

Ứng dụng được thiết kế tối ưu hóa dung lượng và hiệu năng, tương thích với hầu hết máy tính văn phòng tại các cơ quan, đơn vị, địa phương:

| Thành phần | Yêu cầu Tối thiểu | Khuyến nghị Tối ưu | Ghi chú |
| :--- | :--- | :--- | :--- |
| **Hệ điều hành** | Windows 10 (64-bit) | Windows 10 / 11 (64-bit) hoặc Windows Server 2016+ | Tương thích Windows 7 SP1 (64-bit) nếu đã cài VC++ Redistributable. |
| **Bộ xử lý (CPU)** | Intel Core i3 / AMD Ryzen 3 trở lên (x86_64) | Intel Core i5 / AMD Ryzen 5 hoặc tương đương | Tận dụng tập lệnh AVX/AVX2 để chạy mô hình AI OCR siêu tốc. |
| **Bộ nhớ (RAM)** | 2 GB RAM | 4 GB – 8 GB RAM | 2 GB đủ dùng cho soạn thảo; 4 GB+ tối ưu khi xử lý file scan nhiều trang. |
| **Dung lượng đĩa** | 500 MB dung lượng trống | 1 GB SSD trống | Chứa bộ ứng dụng, mô hình AI OCR ngoại tuyến và lưu tệp tạm. |
| **Màn hình** | Độ phân giải 1366 x 768 | Full HD (1920 x 1080) trở lên | Hỗ trợ High-DPI Scaling tự động (100%, 125%, 150%). |
| **Phần mềm Office** | **Không bắt buộc** | Microsoft Word 2013 / 2016 / 2019 / 2021 / Microsoft 365 | Đã tích hợp sẵn bộ máy OpenXML & PDF độc lập; tự kích hoạt Word COM nếu máy có sẵn Office. |
| **Thiết bị ký số** | Không bắt buộc | USB Token Ban Cơ yếu Chính phủ hoặc Token CA công cộng | Driver Token đã cài đặt trên Windows để thực hiện ký số. |
| **Kết nối mạng** | **Không cần Internet (Offline 100%)** | **Offline 100% (Air-gapped)** | Tuyệt đối an toàn đối với tài liệu mật và tài liệu nội bộ. |

---

## 3. TẢI VỀ LÀ SỬ DỤNG ĐƯỢC NGAY (ZERO-INSTALL / PORTABLE)

> [!TIP]
> Ứng dụng đã được đóng gói nhị phân độc lập hoàn chỉnh theo cấu trúc **Onedir Portable**. Cán bộ, chuyên viên **không cần cài đặt Python, không cần cài đặt thư viện phụ trợ, không cần quyền quản trị viên Administrator** trên máy tính cơ quan!

### Cách 1: Sử dụng Bản nén Di động (Khuyến nghị cho USB)
1. Tải tệp nén **`AutoFormatter_v2.0.zip`** về máy tính.
2. Nhấp chuột phải vào tệp zip chọn **Extract All... (Giải nén)** vào một thư mục bất kỳ (ví dụ: `D:\AutoFormatter` hoặc trực tiếp trong USB).
3. Mở thư mục vừa giải nén và nhấp đúp vào tệp **`AutoFormatter.exe`** để mở ứng dụng ngay trong vòng **< 1 giây**.

### Cách 2: Sử dụng Bộ cài đặt Tự động Windows
1. Tải và chạy tệp **`AutoFormatter_v2.0_Setup.exe`**.
2. Nhấn **Next -> Next -> Finish** để chương trình tự động tạo lối tắt (Shortcut) có biểu tượng trên Màn hình nền (Desktop) và Menu Start.

---

## 4. HƯỚNG DẪN SỬ DỤNG CÁC TÍNH NĂNG CHÍNH

### 4.1. Chuẩn hóa Thể thức Hàng loạt (Batch Formatting)
Dành cho tình huống tiếp nhận các tệp văn bản Word hoặc PDF scan cần rà soát và định dạng lại nhanh chóng:

1. **Nạp tệp tài liệu:**
   - **Kéo - thả** một hoặc nhiều tệp `.docx`, `.pdf` vào vùng thả giữa màn hình.
   - Hoặc nhấn nút **Thêm tệp...** / **Thêm thư mục...** để chọn tệp từ máy tính.
2. **Chọn Quy chuẩn Thể thức:**
   - Chọn **Hành chính (NĐ 30/2020)** cho văn bản Ủy ban nhân dân, Hội đồng nhân dân, Sở, Ban, Ngành...
   - Chọn **Cơ quan Đảng (HD 36-HD/VPTW)** cho văn bản Tỉnh ủy, Huyện ủy, Đảng ủy xã, Chi bộ...
3. **Chọn Hồ sơ Cơ quan:**
   - Chọn hồ sơ phù hợp trong danh mục (ví dụ: *Ủy ban nhân dân xã / thị trấn*, *Hội đồng nhân dân xã*, *Đảng ủy xã*...) hoặc bấm **Quản lý hồ sơ cơ quan...** để tạo mới.
4. **Bắt đầu xử lý:**
   - Nhấn **▶ BẮT ĐẦU CHUẨN HÓA THỂ THỨC**.
   - Quan sát bảng 4 thẻ lọc lỗi tự động: Số đoạn đã căn lề, khoảng trắng thừa đã dọn, dấu câu đã sửa, cảnh báo viết hoa.
   - Nhấn **Mở file Word** để xem tệp kết quả (`_ChuanHoa.docx`) hoặc **Mở PDF** để in ấn.

---

### 4.2. Trình Soạn thảo Hành chính WYSIWYG Nhúng Toàn diện (Word Editor)
Phiên bản 2.0 tích hợp trực tiếp một Trình Soạn thảo Hành chính hoàn chỉnh bên trong ứng dụng:

1. **Khởi chạy trình soạn thảo:**
   - Bấm nút **Soạn thảo tích hợp** trên thanh công cụ sau khi chọn tệp, hoặc nhấn **Ctrl+N** để tạo văn bản mới.
   - Nút **Undo / Redo** nhanh đặt ngay trên thanh công cụ; cho phép **đổi tên tệp trực tiếp** bằng cách nhấp chuột vào tên tệp trên thanh tiêu đề.
2. **Khung vẽ A4 rời rạc (Discrete Multi-page Canvas):**
   - Không bị cuộn dài vô tận: Văn bản được kết xuất thành các tờ giấy trắng A4 rời rạc (210 x 297 mm) với bóng đổ thực tế và ranh giới lề in chuẩn xác.
   - Nhấn **Ctrl + Enter** tại bất kỳ vị trí nào để ngắt sang trang A4 kế tiếp chuẩn xác.
3. **Độc lập định hướng khổ giấy (Independent Orientation):**
   - **Trang 1:** Luôn khóa cố định khổ dọc (Portrait) theo quy chuẩn NĐ 30 & HD 36 để đảm bảo tính trang nghiêm của Quốc hiệu và Cơ quan ban hành.
   - **Từ trang 2 trở đi:** Tùy chọn độc lập xoay ngang (Landscape 297 x 210 mm) dành riêng cho các trang biểu mẫu, phụ lục, danh sách hoặc bảng biểu nhiều cột mà **hoàn toàn không làm ảnh hưởng đến khổ dọc của trang 1**.
4. **Bảo toàn 100% Hình ảnh & Đồ họa (Zero Image Loss):**
   - Mọi hình ảnh, biểu đồ, sơ đồ trong tệp Word được bảo toàn nguyên vẹn, tự động căn giữa trang, triệt tiêu thụt lề đầu dòng, tự co giãn theo vùng in khổ A4 (không bị tràn mép giấy).
   - Tự động nhận diện và định dạng các dòng chú thích hình ảnh (11pt in nghiêng, căn giữa).
5. **Dải băng điều khiển (Administrative Ribbon):**
   - Tab **Soạn thảo:** Phông chữ chuẩn Times New Roman Unicode, cỡ chữ 13–14 pt, căn đều hai lề (Justify), giãn dòng 1.2 lines, giãn đoạn 6 pt, thụt lề đầu dòng chuẩn 1.27 cm (1 Tab). Đánh số tự động thông minh (tiếp nối số thứ tự, đánh số phân cấp La Mã, chữ cái, thập phân).
   - Tab **Bố cục & Lề in:** Thiết lập lề trái 30 mm (đóng gáy), lề phải 15 mm, trên/dưới 20 mm; hiển thị thước đo milimét (Horizontal Ruler).
   - Tab **Bảng biểu:** Quản lý hàng/cột, gộp ô, tách ô, tự động căn giữa bảng, bảo vệ bảng không bị ngắt đôi qua trang (`cantSplit`) và lặp lại dòng tiêu đề trên từng trang (`tblHeader`).
   - Nút **Sửa thông tin thể thức:** Mở hộp thoại siêu dữ liệu để cập nhật nhanh Cơ quan ban hành, Quốc hiệu, Số hiệu, Trích yếu, Thẩm quyền ký và Nơi nhận.

---

### 4.3. Rà soát Tuân thủ Pháp lý & Sửa nhanh 1 chạm (Compliance Audit)
Đánh giá độ tuân thủ thể thức của văn bản đối chiếu với Nghị định 30 hoặc Hướng dẫn 36:

1. Trong Trình Soạn thảo, nhấn nút **Rà soát thể thức** (hoặc phím tắt **F7**).
2. Hộp thoại **Kết quả Rà soát Thể thức Pháp lý** xuất hiện với giao diện hiện đại:
   - **Thang điểm số nguyên:** Chấm điểm mức độ hoàn thiện của văn bản từ 0 đến 100 điểm.
   - **Bảng Checklist quy chuẩn:** Rà soát tự động 6 tiêu chuẩn cốt lõi:
     - Khổ giấy A4 & Lề in pháp lý.
     - Phông chữ chuẩn Times New Roman Unicode.
     - Quốc hiệu & Tiêu ngữ (hoặc Tiêu đề Đảng) đầy đủ đường gạch ngang.
     - Tiền tố thẩm quyền người ký (`TM.`, `KT.` đối với Nhà nước; `T/M`, `K/T` đối với Đảng).
     - Cú pháp dấu câu căn cứ pháp lý (chấm phẩy `;` ở các căn cứ, dấu phẩy `,` hoặc chấm `.` ở căn cứ cuối).
     - Phát hiện các ký tự giữ chỗ chưa hoàn thiện (`...`, `___`).
3. **Tính năng Sửa nhanh 1 chạm:**
   - Nhấn nút **[Sửa nhanh]** bên cạnh từng điều khoản, hoặc nhấn **[Tự động sửa nhanh toàn bộ]** ở góc dưới. Hệ thống sẽ tự động hiệu chỉnh toàn bộ lề in, chuyển đổi phông chữ và chuẩn hóa dấu câu ngay lập tức để đạt điểm tối đa **100 / 100 điểm**.

---

### 4.4. Ký số Điện tử Trực quan (Digital Signature NĐ 30)
Ký số điện tử bằng USB Token Ban Cơ yếu Chính phủ hoặc chứng thư số cơ quan/cá nhân:

1. Nhấn nút **Ký số...** (hoặc phím tắt **Ctrl+K**).
2. Cửa sổ **Ký số Điện tử Nghị định 30** hiện ra kèm khung xem trước PDF thời gian thực.
3. **Định vị khối ký thông minh (Smart Signature Locator):**
   - Hệ thống tự động quét trang cuối của văn bản chính (tự động bỏ qua các trang phụ lục biểu mẫu đính kèm).
   - Khung ký được tự động đặt ngay ngắn tại vị trí giữa chức danh và họ tên người ký.
   - Cán bộ có thể **nhấp chuột vào bất kỳ vị trí nào** trên trang hoặc **kéo thả chuột** để vi chỉnh từng milimét.
4. **Thiết lập chữ ký:**
   - Chọn kiểu ký: *Chữ ký cơ quan (Dấu đỏ Điều 13 NĐ 30 trùm 1/3 chữ ký)* hoặc *Chữ ký cá nhân (Tên và chức danh màu xanh)*.
5. Cắm USB Token vào máy tính, chọn chứng thư số tương ứng trong danh sách và bấm **THỰC HIỆN KÝ SỐ VĂN BẢN**.
6. Nhập mã PIN của USB Token trên hộp thoại bảo mật của hãng. Văn bản PDF xuất ra sẽ có **dấu tích xanh hợp lệ (Signature Valid)** trên Adobe Acrobat Reader và các hệ thống Quản lý văn bản điều hành.

---

### 4.5. Nhận dạng Chữ OCR Ngoại tuyến từ Ảnh Scan / PDF Scan
Xử lý các tài liệu cũ, văn bản chụp bằng điện thoại hoặc tệp scan không có lớp chữ:

1. Thêm tệp ảnh hoặc tệp `.pdf` scan vào ứng dụng.
2. Nhấn nút **Nhận dạng OCR**.
3. Mô hình AI RapidOCR (ONNX Runtime) chạy trực tiếp trên CPU máy tính:
   - Tự động nắn góc nghiêng tài liệu (Deskewing) nếu tài liệu bị chụp xéo.
   - Tự động tách con dấu đỏ (HSV Filter) để bóc tách rõ nét lớp chữ bên dưới.
   - Tái tạo lại tài liệu thành các đoạn văn Word (`.docx`) hoàn chỉnh và tự động chuyển giao vào luồng chuẩn hóa thể thức.

---

### 4.6. Khởi tạo Nhanh Văn bản Mẫu (Template Wizard)
Dành cho cán bộ cần soạn thảo văn bản mới từ đầu theo đúng thẩm quyền:

1. Bấm nút **Soạn thảo mới** trên giao diện chính.
2. Chọn loại văn bản: *Quyết định cá biệt, Quyết định quy phạm, Nghị quyết, Báo cáo, Kế hoạch, Tờ trình, Công văn hành chính...*
3. Nhập trích yếu và nơi nhận. Hệ thống sẽ tự động nạp cấu trúc chuẩn mực với đầy đủ Quốc hiệu, Tiêu ngữ, số ký hiệu mẫu, căn cứ pháp lý và khối chữ ký hợp chuẩn.

---

## 5. DÀNH CHO LẬP TRÌNH VIÊN & QUẢN TRỊ VIÊN HỆ THỐNG

### 5.1. Cài đặt Môi trường Phát triển (Source Code)
Nếu muốn đóng góp mã nguồn hoặc tùy biến hệ thống:
```bash
# 1. Yêu cầu Python 3.10 hoặc 3.11 (64-bit) trên Windows
git clone https://github.com/your-org/AutoFormatter.git
cd AutoFormatter

# 2. Cài đặt các thư viện phụ thuộc
pip install -r requirements.txt

# 3. Chạy ứng dụng trực tiếp từ mã nguồn
python main.py
```

### 5.2. Chạy Bộ Kiểm thử Tự động (Automated Test Suite)
Chạy toàn bộ 68 kịch bản kiểm thử tự động để bảo đảm tính toàn vẹn của hệ thống:
```bash
python -m unittest discover tests
```

### 5.3. Chế độ Dòng lệnh (CLI) Cho Kịch bản Tự động hóa
Chuẩn hóa nhanh tài liệu hoặc tích hợp vào các script tự động hóa mà không cần mở giao diện đồ họa:
```bash
# Chuẩn hóa tệp Word theo Nghị định 30 và tự động xuất thêm tệp PDF
python main.py --cli -f "du_thao_to_trinh.docx" -s ND30 --pdf

# Chuẩn hóa tệp theo Hướng dẫn 36 (Khối Đảng)
python main.py --cli -f "du_thao_nghi_quyet.docx" -s HD36

# Tự động nhận dạng OCR từ tệp PDF scan rồi chuẩn hóa
python main.py --cli -f "van_ban_scan.pdf" -s ND30 --ocr
```

### 5.4. Quy trình Biên dịch Đóng gói Bản Thực thi (.exe)
Ứng dụng sử dụng cấu hình tối ưu sẵn tại [`build_exe.spec`](file:///d:/AutoFormatter/build_exe.spec) để đóng gói:
```bash
# Biên dịch gói Onedir khởi động tức thì (< 1s)
pyinstaller build_exe.spec --noconfirm
```
Sản phẩm đóng gói hoàn tất sẽ nằm tại thư mục `dist/AutoFormatter/`. Để tạo bộ cài đặt setup `.exe`, mở tệp [`installer.iss`](file:///d:/AutoFormatter/installer.iss) bằng phần mềm **Inno Setup Compiler** và nhấn **Compile (F9)**.

---

## 6. DANH MỤC CẤU TRÚC THƯ MỤC MÃ NGUỒN

```
AutoFormatter/
├── main.py                         # Điểm khởi chạy chính (Desktop GUI & CLI)
├── requirements.txt                # Danh sách thư viện Python cần thiết
├── build_exe.spec                  # Cấu hình biên dịch đóng gói PyInstaller Onedir
├── installer.iss                   # Kịch bản Inno Setup tạo bộ cài đặt Windows Setup
├── config/                         # Cấu hình quy chuẩn & từ điển pháp lý
│   ├── nd30_rules.json             # Quy chuẩn Nghị định 30/2020/NĐ-CP (Nhà nước)
│   ├── hd36_rules.json             # Quy chuẩn Hướng dẫn 36-HD/VPTW (Khối Đảng)
│   ├── agency_profiles.json        # Danh mục hồ sơ cơ quan (UBND, HĐND, Đảng ủy xã...)
│   ├── capitalization_dict.txt     # Từ điển danh từ riêng & tổ chức nhà nước
│   └── capitalization_rules.json   # Quy tắc kiểm tra viết hoa theo Phụ lục II NĐ 30
├── core/                           # Tầng logic nghiệp vụ cốt lõi (Core Engine)
│   ├── formatter_engine.py         # Bộ máy chuẩn hóa khổ giấy, lề, font, bảng biểu
│   ├── doc_parser.py               # Phân tích cú pháp AST, tách cột, bóc tách cấu trúc
│   ├── header_footer_builder.py    # Dựng bảng ẩn 2 cột cho Header và Footer
│   ├── administrative_sanitizer.py # Dọn dẹp khoảng trắng, chuẩn hóa dấu câu & ngày tháng
│   ├── legal_constraint_guard.py   # Giám sát tuân thủ thể thức & chấm điểm 0-100
│   ├── signature_engine.py         # Bộ máy ký số CAPI USB Token & định vị Điều 13 NĐ 30
│   ├── ocr_engine.py               # Bộ máy OCR RapidOCR ngoại tuyến & Deskew ảnh scan
│   ├── spell_linter.py             # Bộ rà soát lỗi chính tả & quy tắc viết hoa
│   ├── win32_helper.py             # Điều khiển MS Word COM Automation & xuất PDF
│   └── editor_core/                # Động cơ xử lý OpenXML nội bộ cho trình soạn thảo
│       ├── openxml_builder.py      # Sinh file DOCX từ siêu dữ liệu (Content Control Lock)
│       └── table_layout_manager.py # Quản trị bố cục bảng ẩn viền và bảo vệ ngắt trang
├── ui/                             # Tầng giao diện người dùng Desktop (PyQt6)
│   ├── main_window.py              # Cửa sổ chính điều khiển tác vụ chuẩn hóa hàng loạt
│   ├── signature_dialog.py         # Hộp thoại ký số trực quan tương tác trên trang PDF
│   ├── profile_dialog.py           # Hộp thoại quản lý danh mục hồ sơ cơ quan ban hành
│   ├── quick_wizard_dialog.py      # Hộp thoại tạo nhanh văn bản mẫu theo loại hình
│   ├── styles.py                   # Bảng định dạng QSS hiện đại chuẩn công vụ
│   └── editor/                     # Phân hệ Trình Soạn thảo Hành chính WYSIWYG
│       ├── word_editor_window.py   # Cửa sổ soạn thảo văn bản tích hợp toàn diện
│       ├── document_canvas.py      # Khung vẽ A4 rời rạc, hỗ trợ độc lập xoay ngang
│       ├── administrative_ribbon.py# Dải băng điều khiển Ribbon Bar chuẩn công vụ
│       ├── compliance_audit_dlg.py # Hộp thoại rà soát thể thức & nút bấm sửa nhanh 1 chạm
│       ├── metadata_editor_dlg.py  # Hộp thoại chỉnh sửa siêu dữ liệu văn bản 2 chiều
│       └── document_ruler.py       # Thước căn lề milimét tương tác
├── templates/                      # Tệp Word mẫu chuẩn cho Đảng và Nhà nước
├── assets/                         # Biểu tượng và tài nguyên đồ họa High-DPI
├── tests/                          # Bộ 68 kịch bản kiểm thử tự động
└── dist/                           # Sản phẩm xuất xưởng sẵn sàng phân phối
    ├── AutoFormatter/              # Thư mục ứng dụng Portable Onedir (< 1s mở)
    └── AutoFormatter_v2.0.zip      # Gói nén di động phân phối (~125 MB)
```

---

## 7. CHÍNH SÁCH BẢO MẬT & ĐỘC LẬP TÁC CHIẾN

- **Bảo mật tuyệt đối (100% Offline):** Ứng dụng không thực hiện bất kỳ kết nối mạng ra bên ngoài, không sử dụng dịch vụ đám mây, không thu thập nhật ký (telemetry) của người dùng.
- **An toàn khóa bí mật (Private Key Security):** Khóa riêng của chứng thư số luôn nằm an toàn trong phần cứng USB Token chuẩn FIPS 140-2 Level 2/3. Toàn bộ quá trình ký số được ủy thác trực tiếp cho bộ điều khiển mật mã của hệ điều hành Windows CryptoAPI, bảo đảm tính pháp lý cao nhất theo Luật Giao dịch điện tử.
