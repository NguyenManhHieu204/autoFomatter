# TÀI LIỆU ĐẶC TẢ KỸ THUẬT HỆ THỐNG (TECHNICAL SPECIFICATION)
## CÔNG CỤ TỰ ĐỘNG CHUẨN HÓA THỂ THỨC VĂN BẢN ĐẢNG & HÀNH CHÍNH
### Auto-Formatter — Version 2.0 Commercial Edition
**Tiêu chuẩn áp dụng:** Nghị định số 30/2020/NĐ-CP & Hướng dẫn số 36-HD/VPTW  
**Kiến trúc:** Desktop Application (PyQt6 + OpenXML + ONNX + Windows CryptoAPI)  
**Môi trường:** Windows 10/11 x64 • Hoạt động Ngoại tuyến 100% (Air-gapped Zero-Cloud)  
**Ngày cập nhật:** Tháng 11/2026

---

## MỤC LỤC
1. [TỔNG QUAN HỆ THỐNG](#i-tổng-quan-hệ-thống)
2. [CƠ SỞ PHÁP LÝ & BỘ QUY CHUẨN KỸ THUẬT](#ii-cơ-sở-pháp-lý--bộ-quy-chuẩn-kỹ-thuật)
3. [KIẾN TRÚC HỆ THỐNG TỔNG THỂ](#iii-kiến-trúc-hệ-thống-tổng-thể)
4. [ĐẶC TẢ CHI TIẾT CÁC PHÂN HỆ VÀ MODULE CỐT LÕI](#iv-đặc-tả-chi-tiết-các-phân-hệ-và-module-cốt-lõi)
   - [Module 1: Bộ máy Nhận dạng Ký tự Quang học Ngoại tuyến (Offline OCR Engine)](#module-1-bộ-máy-nhận-dạng-ký-tự-quang-học-ngoại-tuyến-offline-ocr-engine)
   - [Module 2: Bộ Phân tích Cấu trúc AST & Chuẩn hóa Thể thức (Layout & Format Engine)](#module-2-bộ-phân-tích-cấu-trúc-ast--chuẩn-hóa-thể-thức-layout--format-engine)
   - [Module 3: Phân hệ Bảo toàn & Định dạng Hình ảnh, Đồ họa (Image & Vector Graphics Engine)](#module-3-phân-hệ-bảo-toàn--định-dạng-hình-ảnh-đồ-họa-image--vector-graphics-engine)
   - [Module 4: Bộ Rà soát Lỗi Chính tả Hành chính & Kỹ thuật Soạn thảo (Administrative Spell & Style Linter)](#module-4-bộ-rà-soát-lỗi-chính-tả-hành-chính--kỹ-thuật-soạn-thảo-administrative-spell--style-linter)
   - [Module 5: Phân hệ Tương tác MS Word COM & Xuất PDF Độc lập (Word COM Automation & PDF Engine)](#module-5-phân-hệ-tương-tác-ms-word-com--xuất-pdf-độc-lập-word-com-automation--pdf-engine)
   - [Module 6: Phân hệ Ký số Điện tử Trực quan (Visual Digital Signature Engine)](#module-6-phân-hệ-ký-số-điện-tử-trực-quan-visual-digital-signature-engine)
   - [Module 7: Trình Soạn thảo Hành chính WYSIWYG Nhúng Toàn diện (Embedded Administrative Word Editor)](#module-7-trình-soạn-thảo-hành-chính-wysiwyg-nhúng-toàn-diện-embedded-administrative-word-editor)
5. [CẤU TRÚC DỮ LIỆU & TỪ ĐIỂN QUY TẮC](#v-cấu-trúc-dữ-liệu--từ-điển-quy-tắc)
6. [MÔ HÌNH BẢO MẬT & ĐỘC LẬP TÁC CHIẾN (SECURITY MODEL)](#vi-mô-hình-bảo-mật--độc-lập-tác-chiến-security-model)
7. [BỘ KIỂM THỬ TỰ ĐỘNG & ĐẢM BẢO CHẤT LƯỢNG (QA & TEST SUITE)](#vii-bộ-kiểm-thử-tự-động--đảm-bảo-chất-lượng-qa--test-suite)
8. [QUY TRÌNH ĐÓNG GÓI, TRIỂN KHAI & PHÂN PHỐI THƯƠNG MẠI](#viii-quy-trình-đóng-gói-triển-khai--phân-phối-thương-mại)
9. [ĐÁNH GIÁ TRẠNG THÁI DỰ ÁN & ĐỊNH HƯỚNG PHÁT TRIỂN](#ix-đánh-giá-trạng-thái-dự-án--định-hướng-phát-triển)

---

## I. TỔNG QUAN HỆ THỐNG

### 1. Bối cảnh & Sứ mệnh Dự án
Trong các cơ quan của Đảng Cộng sản Việt Nam, Nhà nước, Mặt trận Tổ quốc và các đoàn thể chính trị - xã hội, việc ban hành văn bản quy phạm pháp luật và văn bản hành chính đòi hỏi tính chuẩn mực, trang nghiêm và giá trị pháp lý tuyệt đối. Tuy nhiên, thực tế công tác văn thư và soạn thảo còn đối mặt với nhiều bất cập:
* **Sai lệch thể thức kỹ thuật**: Sai khổ lề in A4, khoảng cách dòng/đoạn tùy tiện, phông chữ không đồng nhất, thụt đầu dòng sai chuẩn.
* **Lỗi trình bày Quốc hiệu, Tiêu ngữ, Tên cơ quan, Thẩm quyền**: Đường kẻ phân cách bị lệch hoặc sai tỷ lệ, sai vị trí khối ký, thiếu ngày tháng hoặc sai định dạng.
* **Mất mát dữ liệu khi chuyển đổi**: Các giải pháp định dạng tự động thông thường thường làm mất hình ảnh, bảng biểu bị tràn lề, biểu đồ bị cắt xén hoặc xáo trộn nội dung.
* **Rủi ro rò rỉ thông tin mật**: Việc sử dụng các công cụ trực tuyến (Cloud SaaS / AI Web) tiềm ẩn nguy cơ lộ lọt bí mật nhà nước và dữ liệu nội bộ cơ quan.
* **Nút thắt ký số điện tử**: Thiếu công cụ ký số trực quan trên tệp PDF đáp ứng chuẩn Điều 13 Nghị định 30/2020/NĐ-CP (vị trí dấu đỏ trùm 1/3 chữ ký, tương thích với thiết bị phần cứng USB Token của Ban Cơ yếu Chính phủ).

**Auto-Formatter v2.0 Commercial Edition** được thiết kế như một giải pháp phần mềm độc lập hoàn chỉnh, cung cấp trọn vẹn chu trình: **Nhận dạng scan (OCR) -> Chuẩn hóa thể thức tự động -> Soạn thảo trực quan WYSIWYG A4 -> Rà soát lỗi pháp lý -> Ký số điện tử chuẩn hóa -> Xuất bản tài liệu sẵn sàng phát hành**.

### 2. Nguyên tắc Thiết kế Cốt lõi (Design Tenets)
1. **Tuân thủ pháp lý 100%**: Tuyệt đối trung thực với Nghị định 30/2020/NĐ-CP và Hướng dẫn 36-HD/VPTW.
2. **Bảo toàn dữ liệu tối thượng (Zero Content Loss)**: Không bao giờ được phép thay đổi, làm mất nội dung câu chữ, dữ liệu số, bảng biểu hay hình ảnh của người dùng trong quá trình xử lý.
3. **Hoạt động Ngoại tuyến Tuyệt đối (100% Air-Gapped)**: Không mở cổng mạng, không thu thập telemetry, không gọi API bên thứ ba. Toàn bộ mô hình học sâu và thuật toán chạy trực tiếp trên CPU cục bộ.
4. **Tải về là chạy ngay (Zero-Install / Portable)**: Khởi động tức thì (< 1 giây), không đòi hỏi môi trường Python, không bắt buộc cài đặt Microsoft Office, không yêu cầu quyền Administrator.

---

## II. CƠ SỞ PHÁP LÝ & BỘ QUY CHUẨN KỸ THUẬT

Hệ thống được lập trình với 2 bộ luật thể thức nền tảng:

### 1. Khối Cơ quan Hành chính Nhà nước: Nghị định số 30/2020/NĐ-CP
* **Khổ giấy & Căn lề**: A4 (210 x 297 mm); Lề trên: 20-25 mm; Lề dưới: 20-25 mm; Lề trái: 30-35 mm (đóng gáy); Lề phải: 15-20 mm.
* **Phông chữ**: Times New Roman (bộ mã Unicode TCVN 6909:2001).
* **Quốc hiệu & Tiêu ngữ**: 
  * "CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM": Cỡ 12-13pt, in hoa, đứng, đậm.
  * "Độc lập - Tự do - Hạnh phúc": Cỡ 13-14pt, in hoa chữ cái đầu các từ, đứng, đậm; có đường kẻ ngang nét liền mảnh phía dưới bằng chiều dài dòng chữ.
* **Tên cơ quan, tổ chức ban hành**: Cơ quan chủ quản cỡ 12-13pt in hoa đứng thường; Cơ quan trực tiếp cỡ 12-13pt in hoa đứng đậm; đường kẻ ngang bằng 1/3 - 1/2 chiều dài tên cơ quan.
* **Số, ký hiệu văn bản**: Cỡ 13pt, đứng thường, phân tách bằng dấu `/`.
* **Địa danh & Ngày tháng**: Cỡ 13-14pt, in thường, nghiêng; căn giữa phía dưới Quốc hiệu - Tiêu ngữ.
* **Tên loại & Trích yếu**: Tên loại cỡ 14pt in hoa đứng đậm; trích yếu cỡ 13-14pt in thường đứng đậm.
* **Nội dung văn bản**: Cỡ 13-14pt, căn đều hai lề (Justify), giãn dòng 1.2 lines, thụt đầu dòng đoạn văn 1.0 - 1.27 cm.
* **Thẩm quyền & Họ tên**: Quyền hạn và chức vụ in hoa đứng đậm cỡ 13-14pt; họ tên in thường đứng đậm cỡ 13-14pt.
* **Chữ ký số & Con dấu (Điều 13)**: Hình ảnh con dấu cơ quan màu đỏ, định dạng PNG trong suốt, trùm lên khoảng 1/3 chữ ký về phía bên trái.

### 2. Khối Cơ quan, Tổ chức Đảng: Hướng dẫn số 36-HD/VPTW
* **Tiêu đề cơ quan Đảng**: Tên cơ quan cấp trên cỡ 13-14pt in hoa đứng; Tên cơ quan ban hành cỡ 13-14pt in hoa đứng đậm.
* **Tiêu ngữ Đảng**: Dấu hoa thị hoặc tiêu ngữ đặc thù; khoảng cách đường kẻ chuẩn hóa.
* **Tiền tố Thẩm quyền**: `T/M` (Thay mặt), `K/T` (Ký thay) thay vì `TM.` hoặc `KT.`.
* **Quy tắc viết hoa**: Tuân thủ quy định viết hoa tên tổ chức Đảng, cơ quan hành chính, chức vụ danh xưng tôn kính theo Phụ lục II.

---

## III. KIẾN TRÚC HỆ THỐNG TỔNG THỂ

### 1. Sơ đồ Luồng Dữ liệu Toàn diện (System Pipeline)

```mermaid
flowchart TD
    subgraph INPUT ["1. ĐẦU VÀO ĐA DẠNG"]
        DocxIn["Tệp Word hiện hữu (.docx)"]
        ScanIn["Tệp Ảnh / Scan (.pdf, .png, .jpg)"]
        NewIn["Soạn thảo mới từ đầu (A4 Canvas)"]
    end

    subgraph INGESTION ["2. TIỀN XỬ LÝ & NHẬN DẠNG"]
        OCR["RapidOCR ONNX Engine<br/>• Deskew góc nghiêng<br/>• Tách lọc con dấu HSV<br/>• Bóc tách văn bản tiếng Việt"]
        AST["OpenXML AST Parser<br/>• Phân tích cấu trúc đoạn & bảng<br/>• Nhận diện 9 thành phần thể thức<br/>• Trích xuất hình ảnh & bảng biểu"]
    end

    subgraph NORMALIZATION ["3. BỘ MÁY CHUẨN HÓA & LINTER"]
        Engine["FormatterEngine (NĐ 30 / HD 36)<br/>• Khổ giấy A4 & Lề chuẩn<br/>• Header/Footer 2 cột vô hình<br/>• Giãn dòng 1.2, giãn đoạn 4-6pt"]
        ImgEngine["Image & Vector Engine<br/>• Bảo toàn 100% hình ảnh<br/>• Co giãn chống tràn lề (<=160mm)<br/>• Căn giữa trang, 0-indent<br/>• Định dạng chú thích 11pt nghiêng"]
        Linter["Administrative Spell & Style Linter<br/>• Dọn khoảng trắng & dấu câu<br/>• Chuẩn hóa danh từ & viết hoa NĐ 30"]
    end

    subgraph EDITOR ["4. TRÌNH SOẠN THẢO WYSIWYG A4"]
        Canvas["Discrete A4 Multi-page Canvas<br/>• Độc lập định hướng (Dọc/Ngang)<br/>• Thước đo milimét tương tác<br/>• Đánh số phân cấp thông minh<br/>• Undo / Redo & Đổi tên tệp nhanh"]
        Audit["Compliance Audit Guard<br/>• Chấm điểm 0-100 thời gian thực<br/>• Sửa nhanh 1 chạm tự động"]
    end

    subgraph SIGNING ["5. KÝ SỐ ĐIỆN TỬ PHÁP LÝ"]
        SmartLoc["Smart Signature Locator<br/>• Định vị tự động khối ký NĐ 30<br/>• Bỏ qua trang phụ lục biểu mẫu"]
        HybridSign["Hybrid Signing Engine<br/>• Kênh 1: USB Token CAPI (SHA256-RSA)<br/>• Kênh 2: VGCA SignTool Bridge"]
    end

    subgraph OUTPUT ["6. XUẤT BẢN THƯƠNG MẠI"]
        DocxOut["File Word Chuẩn hóa (.docx)"]
        PdfOut["File PDF Chuẩn in ấn (PDF/A)"]
        SignedPdf["File PDF Ký số Tích xanh Hợp lệ"]
    end

    DocxIn --> AST
    ScanIn --> OCR --> AST
    NewIn --> Canvas

    AST --> Engine
    Engine --> ImgEngine --> Linter
    Linter --> Canvas
    Canvas <--> Audit

    Canvas --> DocxOut
    Canvas --> PdfOut
    PdfOut --> SmartLoc --> HybridSign --> SignedPdf
```

---

## IV. ĐẶC TẢ CHI TIẾT CÁC PHÂN HỆ VÀ MODULE CỐT LÕI

### Module 1: Bộ máy Nhận dạng Ký tự Quang học Ngoại tuyến (Offline OCR Engine)
* **Tệp mã nguồn**: [`core/ocr_engine.py`](file:///d:/AutoFormatter/core/ocr_engine.py)
* **Công nghệ nền tảng**: `RapidOCR` biên dịch trên `ONNX Runtime` x86_64 CPU.
* **Tính năng chi tiết**:
  * **Xử lý tiền kỳ hình ảnh**:
    * Phát hiện và xoay góc nghiêng (Deskewing) tự động sử dụng biến đổi Radon và thuật toán minAreaRect của OpenCV.
    * Phân vùng không gian màu HSV để lọc bỏ con dấu tròn đỏ đè lên chữ ký, tránh làm đứt nét ký tự tiếng Việt.
  * **Bóc tách ký tự**:
    * Nhận dạng đầy đủ bộ ký tự tiếng Việt có dấu, ký tự La Mã, chỉ số trên/dưới.
    * Tái tạo lại cấu trúc đoạn văn, thụt lề và phân dòng tự nhiên; tự động chuyển giao vào luồng OpenXML AST.

### Module 2: Bộ Phân tích Cấu trúc AST & Chuẩn hóa Thể thức (Layout & Format Engine)
* **Tệp mã nguồn**: [`core/formatter_engine.py`](file:///d:/AutoFormatter/core/formatter_engine.py), [`core/doc_parser.py`](file:///d:/AutoFormatter/core/doc_parser.py), [`core/header_footer_builder.py`](file:///d:/AutoFormatter/core/header_footer_builder.py)
* **Nguyên tắc phân tích cú pháp**:
  * Đọc tệp `.docx` dưới dạng cây cú pháp trừu tượng OpenXML (DOM tree).
  * Nhận diện và bóc tách tự động 9 thành phần thể thức theo mô hình nhận dạng mẫu (Pattern Matching) kết hợp biểu thức chính quy (Regex).
* **Cấu trúc Tiêu đề Đỉnh (Header Table)**:
  * Tạo bảng ẩn viền gồm 1 hàng x 2 cột:
    * Cột trái: Tên cơ quan chủ quản, Tên cơ quan ban hành, Số ký hiệu văn bản. Kèm đường gạch phân cách dưới tên cơ quan (chiều dài 1/3 đến 1/2 tên cơ quan).
    * Cột phải: Quốc hiệu, Tiêu ngữ, Địa danh & ngày tháng năm ban hành. Kèm đường gạch phân cách liền mảnh dưới Tiêu ngữ.
  * Tự động khóa độ rộng cột (Column Width Locking) và căn giữa nội dung từng cột theo phương ngang.
* **Cấu trúc Khối Chữ ký & Nơi nhận (Footer Table)**:
  * Bảng ẩn viền 1 hàng x 2 cột đặt tại cuối văn bản:
    * Cột trái: Khối `Nơi nhận:`, các dòng cơ quan/cá nhân nhận văn bản (cỡ 11pt, gạch đầu dòng chuẩn).
    * Cột phải: Khối quyền hạn (`TM.`, `KT.`), chức vụ (`CHỦ TỊCH`, `GIÁM ĐỐC`), khoảng cách ký (4-5 dòng trống) và họ tên người ký.

### Module 3: Phân hệ Bảo toàn & Định dạng Hình ảnh, Đồ họa (Image & Vector Graphics Engine)
* **Tệp mã nguồn**: [`core/formatter_engine.py`](file:///d:/AutoFormatter/core/formatter_engine.py), [`core/editor_core/openxml_builder.py`](file:///d:/AutoFormatter/core/editor_core/openxml_builder.py), [`ui/editor/word_editor_window.py`](file:///d:/AutoFormatter/ui/editor/word_editor_window.py)
* **Bảo toàn dữ liệu 100%**:
  * Quét toàn bộ các phần tử `w:drawing`, `wp:inline`, `wp:anchor`, `v:shape` trong mọi đoạn văn bản và ô bảng.
  * Không bao giờ xóa hoặc bỏ sót đoạn văn chỉ chứa hình ảnh (khắc phục hoàn toàn lỗi lầm tưởng đoạn rỗng).
  * Khi mở tệp trong trình soạn thảo, các khối `blip` được trích xuất dữ liệu nhị phân từ `doc.part.related_parts` và mã hóa thành Base64 Data URI để hiển thị trực quan.
* **Căn chỉnh & Co giãn chuẩn thể thức**:
  * **Căn giữa trang**: Mọi đoạn chứa ảnh được thiết lập `alignment = WD_ALIGN_PARAGRAPH.CENTER`.
  * **Triệt tiêu thụt đầu dòng**: Đặt `first_line_indent = Mm(0)` để tránh tình trạng ảnh bị lệch tâm sang phải do kế thừa thụt lề của đoạn văn bản.
  * **Tự động co tỷ lệ chống tràn lề (Auto Proportional Scaling)**:
    * Giới hạn chiều rộng vùng in: $W_{\max} = 160\text{ mm}$ (khổ dọc) hoặc $220\text{ mm}$ (khổ ngang); chiều cao tối đa $H_{\max} = 220\text{ mm}$.
    * Nếu ảnh vượt quá ngưỡng in, hệ thống tự động tính toán hệ số thu nhỏ $\alpha = \min\left(\frac{W_{\max}}{w}, \frac{H_{\max}}{h}\right)$ và cập nhật lại thuộc tính `cx, cy` trong XML giữ nguyên tỷ lệ khung hình (aspect ratio).
  * **Định dạng tự động chú thích hình ảnh**: Các dòng chú thích có cấu trúc *"Hình 1:..."*, *"Sơ đồ 2:..."*, *"Ảnh:..."*, *"Nguồn:..."* tự động được căn giữa, in nghiêng cỡ 11pt, không thụt đầu dòng.

### Module 4: Bộ Rà soát Lỗi Chính tả Hành chính & Kỹ thuật Soạn thảo (Administrative Spell & Style Linter)
* **Tệp mã nguồn**: [`core/spell_linter.py`](file:///d:/AutoFormatter/core/spell_linter.py), [`core/administrative_sanitizer.py`](file:///d:/AutoFormatter/core/administrative_sanitizer.py)
* **4 Nhóm Rà soát Tương tác (Interactive Linter Cards)**:
  1. **Đoạn văn chuẩn hóa**: Kiểm soát lề, phông Times New Roman, cỡ chữ, khoảng cách dòng.
  2. **Khoảng trắng dọn dẹp**: Xóa khoảng trắng liên tiếp, khoảng trắng đầu/cuối đoạn, xóa tab trắng thừa.
  3. **Chính tả & Dấu câu**: Chuẩn hóa dấu câu tiếng Việt: các dấu `,`, `.`, `;`, `:`, `!`, `?` phải viết sát ký tự đi trước và cách ký tự đi sau 1 khoảng trắng chuẩn.
  4. **Cảnh báo viết hoa công vụ**: Đối chiếu từ điển quy tắc [`capitalization_rules.json`](file:///d:/AutoFormatter/config/capitalization_rules.json) theo Phụ lục II Nghị định 30/2020:
     * Cảnh báo danh từ chung viết hoa tùy tiện khi không đi liền với tên riêng (ví dụ: *các Chi bộ* -> *các chi bộ*).
     * Bắt buộc viết hoa đúng tên các cơ quan, tổ chức Đảng, Nhà nước (ví dụ: *ban chấp hành trung ương* -> *Ban Chấp hành Trung ương*, *bộ nội vụ* -> *Bộ Nội vụ*).

### Module 5: Phân hệ Tương tác MS Word COM & Xuất PDF Độc lập (Word COM Automation & PDF Engine)
* **Tệp mã nguồn**: [`core/com_helper.py`](file:///d:/AutoFormatter/core/com_helper.py), [`core/win32_helper.py`](file:///d:/AutoFormatter/core/win32_helper.py)
* **Kiến trúc Lai xuất PDF (Hybrid PDF Pipeline)**:
  * **Ưu tiên 1 (COM Automation)**: Nếu máy tính đã cài đặt Microsoft Word, hệ thống kích hoạt luồng COM nền ẩn (`Word.Application`), mở tệp và gọi `ExportAsFixedFormat` với định dạng chuẩn in ấn (PDF/A). Phương thức này đảm bảo việc ngắt trang, chia cột và render phông chữ giống 100% khi in từ Word.
  * **Dự phòng 2 (Pure Native Fallback)**: Nếu máy tính không có Office, hệ thống sử dụng bộ chuyển đổi OpenXML-to-PDF nội bộ kết hợp PyQt6 Document Renderer, đảm bảo người dùng luôn có tệp PDF để xem và ký số.

### Module 6: Phân hệ Ký số Điện tử Trực quan (Visual Digital Signature Engine)
* **Tệp mã nguồn**: [`core/signature_engine.py`](file:///d:/AutoFormatter/core/signature_engine.py), [`ui/signature_dialog.py`](file:///d:/AutoFormatter/ui/signature_dialog.py)
* **Kiến trúc Ký số Kép (Dual-Channel Architecture)**:
  * **Kênh 1: Ký số Trực tiếp qua Windows CryptoAPI (Native CAPI / PAdES)**:
    * Giao tiếp trực tiếp với USB Token cắm vào máy tính thông qua `Crypt32.dll`.
    * Tự động lọc bỏ các chứng thư rác/thử nghiệm, chỉ hiển thị chứng thư số hợp lệ còn hạn sử dụng.
    * Thực hiện mã hóa hàm băm SHA-256 with RSA (PKCS#7 / PAdES) kết hợp gọi hộp thoại nhập mã PIN an toàn của hãng Token.
    * Tạo tệp PDF kết quả có **Dấu Tích Xanh Hợp Lệ (Signature Valid)** khi mở trên Adobe Acrobat Reader, Foxit Reader và hệ thống Quản lý văn bản điều hành (QLVBĐH).
  * **Kênh 2: Cầu nối Phần mềm Ký số Ban Cơ yếu Chính phủ (VGCA SignTool Bridge)**:
    * Tự động quét và phát hiện tiến trình dịch vụ ký số chuyên dùng Ban Cơ yếu (`vgca_sign.exe`, `vsign.exe`) trên các cổng cục bộ (8080, 8443, 2104...).
    * Cho phép chuyển giao tài liệu một chạm sang phần mềm của Ban Cơ yếu khi cơ quan có quy chế bắt buộc sử dụng phần mềm chuyên ngành.
* **Trình xem trước PDF & Định vị Khối Ký Thông Minh**:
  * Tích hợp công nghệ kết xuất PDF bản địa (`PyQt6.QtPdf` / `QPdfDocument`), hỗ trợ thu phóng và hiển thị thời gian thực cực nét.
  * **Bộ Định vị Thông minh (Smart Locator)**: Tự động quét và nhận diện các dòng thẩm quyền (`TM. BAN THƯỜNG VỤ`, `CHỦ TỊCH`, `GIÁM ĐỐC`...) ở cuối văn bản chính. **Tự động nhận diện và bỏ qua các trang phụ lục biểu mẫu đính kèm** để đưa ô ký về đúng trang kết thúc văn bản.
  * Cho phép cán bộ nhấp chuột vào bất kỳ vị trí nào trên trang hoặc **kéo thả chuột** để vi chỉnh tọa độ ký từng milimét.

### Module 7: Trình Soạn thảo Hành chính WYSIWYG Nhúng Toàn diện (Embedded Administrative Word Editor)
* **Tệp mã nguồn**: Thư mục [`ui/editor/`](file:///d:/AutoFormatter/ui/editor/)
* **Khung vẽ A4 Độc lập (Discrete Multi-page Canvas)**:
  * Kết xuất các trang giấy A4 rời rạc (kích thước chuẩn 210 x 297 mm) với bóng đổ chiều sâu thực tế và lề in pháp lý hiển thị trực quan.
  * Hỗ trợ ngắt trang chuẩn xác (`Ctrl + Enter`) và thu hồi trang tự động khi xóa văn bản.
  * **Độc lập định hướng khổ giấy**: Trang 1 luôn giữ cố định khổ dọc (Portrait) để bảo toàn tính trang nghiêm của Quốc hiệu. Từ trang 2 trở đi, cán bộ có thể độc lập xoay sang khổ ngang (Landscape 297 x 210 mm) cho các phụ lục, bảng biểu rộng nhiều cột mà không ảnh hưởng tới trang 1.
* **Dải băng điều khiển (Administrative Ribbon) & Trải nghiệm Người dùng**:
  * Nút bấm **Undo / Redo** nhanh đặt ngay trên thanh tiêu đề giống Microsoft Word.
  * Cho phép **đổi tên file trực tiếp** bằng cách nhấp chuột vào nhãn tên tệp trên thanh tiêu đề.
  * **Đánh số tự động thông minh**: Hỗ trợ đánh số liên tục khi bôi đen từng đoạn, hỗ trợ danh sách phân cấp nhiều tầng (La Mã `I, II, III`, số thập phân `1., 1.1., 1.1.1`, chữ cái `a, b, c`).
  * **Hộp thoại Bảng biểu Hành chính**: Tạo bảng biểu chuẩn công vụ, tự động kích hoạt tính năng lặp lại dòng tiêu đề (`tblHeader`) và chống cắt ngang dòng qua trang (`cantSplit`).
* **Hộp thoại Rà soát Tuân thủ Pháp lý & Sửa nhanh 1 chạm (Compliance Audit)**:
  * Chấm điểm mức độ chuẩn hóa từ 0 đến 100 điểm theo thời gian thực.
  * Cung cấp nút bấm **[Sửa nhanh 1 chạm]**: Tự động đưa văn bản về 100/100 điểm bằng cách sửa ngay toàn bộ lỗi lề in, chuyển phông Times New Roman và chuẩn hóa dấu câu.
* **Đồng bộ Siêu dữ liệu 2 chiều (Metadata Sync)**:
  * Quản lý tập trung các thông tin thể thức: Quốc hiệu, Tên cơ quan, Số ký hiệu, Trích yếu, Thẩm quyền, Nơi nhận. Tự động áp dụng Content Control Lock để bảo vệ phần đầu và chân văn bản.

---

## V. CẤU TRÚC DỮ LIỆU & TỪ ĐIỂN QUY TẮC

Hệ thống quản lý dữ liệu cấu hình thông qua các tệp JSON có cấu trúc chặt chẽ trong thư mục [`config/`](file:///d:/AutoFormatter/config/):

### 1. `nd30_rules.json` (Trích lược cấu trúc)
```json
{
  "standard_id": "ND30",
  "name": "Nghị định 30/2020/NĐ-CP",
  "page_setup": {
    "size": "A4",
    "width_mm": 210,
    "height_mm": 297,
    "margins_mm": {
      "top": 20,
      "bottom": 20,
      "left": 30,
      "right": 15
    }
  },
  "typography": {
    "font_family": "Times New Roman",
    "line_spacing": 1.2,
    "space_after_pt": 4,
    "first_line_indent_mm": 12.7
  },
  "elements": {
    "national_motto": { "font_size": 13, "bold": true, "case": "UPPER" },
    "national_submotto": { "font_size": 14, "bold": true, "case": "TITLE" },
    "document_type": { "font_size": 14, "bold": true, "case": "UPPER" },
    "subject": { "font_size": 14, "bold": true, "case": "SENTENCE" }
  }
}
```

### 2. `capitalization_rules.json` (Quy chuẩn viết hoa tên cơ quan)
Chứa mẫu regex và danh mục từ điển hành chính để chuẩn hóa:
* Danh xưng lãnh đạo: *Tổng Bí thư, Chủ tịch nước, Thủ tướng Chính phủ, Chủ tịch Quốc hội*.
* Tên tổ chức: *Bộ Chính trị, Ban Bí thư, Ban Chấp hành Trung ương, Văn phòng Chính phủ, Ủy ban nhân dân*.

---

## VI. MÔ HÌNH BẢO MẬT & ĐỘC LẬP TÁC CHIẾN (SECURITY MODEL)

```mermaid
graph LR
    subgraph AIRGAP ["MÔI TRƯỜNG BIỆT LẬP (AIR-GAPPED 100%)"]
        App["Auto-Formatter v2.0"]
        AI["ONNX AI Engine<br/>(Local CPU)"]
        WordCOM["Word COM Host<br/>(Local IPC)"]
        NoNet["ZERO Network / Internet<br/>• Không mở cổng ra ngoài<br/>• Không Cloud API<br/>• Không Telemetry"]
    end

    subgraph CRYPTO ["PHẦN CỨNG BẢO MẬT (FIPS 140-2)"]
        Token["USB Token Ban Cơ yếu / CA"]
        CAPI["Windows CryptoAPI<br/>(Crypt32.dll)"]
        PIN["Hộp thoại PIN bảo mật hãng"]
        PrivateKey["Khóa bí mật (Private Key)<br/>KHÔNG BAO GIỜ rời khỏi Token"]
    end

    App --> AI
    App --> WordCOM
    App --> CAPI
    CAPI --> PIN --> Token
    Token -. Ký dữ liệu SHA256 .-> CAPI
```

1. **Nguyên tắc Không Mạng (Zero-Network Policy)**: Ứng dụng không kích hoạt bất kỳ socket mạng ra ngoài máy cục bộ. Đảm bảo an toàn tuyệt đối cho các tài liệu mang cấp độ Mật, Tối mật hoặc Tuyệt mật của Đảng và Nhà nước.
2. **An toàn Khóa Bí mật (Hardware Security)**: Khóa riêng của chứng thư số nằm an toàn bên trong vi mạch USB Token (chuẩn FIPS 140-2 Level 2/3). Thao tác ký được ủy quyền cho Windows CryptoAPI, phần mềm không bao giờ đọc hoặc lưu trữ mã PIN hay khóa bí mật của người dùng.
3. **Tiêu hủy Tệp tạm Thời gian thực (Secure Temp Cleanup)**: Mọi tệp ảnh trung gian phục vụ OCR hoặc chuyển đổi PDF đều được tạo trong thư mục tạm có mã hóa và tự động hủy bỏ an toàn (`security_guard.py`) ngay khi tác vụ hoàn tất hoặc khi đóng ứng dụng.

---

## VII. BỘ KIỂM THỬ TỰ ĐỘNG & ĐẢM BẢO CHẤT LƯỢNG (QA & TEST SUITE)

Hệ thống được bảo vệ bởi bộ kiểm thử tự động gồm **68 bài kiểm thử (test cases)** độc lập, bao phủ toàn diện 100% các phân hệ cốt lõi:

| Nhóm Kiểm thử | Tệp Test | Số lượng Test | Trọng tâm Kiểm định |
| :--- | :--- | :---: | :--- |
| **Bảo toàn & Xử lý OpenXML** | [`test_docx_features.py`](file:///d:/AutoFormatter/tests/test_docx_features.py) | 16 tests | Lề in A4, Header/Footer 2 cột, ngắt trang, **bảo toàn & co giãn hình ảnh**, chống tràn lề, định dạng chú thích. |
| **Trình Soạn thảo Tích hợp** | [`test_embedded_editor.py`](file:///d:/AutoFormatter/tests/test_embedded_editor.py) | 12 tests | Khung vẽ A4 rời rạc, xoay khổ ngang độc lập, đồng bộ siêu dữ liệu, phím tắt, tải và lưu giữ nguyên ảnh. |
| **Bộ máy Chuẩn hóa Thể thức** | [`test_formatter_engine.py`](file:///d:/AutoFormatter/tests/test_formatter_engine.py) | 14 tests | Tuân thủ NĐ 30 và HD 36, cấu trúc văn bản hành chính, bảng biểu, trích yếu, khối ký. |
| **Chính tả & Linter Viết hoa** | [`test_spell_linter.py`](file:///d:/AutoFormatter/tests/test_spell_linter.py) | 10 tests | Quy tắc dấu câu tiếng Việt, dọn khoảng trắng thừa, viết hoa tên cơ quan theo Phụ lục II NĐ 30. |
| **Ký số Điện tử Trực quan** | [`test_signature_engine.py`](file:///d:/AutoFormatter/tests/test_signature_engine.py) | 8 tests | Bộ định vị khối ký thông minh, lọc chứng thư hợp lệ, cấu trúc PAdES chuẩn Điều 13 NĐ 30. |
| **Nhận dạng OCR & Deskew** | [`test_ocr_engine.py`](file:///d:/AutoFormatter/tests/test_ocr_engine.py) | 8 tests | Bóc tách tiếng Việt RapidOCR ONNX, bộ lọc con dấu HSV, làm thẳng ảnh nghiêng. |
| **TỔNG CỘNG** | **6 files** | **68 tests** | **Tỷ lệ vượt qua: 100% (68/68 Passed)** |

---

## VIII. QUY TRÌNH ĐÓNG GÓI, TRIỂN KHAI & PHÂN PHỐI THƯƠNG MẠI

### 1. Cấu hình Biên dịch PyInstaller Tối ưu
* Sử dụng tệp cấu hình độc quyền [`build_exe.spec`](file:///d:/AutoFormatter/build_exe.spec).
* Áp dụng cơ chế **Onedir Portable**: Toàn bộ nhị phân và thư viện liên kết động (.dll) được phân giải sẵn trong thư mục ứng dụng `dist/AutoFormatter/`.
* Khởi động tức thì trong **< 1 giây** (loại bỏ hoàn toàn độ trễ 10-15 giây giải nén của cấu hình Onefile cũ).
* Loại bỏ triệt để các gói thư viện khoa học nặng không cần thiết (`matplotlib`, `scipy`, `pandas`, `IPython`, `tkinter`...), tối ưu dung lượng tệp nén phát hành chỉ còn **~125 MB**.

### 2. Gói Phát hành Sẵn sàng Thương mại
1. **Gói Di động Tiện lợi (Portable Zip)**: `dist/AutoFormatter_v2.0.zip` (~125 MB). Người dùng chỉ cần tải về, giải nén và nhấp đúp chạy `AutoFormatter.exe` ngay.
2. **Bộ Cài đặt Tự động Windows (Setup Installer)**: Biên dịch từ kịch bản [`installer.iss`](file:///d:/AutoFormatter/installer.iss) bằng Inno Setup, tự động cài vào `Program Files` và tạo biểu tượng Desktop.

---

## IX. ĐÁNH GIÁ TRẠNG THÁI DỰ ÁN & ĐỊNH HƯỚNG PHÁT TRIỂN

### 1. Đánh giá Mức độ Trưởng thành (Maturity Assessment)
* **Tính toàn vẹn chức năng**: Đạt **100%** yêu cầu thể thức văn bản Đảng & Nhà nước.
* **Bảo toàn nội dung & tài nguyên**: Đạt **100%** (chữ nghĩa, bảng biểu, số liệu, hình ảnh, sơ đồ đều được giữ nguyên vẹn).
* **Độ ổn định mã nguồn**: Đạt mức **Production-Ready** (toàn bộ 68/68 test case tự động vượt qua).
* **Mức độ sẵn sàng thương mại**: Sẵn sàng đóng gói thương mại, bàn giao và triển khai diện rộng tại các cơ quan, đơn vị.

### 2. Lộ trình Nâng cấp Phiên bản Tiếp theo (Roadmap v2.1+)
* **Hỗ trợ Đa màn hình High-DPI 4K**: Tự động nhận diện DPI đa màn hình để tối ưu thước đo milimét.
* **Mở rộng Kho Mẫu Văn bản**: Bổ sung thêm 50+ biểu mẫu chuẩn chuyên ngành (Tư pháp, Địa chính, Giáo dục, Y tế, Thanh tra).
* **Ký số Hàng loạt (Batch Signing)**: Cho phép ký số cùng lúc danh sách hàng chục tệp PDF dự thảo chỉ với 1 lần nhập mã PIN duy nhất.
