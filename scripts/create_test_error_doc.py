"""
Tạo văn bản mẫu chứa đầy đủ các loại lỗi thể thức, chính tả hành chính, lề trang, phông chữ
để người dùng kiểm thử ứng dụng Auto-Formatter.
"""

import os
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE


def create_error_document(output_path: str = "Van_Ban_Mau_Loi_Kiem_Tra.docx"):
    doc = docx.Document()

    # 1. CỐ TÌNH CẤU HÌNH SAI KHỔ GIẤY VÀ LỀ TRANG
    # Khổ giấy Letter (thay vì A4 tiêu chuẩn 210 x 297 mm)
    section = doc.sections[0]
    section.page_width = Inches(8.5)
    section.page_height = Inches(11.0)

    # Lề sai quy định: Trên 1.0cm, Dưới 1.0cm, Trái 1.5cm, Phải 3.5cm
    # (Chuẩn NĐ30/HD36: Trên: 2.0-2.5cm, Dưới: 2.0-2.5cm, Trái: 3.0-3.5cm, Phải: 1.5-2.0cm)
    section.top_margin = Cm(1.0)
    section.bottom_margin = Cm(1.0)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(3.5)

    # 2. CỐ TÌNH DÙNG PHÔNG CHỮ SAI (Calibri, Arial, Tahoma thay vì Times New Roman)
    # VÀ PHẦN ĐẦU VĂN BẢN ĐÁNH THỦ CÔNG BẰNG DẤU CÁCH / PHÍM TAB LỆCH DÒNG

    p1 = doc.add_paragraph()
    r1 = p1.add_run("UBND TỈNH QUẢNG NINH                                CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM")
    r1.font.name = "Calibri"
    r1.font.size = Pt(11)
    r1.font.color.rgb = RGBColor(30, 41, 59)

    p2 = doc.add_paragraph()
    r2 = p2.add_run("SỞ NỘI VỤ                                                            Độc lập - Tự do - Hạnh phúc")
    r2.font.name = "Arial"
    r2.font.size = Pt(11)
    r2.font.bold = True

    # Ngày tháng sai: ngày 5 (thiếu số 0), tháng 04 (thừa số 0 theo NĐ 30), cách dấu phẩy sai
    p3 = doc.add_paragraph()
    r3 = p3.add_run("Số:  45 /SNV - VP                                  Quảng Ninh ,ngày 5 tháng 04 năm 2026")
    r3.font.name = "Tahoma"
    r3.font.size = Pt(10)
    r3.font.italic = True

    # 3. TIÊU ĐỀ VĂN BẢN: Cỡ chữ quá to (18pt), phông Arial, thừa nhiều khoảng trắng
    doc.add_paragraph()  # đoạn trống
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("BÁO   CÁO")
    r_title.font.name = "Arial"
    r_title.font.size = Pt(18)
    r_title.font.bold = True

    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_sub = p_sub.add_run("Về    việc rà  soát ,đánh  giá công tác cải cách hành chính và chuyển đổi số")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.italic = True

    # 4. NỘI DUNG VĂN BẢN (BODY PARAGRAPHS) VỚI CÁC LOẠI LỖI TOÀN DIỆN:
    # Lỗi 1: Khoảng trắng thừa, dính dấu câu, không viết hoa cơ quan nhà nước (ủy ban nhân dân),
    # viết hoa tùy tiện danh từ chung (các Cơ quan, các Đơn vị).
    p_body1 = doc.add_paragraph()
    p_body1.paragraph_format.line_spacing = 1.0  # Dòng đơn quá sít
    p_body1.paragraph_format.space_before = Pt(0)
    p_body1.paragraph_format.space_after = Pt(0)
    p_body1.paragraph_format.first_line_indent = Pt(0)  # Không thụt đầu dòng
    r_b1 = p_body1.add_run(
        "Thực hiện ý kiến chỉ đạo của ủy ban nhân dân tỉnh tại Công văn số 123/UBND-TH  ,Sở Nội vụ đã chủ trì  ,phối hợp "
        "với các Cơ quan  ,Đơn vị tiến hành kiểm tra tình hình thực tế  .Qua rà soát  ,nhìn chung các Đơn vị đã nghiêm túc "
        "triển khai nhiệm vụ được giao  ."
    )
    r_b1.font.name = "Calibri"
    r_b1.font.size = Pt(11)

    # Lỗi 2: Dấu ngoặc kép thẳng "", ngày tháng sai quy chuẩn (ngày 3 tháng 05),
    # không viết hoa cơ quan chính trị (đảng cộng sản việt nam, chính phủ),
    # viết hoa tùy tiện danh từ chung (toàn thể Cán bộ, Công chức, Các Phòng, Kế hoạch).
    p_body2 = doc.add_paragraph()
    p_body2.paragraph_format.line_spacing = 2.0  # Dòng đôi quá thưa
    p_body2.paragraph_format.space_before = Pt(18)  # Khoảng cách trước quá lớn
    p_body2.paragraph_format.space_after = Pt(18)
    r_b2 = p_body2.add_run(
        "Trong giai đoạn từ ngày 3 tháng 05 năm 2026 đến ngày 9 tháng 8 năm 2026  ,toàn thể Cán bộ  ,Công chức đã tích cực "
        "tham gia phong trào \"Học tập và làm theo tư tưởng, đạo đức, phong cách Hồ Chí Minh\" do đảng cộng sản việt nam và "
        "chính phủ phát động  .Các Phòng chuyên môn đã hoàn thành 100% Kế hoạch đề ra  ."
    )
    r_b2.font.name = "Arial"
    r_b2.font.size = Pt(16)  # Cỡ chữ 16pt quá lớn cho nội dung

    # Lỗi 3: Dùng phím Space thủ công thay vì Tab/Thụt đầu dòng, dính dấu hai chấm / chấm phẩy,
    # viết hoa sai (Đồng chí Đảng viên, Chi bộ), quên viết hoa (ban thường vụ tỉnh ủy, thủ tướng chính phủ).
    p_body3 = doc.add_paragraph()
    r_b3 = p_body3.add_run(
        "          Tuy nhiên  ,bên cạnh những kết quả đạt được  ,vẫn còn một số tồn tại  :Một số Đồng chí Đảng viên chưa thực "
        "sự chủ động trong việc ứng dụng công nghệ thông tin  ;sinh hoạt Chi bộ đôi khi còn hình thức  .Đề nghị ban thường vụ "
        "tỉnh ủy và thủ tướng chính phủ tiếp tục quan tâm  ,chỉ đạo sát sao hơn nữa trong thời gian tới  ."
    )
    r_b3.font.name = "Tahoma"
    r_b3.font.size = Pt(10.5)

    # 5. KHUNG CUỐI VĂN BẢN (NƠI NHẬN & CHỮ KÝ) ĐÁNH LỆCH LẠC, THIẾU KHOẢNG KÝ
    doc.add_paragraph()
    p_sign1 = doc.add_paragraph()
    r_s1 = p_sign1.add_run("Nơi nhận :                                                          TM .UBND TỈNH")
    r_s1.font.name = "Calibri"
    r_s1.font.size = Pt(10)

    p_sign2 = doc.add_paragraph()
    r_s2 = p_sign2.add_run("- Như trên ;                                                         KT .CHỦ TỊCH")
    r_s2.font.name = "Calibri"
    r_s2.font.size = Pt(10)

    p_sign3 = doc.add_paragraph()
    r_s3 = p_sign3.add_run("- Lưu :VT ,SNV .                                                    PHÓ CHỦ TỊCH")
    r_s3.font.name = "Calibri"
    r_s3.font.size = Pt(10)

    # Tên người ký viết thường, phông chữ lệch
    p_sign4 = doc.add_paragraph()
    r_s4 = p_sign4.add_run("                                                                    nguyễn văn an")
    r_s4.font.name = "Calibri"
    r_s4.font.size = Pt(11)

    abs_output = os.path.abspath(output_path)
    doc.save(abs_output)
    print(f"[OK] Da tao thanh cong tep van ban mau loi tai: {abs_output}")
    return abs_output


if __name__ == "__main__":
    create_error_document()
