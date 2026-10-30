"""
scripts/create_extreme_error_doc.py:
Tạo tài liệu mẫu chứa toàn bộ các loại lỗi thể thức điển hình:
1. Lỗi Lề trang & Khổ giấy (Letter thay vì A4, lề lệch chuẩn)
2. Lỗi Quốc hiệu (viết thường lộn xộn, không in hoa đứng đậm)
3. Lỗi Tiêu ngữ (viết in hoa toàn bộ, dấu phẩy thay vì gạch nối, đường kẻ dấu chấm)
4. Lỗi Tên cơ quan ban hành (viết chữ thường, thụt lệch)
5. Lỗi Số, ký hiệu & Ngày tháng (dính dấu, sai số 0, viết hoa địa danh)
6. Lỗi Tên loại & Trích yếu văn bản (viết chữ thường, phông chữ lệch)
7. Lỗi Thụt lề, căn chỉnh, phông chữ (Calibri, Arial, Tahoma thay vì Times New Roman)
8. Lỗi Khoảng trắng thừa & dấu câu (dính dấu phẩy, chấm, chấm phẩy)
9. Lỗi Viết hoa tùy tiện danh từ chung (các Chi bộ, các Đồng chí, các Đơn vị...)
10. Lỗi Quên viết hoa cơ quan Đảng & Nhà nước (ủy ban nhân dân, bộ y tế, thủ tướng chính phủ...)
11. Lỗi Khối chữ ký & Nơi nhận (viết thường toàn bộ, lệch cột)
"""

import os
import sys
import docx
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH


def create_comprehensive_error_doc(output_path: str = "Tai_Lieu_Mau_Test/Van_Ban_Loi_Toan_Dien_Quoc_Hieu_The_Thuc.docx"):
    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    doc = docx.Document()

    # 1. SAI KHỔ GIẤY VÀ LỀ TRANG
    section = doc.sections[0]
    section.page_width = Inches(8.5)   # Khổ giấy Letter
    section.page_height = Inches(11.0)
    section.top_margin = Cm(1.0)      # Lề trên quá hẹp (1.0cm thay vì 2.0-2.5cm)
    section.bottom_margin = Cm(1.2)   # Lề dưới quá hẹp
    section.left_margin = Cm(1.5)     # Lề trái quá hẹp (1.5cm thay vì 3.0-3.5cm)
    section.right_margin = Cm(3.8)    # Lề phải quá rộng (3.8cm thay vì 1.5-2.0cm)

    # 2. LỖI QUỐC HIỆU VÀ CƠ QUAN CẤP TRÊN (DÒNG 1)
    # Cơ quan cấp trên viết thường; Quốc hiệu viết thường không in hoa, phông Arial
    p1 = doc.add_paragraph()
    p1.paragraph_format.space_before = Pt(0)
    p1.paragraph_format.space_after = Pt(0)
    p1.paragraph_format.line_spacing = 1.0

    r_cq = p1.add_run("ubnd tỉnh quảng ninh                                             ")
    r_cq.font.name = "Calibri"
    r_cq.font.size = Pt(11)

    r_qh = p1.add_run("cộng hòa xã hội chủ nghĩa việt nam")
    r_qh.font.name = "Arial"
    r_qh.font.size = Pt(12)
    r_qh.font.bold = False

    # 3. LỖI TIÊU NGỮ VÀ CƠ QUAN BAN HÀNH (DÒNG 2)
    # Tiêu ngữ viết in hoa toàn bộ, dùng dấu phẩy ngăn cách thay vì gạch nối, phông Tahoma
    p2 = doc.add_paragraph()
    p2.paragraph_format.space_before = Pt(0)
    p2.paragraph_format.space_after = Pt(0)
    p2.paragraph_format.line_spacing = 1.0

    r_dv = p2.add_run("sở tài chính                                                               ")
    r_dv.font.name = "Calibri"
    r_dv.font.size = Pt(11)

    r_tn = p2.add_run("ĐỘC LẬP, TỰ DO, HẠNH PHÚC")
    r_tn.font.name = "Tahoma"
    r_tn.font.size = Pt(12)
    r_tn.font.italic = True

    # 4. ĐƯỜNG KẺ TIÊU NGỮ VÀ CƠ QUAN SAI (DÒNG 3)
    p_line = doc.add_paragraph()
    p_line.paragraph_format.space_before = Pt(0)
    p_line.paragraph_format.space_after = Pt(4)
    r_line = p_line.add_run("  -------------                                                                    ..................................................")
    r_line.font.name = "Courier New"
    r_line.font.size = Pt(9)

    # 5. SỐ KÝ HIỆU VÀ ĐỊA DANH NGÀY THÁNG SAI (DÒNG 4)
    # Số ký hiệu dính khoảng trắng, viết chữ thường: Số : 89 / stc - khtc
    # Địa danh in hoa toàn bộ, ngày 5 thiếu số 0, tháng 09 thừa số 0, dính dấu phẩy
    p3 = doc.add_paragraph()
    p3.paragraph_format.space_before = Pt(0)
    p3.paragraph_format.space_after = Pt(12)

    r_so = p3.add_run("Số :  89 / stc - khtc                                            ")
    r_so.font.name = "Calibri"
    r_so.font.size = Pt(10)

    r_ngay = p3.add_run("QUẢNG NINH ,ngày 5 tháng 09 năm 2026")
    r_ngay.font.name = "Arial"
    r_ngay.font.size = Pt(11)

    # 6. TÊN LOẠI VĂN BẢN VÀ TRÍCH YẾU NỘI DUNG SAI
    # Tên loại viết thường, phông Courier New, cỡ chữ quá to (18pt)
    p_t = doc.add_paragraph()
    p_t.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_t.paragraph_format.space_before = Pt(10)
    p_t.paragraph_format.space_after = Pt(4)
    r_t = p_t.add_run("tờ   trình")
    r_t.font.name = "Courier New"
    r_t.font.size = Pt(18)
    r_t.font.bold = True

    # Trích yếu căn trái lệch lạc, cỡ 13pt
    p_sub = doc.add_paragraph()
    p_sub.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_sub.paragraph_format.space_before = Pt(0)
    p_sub.paragraph_format.space_after = Pt(14)
    r_sub = p_sub.add_run("Về việc   phân bổ dự toán kinh phí triển khai chuyển đổi số và cải cách hành chính năm 2026")
    r_sub.font.name = "Calibri"
    r_sub.font.size = Pt(13)
    r_sub.font.bold = True
    r_sub.font.italic = True

    # 7. NỘI DUNG VĂN BẢN - ĐOẠN 1
    # - Không thụt đầu dòng
    # - Dòng đơn 1.0 sít sao
    # - Dính dấu chấm, dấu phẩy
    # - Không viết hoa cơ quan nhà nước (ủy ban nhân dân, sở nội vụ)
    # - Viết hoa tùy tiện danh từ chung (các Cơ quan, các Đơn vị, các Chi bộ)
    p_b1 = doc.add_paragraph()
    p_b1.paragraph_format.line_spacing = 1.0
    p_b1.paragraph_format.space_before = Pt(0)
    p_b1.paragraph_format.space_after = Pt(0)
    p_b1.paragraph_format.first_line_indent = Pt(0)
    r_b1 = p_b1.add_run(
        "Căn cứ ý kiến chỉ đạo của ủy ban nhân dân tỉnh tại Thông báo số 45/TB-UBND  ;sở tài chính đã chủ trì  ,phối hợp "
        "với sở nội vụ và các Cơ quan  ,Đơn vị trên địa bàn tiến hành rà soát nhu cầu kinh phí  .Thời gian qua  ,"
        "các Chi bộ và các Ban ngành trực thuộc đã tích cực lập dự toán  ,tuy nhiên nhiều Đơn vị vẫn chưa hoàn thiện hồ sơ  ."
    )
    r_b1.font.name = "Calibri"
    r_b1.font.size = Pt(11)

    # 8. NỘI DUNG VĂN BẢN - ĐOẠN 2
    # - Dòng đôi 2.0 quá thưa
    # - Dùng 15 dấu cách space thủ công thay cho thụt đầu dòng
    # - Không viết hoa cơ quan Đảng & Chính phủ (bộ y tế, ban thường vụ tỉnh ủy, chính phủ, ban chấp hành trung ương)
    # - Viết hoa sai (các Đồng chí, Lãnh đạo các Sở)
    p_b2 = doc.add_paragraph()
    p_b2.paragraph_format.line_spacing = 2.0
    p_b2.paragraph_format.space_before = Pt(14)
    p_b2.paragraph_format.space_after = Pt(14)
    r_b2 = p_b2.add_run(
        "               Theo hướng dẫn của bộ y tế và ban thường vụ tỉnh ủy  ,việc bảo đảm cơ sở vật chất cần có sự tham gia "
        "trách nhiệm của các Đồng chí Lãnh đạo các Sở  ,Ngành  .Để kịp thời báo cáo ban chấp hành trung ương và chính phủ  ,"
        "kính đề nghị ban cán sự đảng ủy ban nhân dân tỉnh xem xét  ,cho chủ trương thực hiện  ."
    )
    r_b2.font.name = "Arial"
    r_b2.font.size = Pt(15)

    # 9. NỘI DUNG VĂN BẢN - ĐOẠN 3
    # - Sai quy tắc viết hoa: thủ tướng chính phủ, bộ tài chính
    # - Viết hoa sai: các Phòng chuyên môn, Kế hoạch
    p_b3 = doc.add_paragraph()
    p_b3.paragraph_format.line_spacing = 1.15
    p_b3.paragraph_format.space_before = Pt(0)
    p_b3.paragraph_format.space_after = Pt(12)
    r_b3 = p_b3.add_run(
        "Sở Tài chính kính trình ủy ban nhân dân tỉnh xem xét  ,báo cáo thủ tướng chính phủ và bộ tài chính "
        "bổ sung nguồn lực  .Giao các Phòng chuyên môn đôn đốc các Đơn vị thực hiện nghiêm túc Kế hoạch này  ."
    )
    r_b3.font.name = "Tahoma"
    r_b3.font.size = Pt(10.5)

    # 10. KHỐI CHỮ KÝ & NƠI NHẬN - SAI THỂ THỨC TOÀN DIỆN
    # Viết chữ thường lộn xộn, thụt lệch dòng, không có khoảng cách ký
    p_k1 = doc.add_paragraph()
    p_k1.paragraph_format.space_before = Pt(8)
    p_k1.paragraph_format.space_after = Pt(2)
    r_k1 = p_k1.add_run("Nơi nhận:                                                                                 tm. ủy ban nhân dân tỉnh")
    r_k1.font.name = "Calibri"
    r_k1.font.size = Pt(10)

    p_k2 = doc.add_paragraph()
    p_k2.paragraph_format.space_before = Pt(0)
    p_k2.paragraph_format.space_after = Pt(2)
    r_k2 = p_k2.add_run("- như trên;                                                                               kt. chủ tịch")
    r_k2.font.name = "Calibri"
    r_k2.font.size = Pt(10)

    p_k3 = doc.add_paragraph()
    p_k3.paragraph_format.space_before = Pt(0)
    p_k3.paragraph_format.space_after = Pt(2)
    r_k3 = p_k3.add_run("- lưu: vt, khtc.                                                                         phó chủ tịch")
    r_k3.font.name = "Calibri"
    r_k3.font.size = Pt(10)

    p_k4 = doc.add_paragraph()
    p_k4.paragraph_format.space_before = Pt(0)
    p_k4.paragraph_format.space_after = Pt(0)
    r_k4 = p_k4.add_run("                                                                                                  trần văn hoàng")
    r_k4.font.name = "Calibri"
    r_k4.font.size = Pt(11)

    abs_out = os.path.abspath(output_path)
    doc.save(abs_out)
    return abs_out


if __name__ == "__main__":
    out_docx = create_comprehensive_error_doc()
    print("SUCCESS_DOCX:" + out_docx)
