"""
Tạo các văn bản mẫu chuẩn (reference templates) theo Hướng dẫn 36-HD/VPTW và Nghị định 30/2020/NĐ-CP
"""

import os
import docx
from core.formatter_engine import FormatterEngine


def generate_templates():
    tmpl_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "templates")
    os.makedirs(tmpl_dir, exist_ok=True)
    engine = FormatterEngine()

    # 1. Mẫu Nghị định 30/2020/NĐ-CP
    doc_nd30 = docx.Document()
    doc_nd30.add_paragraph("QUYẾT ĐỊNH")
    doc_nd30.add_paragraph("Về việc ban hành quy chế công tác văn thư và lưu trữ")
    doc_nd30.add_paragraph("Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;")
    doc_nd30.add_paragraph("Căn cứ Nghị định số 30/2020/NĐ-CP ngày 05 tháng 3 năm 2020 của Chính phủ về công tác văn thư;")
    doc_nd30.add_paragraph("Theo đề nghị của Giám đốc Sở Nội vụ tại Tờ trình số 15/TTr-SNV ngày 10 tháng 3 năm 2026.")
    doc_nd30.add_paragraph("Điều 1. Ban hành kèm theo Quyết định này “Quy chế công tác văn thư và lưu trữ điện tử”.")
    doc_nd30.add_paragraph("Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.")
    doc_nd30.add_paragraph("Điều 3. Chánh Văn phòng, Giám đốc các Sở, Thủ trưởng các cơ quan, đơn vị chịu trách nhiệm thi hành Quyết định này./.")

    draft_nd30 = os.path.join(tmpl_dir, "temp_nd30.docx")
    target_nd30 = os.path.join(tmpl_dir, "Mau_Van_Ban_Hanh_Chinh_ND30.docx")
    doc_nd30.save(draft_nd30)

    engine.format_document(
        input_docx_path=draft_nd30,
        output_docx_path=target_nd30,
        standard_id="ND30",
        profile_override={
            "superior_agency": "ỦY BAN NHÂN DÂN",
            "issuing_agency": "THÀNH PHỐ HÀ NỘI",
            "doc_number_symbol": "Số: 01/2026/QĐ-UBND",
            "location": "Hà Nội",
            "recipients": [
                "- Như Điều 3;",
                "- Thường trực HĐND thành phố;",
                "- Chủ tịch, các Phó Chủ tịch UBND TP;",
                "- Cổng Thông tin điện tử thành phố;",
                "- Lưu: VT, TH."
            ],
            "authority_prefix": "TM. ỦY BAN NHÂN DÂN",
            "signer_position": "CHỦ TỊCH",
            "signer_name": "Nguyễn Văn An"
        }
    )
    if os.path.exists(draft_nd30):
        os.remove(draft_nd30)
    if os.path.exists(f"{draft_nd30}.bak"):
        os.remove(f"{draft_nd30}.bak")

    # 2. Mẫu Hướng dẫn 36-HD/VPTW (Văn bản Đảng)
    doc_hd36 = docx.Document()
    doc_hd36.add_paragraph("NGHỊ QUYẾT")
    doc_hd36.add_paragraph("Về tăng cường sự lãnh đạo của Đảng đối với công tác chuyển đổi số")
    doc_hd36.add_paragraph("Căn cứ Điều lệ Đảng Cộng sản Việt Nam;")
    doc_hd36.add_paragraph("Căn cứ Quy chế làm việc của Ban Chấp hành Đảng bộ tỉnh nhiệm kỳ 2025 - 2030;")
    doc_hd36.add_paragraph("Sau khi xem xét Tờ trình của Ban Thường vụ Tỉnh ủy, Ban Chấp hành Đảng bộ quyết nghị:")
    doc_hd36.add_paragraph("I. TÌNH HÌNH VÀ NGUYÊN NHÂN")
    doc_hd36.add_paragraph("Trong những năm qua, công tác ứng dụng công nghệ thông tin và chuyển đổi số trong các cơ quan đảng đã đạt nhiều kết quả quan trọng.")
    doc_hd36.add_paragraph("II. MỤC TIÊU VÀ NHIỆM VỤ TRỌNG TÂM")
    doc_hd36.add_paragraph("1. Mục tiêu tổng quát: Xây dựng hệ thống thông tin phục vụ công tác lãnh đạo, chỉ đạo thông suốt, an toàn và bảo mật tuyệt đối.")
    doc_hd36.add_paragraph("2. Nhiệm vụ, giải pháp: Đẩy mạnh chuẩn hóa thể thức văn bản theo Hướng dẫn 36-HD/VPTW của Văn phòng Trung ương Đảng.")
    doc_hd36.add_paragraph("III. TỔ CHỨC THỰC HIỆN")
    doc_hd36.add_paragraph("Các huyện ủy, thị ủy, thành ủy và đảng ủy trực thuộc quán triệt sâu rộng Nghị quyết này đến toàn thể cán bộ, đảng viên./.")

    draft_hd36 = os.path.join(tmpl_dir, "temp_hd36.docx")
    target_hd36 = os.path.join(tmpl_dir, "Mau_Van_Ban_Dang_HD36.docx")
    doc_hd36.save(draft_hd36)

    engine.format_document(
        input_docx_path=draft_hd36,
        output_docx_path=target_hd36,
        standard_id="HD36",
        profile_override={
            "superior_agency": "TỈNH ỦY HÀ NỘI",
            "issuing_agency": "BAN CHẤP HÀNH",
            "doc_number_symbol": "Số 05-NQ/TU",
            "location": "Hà Nội",
            "recipients": [
                "- Ban Bí thư Trung ương Đảng (để b/c);",
                "- Các ban đảng Trung ương;",
                "- Thường trực Tỉnh ủy;",
                "- Các huyện ủy, thị ủy, đảng ủy trực thuộc;",
                "- Lưu: VT, TH."
            ],
            "authority_prefix": "T/M BAN CHẤP HÀNH",
            "signer_position": "BÍ THƯ",
            "signer_name": "Trần Văn Bình"
        }
    )
    if os.path.exists(draft_hd36):
        os.remove(draft_hd36)
    if os.path.exists(f"{draft_hd36}.bak"):
        os.remove(f"{draft_hd36}.bak")

    print("Generated templates successfully!")


if __name__ == "__main__":
    generate_templates()
