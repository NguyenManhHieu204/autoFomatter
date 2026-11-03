"""
core/editor_core/openxml_builder.py: Bộ sinh mã cấu trúc OpenXML / DOCX chuẩn hóa
- Sinh tệp .docx chuẩn từ siêu dữ liệu (Metadata) hành chính & Đảng
- Dựng Khung Tiêu đề & Quốc hiệu (Header) và Khung Nơi nhận & Chữ ký (Footer)
- Áp dụng cơ chế Khóa cấu trúc thể thức (w:sdt với w:lock w:val="sdtContentLocked")
- Định dạng chuẩn A4, lề in pháp lý (NĐ 30: 20-20-30-15 mm), font Times New Roman Unicode
"""

import os
import re
from typing import Dict, Any, List, Optional
import docx
from docx.shared import Mm, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

from core.editor_core.table_layout_manager import TableLayoutManager


def format_run_font(run, font_name: str = "Times New Roman", size_pt: float = 13.0, bold: bool = False, italic: bool = False, underline: bool = False):
    """Định dạng phông chữ chuẩn TCVN / Unicode"""
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline

    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


class OpenXmlBuilder:
    """
    Sinh tài liệu Word OpenXML từ Metadata mẫu với cấu trúc bảo vệ (Content Control Lock)
    """

    @classmethod
    def apply_page_setup(cls, doc: docx.Document, standard_id: str = "ND30"):
        """Thiết lập khổ giấy A4 và lề in theo Nghị định 30 / Hướng dẫn 36"""
        for section in doc.sections:
            section.page_width = Mm(210.0)
            section.page_height = Mm(297.0)
            section.top_margin = Mm(20.0)
            section.bottom_margin = Mm(20.0)
            section.left_margin = Mm(30.0)
            section.right_margin = Mm(15.0)

            # Đánh số trang: Ẩn trang đầu, đánh số từ trang 2 ở giữa Header
            section.different_first_page_header_footer = True
            header = section.header
            p_hdr = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
            p_hdr.text = ""
            p_hdr.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_hdr.paragraph_format.space_before = Pt(0)
            p_hdr.paragraph_format.space_after = Pt(0)

            run_pn = p_hdr.add_run()
            format_run_font(run_pn, font_name="Times New Roman", size_pt=13.5)
            fld_xml = parse_xml(r'<w:fldSimple %s w:instr="PAGE"/>' % nsdecls('w'))
            run_pn._r.append(fld_xml)

            # Trang đầu tiên để trống header
            first_header = section.first_page_header
            if first_header.paragraphs:
                first_header.paragraphs[0].text = ""

    @classmethod
    def wrap_element_in_locked_sdt(cls, element: Any, tag_name: str = "AdministrativeLockedRegion"):
        """
        Bọc phần tử XML (Table hoặc Paragraph) vào thẻ Content Control w:sdt có khóa:
        <w:sdt>
          <w:sdtPr>
            <w:lock w:val="sdtContentLocked"/>
            <w:tag w:val="{tag_name}"/>
          </w:sdtPr>
          <w:sdtContent>
            ...
          </w:sdtContent>
        </w:sdt>
        """
        parent = element.getparent()
        if parent is None:
            return

        sdt = parse_xml(
            f'<w:sdt {nsdecls("w")}>\n'
            f'  <w:sdtPr>\n'
            f'    <w:lock w:val="sdtContentLocked"/>\n'
            f'    <w:tag w:val="{tag_name}"/>\n'
            f'  </w:sdtPr>\n'
            f'  <w:sdtContent/>\n'
            f'</w:sdt>'
        )
        sdt_content = sdt.find(qn('w:sdtContent'))
        parent.insert(parent.index(element), sdt)
        parent.remove(element)
        sdt_content.append(element)

    @classmethod
    def create_scaffolded_document(cls, meta: Dict[str, Any], lock_scaffolding: bool = True) -> docx.Document:
        """
        Khởi tạo tài liệu Word hoàn chỉnh với khung thể thức chuẩn hóa:
        - Khung Header (Quốc hiệu/Tiêu ngữ hoặc Tiêu đề Đảng)
        - Tiêu đề loại văn bản & Trích yếu
        - Khối nội dung mẫu theo loại văn bản (Thân văn bản mở)
        - Khung Footer (Nơi nhận & Chữ ký)
        """
        standard_id = meta.get("standard_id", "ND30").upper()
        doc = docx.Document()
        cls.apply_page_setup(doc, standard_id=standard_id)

        # 1. BẢNG KHUNG ĐẦU VĂN BẢN (HEADER TABLE)
        header_table = doc.add_table(rows=1, cols=2)
        header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        TableLayoutManager.make_table_borders_invisible(header_table)
        TableLayoutManager.set_column_widths(header_table, [68.0, 97.0])

        cell_left = header_table.cell(0, 0)
        cell_right = header_table.cell(0, 1)
        TableLayoutManager.set_cell_margins(cell_left, top=30, bottom=30, left=30, right=30)
        TableLayoutManager.set_cell_margins(cell_right, top=30, bottom=30, left=30, right=30)

        # Cột Trái Header: Cơ quan ban hành & Số hiệu
        p_l = cell_left.paragraphs[0]
        p_l.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l.paragraph_format.space_before = Pt(0)
        p_l.paragraph_format.space_after = Pt(2)
        p_l.paragraph_format.line_spacing = 1.1

        sup_agency = meta.get("superior_agency", "").strip().upper()
        if sup_agency:
            r_sup = p_l.add_run(sup_agency + "\n")
            format_run_font(r_sup, size_pt=12.0, bold=False)

        iss_agency = meta.get("issuing_agency", "").strip().upper() or ("BAN THƯỜNG VỤ" if standard_id == "HD36" else "ỦY BAN NHÂN DÂN")
        r_iss = p_l.add_run(iss_agency)
        format_run_font(r_iss, size_pt=12.5, bold=True)

        p_l_div = cell_left.add_paragraph()
        p_l_div.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l_div.paragraph_format.space_before = Pt(1)
        p_l_div.paragraph_format.space_after = Pt(4)
        div_symbol = "———*———" if standard_id == "HD36" else "—————"
        r_div = p_l_div.add_run(div_symbol)
        format_run_font(r_div, size_pt=9.0, bold=True)

        p_l_num = cell_left.add_paragraph()
        p_l_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l_num.paragraph_format.space_before = Pt(2)
        p_l_num.paragraph_format.space_after = Pt(0)

        raw_num = meta.get("doc_number", "").strip()
        num_prefix = "Số " if standard_id == "HD36" else "Số: "
        if not raw_num:
            raw_num = "....-QĐ/...." if standard_id == "HD36" else "..../QĐ-...."
        clean_num = re.sub(r'^(số|so)\s*[:\.]?\s*', '', raw_num, flags=re.IGNORECASE)
        r_num = p_l_num.add_run(f"{num_prefix}{clean_num}")
        format_run_font(r_num, size_pt=12.0, bold=False)

        # Cột Phải Header: Quốc hiệu / Tiêu ngữ hoặc Đảng
        p_r = cell_right.paragraphs[0]
        p_r.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r.paragraph_format.space_before = Pt(0)
        p_r.paragraph_format.space_after = Pt(2)
        p_r.paragraph_format.line_spacing = 1.1

        if standard_id == "HD36":
            r_party = p_r.add_run("ĐẢNG CỘNG SẢN VIỆT NAM")
            format_run_font(r_party, size_pt=12.5, bold=True)

            p_r_line = cell_right.add_paragraph()
            p_r_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_r_line.paragraph_format.space_before = Pt(1)
            p_r_line.paragraph_format.space_after = Pt(6)
            r_line = p_r_line.add_run("———————————————")
            format_run_font(r_line, size_pt=10.0, bold=True)
        else:
            r_c1 = p_r.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
            format_run_font(r_c1, size_pt=12.5, bold=True)
            r_c2 = p_r.add_run("Độc lập - Tự do - Hạnh phúc")
            format_run_font(r_c2, size_pt=13.0, bold=True)

            p_r_line = cell_right.add_paragraph()
            p_r_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_r_line.paragraph_format.space_before = Pt(1)
            p_r_line.paragraph_format.space_after = Pt(6)
            r_line = p_r_line.add_run("———————————————")
            format_run_font(r_line, size_pt=10.0, bold=True)

        # Địa danh, ngày tháng năm
        p_r_date = cell_right.add_paragraph()
        p_r_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r_date.paragraph_format.space_before = Pt(4)
        p_r_date.paragraph_format.space_after = Pt(0)
        loc = meta.get("location", "Hà Nội").strip() or "Hà Nội"
        date_str = meta.get("date_str", "").strip() or f"{loc}, ngày ... tháng ... năm 2026"
        r_date = p_r_date.add_run(date_str)
        format_run_font(r_date, size_pt=13.0, italic=True)

        if lock_scaffolding:
            cls.wrap_element_in_locked_sdt(header_table._tbl, "AdministrativeHeader")

        # 2. TIÊU ĐỀ LOẠI VĂN BẢN VÀ TRÍCH YẾU
        doc_type = meta.get("doc_type", "QUYẾT ĐỊNH").strip().upper()
        p_type = doc.add_paragraph()
        p_type.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_type.paragraph_format.space_before = Pt(14)
        p_type.paragraph_format.space_after = Pt(4)
        r_type = p_type.add_run(doc_type)
        format_run_font(r_type, size_pt=14.0, bold=True)

        doc_subject = meta.get("doc_subject", "").strip()
        if doc_subject:
            p_subj = doc.add_paragraph()
            p_subj.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_subj.paragraph_format.space_before = Pt(2)
            p_subj.paragraph_format.space_after = Pt(12)
            r_subj = p_subj.add_run(doc_subject)
            format_run_font(r_subj, size_pt=13.5, bold=True)

        # 3. THÂN VĂN BẢN (BODY PARAGRAPHS - VÙNG BIÊN TẬP MỞ)
        body_paras = meta.get("body_paragraphs", [])
        if not body_paras:
            # Mẫu mặc định nếu người dùng chưa điền
            if "QUYẾT ĐỊNH" in doc_type:
                body_paras = [
                    "Căn cứ Luật Tổ chức chính quyền địa phương ngày 19 tháng 6 năm 2015;",
                    "Theo đề nghị của Chánh Văn phòng,",
                    "QUYẾT ĐỊNH:",
                    "Điều 1. Ban hành kèm theo Quyết định này Kế hoạch công tác năm 2026.",
                    "Điều 2. Quyết định này có hiệu lực thi hành kể từ ngày ký.",
                    "Điều 3. Chánh Văn phòng, Thủ trưởng các cơ quan, đơn vị có liên quan chịu trách nhiệm thi hành Quyết định này."
                ]
            elif "NGHỊ QUYẾT" in doc_type:
                body_paras = [
                    "Căn cứ Điều lệ Đảng Cộng sản Việt Nam;",
                    "BAN CHẤP HÀNH ĐẢNG BỘ QUYẾT NGHỊ:",
                    "I. TÌNH HÌNH THỰC HIỆN NHIỆM VỤ",
                    "Đảng bộ đã lãnh đạo, chỉ đạo toàn diện các mặt công tác, đạt được nhiều kết quả quan trọng.",
                    "II. MỤC TIÊU VÀ NHIỆM VỤ TRỌNG TÂM",
                    "1. Tiếp tục đổi mới phương thức lãnh đạo của Đảng.",
                    "2. Nâng cao hiệu lực, hiệu quả quản lý nhà nước trên địa bàn.",
                    "III. TỔ CHỨC THỰC HIỆN",
                    "Các cấp ủy trực thuộc, các cơ quan ban ngành xây dựng kế hoạch cụ thể để triển khai thực hiện thắng lợi Nghị quyết."
                ]
            elif "CÔNG VĂN" in doc_type:
                body_paras = [
                    "Kính gửi: Các cơ quan, ban, ngành, đoàn thể.",
                    "Thực hiện ý kiến chỉ đạo của cấp có thẩm quyền về việc tăng cường công tác cải cách hành chính và chuyển đổi số;",
                    "Đề nghị các cơ quan, đơn vị phối hợp chặt chẽ, hoàn thành các nội dung báo cáo gửi về trước ngày 15 hàng tháng để tổng hợp.",
                    "Rất mong nhận được sự phối hợp của các đồng chí."
                ]
            elif "TỜ TRÌNH" in doc_type:
                body_paras = [
                    "Kính gửi: Lãnh đạo cấp có thẩm quyền.",
                    "I. SỰ CẦN THIẾT BAN HÀNH",
                    "Để đáp ứng yêu cầu thực tiễn và hoàn thiện cơ sở pháp lý cho công tác quản lý trên địa bàn;",
                    "II. NỘI DUNG CHÍNH CỦA ĐỀ ÁN",
                    "Đề án tập trung vào các giải pháp nâng cao hiệu quả làm việc và hiện đại hóa công sở.",
                    "Kính trình xem xét, phê duyệt."
                ]
            else: # Thông báo
                body_paras = [
                    "Căn cứ kết luận cuộc họp giao ban công tác;",
                    "THÔNG BÁO:",
                    "1. Thống nhất triển khai đồng bộ các nhiệm vụ trọng tâm quý tới.",
                    "2. Giao Văn phòng đôn đốc, theo dõi tiến độ thực hiện của các đơn vị."
                ]

        for p_text in body_paras:
            t_strip = p_text.strip()
            if not t_strip:
                continue

            # Kiểm tra ngắt trang thực tế (Không in dòng chữ ra văn bản Word)
            if any(k in t_strip.upper() for k in ["[PAGE_BREAK]", "PAGE BREAK", "NGẮT TRANG", "PAGE_BREAK"]):
                doc.add_page_break()
                continue

            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.line_spacing = 1.2
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            # Tiêu đề lệnh (QUYẾT ĐỊNH:, QUYẾT NGHỊ:, THÔNG BÁO:) -> Căn giữa, in đậm
            if t_strip.endswith(":") and any(k in t_strip.upper() for k in ["QUYẾT ĐỊNH", "QUYẾT NGHỊ", "THÔNG BÁO"]):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                p.paragraph_format.first_line_indent = Mm(0)
                p.paragraph_format.space_before = Pt(8)
                p.paragraph_format.space_after = Pt(6)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.5, bold=True)
            # Căn cứ pháp lý -> In nghiêng
            elif re.match(r'^[-–—•*]?\s*(căn\s+cứ|theo\s+đề\s+nghị)\b', t_strip, re.IGNORECASE):
                p.paragraph_format.first_line_indent = Mm(12.7)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.0, italic=True)
            # Tiêu đề La Mã / Số (I., II., 1., 2.)
            elif re.match(r'^[IVXLCDM]+\.\s+', t_strip) or re.match(r'^\d+\.\s+[A-ZÀ-Ỹ]', t_strip):
                p.paragraph_format.first_line_indent = Mm(12.7)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.0, bold=True)
            # Điều khoản (Điều 1., Điều 2.)
            elif re.match(r'^(điều|dieu)\s+\d+', t_strip, re.IGNORECASE):
                p.paragraph_format.first_line_indent = Mm(12.7)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.0, bold=True)
            # Kính gửi:
            elif t_strip.lower().startswith("kính gửi"):
                p.paragraph_format.first_line_indent = Mm(12.7)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.5, bold=True)
            else:
                p.paragraph_format.first_line_indent = Mm(12.7)
                r = p.add_run(t_strip)
                format_run_font(r, size_pt=13.0, bold=False)

        # 4. BẢNG KHUNG CUỐI VĂN BẢN (FOOTER TABLE: NƠI NHẬN & CHỮ KÝ)
        footer_table = doc.add_table(rows=1, cols=2)
        footer_table.alignment = WD_TABLE_ALIGNMENT.CENTER
        TableLayoutManager.make_table_borders_invisible(footer_table)
        TableLayoutManager.set_column_widths(footer_table, [65.0, 100.0])

        f_cell_left = footer_table.cell(0, 0)
        f_cell_right = footer_table.cell(0, 1)
        TableLayoutManager.set_cell_margins(f_cell_left, top=30, bottom=30, left=30, right=30)
        TableLayoutManager.set_cell_margins(f_cell_right, top=30, bottom=30, left=30, right=30)

        # Cột Trái Footer: Nơi nhận
        p_fl = f_cell_left.paragraphs[0]
        p_fl.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_fl.paragraph_format.space_before = Pt(0)
        p_fl.paragraph_format.space_after = Pt(2)
        r_rec_hdr = p_fl.add_run("Nơi nhận:\n")
        format_run_font(r_rec_hdr, size_pt=12.0, bold=True, italic=True)

        recipients = meta.get("recipients", [])
        if not recipients:
            if standard_id == "HD36":
                recipients = ["- Thường trực cấp ủy (b/c);", "- Các ban Đảng, đoàn thể;", "- Lưu: VP, Ban XD Đảng."]
            else:
                recipients = ["- Như Điều 3;", "- Chủ tịch, các PCT UBND;", "- Lưu: VT, TH."]

        for rec in recipients:
            p_item = f_cell_left.add_paragraph()
            p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_item.paragraph_format.space_before = Pt(0)
            p_item.paragraph_format.space_after = Pt(1)
            p_item.paragraph_format.first_line_indent = Mm(0)
            p_item.paragraph_format.line_spacing = 1.05
            r_it = p_item.add_run(rec)
            format_run_font(r_it, size_pt=11.0, italic=False)

        # Cột Phải Footer: Thẩm quyền ký, Chức vụ và Họ tên
        p_fr = f_cell_right.paragraphs[0]
        p_fr.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_fr.paragraph_format.space_before = Pt(0)
        p_fr.paragraph_format.space_after = Pt(2)
        p_fr.paragraph_format.line_spacing = 1.1

        auth_prefix = meta.get("authority_prefix", "").strip().upper()
        if not auth_prefix:
            auth_prefix = "T/M BAN THƯỜNG VỤ" if standard_id == "HD36" else "TM. ỦY BAN NHÂN DÂN"
        r_auth = p_fr.add_run(auth_prefix + "\n")
        format_run_font(r_auth, size_pt=13.0, bold=True)

        signer_pos = meta.get("signer_position", "").strip().upper()
        if not signer_pos:
            signer_pos = "BÍ THƯ" if standard_id == "HD36" else "CHỦ TỊCH"
        r_pos = p_fr.add_run(signer_pos + "\n\n\n\n")  # Khoảng trống ký 40-50pt
        format_run_font(r_pos, size_pt=13.0, bold=True)

        signer_name = meta.get("signer_name", "").strip() or "Nguyễn Văn A"
        r_name = p_fr.add_run(signer_name)
        format_run_font(r_name, size_pt=13.5, bold=True)

        if lock_scaffolding:
            cls.wrap_element_in_locked_sdt(footer_table._tbl, "AdministrativeFooter")

        return doc
