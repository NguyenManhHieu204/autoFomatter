"""
Module 2 & 3: Tự động Dựng Khung Đầu Văn bản & Khung Cuối Văn bản
(Header Table & Footer Signature / Recipients Builder)
Sử dụng cấu trúc Bảng ẩn viền (Invisible Border Table - 2 cột) để định vị tuyệt đối:
- Khối Đảng: Hướng dẫn 36-HD/VPTW
- Khối Hành chính Nhà nước: Nghị định 30/2020/NĐ-CP
"""

import re
from typing import List, Dict, Any, Optional
import docx
from docx.shared import Pt, Mm, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


def set_cell_margins(cell, top=0, bottom=0, left=60, right=60):
    """Thiết lập padding cho ô bảng để nội dung gọn gàng, chuẩn xác"""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>\n'
        f'  <w:top w:w="{top}" w:type="dxa"/>\n'
        f'  <w:left w:w="{left}" w:type="dxa"/>\n'
        f'  <w:bottom w:w="{bottom}" w:type="dxa"/>\n'
        f'  <w:right w:w="{right}" w:type="dxa"/>\n'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def make_table_borders_invisible(table):
    """Xóa toàn bộ đường viền hiển thị của bảng (Bảng ẩn viền)"""
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:left w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:bottom w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:right w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'  <w:insideV w:val="none" w:sz="0" w:space="0" w:color="auto"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)


def format_run(run, font_name="Times New Roman", size_pt=13.0, bold=False, italic=False, underline=False):
    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.bold = bold
    run.font.italic = italic
    run.font.underline = underline
    # Đảm bảo phông chữ áp dụng cho toàn bộ bảng mã Unicode TCVN một cách tối ưu
    rPr = run._r.get_or_add_rPr()
    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is None:
        rFonts = OxmlElement('w:rFonts')
        rPr.append(rFonts)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rFonts.set(qn('w:eastAsia'), font_name)
    rFonts.set(qn('w:cs'), font_name)


class HeaderFooterBuilder:
    def __init__(self, rules_hd36: Dict[str, Any], rules_nd30: Dict[str, Any]):
        self.rules_hd36 = rules_hd36
        self.rules_nd30 = rules_nd30

    def get_rules(self, standard_id: str) -> Dict[str, Any]:
        return self.rules_hd36 if standard_id.upper() == "HD36" else self.rules_nd30

    def standardize_authority_prefix(self, raw_prefix: str, standard_id: str) -> str:
        """
        Chuẩn hóa từ viết tắt quyền hạn ký theo quy định:
        - Khối Đảng (HD 36): Có dấu gạch chéo (T/M, K/T, T/L)
        - Khối Nhà nước (NĐ 30): Có dấu chấm, không có gạch chéo (TM., KT., TL., TUQ.)
        """
        if not raw_prefix:
            return ""

        def _standardize_single(prefix_line: str) -> str:
            line = prefix_line.strip().upper()
            if not line:
                return ""
            if standard_id.upper() == "ND30":
                line = re.sub(r'\bT\s*[\./]?\s*M\s*\.?\b|\bTM\s*\.?\b', 'TM.', line, flags=re.IGNORECASE)
                line = re.sub(r'\bK\s*[\./]?\s*T\s*\.?\b|\bKT\s*\.?\b', 'KT.', line, flags=re.IGNORECASE)
                line = re.sub(r'\bT\s*[\./]?\s*L\s*\.?\b|\bTL\s*\.?\b', 'TL.', line, flags=re.IGNORECASE)
                line = re.sub(r'\bT\s*[\./]?\s*U\s*[\./]?\s*Q\s*\.?\b|\bTUQ\s*\.?\b', 'TUQ.', line, flags=re.IGNORECASE)
                line = re.sub(r'\.{2,}', '.', line)
                line = re.sub(r'^(TM\.|KT\.|TL\.|TUQ\.)\s*', r'\1 ', line)
                line = re.sub(r'[ \t]+', ' ', line).strip()
            else: # HD36
                line = re.sub(r'\bT\s*[\./]?\s*M\s*\.?\b|\bTM\s*\.?\b', 'T/M', line, flags=re.IGNORECASE)
                line = re.sub(r'\bK\s*[\./]?\s*T\s*\.?\b|\bKT\s*\.?\b', 'K/T', line, flags=re.IGNORECASE)
                line = re.sub(r'\bT\s*[\./]?\s*L\s*\.?\b|\bTL\s*\.?\b', 'T/L', line, flags=re.IGNORECASE)
                line = re.sub(r'\bT\s*[\./]?\s*U\s*[\./]?\s*Q\s*\.?\b|\bTUQ\s*\.?\b', 'TUQ', line, flags=re.IGNORECASE)
                line = re.sub(r'^(T/M|K/T|T/L|TUQ)\s*', r'\1 ', line)
                line = re.sub(r'[ \t]+', ' ', line).strip()
            return line

        lines = [_standardize_single(l) for l in raw_prefix.split("\n")]
        return "\n".join([l for l in lines if l])

    def build_header_table(
        self,
        table: docx.table.Table,
        standard_id: str,
        superior_agency: str,
        issuing_agency: str,
        doc_number: str,
        location: str,
        date_str: str
    ):
        """
        Điền và định dạng dữ liệu vào Bảng Khung Đầu Văn bản (Module 2)
        """
        rules = self.get_rules(standard_id)
        from core.administrative_sanitizer import AdministrativeSanitizer
        superior_agency, issuing_agency = AdministrativeSanitizer.split_and_normalize_agency_header(
            superior_agency, issuing_agency
        )
        make_table_borders_invisible(table)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        # Đặt độ rộng 2 cột (Tổng 165mm trên trang A4)
        col_left_width = Mm(68.0)
        col_right_width = Mm(97.0)
        for row in table.rows:
            row.cells[0].width = col_left_width
            row.cells[1].width = col_right_width
            set_cell_margins(row.cells[0], top=40, bottom=40, left=40, right=40)
            set_cell_margins(row.cells[1], top=40, bottom=40, left=40, right=40)

        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)

        # Xóa các đoạn văn mặc định trong ô
        for p in cell_left.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        for p in cell_right.paragraphs[1:]:
            p._element.getparent().remove(p._element)

        p_l1 = cell_left.paragraphs[0]
        p_l1.text = ""
        p_l1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l1.paragraph_format.space_before = Pt(0)
        p_l1.paragraph_format.space_after = Pt(2)
        p_l1.paragraph_format.line_spacing = 1.1

        # Cột Trái 1: Cơ quan chủ quản (nếu có)
        if superior_agency:
            r_sup = p_l1.add_run(superior_agency.upper().strip() + "\n")
            format_run(
                r_sup,
                size_pt=rules["header"]["superior_agency_font_size_pt"],
                bold=rules["header"]["superior_agency_bold"]
            )

        # Cột Trái 2: Cơ quan ban hành (In hoa, đậm)
        issuing_text = issuing_agency.upper().strip() if issuing_agency else "CƠ QUAN BAN HÀNH"
        r_iss = p_l1.add_run(issuing_text)
        format_run(
            r_iss,
            size_pt=rules["header"]["issuing_agency_font_size_pt"],
            bold=rules["header"]["issuing_agency_bold"]
        )

        # Cột Trái 3: Đường kẻ ngang dưới tên cơ quan ban hành (khoảng 1/3 đến 1/2 chiều dài chữ)
        p_l_divider = cell_left.add_paragraph()
        p_l_divider.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l_divider.paragraph_format.space_before = Pt(1)
        p_l_divider.paragraph_format.space_after = Pt(4)
        p_l_divider.paragraph_format.line_spacing = 1.0

        if standard_id.upper() == "HD36":
            # Đảng: đường kẻ ngắn hoặc dấu sao / gạch ngang
            r_div = p_l_divider.add_run("———*———")
        else:
            # Nhà nước: đường kẻ liền dài 1/3 - 1/2 độ dài dòng chữ
            r_div = p_l_divider.add_run("—————")
        format_run(r_div, size_pt=9.0, bold=True)

        # Cột Trái 4: Số, ký hiệu văn bản (NĐ 30 dùng 'Số: .../...', HĐ 36 dùng 'Số ...-.../...')
        p_l_num = cell_left.add_paragraph()
        p_l_num.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_l_num.paragraph_format.space_before = Pt(2)
        p_l_num.paragraph_format.space_after = Pt(0)

        num_prefix = rules["header"]["doc_number_prefix"]
        if not doc_number:
            if standard_id.upper() == "HD36":
                doc_number = f"{num_prefix}...-QĐ/..."
            else:
                doc_number = f"{num_prefix}.../QĐ-..."
        else:
            clean_num = re.sub(r'^(số|so)\s*[:\.]?\s*', '', doc_number.strip(), flags=re.IGNORECASE)
            clean_num = re.sub(r'\s*/\s*', '/', clean_num)
            clean_num = re.sub(r'\s*-\s*', '-', clean_num)
            doc_number = f"{num_prefix}{clean_num}"

        r_num = p_l_num.add_run(doc_number)
        format_run(r_num, size_pt=rules["header"]["doc_number_font_size_pt"], bold=False)

        # ================== CỘT PHẢI ==================
        p_r1 = cell_right.paragraphs[0]
        p_r1.text = ""
        p_r1.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r1.paragraph_format.space_before = Pt(0)
        p_r1.paragraph_format.space_after = Pt(2)
        p_r1.paragraph_format.line_spacing = 1.1

        if standard_id.upper() == "HD36":
            # Khối Đảng: "ĐẢNG CỘNG SẢN VIỆT NAM"
            r_party = p_r1.add_run("ĐẢNG CỘNG SẢN VIỆT NAM")
            format_run(
                r_party,
                size_pt=rules["header"]["party_heading_font_size_pt"],
                bold=True
            )

            # Đường kẻ liền bằng đúng độ dài dòng chữ
            p_r_line = cell_right.add_paragraph()
            p_r_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_r_line.paragraph_format.space_before = Pt(1)
            p_r_line.paragraph_format.space_after = Pt(6)
            r_line = p_r_line.add_run("———————————————")
            format_run(r_line, size_pt=10.0, bold=True)

        else:
            # Khối Hành chính Nhà nước:
            # Dòng 1: CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM (12-13 pt, đậm)
            r_c1 = p_r1.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
            format_run(
                r_c1,
                size_pt=rules["header"]["motto_country_font_size_pt"],
                bold=True
            )

            # Dòng 2: Độc lập - Tự do - Hạnh phúc (13-14 pt, đậm)
            r_c2 = p_r1.add_run("Độc lập - Tự do - Hạnh phúc")
            format_run(
                r_c2,
                size_pt=rules["header"]["motto_slogan_font_size_pt"],
                bold=True
            )

            # Đường kẻ liền bằng đúng độ dài dòng chữ Tiêu ngữ
            p_r_line = cell_right.add_paragraph()
            p_r_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p_r_line.paragraph_format.space_before = Pt(1)
            p_r_line.paragraph_format.space_after = Pt(6)
            r_line = p_r_line.add_run("———————————————")
            format_run(r_line, size_pt=10.0, bold=True)

        # Cột Phải: Địa danh và ngày tháng năm
        p_r_date = cell_right.add_paragraph()
        p_r_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_r_date.paragraph_format.space_before = Pt(4)
        p_r_date.paragraph_format.space_after = Pt(0)

        from core.administrative_sanitizer import AdministrativeSanitizer

        # 1. Trích xuất địa danh và chuỗi ngày tháng
        extracted_loc, raw_date = AdministrativeSanitizer.extract_location_and_date(date_str)
        loc = location.strip() if location else ""
        if extracted_loc:
            loc = extracted_loc
        if not loc or loc in ["...", "…"]:
            loc = AdministrativeSanitizer.infer_location_from_agencies(issuing_agency, superior_agency)

        final_date = raw_date if raw_date else (date_str.strip() if date_str else "")
        if not final_date:
            final_date = "ngày ... tháng ... năm 2026"

        # Chuẩn hóa ngày tháng nếu đã có số cụ thể (ngày < 10 có 0, tháng 1, 2 có 0 per NĐ 30)
        final_date = AdministrativeSanitizer.format_date_nd30(final_date)

        if loc and loc not in ["...", "…"]:
            if final_date.lower().startswith(loc.lower()):
                full_date_text = final_date
            else:
                full_date_text = f"{loc}, {final_date}"
        else:
            if final_date.startswith("...") or final_date.startswith("…"):
                full_date_text = final_date
            else:
                full_date_text = f"..., {final_date}"

        full_date_text = re.sub(r'\s+,', ',', full_date_text)
        r_date = p_r_date.add_run(full_date_text)
        format_run(r_date, size_pt=rules["header"]["date_font_size_pt"], italic=True)

    def build_footer_table(
        self,
        table: docx.table.Table,
        standard_id: str,
        recipients_list: List[str],
        authority_prefix: str,
        signer_position: str,
        signer_name: str
    ):
        """
        Điền và định dạng dữ liệu vào Bảng Khung Cuối Văn bản (Nơi nhận & Chữ ký - Module 3)
        """
        rules = self.get_rules(standard_id)
        make_table_borders_invisible(table)
        table.alignment = WD_TABLE_ALIGNMENT.CENTER

        col_left_width = Mm(72.0)
        col_right_width = Mm(93.0)
        for row in table.rows:
            row.cells[0].width = col_left_width
            row.cells[1].width = col_right_width
            set_cell_margins(row.cells[0], top=20, bottom=20, left=30, right=30)
            set_cell_margins(row.cells[1], top=20, bottom=20, left=30, right=30)
            try:
                trPr = row._tr.get_or_add_trPr()
                if trPr.find(qn('w:cantSplit')) is None:
                    from docx.oxml import parse_xml
                    from docx.oxml.ns import nsdecls
                    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            except Exception:
                pass

        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)

        # Xóa các đoạn văn cũ
        for p in cell_left.paragraphs[1:]:
            p._element.getparent().remove(p._element)
        for p in cell_right.paragraphs[1:]:
            p._element.getparent().remove(p._element)

        # ================== CỘT TRÁI: NƠI NHẬN ==================
        p_rec_head = cell_left.paragraphs[0]
        p_rec_head.text = ""
        p_rec_head.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p_rec_head.paragraph_format.space_before = Pt(4)
        p_rec_head.paragraph_format.space_after = Pt(2)
        p_rec_head.paragraph_format.line_spacing = 1.0

        r_rh = p_rec_head.add_run(rules["recipients"]["header_text"])
        format_run(
            r_rh,
            size_pt=rules["recipients"]["header_font_size_pt"],
            bold=rules["recipients"]["header_bold"],
            italic=rules["recipients"]["header_italic"]
        )

        # Các dòng nơi nhận
        items = [item.strip() for item in recipients_list if item.strip()]
        if not items:
            items = [
                "- Như trên,",
                rules["recipients"]["default_last_item"]
            ]
        else:
            # Đảm bảo có dòng lưu trữ văn thư
            has_archive = any("Lưu" in it or "VT" in it for it in items)
            if not has_archive:
                items.append(rules["recipients"]["default_last_item"])

        try:
            from core.spell_linter import SpellLinter
            _linter = SpellLinter()
            items = [_linter.lint_text(it, standard_id=standard_id).cleaned_text for it in items]
        except Exception:
            pass

        for item in items:
            p_item = cell_left.add_paragraph()
            p_item.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p_item.paragraph_format.space_before = Pt(0)
            p_item.paragraph_format.space_after = Pt(rules["recipients"]["item_space_after_pt"])
            p_item.paragraph_format.line_spacing = rules["recipients"]["item_line_spacing"]

            # Định dạng gạch đầu dòng
            line_text = item if item.startswith("-") else f"- {item}"
            r_it = p_item.add_run(line_text)
            format_run(r_it, size_pt=rules["recipients"]["item_font_size_pt"])

        # ================== CỘT PHẢI: QUYỀN HẠN & CHỮ KÝ ==================
        p_auth = cell_right.paragraphs[0]
        p_auth.text = ""
        p_auth.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_auth.paragraph_format.space_before = Pt(4)
        p_auth.paragraph_format.space_after = Pt(2)
        p_auth.paragraph_format.line_spacing = 1.1

        clean_auth = self.standardize_authority_prefix(authority_prefix, standard_id)
        if clean_auth:
            r_auth = p_auth.add_run(clean_auth + "\n")
            format_run(r_auth, size_pt=rules["signing"]["position_font_size_pt"], bold=True)

        # Chức vụ người ký (chỉ thêm nếu chưa xuất hiện ở dòng cuối cùng của clean_auth)
        pos_text = signer_position.strip().upper() if signer_position else ""
        if not pos_text and not clean_auth:
            pos_text = "CHỦ TỊCH"
        if pos_text and not clean_auth.endswith(pos_text):
            r_pos = p_auth.add_run(pos_text)
            format_run(r_pos, size_pt=rules["signing"]["position_font_size_pt"], bold=True)

        # Khoảng trống chừa ký tay và đóng dấu (~28-30mm khoảng trống ký chuẩn NĐ 30, tối ưu không tràn trang)
        p_gap = cell_right.add_paragraph()
        p_gap.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_gap.paragraph_format.space_before = Pt(22)
        p_gap.paragraph_format.space_after = Pt(0)
        p_gap.paragraph_format.line_spacing = 1.0

        # Họ và tên người ký
        p_name = cell_right.add_paragraph()
        p_name.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p_name.paragraph_format.space_before = Pt(0)
        p_name.paragraph_format.space_after = Pt(0)

        # Đảm bảo khối chữ ký không bị tách đôi qua trang (Keep with next & CantSplit)
        p_auth.paragraph_format.keep_with_next = True
        p_gap.paragraph_format.keep_with_next = True
        p_name.paragraph_format.keep_with_next = True

        for row in table.rows:
            try:
                from docx.oxml import parse_xml
                from docx.oxml.ns import nsdecls
                trPr = row._tr.get_or_add_trPr()
                trPr.append(parse_xml(r'<w:cantSplit %s/>' % nsdecls('w')))
            except Exception:
                pass

        name_text = " ".join([w.capitalize() for w in signer_name.strip().split()]) if signer_name else "Nguyễn Văn A"
        r_name = p_name.add_run(name_text)
        format_run(
            r_name,
            size_pt=rules["signing"]["signer_name_font_size_pt"],
            bold=rules["signing"]["signer_name_bold"]
        )
