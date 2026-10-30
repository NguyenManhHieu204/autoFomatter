"""
core/doc_parser.py: Đọc, bóc tách và phân tích các thành phần cấu trúc của văn bản Word (.docx)
Bao gồm:
- Phát hiện & trích xuất Tiêu đề / Quốc hiệu / Cơ quan ban hành (Header)
- Phát hiện Tên loại văn bản & Trích yếu nội dung (Subject)
- Phát hiện các đoạn văn phần Nội dung chính (Body) & Bảng biểu (Tables)
- Phát hiện Nơi nhận & Khung quyền hạn ký (Footer)
"""

import os
import re
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict, Any
import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


from core.administrative_sanitizer import AdministrativeSanitizer


@dataclass
class ExtractedHeader:
    superior_agency: str = ""
    issuing_agency: str = ""
    doc_number: str = ""
    motto_country: str = ""
    motto_slogan: str = ""
    party_heading: str = ""
    location_date: str = ""
    has_existing_table: bool = False
    raw_paragraphs_count: int = 0


@dataclass
class ExtractedFooter:
    recipients: List[str] = field(default_factory=list)
    authority_prefix: str = ""
    signer_position: str = ""
    signer_name: str = ""
    has_existing_table: bool = False


@dataclass
class ParsedDocStructure:
    doc_path: str = ""             # Đường dẫn file gốc nếu có
    doc_type: str = ""             # Ví dụ: QUYẾT ĐỊNH, CÔNG VĂN, BÁO CÁO, KẾ HOẠCH...
    doc_subject: str = ""          # Trích yếu nội dung (Về việc...)
    header: ExtractedHeader = field(default_factory=ExtractedHeader)
    footer: ExtractedFooter = field(default_factory=ExtractedFooter)
    body_paragraphs_indices: List[int] = field(default_factory=list)
    title_paragraphs_indices: List[int] = field(default_factory=list)
    footer_para_indices: List[int] = field(default_factory=list)
    detected_standard: str = "ND30"
    total_paragraphs: int = 0
    total_tables: int = 0
    is_scanned_image: bool = False # Cờ nhận diện tài liệu dạng ảnh scan
    image_count: int = 0          # Số lượng hình ảnh scan / khung vẽ
    text_word_count: int = 0      # Tổng số từ có thể chỉnh sửa


DRAWING_XPATH = './/*[local-name()="drawing" or local-name()="pict" or local-name()="object" or local-name()="AlternateContent"]'


def paragraph_has_drawings_or_images(p: Paragraph) -> bool:
    """Kiểm tra đoạn văn có chứa hình ảnh, khung vẽ, đối tượng đồ họa không"""
    if p is None or p._element is None:
        return False
    drawings = p._element.xpath(DRAWING_XPATH)
    return len(drawings) > 0


def run_has_drawings_or_images(run) -> bool:
    """Kiểm tra một run có chứa hình ảnh hoặc khung vẽ không"""
    if run is None or run._r is None:
        return False
    drawings = run._r.xpath(DRAWING_XPATH)
    return len(drawings) > 0


DOC_TYPES_PATTERN = re.compile(
    r'(?:^|[\s\b])(NGHỊ QUYẾT|QUYẾT ĐỊNH|CHỈ THỊ|KẾT LUẬN|QUY CHẾ|QUY ĐỊNH|THÔNG TRI|'
    r'HƯỚNG DẪN|CÔNG VĂN|BÁO CÁO|TỜ TRÌNH|KẾ HOẠCH|THÔNG BÁO|BIÊN BẢN|GIẤY MỜI|'
    r'CHƯƠNG TRÌNH|ĐỀ ÁN|PHƯƠNG ÁN|QUYET DINH|NGHI QUYET|QUYET NGHI)\b',
    re.IGNORECASE
)

AUTHORITY_PREFIX_PATTERN = re.compile(
    r'^(TM\.|KT\.|TL\.|TUQ\.|T/M|K/T|T/L|T\.M\.|K\.T\.|T\.L\.)\s*',
    re.IGNORECASE
)

DATE_PATTERN = re.compile(
    r'([A-Za-zÀ-ỹ\s]+,)?\s*ngày\s+\d{1,2}\s+tháng\s+\d{1,2}\s+năm\s+\d{4}',
    re.IGNORECASE
)


class DocParser:
    def __init__(self, doc_path_or_doc: Any):
        if isinstance(doc_path_or_doc, str):
            from core.security_guard import SecurityGuard
            ok, msg = SecurityGuard.validate_docx_package(doc_path_or_doc)
            if not ok:
                raise ValueError(f"Tệp không vượt qua kiểm tra an toàn: {msg}")
            self.doc_path = doc_path_or_doc
            self.doc = docx.Document(doc_path_or_doc)
        else:
            self.doc_path = ""
            self.doc = doc_path_or_doc

    def parse(self) -> ParsedDocStructure:
        structure = ParsedDocStructure()
        structure.doc_path = self.doc_path
        structure.total_paragraphs = len(self.doc.paragraphs)
        structure.total_tables = len(self.doc.tables)

        # 0. Phát hiện tài liệu dạng ảnh scan (Scanned Document Detection - Tối ưu bộ nhớ)
        drawings = self.doc.element.xpath(DRAWING_XPATH)
        structure.image_count = len(drawings)

        text_parts = [p.text for p in self.doc.paragraphs if p.text]
        for tbl in self.doc.tables:
            for row in tbl.rows:
                for cell in row.cells:
                    if cell.text:
                        text_parts.append(cell.text)
        structure.text_word_count = sum(len(part.split()) for part in text_parts)

        if structure.image_count > 0 and structure.text_word_count < 30:
            structure.is_scanned_image = True

        full_text = "\n".join(text_parts)
        structure.detected_standard = AdministrativeSanitizer.detect_document_standard(full_text)

        # 1. Kiểm tra bảng đầu tiên xem có phải Header Table không
        self._analyze_header(structure)

        # 2. Tìm Tên loại văn bản & Trích yếu
        self._analyze_title_and_subject(structure)

        # 3. Phân tích phần Chân trang (Nơi nhận & Chữ ký)
        self._analyze_footer(structure)

        # 4. Xác định danh sách các đoạn văn Body
        header_para_max = max(structure.title_paragraphs_indices) if structure.title_paragraphs_indices else structure.header.raw_paragraphs_count
        body_indices = []
        for i in range(len(self.doc.paragraphs)):
            # Bỏ qua các đoạn thuộc header hoặc title đã nhận diện
            if i in structure.title_paragraphs_indices:
                continue
            if structure.header.raw_paragraphs_count > 0 and i < structure.header.raw_paragraphs_count:
                continue
            # Bỏ qua các đoạn thuộc footer nếu được nhận diện dạng paragraph
            if self._is_footer_paragraph(i, structure):
                continue
            body_indices.append(i)

        structure.body_paragraphs_indices = body_indices
        return structure

    def _analyze_header(self, structure: ParsedDocStructure):
        """
        Kiểm tra nếu văn bản đã có Header dạng bảng hoặc dạng các đoạn văn đầu tiên.
        """
        if self.doc.tables:
            first_table = self.doc.tables[0]
            # Nếu bảng có 2 cột và nằm ở đầu văn bản (trước khi có nhiều paragraph text)
            if len(first_table.columns) == 2 and len(first_table.rows) >= 1:
                col0_text = "\n".join([cell.text for cell in first_table.columns[0].cells]).strip()
                col1_text = "\n".join([cell.text for cell in first_table.columns[1].cells]).strip()

                is_gov_header = any(k in col1_text.upper() for k in ["CỘNG HÒA XÃ HỘI", "CONG HOA XA HOI", "ĐỘC LẬP", "DOC LAP"])
                is_party_header = any(k in (col0_text + " " + col1_text).upper() for k in ["ĐẢNG CỘNG SẢN", "DANG CONG SAN", "ĐẢNG BỘ", "DANG BO", "TỈNH ỦY", "TINH UY"])

                if is_gov_header or is_party_header:
                    structure.header.has_existing_table = True
                    self._extract_from_header_table(first_table, structure.header)
                    return

        # Nếu không có bảng, quét các đoạn văn đầu tiên (tối đa 8 đoạn)
        header_para_count = 0
        for i, p in enumerate(self.doc.paragraphs[:8]):
            text = p.text.strip()
            text_u = text.upper()
            if not text:
                header_para_count += 1
                continue

            # Nếu gặp Căn cứ, Điều, hoặc gạch đầu dòng -> Đã vào phần Thân bài, DỪNG QUÉT HEADER!
            if re.match(r'^[-–—•*]?\s*(căn\s+cứ|can\s+cur|can\s+cu|theo\s+đề\s+nghị|theo\s+de\s+nghi|điều|dieu)\b', text, re.IGNORECASE):
                break

            if "CỘNG HÒA XÃ HỘI" in text_u or "CONG HOA XA HOI" in text_u:
                structure.header.motto_country = text
                header_para_count = i + 1
            elif "ĐỘC LẬP" in text_u or "DOC LAP" in text_u:
                structure.header.motto_slogan = text
                header_para_count = i + 1
            elif "ĐẢNG CỘNG SẢN" in text_u or "DANG CONG SAN" in text_u:
                structure.header.party_heading = text
                header_para_count = i + 1
            elif text_u.startswith("SỐ") or "Số:" in text or "Số " in text or text_u.startswith("SO ") or text_u.startswith("S6 "):
                structure.header.doc_number = text
                header_para_count = i + 1
            elif DATE_PATTERN.search(text):
                structure.header.location_date = text
                header_para_count = i + 1
            elif DOC_TYPES_PATTERN.search(text):
                # Gặp tên loại văn bản -> Dừng quét header
                break
            else:
                # Chỉ coi là cơ quan ban hành nếu dòng ngắn (dưới 10 từ) và không chứa dấu chấm phẩy
                if len(text.split()) <= 10 and ";" not in text and not text.startswith("-"):
                    if not structure.header.superior_agency:
                        structure.header.superior_agency = text
                    elif not structure.header.issuing_agency:
                        structure.header.issuing_agency = text
                    header_para_count = i + 1
                else:
                    break

        structure.header.raw_paragraphs_count = header_para_count

    def _extract_from_header_table(self, table: Table, header: ExtractedHeader):
        cell_0 = table.cell(0, 0)
        cell_1 = table.cell(0, 1)

        lines_0 = [line.strip() for line in cell_0.text.split("\n") if line.strip()]
        lines_1 = [line.strip() for line in cell_1.text.split("\n") if line.strip()]

        for line in lines_0:
            if line.upper().startswith("SỐ") or "Số:" in line or "Số " in line:
                header.doc_number = line
            elif not header.superior_agency:
                header.superior_agency = line
            elif not header.issuing_agency:
                header.issuing_agency = line

        for line in lines_1:
            if "CỘNG HÒA XÃ HỘI" in line.upper():
                header.motto_country = line
            elif "ĐỘC LẬP" in line.upper():
                header.motto_slogan = line
            elif "ĐẢNG CỘNG SẢN" in line.upper():
                header.party_heading = line
            elif DATE_PATTERN.search(line):
                header.location_date = line

    def _analyze_title_and_subject(self, structure: ParsedDocStructure):
        """
        Xác định đoạn văn Tên loại văn bản (QUYẾT ĐỊNH, CÔNG VĂN...) và Trích yếu nội dung
        """
        start_idx = structure.header.raw_paragraphs_count
        search_limit = min(start_idx + 6, len(self.doc.paragraphs))

        for i in range(start_idx, search_limit):
            p = self.doc.paragraphs[i]
            text = p.text.strip()
            match = DOC_TYPES_PATTERN.match(text)
            if match:
                structure.doc_type = match.group(0).upper()
                structure.title_paragraphs_indices.append(i)

                # Kiểm tra đoạn tiếp theo có phải trích yếu không (Về việc...)
                if i + 1 < len(self.doc.paragraphs):
                    next_text = self.doc.paragraphs[i + 1].text.strip()
                    if next_text.startswith("Về việc") or next_text.startswith("về việc") or "V/v" in next_text:
                        structure.doc_subject = next_text
                        structure.title_paragraphs_indices.append(i + 1)
                break

    def _analyze_footer(self, structure: ParsedDocStructure):
        """
        Xác định phần Nơi nhận và Quyền hạn ký (Chân văn bản)
        """
        # Kiểm tra bảng cuối cùng xem có phải bảng Footer không
        if self.doc.tables:
            last_table = self.doc.tables[-1]
            if len(last_table.columns) == 2 and len(last_table.rows) >= 1:
                left_text = "\n".join([c.text for c in last_table.columns[0].cells]).strip()
                right_text = "\n".join([c.text for c in last_table.columns[1].cells]).strip()

                if any(k in left_text.upper() for k in ["NƠI NHẬN", "NOI NHAN"]):
                    structure.footer.has_existing_table = True
                    self._extract_from_footer_table(last_table, structure.footer)
                    return

        # Nếu không có bảng cuối, kiểm tra xem footer có nằm trong các đoạn văn không
        paras_text = [p.text for p in self.doc.paragraphs]
        footer_info = AdministrativeSanitizer.parse_in_text_footer(paras_text)
        if footer_info:
            structure.footer.recipients = footer_info["recipients"]
            structure.footer.authority_prefix = footer_info["authority_prefix"]
            structure.footer.signer_position = footer_info["signer_position"]
            structure.footer.signer_name = footer_info["signer_name"]
            structure.footer_para_indices = list(range(footer_info["start_idx"], footer_info["end_idx"]))

    def _extract_from_footer_table(self, table: Table, footer: ExtractedFooter):
        cell_left = table.cell(0, 0)
        cell_right = table.cell(0, 1)

        # Nơi nhận
        for line in cell_left.text.split("\n"):
            line = line.strip()
            if line and not line.startswith("Nơi nhận") and not line.startswith("NƠI NHẬN"):
                footer.recipients.append(line)

        # Chữ ký & quyền hạn
        right_lines = [l.strip() for l in cell_right.text.split("\n") if l.strip()]
        for line in right_lines:
            if AUTHORITY_PREFIX_PATTERN.match(line):
                footer.authority_prefix = line
            elif not footer.signer_position and line.isupper():
                footer.signer_position = line
            else:
                footer.signer_name = line

    def _is_footer_paragraph(self, idx: int, structure: ParsedDocStructure) -> bool:
        # Nếu đã có footer table thì các paragraph bình thường không phải footer
        if structure.footer.has_existing_table:
            return False
        if structure.footer_para_indices and idx in structure.footer_para_indices:
            return True
        total = len(self.doc.paragraphs)
        if idx >= total - 6:
            text = self.doc.paragraphs[idx].text.strip()
            if text.startswith("Nơi nhận") or text.startswith("- ") or text.startswith("TM.") or text.startswith("KT."):
                return True
        return False
