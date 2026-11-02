"""
core/editor_core/table_layout_manager.py: Quản trị Bố cục Bảng biểu Chuẩn Thể thức
- Tính toán kích thước cột theo hệ mét (Millimeters -> Twips/Dxa)
- Định dạng bảng ẩn viền cho Quốc hiệu - Tên cơ quan (Header) và Nơi nhận - Chữ ký (Footer)
- Định dạng bảng biểu nội dung chuẩn (Viền khung mảnh, lặp tiêu đề trang tblHeader, chống gãy dòng cantSplit)
"""

from typing import List, Optional, Tuple, Any
import docx
from docx.shared import Mm, Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn


class TableLayoutManager:
    # Kích thước in ấn chuẩn A4 (210mm x 297mm)
    PAGE_WIDTH_MM = 210.0
    PAGE_HEIGHT_MM = 297.0
    
    # Lề in chuẩn NĐ 30: Trái 30mm, Phải 15mm -> Chiều rộng khả dụng = 165mm
    PRINTABLE_WIDTH_MM = 165.0

    @classmethod
    def make_table_borders_invisible(cls, table: docx.table.Table):
        """Xóa toàn bộ viền của bảng (dùng cho Khung Tiêu đề và Khung Chữ ký)"""
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

    @classmethod
    def apply_standard_grid_borders(cls, table: docx.table.Table, border_color: str = "000000", sz: str = "4"):
        """Gán viền đen mảnh chuẩn văn thư hành chính (1/2 pt = sz="4") cho toàn bộ bảng"""
        tblPr = table._tbl.tblPr
        borders = parse_xml(
            f'<w:tblBorders {nsdecls("w")}>\n'
            f'  <w:top w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'  <w:left w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'  <w:bottom w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'  <w:right w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'  <w:insideH w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'  <w:insideV w:val="single" w:sz="{sz}" w:space="0" w:color="{border_color}"/>\n'
            f'</w:tblBorders>'
        )
        tblPr.append(borders)

    @classmethod
    def set_cell_margins(cls, cell: docx.table._Cell, top: int = 40, bottom: int = 40, left: int = 60, right: int = 60):
        """Thiết lập padding đệm bên trong ô bảng (đơn vị dxa, 20 dxa = 1 pt)"""
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

    @classmethod
    def apply_multipage_safety(cls, table: docx.table.Table):
        """
        Bảo vệ bảng biểu nhiều trang:
        - Lặp lại dòng tiêu đề (tblHeader)
        - Không cắt đôi hàng qua trang (cantSplit)
        """
        if len(table.rows) > 0:
            try:
                trPr0 = table.rows[0]._tr.get_or_add_trPr()
                if trPr0.find(qn('w:tblHeader')) is None:
                    trPr0.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
            except Exception:
                pass

        for row in table.rows:
            try:
                trPr = row._tr.get_or_add_trPr()
                if trPr.find(qn('w:cantSplit')) is None:
                    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
            except Exception:
                pass

    @classmethod
    def set_column_widths(cls, table: docx.table.Table, widths_mm: List[float]):
        """Thiết lập độ rộng cho từng cột theo milimét chuẩn"""
        table.autofit = False
        for row in table.rows:
            for idx, w_mm in enumerate(widths_mm):
                if idx < len(row.cells):
                    row.cells[idx].width = Mm(w_mm)
                    # Gán trực tiếp tcW dxa để tương thích tuyệt đối MS Word
                    tcPr = row.cells[idx]._tc.get_or_add_tcPr()
                    tcW = parse_xml(f'<w:tcW {nsdecls("w")} w:w="{int(w_mm * 56.6929)}" w:type="dxa"/>')
                    tcPr.append(tcW)
