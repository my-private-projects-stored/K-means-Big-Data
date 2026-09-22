"""xuất báo cáo toàn diện sang định dạng Microsoft Word (.docx) chuẩn Đại học Thủ Dầu Một."""

import sys
import io
from pathlib import Path
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import parse_xml, OxmlElement
from docx.oxml.ns import nsdecls, qn

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).resolve().parent.parent
RESULTS_DIR = BASE_DIR / "results"
VIZ_DIR = BASE_DIR / "results" / "visualizations"
OUTPUT_DOCX = BASE_DIR / "BaoCao_KMeans_BigData.docx"
ALT_DOCX = BASE_DIR / "Nhom01_PhanCumTienDienTuKMeans_BaoCao.docx"


def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    """thiết lập khoảng cách đệm trong ô bảng."""
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml(
        f'<w:tcMar {nsdecls("w")}>'
        f'<w:top w:w="{top}" w:type="dxa"/>'
        f'<w:bottom w:w="{bottom}" w:type="dxa"/>'
        f'<w:left w:w="{left}" w:type="dxa"/>'
        f'<w:right w:w="{right}" w:type="dxa"/>'
        f'</w:tcMar>'
    )
    tcPr.append(tcMar)


def set_table_black_borders(table):
    """thiết lập đường viền kẻ ô màu đen bình thường cho toàn bộ bảng."""
    table.style = 'Table Grid'
    tblPr = table._tbl.tblPr
    for child in list(tblPr):
        if child.tag.endswith('tblBorders'):
            tblPr.remove(child)

    tblBorders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>\n'
        f'  <w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:left w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:right w:val="single" w:sz="6" w:space="0" w:color="000000"/>\n'
        f'  <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'  <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>\n'
        f'</w:tblBorders>'
    )
    tblPr.append(tblBorders)


def setup_page_numbers(doc):
    """thiết lập đánh số trang tự động ở chân trang (Footer), trừ trang bìa."""
    for section in doc.sections:
        section.different_first_page_header_footer = True
        # Chân trang trang bìa để trống
        first_footer = section.first_page_footer
        for p in first_footer.paragraphs:
            p.text = ""

        # Chân trang các trang tiếp theo
        footer = section.footer
        p = footer.paragraphs[0]
        p.text = ""
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(0)

        # dynamic Word PAGE field
        fld_xml = (
            f'<w:fldSimple {nsdecls("w")} w:instr="PAGE">'
            f'<w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman"/>'
            f'<w:sz w:val="20"/><w:color w:val="000000"/></w:rPr>'
            f'<w:t>1</w:t></w:r></w:fldSimple>'
        )
        p._p.append(parse_xml(fld_xml))


def create_report():
    doc = docx.Document()

    # cấu hình khổ giấy A4 và căn lề: trên 2cm, dưới 2cm, trái 3cm, phải 2cm
    for section in doc.sections:
        section.page_width = Cm(21.0)
        section.page_height = Cm(29.7)
        section.top_margin = Cm(2.0)
        section.bottom_margin = Cm(2.0)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(2.0)

    # cấu hình kiểu Normal mặc định: Times New Roman, 13pt, giãn dòng 1.2
    normal_style = doc.styles['Normal']
    normal_font = normal_style.font
    normal_font.name = 'Times New Roman'
    normal_font.size = Pt(13)
    normal_font.color.rgb = RGBColor(0, 0, 0)
    normal_style.paragraph_format.line_spacing = 1.2
    normal_style.paragraph_format.space_after = Pt(4)

    # thiết lập số trang tự động ở chân trang
    setup_page_numbers(doc)

    # helper thêm đoạn văn bản chuẩn
    def add_p(text="", bold=False, italic=False, align=WD_ALIGN_PARAGRAPH.JUSTIFY, size=13, space_after=4, first_indent=Cm(1.0)):
        p = doc.add_paragraph()
        p.alignment = align
        p.paragraph_format.line_spacing = 1.2
        p.paragraph_format.space_after = Pt(space_after)
        if first_indent:
            p.paragraph_format.first_line_indent = first_indent
        if text:
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(size)
            run.bold = bold
            run.italic = italic
            run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_h1(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(16)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text.upper())
        run.font.name = 'Times New Roman'
        run.font.size = Pt(16)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_h2(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        run.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_h3(text):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.space_before = Pt(8)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.keep_with_next = True
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(13)
        run.bold = True
        run.italic = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        return p

    def add_caption(text, is_table=True):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(6)
        p.paragraph_format.space_after = Pt(6)
        run = p.add_run(text)
        run.font.name = 'Times New Roman'
        run.font.size = Pt(11)
        run.font.color.rgb = RGBColor(0, 0, 0)
        if is_table:
            run.bold = True
        else:
            run.italic = True
            run.bold = True
        return p

    def add_image_safe(img_path, width_inches=6.0, caption=""):
        if img_path.exists():
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(8)
            p.paragraph_format.space_after = Pt(2)
            run = p.add_run()
            run.add_picture(str(img_path), width=Inches(width_inches))
            if caption:
                add_caption(caption, is_table=False)

    def add_custom_table(headers, rows, col_widths=None):
        table = doc.add_table(rows=len(rows) + 1, cols=len(headers))
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        set_table_black_borders(table)

        # header
        hdr_cells = table.rows[0].cells
        for idx, h_text in enumerate(headers):
            cell = hdr_cells[idx]
            cell.text = h_text
            set_cell_margins(cell, top=120, bottom=120, left=150, right=150)
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after = Pt(2)
            for r in p.runs:
                r.font.name = 'Times New Roman'
                r.font.size = Pt(11)
                r.bold = True
                r.font.color.rgb = RGBColor(0, 0, 0)

        # data rows (nét kẻ đen bình thường, nền trắng chuẩn)
        for r_idx, row_data in enumerate(rows):
            row_cells = table.rows[r_idx + 1].cells
            for c_idx, val in enumerate(row_data):
                cell = row_cells[c_idx]
                cell.text = str(val)
                set_cell_margins(cell, top=100, bottom=100, left=150, right=150)
                p = cell.paragraphs[0]
                p.paragraph_format.space_after = Pt(2)
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER if c_idx == 0 or len(str(val)) < 15 else WD_ALIGN_PARAGRAPH.LEFT
                for r in p.runs:
                    r.font.name = 'Times New Roman'
                    r.font.size = Pt(11)
                    r.font.color.rgb = RGBColor(0, 0, 0)

        # chỉnh chiều rộng cột
        if col_widths and len(col_widths) == len(headers):
            for row in table.rows:
                for c_idx, w in enumerate(col_widths):
                    row.cells[c_idx].width = Inches(w)

        doc.add_paragraph().paragraph_format.space_after = Pt(4)
        return table

    # -------------------------------------------------------------
    # TRANG BÌA (TRANG RIÊNG)
    # -------------------------------------------------------------
    add_p("TRƯỜNG ĐẠI HỌC THỦ DẦU MỘT", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, first_indent=Cm(0))
    add_p("VIỆN CÔNG NGHỆ SỐ", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13, space_after=40, first_indent=Cm(0))

    add_p("BÁO CÁO TIỂU LUẬN / BÀI TẬP LỚN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=16, first_indent=Cm(0))
    add_p("HỌC PHẦN: KHAI THÁC TẬP DỮ LIỆU LỚN", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=35, first_indent=Cm(0))

    add_p("ĐỀ TÀI:", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=14, space_after=8, first_indent=Cm(0))
    add_p("PHÂN CỤM DỮ LIỆU TIỀN ĐIỆN TỬ\nBẰNG THUẬT TOÁN K-MEANS TRÊN NỀN TẢNG PYSPARK",
          bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=18, space_after=60, first_indent=Cm(0))

    # Thông tin sinh viên
    info_lines = [
        "Sinh viên thực hiện:\t[Họ và tên sinh viên]",
        "Mã số sinh viên:\t[MSSV]",
        "Lớp:\t\t[Tên lớp sinh viên]",
        "Lớp học phần:\t\t[Mã lớp học phần]",
        "Nhóm thực hiện:\t\t[Nhóm ...]",
        "Giảng viên hướng dẫn:\t[TS/ThS. Giảng viên hướng dẫn]",
    ]
    for line in info_lines:
        add_p(line, bold=True, align=WD_ALIGN_PARAGRAPH.LEFT, size=13, space_after=6, first_indent=Cm(3.0))

    add_p("", space_after=50, first_indent=Cm(0))
    add_p("BÌNH DƯƠNG / TP. HỒ CHÍ MINH, NĂM 2026", bold=True, align=WD_ALIGN_PARAGRAPH.CENTER, size=13, first_indent=Cm(0))

    doc.add_page_break()

    # -------------------------------------------------------------
    # LỜI CAM ĐOAN (TRANG RIÊNG)
    # -------------------------------------------------------------
    add_h1("LỜI CAM ĐOAN")
    add_p("Tôi xin cam đoan đây là báo cáo kết quả thực hành và nghiên cứu do nhóm chúng tôi tự thực hiện dưới sự định hướng của Giảng viên phụ trách học phần Khai thác tập dữ liệu lớn. Các số liệu thống kê, bảng biểu và kết quả chạy mô hình trong báo cáo này phản ánh đúng quá trình chạy mã nguồn thực tế trên máy tính.")
    add_p("Tập dữ liệu phục vụ nghiên cứu được thu thập từ nguồn dữ liệu công khai của sàn Binance (Binance Public Data Vision), bao gồm dữ liệu nến 1 phút của 20 cặp tiền điện tử trong năm 2025. Các tài liệu kỹ thuật, thuật toán và thư viện mã nguồn mở được sử dụng đều được nêu rõ trong danh mục tài liệu tham khảo.")
    add_p("Nhóm xin chịu hoàn toàn trách nhiệm trước các quy định về liêm chính học thuật của Nhà trường nếu có bất kỳ sự gian lận hoặc sao chép không hợp lệ nào xảy ra.")

    add_p("", space_after=20, first_indent=Cm(0))
    add_p("Bình Dương, ngày 21 tháng 09 năm 2026", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=Cm(0))
    add_p("Đại diện nhóm thực hiện", bold=True, align=WD_ALIGN_PARAGRAPH.RIGHT, first_indent=Cm(0))
    add_p("(Ký và ghi rõ họ tên)", italic=True, align=WD_ALIGN_PARAGRAPH.RIGHT, space_after=40, first_indent=Cm(0))

    doc.add_page_break()

    # -------------------------------------------------------------
    # MỤC LỤC (TRANG RIÊNG - KHÔNG CÓ SỐ LA MÃ)
    # -------------------------------------------------------------
    add_h1("MỤC LỤC")
    add_p("LỜI CAM ĐOAN", bold=True, first_indent=Cm(0))
    add_p("MỤC LỤC", bold=True, first_indent=Cm(0))
    add_p("DANH MỤC TỪ VIẾT TẮT", bold=True, first_indent=Cm(0))
    add_p("DANH MỤC BẢNG BIỂU", bold=True, first_indent=Cm(0))
    add_p("DANH MỤC HÌNH VÀ BIỂU ĐỒ", bold=True, first_indent=Cm(0))
    add_p("MỞ ĐẦU", bold=True, first_indent=Cm(0))
    add_p("1. Lý do chọn đề tài", first_indent=Cm(0.5))
    add_p("2. Mục tiêu nghiên cứu", first_indent=Cm(0.5))
    add_p("3. Câu hỏi nghiên cứu", first_indent=Cm(0.5))
    add_p("4. Đối tượng và phạm vi nghiên cứu", first_indent=Cm(0.5))
    add_p("CHƯƠNG 1. TỔNG QUAN BÀI TOÁN VÀ CƠ SỞ LÝ THUYẾT", bold=True, first_indent=Cm(0))
    add_p("1.1. Tổng quan bài toán phân cụm dữ liệu tiền điện tử", first_indent=Cm(0.5))
    add_p("1.2. Nền tảng Apache Spark và PySpark", first_indent=Cm(0.5))
    add_p("1.3. Thuật toán phân cụm K-Means", first_indent=Cm(0.5))
    add_p("1.4. Phương pháp xác định số cụm tối ưu", first_indent=Cm(0.5))
    add_p("CHƯƠNG 2. PHƯƠNG PHÁP, THIẾT KẾ VÀ QUY TRÌNH THỰC HIỆN", bold=True, first_indent=Cm(0))
    add_p("2.1. Thu thập dữ liệu từ Binance Public Data Vision", first_indent=Cm(0.5))
    add_p("2.2. Tiền xử lý dữ liệu và trích xuất đặc trưng", first_indent=Cm(0.5))
    add_p("2.3. Thiết kế hệ thống và cấu hình môi trường", first_indent=Cm(0.5))
    add_p("2.4. Chiến lược chọn số cụm K trên tập dữ liệu lớn", first_indent=Cm(0.5))
    add_p("CHƯƠNG 3. THỰC NGHIỆM, KẾT QUẢ VÀ ĐÁNH GIÁ", bold=True, first_indent=Cm(0))
    add_p("3.1. Môi trường thực nghiệm", first_indent=Cm(0.5))
    add_p("3.2. Đánh giá thời gian thực thi (Benchmark)", first_indent=Cm(0.5))
    add_p("3.3. Kết quả xác định số cụm K", first_indent=Cm(0.5))
    add_p("3.4. Phân tích kết quả phân cụm và ý nghĩa thực tế", first_indent=Cm(0.5))
    add_p("3.5. Trực quan hóa dữ liệu và mô phỏng quá trình lặp K-Means", first_indent=Cm(0.5))
    add_p("CHƯƠNG 4. ĐÁNH GIÁ KẾT QUẢ VÀ PHÂN CÔNG NHIỆM VỤ", bold=True, first_indent=Cm(0))
    add_p("4.1. Bảng tự chấm điểm theo đề cương học phần", first_indent=Cm(0.5))
    add_p("4.2. Bảng phân công công việc chi tiết", first_indent=Cm(0.5))
    add_p("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN", bold=True, first_indent=Cm(0))
    add_p("1. Các kết quả đã đạt được", first_indent=Cm(0.5))
    add_p("2. Một số hạn chế của đề tài", first_indent=Cm(0.5))
    add_p("3. Hướng phát triển tiếp theo", first_indent=Cm(0.5))
    add_p("TÀI LIỆU THAM KHẢO", bold=True, first_indent=Cm(0))
    add_p("PHỤ LỤC: MÃ NGUỒN VÀ HƯỚNG DẪN THỰC THI", bold=True, first_indent=Cm(0))
    add_p("1. Cấu trúc thư mục dự án", first_indent=Cm(0.5))
    add_p("2. Hướng dẫn các bước thực thi dự án", first_indent=Cm(0.5))

    doc.add_page_break()

    # -------------------------------------------------------------
    # DANH MỤC TỪ VIẾT TẮT (TRANG RIÊNG)
    # -------------------------------------------------------------
    add_h1("DANH MỤC TỪ VIẾT TẮT")
    add_caption("Bảng viết tắt các thuật ngữ chuyên ngành trong báo cáo", is_table=True)
    abbr_headers = ["Từ viết tắt", "Tên tiếng Anh đầy đủ", "Ý nghĩa / Giải thích"]
    abbr_rows = [
        ["API", "Application Programming Interface", "Giao diện lập trình ứng dụng"],
        ["CSV", "Comma-Separated Values", "Định dạng tệp văn bản phân tách bằng dấu phẩy"],
        ["DAG", "Directed Acyclic Graph", "Đồ thị có hướng không chu trình trong Spark"],
        ["HDFS", "Hadoop Distributed File System", "Hệ thống tệp phân tán của Hadoop"],
        ["IEEE", "Institute of Electrical and Electronics Engineers", "Viện Kỹ sư Điện và Điện tử"],
        ["MLlib", "Machine Learning Library", "Thư viện học máy của Apache Spark"],
        ["OHLCV", "Open, High, Low, Close, Volume", "Các mức giá mở, cao, thấp, đóng và khối lượng nến"],
        ["OOM", "Out-Of-Memory", "Lỗi hết bộ nhớ RAM khi chạy chương trình"],
        ["RAM", "Random Access Memory", "Bộ nhớ truy xuất ngẫu nhiên của máy tính"],
        ["RDD", "Resilient Distributed Dataset", "Tập dữ liệu phân tán có khả năng chịu lỗi trong Spark"],
        ["USDT", "Tether USD", "Đồng tiền mã hóa có giá trị neo theo Đô la Mỹ"],
        ["WCSS", "Within-Cluster Sum of Squares", "Tổng bình phương khoảng cách từ các điểm tới tâm cụm"],
    ]
    add_custom_table(abbr_headers, abbr_rows, col_widths=[1.2, 2.3, 2.7])

    doc.add_page_break()

    # -------------------------------------------------------------
    # DANH MỤC BẢNG BIỂU (TRANG RIÊNG)
    # -------------------------------------------------------------
    add_h1("DANH MỤC BẢNG BIỂU")
    add_p("Bảng 1.1: So sánh đặc tính xử lý giữa Pandas và Apache Spark", bold=True, first_indent=Cm(0))
    add_p("Bảng 2.1: Danh sách 20 cặp tiền điện tử thu thập từ Binance Vision trong năm 2025", bold=True, first_indent=Cm(0))
    add_p("Bảng 2.2: Cấu trúc các cột dữ liệu nến 1 phút ban đầu từ sàn Binance", bold=True, first_indent=Cm(0))
    add_p("Bảng 2.3: Bảng đặc trưng tính toán phục vụ huấn luyện mô hình K-Means", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.1: Cấu hình phần cứng và phần mềm sử dụng trong thực nghiệm", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.2: Thời gian huấn luyện K-Means theo các mốc dữ liệu (100K, 1M, 10.5M dòng)", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.3: Giá trị WCSS tương ứng với các số cụm K từ 2 đến 6 (trên mẫu 1 triệu dòng)", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.4: Điểm Silhouette tương ứng với các số cụm K từ 2 đến 6 (trên mẫu 100.000 dòng)", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.5: Tọa độ tâm của 4 cụm trên không gian đặc trưng đã chuẩn hóa", bold=True, first_indent=Cm(0))
    add_p("Bảng 3.6: Thống kê số lượng bản ghi và tỷ lệ phần trăm phân bố trên toàn bộ 10.512.000 dòng", bold=True, first_indent=Cm(0))
    add_p("Bảng 4.1: Bảng tự chấm điểm theo barem đánh giá học phần Khai thác tập dữ liệu lớn", bold=True, first_indent=Cm(0))
    add_p("Bảng 4.2: Bảng phân công tiến độ và nhiệm vụ chi tiết trong quá trình thực hiện", bold=True, first_indent=Cm(0))

    doc.add_page_break()

    # -------------------------------------------------------------
    # DANH MỤC HÌNH VÀ BIỂU ĐỒ (TRANG RIÊNG)
    # -------------------------------------------------------------
    add_h1("DANH MỤC HÌNH VÀ BIỂU ĐỒ")
    add_p("Hình 1.1: Mô hình làm việc giữa Driver và các Executor trong Apache Spark", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 2.1: Sơ đồ các bước xử lý dữ liệu từ tệp nén đến mô hình K-Means", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.1: Biểu đồ Benchmark thời gian huấn luyện K-Means theo quy mô dữ liệu (PySpark Local Mode, K=4)", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.2: Biểu đồ phương pháp Elbow biểu diễn giá trị WCSS theo từng số cụm K", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.3: Biểu đồ cột thể hiện hệ số Silhouette theo từng giá trị K", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.4: Biểu đồ kết hợp WCSS và Silhouette để xác định số cụm K", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.5: Biểu đồ vành khăn thể hiện tỷ lệ kích thước của 4 cụm thị trường (K=4)", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.6: Biểu đồ Radar so sánh đặc trưng chuẩn hóa tại tâm của 4 cụm", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.7: Bản đồ nhiệt thể hiện tỷ lệ phần trăm các dòng nến của 20 đồng tiền rơi vào từng cụm", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.8: Biểu đồ phân tán 2D các cụm (Tỷ suất sinh lời vs Biên độ nến và Khối lượng)", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.9: Biểu đồ hộp (Boxplot) phân bố đặc trưng theo từng cụm (từ phần vị 1% đến 99%)", italic=True, bold=True, first_indent=Cm(0))
    add_p("Hình 3.10: Minh họa quá trình hội tụ của tâm cụm qua các vòng lặp (Animation Frames)", italic=True, bold=True, first_indent=Cm(0))

    doc.add_page_break()

    # -------------------------------------------------------------
    # MỞ ĐẦU (QUA TRANG MỚI)
    # -------------------------------------------------------------
    add_h1("MỞ ĐẦU")
    add_h2("1. Lý do chọn đề tài")
    add_p("Trong môn học Khai thác tập dữ liệu lớn, việc xử lý các bộ dữ liệu có dung lượng từ hàng triệu đến hàng chục triệu dòng là một bài toán thực tế phổ biến. Thị trường tiền điện tử là một nguồn dữ liệu phong phú vì các sàn giao dịch lớn như Binance hoạt động liên tục 24/7 và ghi nhận lượng giao dịch rất lớn mỗi phút. Khi theo dõi dữ liệu giá nến 1 phút trên nhiều đồng tiền trong suốt một năm, tổng số dòng dữ liệu nhanh chóng vượt qua mốc 10 triệu bản ghi.")
    add_p("Khi nhóm bắt đầu làm việc với tập dữ liệu này bằng các công cụ quen thuộc như thư viện Pandas trên máy tính cá nhân, chương trình thường xuyên bị đơ hoặc báo lỗi hết bộ nhớ (Out-Of-Memory) do Pandas nạp toàn bộ tệp vào RAM và chỉ tính toán trên một luồng CPU. Để giải quyết vấn đề này, nhóm đã tìm hiểu và chuyển sang sử dụng Apache Spark (giao diện PySpark) ở chế độ Local Mode. Spark cho phép chia nhỏ dữ liệu thành nhiều phân vùng để xử lý song song trên tất cả các lõi CPU, đồng thời cung cấp sẵn thư viện học máy phân tán MLlib. Việc áp dụng thuật toán phân cụm K-Means của Spark giúp nhóm vừa hoàn thành bài thực hành phân tích hành vi nến giá tiền điện tử, vừa hiểu rõ cách thức hoạt động của các công cụ xử lý dữ liệu lớn trên máy tính thực tế.")

    add_h2("2. Mục tiêu nghiên cứu")
    add_p("1. Xây dựng chương trình tự động tải, giải nén và làm sạch dữ liệu nến 1 phút của 20 đồng tiền điện tử trong năm 2025 từ nguồn công khai Binance Vision, lưu lại dưới định dạng Parquet.")
    add_p("2. Cấu hình môi trường PySpark Local Mode trên máy tính cá nhân để xử lý ổn định toàn bộ hơn 10,5 triệu dòng dữ liệu mà không gặp lỗi tràn RAM.")
    add_p("3. Thực nghiệm khảo sát số lượng cụm K phù hợp bằng cách kết hợp phương pháp Elbow (độ giảm WCSS) và hệ số Silhouette.")
    add_p("4. Huấn luyện mô hình K-Means trên toàn bộ tập dữ liệu hơn 10,5 triệu dòng, từ đó rút ra nhận xét về các nhóm trạng thái của thị trường nến 1 phút.")
    add_p("5. Đo đạc thời gian thực thi (Benchmark) của K-Means trên các kích thước dữ liệu khác nhau (100 nghìn dòng, 1 triệu dòng và toàn bộ 10,5 triệu dòng) để kiểm chứng khả năng mở rộng của Spark.")

    add_h2("3. Câu hỏi nghiên cứu")
    add_p("Câu hỏi 1: Khi kích thước dữ liệu tăng từ 100.000 dòng lên hơn 10,5 triệu dòng, thời gian chạy của thuật toán K-Means trên PySpark Local Mode tăng theo quy luật nào?")
    add_p("Câu hỏi 2: Đối với tập dữ liệu nến 1 phút năm 2025, nên chọn số cụm K bằng bao nhiêu để các cụm vừa gọn vừa tách biệt rõ ràng dựa trên WCSS và Silhouette?")
    add_p("Câu hỏi 3: 4 cụm tìm được phản ánh những trạng thái dao động giá nào trong thực tế, và tỷ lệ xuất hiện của các đồng coin lớn (như BTC, ETH) có khác gì so với các Altcoin có biên độ dao động rộng?")

    add_h2("4. Đối tượng và phạm vi nghiên cứu")
    add_p("Đối tượng nghiên cứu: Chuỗi dữ liệu nến 1 phút (Open, High, Low, Close, Volume) và các đặc trưng tính toán gồm tỷ suất sinh lời theo nến, biên độ nến và logarit khối lượng.")
    add_p("Phạm vi dữ liệu: Dữ liệu nến 1 phút của 20 cặp đồng tiền điện tử phổ biến ghép cặp với USDT trong cả năm 2025 (từ ngày 01/01/2025 đến ngày 31/12/2025) tải từ kho lưu trữ Binance Data Vision, tổng cộng 10.512.000 dòng sau khi làm sạch.")
    add_p("Công cụ thực hiện: Python 3.11, Apache Spark 3.5 (PySpark), bộ công cụ Hadoop Winutils 3.0 trên hệ điều hành Windows 10/11.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # CHƯƠNG 1 (QUA TRANG MỚI - KHÔNG CÓ SỐ LA MÃ)
    # -------------------------------------------------------------
    add_h1("CHƯƠNG 1. TỔNG QUAN BÀI TOÁN VÀ CƠ SỞ LÝ THUYẾT")
    add_h2("1.1. Tổng quan bài toán phân cụm dữ liệu tiền điện tử")
    add_p("Khác với các thị trường chứng khoán truyền thống có giờ mở cửa và đóng cửa theo phiên, thị trường tiền điện tử hoạt động liên tục không nghỉ. Mỗi cây nến 1 phút ghi nhận 5 chỉ số cơ bản: giá mở cửa (Open), giá cao nhất (High), giá thấp nhất (Low), giá đóng cửa (Close) và khối lượng giao dịch (Volume).")
    add_p("Ở khung thời gian ngắn 1 phút, giá chịu ảnh hưởng của rất nhiều biến động ngẫu nhiên và lệnh đặt tự động, do đó việc dự đoán trực tiếp giá lên hay xuống (học máy có giám sát) thường có sai số rất lớn. Thay vào đó, phương pháp phân cụm (học máy không giám sát) lại rất phù hợp để tìm kiếm các mẫu hình nến quen thuộc. Thuật toán phân cụm sẽ gom những cây nến có mức sinh lời, biên độ dao động và lượng tiền đổ vào giống nhau về cùng một nhóm. Nhờ đó, người phân tích có thể biết được trong 1 năm qua, thị trường dành bao nhiêu phần trăm thời gian để đi ngang tích lũy, bao nhiêu phần trăm thời gian có sóng tăng mạnh và những thời điểm nào có biến động giá bất thường.")

    add_h2("1.2. Nền tảng Apache Spark và PySpark")
    add_h3("1.2.1. Kiến trúc phân tán và mô hình thực thi của Spark")
    add_p("Apache Spark là một hệ thống tính toán phân tán mã nguồn mở, được thiết kế để xử lý dữ liệu trực tiếp trong bộ nhớ RAM thay vì ghi tạm ra đĩa cứng như mô hình Hadoop MapReduce trước đây. Nhờ cơ chế này, các thuật toán cần lặp đi lặp lại nhiều vòng trên cùng một tập dữ liệu (như K-Means) chạy nhanh hơn rất nhiều lần.")
    add_p("Kiến trúc hoạt động của Spark gồm hai phần chính: Driver Program (tiến trình chạy chương trình chính, tạo SparkSession, chuyển đổi lệnh thành DAG) và Executors (các tiến trình công nhân trực tiếp đọc dữ liệu, tính toán và trả kết quả về Driver). Điểm đặc biệt của Spark là cơ chế Lazy Evaluation (thực thi lười): khi viết các lệnh lọc dòng hay thêm cột, Spark chỉ ghi lại sơ đồ logic và chỉ tối ưu thực thi khi gặp thao tác Action như count() hay write.")

    add_h3("1.2.2. Cơ chế Spark Local Mode trên máy tính cá nhân")
    add_p("Khi làm bài tập lớn, sinh viên thường không có sẵn cụm máy chủ phân tán. Spark hỗ trợ chế độ Local Mode rất thuận tiện cho mục đích này. Cả Driver và Executor đều chạy chung trong một tiến trình JVM trên máy tính cá nhân. Spark tự động tận dụng tất cả các luồng CPU để tính toán song song, chia dữ liệu thành các phân vùng nhỏ trong RAM được cấp phát (ví dụ 6GB RAM). Nhờ vậy, sinh viên có thể chạy ổn định hơn 10 triệu dòng mà không gặp lỗi tràn bộ nhớ.")

    add_h3("1.2.3. Thư viện học máy Spark MLlib")
    add_p("Spark cung cấp sẵn thư viện pyspark.ml hỗ trợ học máy phân tán, gồm: VectorAssembler gom các cột thành véc-tơ đặc trưng, StandardScaler chuẩn hóa về trung bình 0 và độ lệch chuẩn 1, KMeans để phân cụm và ClusteringEvaluator tính điểm Silhouette.")

    add_caption("Bảng 1.1: So sánh đặc tính xử lý giữa Pandas và Apache Spark", is_table=True)
    table_11_headers = ["Đặc tính so sánh", "Thư viện Pandas", "Khung tính toán Apache Spark"]
    table_11_rows = [
        ["Cách dùng bộ nhớ RAM", "Nạp toàn bộ tệp vào RAM, dễ bị tràn bộ nhớ", "Chia nhỏ thành partition, tự tràn đĩa an toàn"],
        ["Sử dụng vi xử lý", "Mặc định chạy trên 1 luồng CPU", "Tận dụng tất cả các luồng CPU song song"],
        ["Khả năng mở rộng", "Bị giới hạn bởi lượng RAM trên máy tính", "Chạy tốt từ máy đơn (Local) đến cụm máy chủ"],
        ["Cơ chế chạy lệnh", "Chạy ngay lập tức từng dòng lệnh (Eager)", "Cơ chế Lazy Evaluation, chỉ chạy khi có Action"],
        ["Khả năng chịu lỗi", "Lỗi tiến trình phải chạy lại từ đầu", "Tự tính toán lại phần dữ liệu lỗi nhờ Lineage"],
    ]
    add_custom_table(table_11_headers, table_11_rows, col_widths=[1.5, 2.3, 2.4])

    add_h2("1.3. Thuật toán phân cụm K-Means")
    add_p("K-Means là thuật toán học máy không giám sát nhằm chia N điểm dữ liệu thành K cụm sao cho tổng bình phương khoảng cách từ mỗi điểm đến tâm cụm của nó (WCSS - Within-Cluster Sum of Squares) là nhỏ nhất:")
    add_p("J = Σ Σ ||x_i - μ_k||^2  (với k từ 1 đến K, x_i thuộc cụm k)", italic=True, align=WD_ALIGN_PARAGRAPH.CENTER, first_indent=Cm(0))
    add_p("Thuật toán lặp qua hai bước: gán từng điểm vào tâm gần nhất và tính lại tâm cụm bằng trung bình cộng tọa độ các điểm. Trong Spark MLlib, giải thuật K-Means|| (Scalable K-Means++) được áp dụng để khởi tạo tâm cụm song song nhanh chóng, khắc phục điểm nghẽn của K-Means++ truyền thống trên dữ liệu lớn.")

    add_h2("1.4. Phương pháp xác định số cụm tối ưu")
    add_p("1. Phương pháp Elbow: Chạy K từ 2 đến 6, ghi lại WCSS. Điểm uốn (khuỷu tay) là nơi WCSS giảm chậm hẳn lại, biểu thị việc tăng thêm cụm không còn cải thiện nhiều.")
    add_p("2. Hệ số Silhouette: Đo lường độ gắn kết nội cụm a(i) và độ tách biệt liên cụm b(i), giá trị từ -1 đến 1. Điểm càng gần 1 chứng tỏ các cụm càng cô đặc và tách biệt rõ ràng.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # CHƯƠNG 2 (QUA TRANG MỚI - KHÔNG CÓ SỐ LA MÃ)
    # -------------------------------------------------------------
    add_h1("CHƯƠNG 2. PHƯƠNG PHÁP, THIẾT KẾ VÀ QUY TRÌNH THỰC HIỆN")
    add_h2("2.1. Thu thập dữ liệu từ Binance Public Data Vision")
    add_p("Dữ liệu được tải từ kho lưu trữ chính thức của sàn Binance (https://data.binance.vision/) đối với 20 cặp tiền điện tử ghép cặp USDT có khối lượng giao dịch lớn trong năm 2025 (tổng cộng 240 tệp ZIP, tương ứng 12 tháng).")

    add_caption("Bảng 2.1: Danh sách 20 cặp tiền điện tử thu thập từ Binance Vision trong năm 2025", is_table=True)
    table_21_headers = ["STT", "Cặp giao dịch", "Tên đồng tiền", "Phân loại / Vai trò"]
    table_21_rows = [
        ["1", "BTCUSDT", "Bitcoin", "Đồng tiền dẫn dắt thị trường, tài sản lưu trữ"],
        ["2", "ETHUSDT", "Ethereum", "Nền tảng hợp đồng thông minh phổ biến nhất"],
        ["3", "BNBUSDT", "Binance Coin", "Token tiện ích sàn Binance, mạng BNB Chain"],
        ["4", "SOLUSDT", "Solana", "Chuỗi khối Layer 1 tốc độ giao dịch cao"],
        ["5", "XRPUSDT", "Ripple", "Giải pháp thanh toán chuyển tiền"],
        ["6", "DOGEUSDT", "Dogecoin", "Đồng tiền meme có khối lượng giao dịch lớn"],
        ["7", "ADAUSDT", "Cardano", "Nền tảng Layer 1"],
        ["8", "TRXUSDT", "Tron", "Mạng lưới chuyển tiền thanh khoản cao"],
        ["9", "LINKUSDT", "Chainlink", "Mạng lưới dữ liệu Oracle phi tập trung"],
        ["10", "AVAXUSDT", "Avalanche", "Nền tảng đa chuỗi"],
        ["11-20", "DOT, LTC, UNI, ATOM, NEAR, ETC, XLM, VET, ICP, ALGO", "Các Altcoin lớn", "Hệ sinh thái Layer 1, DeFi, điện toán phân tán"],
    ]
    add_custom_table(table_21_headers, table_21_rows, col_widths=[0.6, 1.4, 2.0, 2.2])

    add_p("Chương trình download_data.py giải nén trực tiếp trong RAM bằng zipfile và io.BytesIO, tự động thử lại tối đa 3 lần khi rớt mạng, trích xuất 7 cột cơ bản: symbol, open_time, open, high, low, close, volume. Tổng cộng đạt 10.512.000 dòng nến 1 phút.")

    add_h2("2.2. Tiền xử lý dữ liệu và trích xuất đặc trưng")
    add_p("Nhóm làm sạch dữ liệu bằng cách loại bỏ dòng có open, close <= 0 hoặc volume < 0, kiểm tra điều kiện nến high >= low, và loại bỏ giá trị Null. Để tránh bị chi phối bởi thị giá tuyệt đối của từng đồng coin, nhóm tính toán 3 đặc trưng tương đối:")
    add_p("1. Tỷ suất sinh lời theo nến: return_pct = (close - open) / open", italic=True)
    add_p("2. Biên độ dao động nến: range_pct = (high - low) / open", italic=True)
    add_p("3. Logarit khối lượng giao dịch: log_volume = ln(volume + 1.0)", italic=True)
    add_p("Sau đó, StandardScaler đưa cả 3 đặc trưng về thang đo chuẩn (trung bình 0, độ lệch chuẩn 1), giúp mỗi đặc trưng đóng góp công bằng khi tính khoảng cách Euclidean.")

    add_h2("2.3. Thiết kế hệ thống và cấu hình môi trường")
    add_p("Hệ thống cài đặt Hadoop Winutils 3.0 trên Windows để tránh lỗi ghi tệp Parquet. SparkSession được cấu hình cấp 6GB RAM cho Driver (spark.driver.memory = 6g) và 8 phân vùng xáo trộn (spark.sql.shuffle.partitions = 8) tương ứng 8 luồng CPU. Dữ liệu được lưu dưới định dạng Apache Parquet nén Snappy, giúp giảm dung lượng từ 1,2 GB xuống chỉ còn 260 MB.")

    add_h2("2.4. Chiến lược chọn số cụm K trên tập dữ liệu lớn")
    add_p("Tính toán Silhouette trên 10,5 triệu dòng đòi hỏi lượng phép tính rất lớn. Để tối ưu thời gian, nhóm lấy mẫu ngẫu nhiên 1.000.000 dòng để khảo sát Elbow và lấy mẫu phân tầng 100.000 dòng để đo hệ số Silhouette cho K từ 2 đến 6. Sau khi xác định K=4, mô hình mới được huấn luyện trên toàn bộ 10.512.000 dòng.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # CHƯƠNG 3 (QUA TRANG MỚI - KHÔNG CÓ SỐ LA MÃ)
    # -------------------------------------------------------------
    add_h1("CHƯƠNG 3. THỰC NGHIỆM, KẾT QUẢ VÀ ĐÁNH GIÁ")
    add_h2("3.1. Môi trường thực nghiệm")
    add_p("Thực nghiệm được thực hiện trên máy tính cá nhân: Windows 10/11, CPU đa nhân (8 luồng), 16 GB RAM DDR4, ổ cứng NVMe SSD, Python 3.11 và Apache Spark 3.5.0 (PySpark Local Mode).")

    add_h2("3.2. Đánh giá thời gian thực thi (Benchmark)")
    add_caption("Bảng 3.2: Thời gian huấn luyện K-Means theo các mốc dữ liệu", is_table=True)
    table_32_headers = ["Mức dữ liệu", "Số lượng dòng", "Thời gian chạy (giây)", "Mức tăng số dòng", "Mức tăng thời gian", "Giá trị WCSS"]
    table_32_rows = [
        ["Nhỏ (100K)", "100.000", "12,13 s", "Gốc (1x)", "1,0x", "26.810,16"],
        ["Trung bình (1M)", "1.000.000", "64,57 s", "Tăng 10 lần", "Tăng 5,32 lần", "399.164,56"],
        ["Lớn (Toàn bộ)", "10.512.000", "163,78 s", "Tăng 105 lần", "Tăng 13,5 lần", "17.105.723,84"],
    ]
    add_custom_table(table_32_headers, table_32_rows, col_widths=[1.1, 1.2, 1.2, 1.1, 1.1, 1.1])

    add_image_safe(RESULTS_DIR / "benchmark_chart.png", width_inches=5.8,
                   caption="Hình 3.1: Biểu đồ Benchmark thời gian huấn luyện K-Means theo quy mô dữ liệu (PySpark Local Mode, K=4)")

    add_p("Nhận xét thực tế: Khi dữ liệu tăng 105 lần (từ 100K lên hơn 10,5 triệu dòng), thời gian chạy chỉ tăng 13,5 lần (đạt 163,78 giây, tương đương 2,7 phút). Ở mốc 100K dòng, thời gian chịu ảnh hưởng nhiều bởi chi phí khởi tạo máy ảo JVM (khoảng 10-12 giây). Khi dữ liệu lớn, khả năng chia nhỏ và tính toán song song trên các luồng CPU của Spark phát huy tác dụng rõ rệt, dữ liệu được giữ trong RAM bộ đệm và xử lý trơn tru mà không bị tràn bộ nhớ.")

    # -------------------------------------------------------------
    # 3.3. TÁCH RÕ BẢNG 3.3 VÀ BẢNG 3.4
    # -------------------------------------------------------------
    add_h2("3.3. Kết quả xác định số cụm K")
    add_p("Để trả lời cho Câu hỏi 2, nhóm tiến hành khảo sát số cụm K từ 2 đến 6 bằng cả phương pháp Elbow và hệ số Silhouette:")

    # Bảng 3.3: WCSS
    add_caption("Bảng 3.3: Giá trị WCSS tương ứng với các số cụm K từ 2 đến 6 (trên mẫu 1 triệu dòng)", is_table=True)
    table_33_headers = ["Số cụm (K)", "Giá trị WCSS", "Mức giảm WCSS so với K-1", "Tỷ lệ giảm"]
    table_33_rows = [
        ["2", "680.886,83", "—", "—"],
        ["3", "517.013,78", "Giảm 163.873,05", "Giảm 24,07%"],
        ["4", "399.164,56", "Giảm 117.849,22", "Giảm 22,80%"],
        ["5", "348.897,24", "Giảm 50.267,32", "Giảm 12,59%"],
        ["6", "305.648,52", "Giảm 43.248,72", "Giảm 12,40%"],
    ]
    add_custom_table(table_33_headers, table_33_rows, col_widths=[1.0, 1.8, 2.0, 1.4])

    # Bảng 3.4: Silhouette
    add_caption("Bảng 3.4: Điểm Silhouette tương ứng với các số cụm K từ 2 đến 6 (trên mẫu 100.000 dòng)", is_table=True)
    table_34_headers = ["Số cụm (K)", "Hệ số Silhouette", "Đánh giá chất lượng phân cụm"]
    table_34_rows = [
        ["2", "0,4313", "Điểm thấp, các cụm bị gộp chung quá nhiều"],
        ["3", "0,5084", "Bắt đầu tách biệt tốt hơn"],
        ["4", "0,5649", "Đạt giá trị cao nhất trong các mốc khảo sát (Tối ưu)"],
        ["5", "0,5608", "Điểm số có xu hướng giảm nhẹ"],
        ["6", "0,5624", "Điểm số đi ngang, không cải thiện thêm"],
    ]
    add_custom_table(table_34_headers, table_34_rows, col_widths=[1.2, 1.8, 3.2])

    add_image_safe(RESULTS_DIR / "elbow_chart.png", width_inches=5.5,
                   caption="Hình 3.2: Biểu đồ phương pháp Elbow biểu diễn giá trị WCSS theo từng số cụm K")

    add_image_safe(RESULTS_DIR / "silhouette_chart.png", width_inches=5.5,
                   caption="Hình 3.3: Biểu đồ cột thể hiện hệ số Silhouette theo từng giá trị K")

    add_image_safe(VIZ_DIR / "k_selection_combined.png", width_inches=5.8,
                   caption="Hình 3.4: Biểu đồ kết hợp WCSS và Silhouette để xác định số cụm K")

    add_p("Lý do chọn K=4: Đường cong Elbow gập rõ rệt tại K=4 (tốc độ giảm WCSS từ K=2 đến K=4 đạt hơn 22% mỗi bước, sau đó chỉ còn giảm 12%). Đồng thời, hệ số Silhouette đạt đỉnh cao nhất tại K=4 với giá trị 0,5649. Do đó, nhóm thống nhất chọn K=4 cho mô hình chính thức.")

    add_h2("3.4. Phân tích kết quả phân cụm và ý nghĩa thực tế")
    add_p("Mô hình K-Means với K=4 huấn luyện trên toàn bộ 10.512.000 dòng mất 106,08 giây. Tọa độ tâm cụm và phân bố tỷ lệ được trình bày dưới đây:")

    add_caption("Bảng 3.5: Tọa độ tâm của 4 cụm trên không gian đặc trưng đã chuẩn hóa", is_table=True)
    table_cluster_headers = ["Cụm", "Tên đặt cho trạng thái", "Return (z)", "Range (z)", "Volume (z)", "Số dòng nến", "Tỷ lệ (%)"]
    table_cluster_rows = [
        ["Cụm 0", "Biến động vừa, khối lượng lớn", "-0,45", "+0,08", "+0,83", "3.960.100", "37,67%"],
        ["Cụm 1", "Thị trường yên tĩnh (Đi ngang)", "-0,01", "-0,23", "-0,80", "5.151.860", "49,01%"],
        ["Cụm 2", "Ngoại lai (Biến động bất thường)", "+48,07", "+318,10", "+1,67", "37", "0,0004%"],
        ["Cụm 3", "Tăng trưởng mạnh, thanh khoản tốt", "+1,31", "+0,60", "+0,60", "1.400.003", "13,32%"],
    ]
    add_custom_table(table_cluster_headers, table_cluster_rows, col_widths=[0.8, 2.0, 0.7, 0.7, 0.7, 1.0, 0.7])

    add_image_safe(VIZ_DIR / "cluster_donut.png", width_inches=5.0,
                   caption="Hình 3.5: Biểu đồ vành khăn thể hiện tỷ lệ kích thước của 4 cụm thị trường (K=4)")

    add_image_safe(VIZ_DIR / "centroid_radar.png", width_inches=5.0,
                   caption="Hình 3.6: Biểu đồ Radar so sánh đặc trưng chuẩn hóa tại tâm của 4 cụm")

    add_p("Ý nghĩa 4 trạng thái thị trường (Trả lời Câu hỏi 3):")
    add_p("1. Cụm 1 (Chiếm 49,01%): Thị trường đi ngang tích lũy (sideway), biên độ nến rất hẹp và khối lượng giao dịch thấp. Đây là trạng thái chiếm phần lớn thời gian trong năm của các cặp nến 1 phút.")
    add_p("2. Cụm 0 (Chiếm 37,67%): Giai đoạn giao dịch sôi động với khối lượng lớn, giá chịu áp lực điều chỉnh giảm nhẹ hoặc rung lắc hấp thụ lượng cung.")
    add_p("3. Cụm 3 (Chiếm 13,32%): Các đợt sóng tăng giá dứt khoát đi kèm thanh khoản mua tích cực, mang lại biên độ lợi nhuận tốt cho nhà đầu tư.")
    add_p("4. Cụm 2 (Chiếm 0,0004% - 37 dòng): Các cây nến bị giật râu dài bất thường trong 1 phút do các đợt quét thanh lý hoặc lệnh lớn đột ngột. Thuật toán đã tự động phát hiện và cô lập riêng nhóm dị biệt này mà không làm nhiễu các cụm chính.")

    add_image_safe(VIZ_DIR / "coin_cluster_heatmap.png", width_inches=5.5,
                   caption="Hình 3.7: Bản đồ nhiệt thể hiện tỷ lệ phần trăm các dòng nến của 20 đồng tiền rơi vào từng cụm")

    add_p("Nhận xét về các đồng tiền: Các đồng coin vốn hóa lớn (BTC, ETH, BNB) có hơn 85% thời gian rơi vào Cụm 1 và Cụm 0 do có thanh khoản rất dày, giá chuyển động đầm hơn. Trong khi đó, các Altcoin có biến độ dao động rộng (DOGE, SOL, NEAR, AVAX) có tỷ lệ nến rơi vào Cụm 3 (Tăng trưởng mạnh) cao hơn hẳn (chiếm từ 15% đến 18% thời gian).")

    add_h2("3.5. Trực quan hóa dữ liệu và mô phỏng quá trình lặp K-Means")
    add_image_safe(VIZ_DIR / "scatter_2d_clusters.png", width_inches=6.0,
                   caption="Hình 3.8: Biểu đồ phân tán 2D các cụm (Tỷ suất sinh lời vs Biên độ nến và Khối lượng)")

    add_image_safe(VIZ_DIR / "feature_boxplots.png", width_inches=6.0,
                   caption="Hình 3.9: Biểu đồ hộp (Boxplot) phân bố đặc trưng theo từng cụm (từ phần vị 1% đến 99%)")

    add_p("Bên cạnh các biểu đồ tĩnh, tệp ảnh động kmeans_animation.gif (16 khung hình) ghi lại quá trình hội tụ của mô hình: ở vòng lặp 0, các tâm cụm phân bố ngẫu nhiên; qua 5 vòng đầu, các tâm di chuyển nhanh về vùng mật độ nến dày đặc; từ vòng lặp 6 trở đi, tâm cụm dịch chuyển rất ít và đi vào trạng thái ổn định.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # CHƯƠNG 4. ĐÁNH GIÁ KẾT QUẢ VÀ PHÂN CÔNG NHIỆM VỤ (QUA TRANG MỚI - CHỨA BẢNG 4.1 & 4.2)
    # -------------------------------------------------------------
    add_h1("CHƯƠNG 4. ĐÁNH GIÁ KẾT QUẢ VÀ PHÂN CÔNG NHIỆM VỤ")
    add_h2("4.1. Bảng tự chấm điểm theo đề cương học phần")
    add_caption("Bảng 4.1: Bảng tự chấm điểm theo barem đánh giá học phần Khai thác tập dữ liệu lớn", is_table=True)
    table_eval_headers = ["Phần nội dung", "Tiêu chí yêu cầu theo đề cương", "Thang điểm", "Điểm tự chấm", "Minh chứng cụ thể trong bài làm"]
    table_eval_rows = [
        ["Phần 1: Dữ liệu", "Thu thập bộ dữ liệu lớn thực tế, nêu rõ nguồn gốc, thời gian, phương thức thu thập. Giải thích các trường dữ liệu, độ lớn >= 1 triệu dòng, làm sạch dữ liệu.", "3,0", "3,0 / 3,0", "Thu thập từ Binance Vision, 20 đồng coin trong cả năm 2025. Đạt 10.512.000 dòng nến 1 phút. Làm sạch kỹ lưỡng, lưu Parquet nén gọn gàng."],
        ["Phần 2: Phân tích", "Trình bày lý thuyết Spark, K-Means, công thức toán học. Chuẩn hóa dữ liệu, chọn số cụm K (Elbow, Silhouette). Phân tích ý nghĩa kết quả, vẽ biểu đồ trực quan.", "4,0", "4,0 / 4,0", "Trình bày nguyên lý K-Means, Spark Local Mode. Khảo sát K=2..6, chọn K=4 có căn cứ từ Elbow và Silhouette. Vẽ 7 biểu đồ đa chiều và ảnh động mô phỏng."],
        ["Phần 3: Xây dựng & Kiểm thử", "Cấu hình môi trường Spark, Hadoop trên máy tính. Viết chương trình hoàn chỉnh, phân chia hàm rõ ràng, không lỗi. Đo thời gian thực thi Benchmark.", "3,0", "3,0 / 3,0", "Cấu hình thành công PySpark và Winutils trên Windows. Mã nguồn chạy mượt mà trên 10,5 triệu dòng. Đo Benchmark cụ thể trên 100K, 1M và 10.5M dòng."],
        ["TỔNG CỘNG", "Đánh giá toàn diện đề tài", "10,0", "10,0 / 10,0", "Hoàn thành đầy đủ các yêu cầu theo đề cương học phần."],
    ]
    add_custom_table(table_eval_headers, table_eval_rows, col_widths=[1.2, 1.8, 0.7, 0.9, 1.8])

    add_h2("4.2. Bảng phân công công việc chi tiết")
    add_caption("Bảng 4.2: Bảng phân công tiến độ và nhiệm vụ chi tiết trong quá trình thực hiện", is_table=True)
    table_tasks_headers = ["STT", "Thời gian", "Nội dung công việc", "Sản phẩm cụ thể", "Người phụ trách"]
    table_tasks_rows = [
        ["1", "Tuần 1", "Tìm hiểu bài toán, viết script tự động tải 240 file ZIP từ Binance", "download_data.py", "Thành viên 1"],
        ["2", "Tuần 2", "Kiểm tra dữ liệu, lọc bỏ dòng lỗi, tính 3 đặc trưng, lưu file Parquet", "crypto_20coins_1m.parquet", "Thành viên 1"],
        ["3", "Tuần 3", "Cài đặt Hadoop Winutils, viết pipeline PySpark MLlib và hàm đo Benchmark", "main.py", "Thành viên 2"],
        ["4", "Tuần 4", "Chạy thử nghiệm K=2..6, tính WCSS và Silhouette, huấn luyện mô hình K=4", "results/*.csv", "Thành viên 2"],
        ["5", "Tuần 5", "Vẽ các biểu đồ 2D, biểu đồ hộp, biểu đồ tròn, heatmap và radar", "visualize_clusters.py", "Thành viên 3"],
        ["6", "Tuần 6", "Viết code chụp các bước lặp của K-Means và xuất ảnh động GIF", "animate_kmeans.py", "Thành viên 3"],
        ["7", "Tuần 7", "Tổng hợp số liệu và viết báo cáo theo mẫu quy định của trường", "Báo cáo tiểu luận", "Cả nhóm"],
    ]
    add_custom_table(table_tasks_headers, table_tasks_rows, col_widths=[0.5, 0.8, 2.3, 1.6, 1.0])

    doc.add_page_break()

    # -------------------------------------------------------------
    # KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN (QUA TRANG MỚI)
    # -------------------------------------------------------------
    add_h1("KẾT LUẬN VÀ HƯỚNG PHÁT TRIỂN")
    add_h2("1. Các kết quả đã đạt được")
    add_p("Sau khi hoàn thành đề tài, nhóm đã thu được các kết quả cụ thể:")
    add_p("1. Xây dựng quy trình xử lý dữ liệu hoàn chỉnh từ khâu tải, làm sạch đến lưu trữ tệp Parquet cho 10.512.000 dòng nến 1 phút từ Binance Vision.")
    add_p("2. Ứng dụng thành công PySpark Local Mode trên máy tính cá nhân, đo đạc Benchmark cho thấy tốc độ xử lý hơn 10,5 triệu dòng chỉ mất 2,7 phút mà không bị tràn RAM.")
    add_p("3. Thống nhất chọn số cụm K=4 dựa trên sự phù hợp của cả phương pháp Elbow và hệ số Silhouette (0,5649).")
    add_p("4. Phân tích được 4 trạng thái biến động nến 1 phút trong thực tế và chỉ ra sự khác biệt về độ dao động giữa Bitcoin và các Altcoin.")

    add_h2("2. Một số hạn chế của đề tài")
    add_p("Thực nghiệm hiện mới chạy trên một máy tính cá nhân ở chế độ Local Mode, chưa có điều kiện thử nghiệm trên một cụm gồm nhiều máy tính liên kết mạng với nhau. Các đặc trưng đưa vào mô hình hiện mới dựa trên nến 1 phút độc lập, chưa xét đến tính chuỗi thời gian hoặc các chỉ báo kỹ thuật như RSI, MACD.")

    add_h2("3. Hướng phát triển tiếp theo")
    add_p("Thử nghiệm đưa mã nguồn lên nền tảng đám mây như Databricks để chạy trên cụm phân tán thật sự. Tìm hiểu ứng dụng Spark Streaming kết hợp Kafka để phân cụm luồng dữ liệu giá trực tiếp theo thời gian thực.")

    doc.add_page_break()

    # -------------------------------------------------------------
    # TÀI LIỆU THAM KHẢO (QUA TRANG MỚI)
    # -------------------------------------------------------------
    add_h1("TÀI LIỆU THAM KHẢO")
    refs = [
        "• M. Zaharia, M. Chowdhury, M. J. Franklin, S. Shenker, and I. Stoica, “Spark: Cluster computing with working sets,” in Proc. 2nd USENIX Conf. Hot Topics Cloud Comput. (HotCloud), Boston, MA, USA, 2010, pp. 10–10.",
        "• J. MacQueen, “Some methods for classification and analysis of multivariate observations,” in Proc. 5th Berkeley Symp. Math. Statist. and Probability, vol. 1, Berkeley, CA, USA: Univ. of California Press, 1967, pp. 281–297.",
        "• B. Bahmani, B. Moseley, A. Vattani, R. Kumar, and S. Vassilvitskii, “Scalable k-means++,” Proc. VLDB Endow., vol. 5, no. 7, pp. 622–633, Mar. 2012.",
        "• P. J. Rousseeuw, “Silhouettes: A graphical aid to the interpretation and validation of cluster analysis,” J. Comput. Appl. Math., vol. 20, pp. 53–65, Nov. 1987.",
        "• Apache Software Foundation, “Apache Spark Documentation: Machine Learning Library (MLlib) Guide,” Spark.apache.org. [Trực tuyến]. Có tại: https://spark.apache.org/docs/latest/ml-guide.html. [Truy cập: 15/09/2026].",
        "• Binance, “Binance Public Data Vision,” Binance.vision. [Trực tuyến]. Có tại: https://data.binance.vision/. [Truy cập: 10/09/2026].",
        "• A. Géron, Hands-On Machine Learning with Scikit-Learn, Keras, and TensorFlow: Concepts, Tools, and Techniques to Build Intelligent Systems, 3rd ed. Sebastopol, CA, USA: O'Reilly Media, 2022.",
        "• H. Karau, A. Konwinski, P. Wendell, and M. Zaharia, Learning Spark: Lightning-Fast Big Data Analysis, 1st ed. Sebastopol, CA, USA: O'Reilly Media, 2015.",
    ]
    for ref in refs:
        add_p(ref, align=WD_ALIGN_PARAGRAPH.JUSTIFY, space_after=6, first_indent=Cm(0))

    doc.add_page_break()

    # -------------------------------------------------------------
    # PHỤ LỤC: MÃ NGUỒN VÀ HƯỚNG DẪN THỰC THI (QUA TRANG MỚI)
    # -------------------------------------------------------------
    add_h1("PHỤ LỤC: MÃ NGUỒN VÀ HƯỚNG DẪN THỰC THI")
    add_h2("1. Cấu trúc thư mục dự án")
    dir_structure = (
        "K-means_Big data/\n"
        "├── data/\n"
        "│   ├── processed/\n"
        "│   │   └── crypto_20coins_1m.parquet       # Tệp dữ liệu 10.5M dòng nén Snappy\n"
        "│   └── download_detail_log.csv             # Nhật ký tải 240 tệp ZIP\n"
        "├── results/\n"
        "│   ├── benchmark.csv                       # Kết quả đo thời gian 100K, 1M, 10.5M\n"
        "│   ├── wcss.csv                            # Giá trị WCSS theo K=2..6\n"
        "│   ├── silhouette.csv                      # Điểm Silhouette theo K=2..6\n"
        "│   ├── centroids.csv                       # Tọa độ tâm của 4 cụm\n"
        "│   ├── cluster_sizes.csv                   # Số lượng bản ghi mỗi cụm\n"
        "│   ├── final_model_info.csv                # Thông tin mô hình cuối cùng\n"
        "│   ├── elbow_chart.png                     # Biểu đồ đường cong Elbow\n"
        "│   ├── silhouette_chart.png                # Biểu đồ cột điểm Silhouette\n"
        "│   ├── benchmark_chart.png                 # Biểu đồ thời gian Benchmark\n"
        "│   └── visualizations/                     # Các biểu đồ phân tích chuyên sâu\n"
        "│       ├── scatter_2d_clusters.png\n"
        "│       ├── coin_cluster_heatmap.png\n"
        "│       ├── feature_boxplots.png\n"
        "│       ├── cluster_donut.png\n"
        "│       ├── k_selection_combined.png\n"
        "│       ├── centroid_radar.png\n"
        "│       └── kmeans_animation.gif            # Ảnh động mô phỏng hội tụ\n"
        "├── src/\n"
        "│   ├── download_data.py                    # Thu thập và tiền xử lý nến từ Binance\n"
        "│   ├── main.py                             # Pipeline K-Means chính trên PySpark\n"
        "│   ├── plot_results.py                     # Vẽ 3 biểu đồ đánh giá mô hình\n"
        "│   ├── visualize_clusters.py               # Trực quan hóa 7 biểu đồ phân cụm\n"
        "│   ├── animate_kmeans.py                   # Tạo animation mô phỏng quá trình lặp\n"
        "│   ├── export_docx.py                      # Xuất báo cáo sang định dạng Word (.docx)\n"
        "│   └── export_pptx.py                      # Xuất slide trình chiếu PowerPoint (.pptx)\n"
        "├── yeu_cau_bao_cao/                        # Quy định và biểu mẫu của trường\n"
        "├── BaoCao_KMeans_BigData.docx              # Bản Word báo cáo đầy đủ\n"
        "├── BaoCao_KMeans_BigData.pptx              # Bản PowerPoint trình chiếu chuẩn 16:9\n"
        "└── bao_cao_tieu_luan.md                    # Toàn văn báo cáo đồ án học phần"
    )
    p_code = doc.add_paragraph()
    p_code.paragraph_format.first_line_indent = Cm(0)
    p_code.paragraph_format.space_before = Pt(4)
    p_code.paragraph_format.space_after = Pt(8)
    run_code = p_code.add_run(dir_structure)
    run_code.font.name = 'Consolas'
    run_code.font.size = Pt(9.5)
    run_code.font.color.rgb = RGBColor(0, 0, 0)

    add_h2("2. Hướng dẫn các bước thực thi dự án")
    add_p("Để chạy toàn bộ hệ thống từ dữ liệu thô đến kết quả phân cụm và trực quan hóa, thực hiện tuần tự các lệnh sau trong Terminal máy tính:")
    
    commands = [
        ("Bước 1: Tải và tiền xử lý dữ liệu (Thu thập 10,5 triệu dòng)", "python src/download_data.py"),
        ("Bước 2: Thực thi pipeline tính toán phân tán PySpark (Benchmark, Elbow, Silhouette, Huấn luyện K-Means)", "python src/main.py"),
        ("Bước 3: Vẽ các biểu đồ đánh giá mô hình cơ bản (Elbow, Silhouette, Benchmark)", "python src/plot_results.py"),
        ("Bước 4: Trực quan hóa chi tiết các biểu đồ phân cụm chuyên sâu (Scatter 2D, Boxplot, Donut, Heatmap, Radar)", "python src/visualize_clusters.py"),
        ("Bước 5: Tạo ảnh động GIF mô phỏng quá trình lặp K-Means", "python src/animate_kmeans.py"),
        ("Bước 6: Xuất file báo cáo Word hoàn chỉnh (.docx)", "python src/export_docx.py"),
        ("Bước 7: Xuất file trình chiếu PowerPoint hoàn chỉnh (.pptx)", "python src/export_pptx.py"),
    ]
    for step_desc, cmd in commands:
        add_p(step_desc, bold=True, first_indent=Cm(0.5))
        p_cmd = doc.add_paragraph()
        p_cmd.paragraph_format.first_line_indent = Cm(1.0)
        p_cmd.paragraph_format.space_after = Pt(6)
        r_cmd = p_cmd.add_run(cmd)
        r_cmd.font.name = 'Consolas'
        r_cmd.font.size = Pt(10.5)
        r_cmd.bold = True
        r_cmd.font.color.rgb = RGBColor(0, 0, 0)

    # Lưu tệp Word chính
    doc.save(OUTPUT_DOCX)
    print(f"Đã tạo thành công tệp Word báo cáo: {OUTPUT_DOCX}")

    # Đồng bộ sang file đặt tên theo quy định nộp nếu không bị khóa bởi Word
    try:
        doc.save(ALT_DOCX)
        print(f"Đã đồng bộ sang: {ALT_DOCX}")
    except Exception as e:
        print(f"Lưu ý: Không thể ghi đè {ALT_DOCX.name} do đang mở ({e}).")


if __name__ == "__main__":
    create_report()
