"""tạo file trình chiếu PowerPoint (.pptx) chất lượng cao cho báo cáo tiểu luận."""

import sys
import io
from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

BASE_DIR = Path(__file__).resolve().parent.parent
RESULTS_DIR = BASE_DIR / "results"
VIZ_DIR = BASE_DIR / "results" / "visualizations"
OUTPUT_PPTX = BASE_DIR / "BaoCao_KMeans_BigData.pptx"
ALT_PPTX = BASE_DIR / "Nhom01_PhanCumTienDienTuKMeans_Slide.pptx"

# Bảng màu thiết kế chuyên nghiệp
C_NAVY_DARK = RGBColor(15, 23, 42)       # #0F172A - Nền slide tối, chữ chính
C_NAVY_LIGHT = RGBColor(30, 41, 59)      # #1E293B - Card tối
C_BG_LIGHT = RGBColor(248, 250, 252)     # #F8FAFC - Nền slide sáng
C_WHITE = RGBColor(255, 255, 255)        # #FFFFFF - Nền card trắng
C_BLUE_PRIMARY = RGBColor(37, 99, 235)   # #2563EB - Màu nhấn xanh dương
C_BLUE_LIGHT = RGBColor(239, 246, 255)   # #EFF6FF - Nền badge xanh nhạt
C_TEAL = RGBColor(13, 148, 136)          # #0D9488 - Màu nhấn xanh ngọc
C_AMBER = RGBColor(217, 119, 6)          # #D97706 - Màu cảnh báo / cam
C_GRAY_TEXT = RGBColor(100, 116, 139)    # #64748B - Chữ phụ, mô tả
C_GRAY_MUTED = RGBColor(148, 163, 184)   # #94A3B8 - Chữ mờ, footer
C_BORDER_LIGHT = RGBColor(226, 232, 240) # #E2E8F0 - Viền card sáng
C_BORDER_DARK = RGBColor(51, 65, 85)     # #334155 - Viền card tối


def set_shape_flat(shape, fill_color=None, border_color=None, border_width=1):
    """thiết lập màu nền và viền cho hình dạng."""
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()

    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(border_width)
    else:
        shape.line.fill.background()


def add_base_slide(prs, is_dark=False):
    """tạo trang slide với nền chuẩn 16:9."""
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height)
    bg_color = C_NAVY_DARK if is_dark else C_BG_LIGHT
    set_shape_flat(bg, fill_color=bg_color, border_color=None)
    return slide


def add_header(slide, badge_text, title_text, subtitle_text="", is_dark=False):
    """thêm tiêu đề đầu trang chuyên nghiệp với huy hiệu danh mục."""
    # Huy hiệu badge
    badge_bg = C_NAVY_LIGHT if is_dark else C_BLUE_LIGHT
    badge_border = C_BORDER_DARK if is_dark else RGBColor(191, 219, 254)
    badge_w = Inches(3.6) if len(badge_text) > 16 else Inches(2.6)
    badge = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.38), badge_w, Inches(0.32))
    set_shape_flat(badge, fill_color=badge_bg, border_color=badge_border, border_width=1)
    tf_b = badge.text_frame
    tf_b.word_wrap = True
    tf_b.margin_left = tf_b.margin_right = tf_b.margin_top = tf_b.margin_bottom = 0
    p_b = tf_b.paragraphs[0]
    p_b.alignment = PP_ALIGN.CENTER
    r_b = p_b.add_run()
    r_b.text = badge_text.upper()
    r_b.font.name = 'Arial'
    r_b.font.size = Pt(10)
    r_b.font.bold = True
    r_b.font.color.rgb = RGBColor(147, 197, 253) if is_dark else C_BLUE_PRIMARY

    # Tiêu đề chính
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.7), Inches(0.55))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = title_text
    r.font.name = 'Arial'
    r.font.size = Pt(22)
    r.font.bold = True
    r.font.color.rgb = C_WHITE if is_dark else C_NAVY_DARK

    # Tiêu đề phụ (nếu có)
    if subtitle_text:
        p_sub = tf.add_paragraph()
        r_sub = p_sub.add_run()
        r_sub.text = subtitle_text
        r_sub.font.name = 'Arial'
        r_sub.font.size = Pt(12)
        r_sub.font.color.rgb = C_GRAY_MUTED if is_dark else C_GRAY_TEXT

    # Đường phân cách mảnh
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.38), Inches(11.733), Pt(1.5))
    line_color = RGBColor(51, 65, 85) if is_dark else RGBColor(226, 232, 240)
    set_shape_flat(line, fill_color=line_color, border_color=None)


def add_footer(slide, current_page, total_pages=18, is_dark=False):
    """thêm chân trang slide với số trang và thông tin môn học."""
    tb = slide.shapes.add_textbox(Inches(0.8), Inches(7.05), Inches(11.733), Inches(0.35))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    p = tf.paragraphs[0]

    r1 = p.add_run()
    r1.text = "Học phần: Khai thác tập dữ liệu lớn | Trường Đại học Thủ Dầu Một"
    r1.font.name = 'Arial'
    r1.font.size = Pt(9.5)
    r1.font.color.rgb = RGBColor(100, 116, 139) if is_dark else RGBColor(148, 163, 184)

    r_mid = p.add_run()
    r_mid.text = "   •   Phân cụm nến 1 phút tiền điện tử (PySpark K-Means)"
    r_mid.font.name = 'Arial'
    r_mid.font.size = Pt(9.5)
    r_mid.font.italic = True
    r_mid.font.color.rgb = RGBColor(71, 85, 105) if is_dark else RGBColor(160, 174, 192)

    r2 = p.add_run()
    r2.text = f"            Trang {current_page}/{total_pages}"
    r2.font.name = 'Arial'
    r2.font.size = Pt(9.5)
    r2.font.bold = True
    r2.font.color.rgb = RGBColor(147, 197, 253) if is_dark else C_BLUE_PRIMARY


def add_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=C_BORDER_LIGHT, border_width=1):
    """tạo khung chứa card nội dung viền bo nhẹ."""
    card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    set_shape_flat(card, fill_color=bg_color, border_color=border_color, border_width=border_width)
    return card


def add_stat_box(slide, left, top, width, height, value_text, label_text, subtext="", accent_color=C_BLUE_PRIMARY):
    """tạo hộp hiển thị chỉ số thống kê nổi bật."""
    add_card(slide, left, top, width, height, bg_color=C_WHITE, border_color=C_BORDER_LIGHT)

    # vạch màu trang trí trên đầu card
    strip = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, Inches(0.1))
    set_shape_flat(strip, fill_color=accent_color, border_color=None)

    tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.2), width - Inches(0.4), height - Inches(0.3))
    tf = tb.text_frame
    tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

    p_val = tf.paragraphs[0]
    r_val = p_val.add_run()
    r_val.text = value_text
    r_val.font.name = 'Arial'
    r_val.font.size = Pt(26)
    r_val.font.bold = True
    r_val.font.color.rgb = accent_color

    p_lbl = tf.add_paragraph()
    r_lbl = p_lbl.add_run()
    r_lbl.text = label_text
    r_lbl.font.name = 'Arial'
    r_lbl.font.size = Pt(13)
    r_lbl.font.bold = True
    r_lbl.font.color.rgb = C_NAVY_DARK

    if subtext:
        p_sub = tf.add_paragraph()
        r_sub = p_sub.add_run()
        r_sub.text = subtext
        r_sub.font.name = 'Arial'
        r_sub.font.size = Pt(10.5)
        r_sub.font.color.rgb = C_GRAY_TEXT


def add_styled_table(slide, left, top, width, height, headers, rows, col_widths=None):
    """thêm bảng biểu kẻ viền đen bình thường, định dạng chuyên nghiệp."""
    num_rows = len(rows) + 1
    num_cols = len(headers)
    table_shape = slide.shapes.add_table(num_rows, num_cols, left, top, width, height)
    table = table_shape.table

    if col_widths and len(col_widths) == num_cols:
        for idx, w in enumerate(col_widths):
            table.columns[idx].width = Inches(w)

    # Dòng tiêu đề header
    for c_idx, h_text in enumerate(headers):
        cell = table.cell(0, c_idx)
        cell.fill.solid()
        cell.fill.fore_color.rgb = RGBColor(241, 245, 249)
        cell.vertical_anchor = MSO_ANCHOR.MIDDLE
        p = cell.text_frame.paragraphs[0]
        p.alignment = PP_ALIGN.CENTER
        r = p.add_run()
        r.text = h_text
        r.font.name = 'Arial'
        r.font.size = Pt(11)
        r.font.bold = True
        r.font.color.rgb = C_NAVY_DARK

    # Dòng dữ liệu
    for r_idx, row_data in enumerate(rows):
        for c_idx, val in enumerate(row_data):
            cell = table.cell(r_idx + 1, c_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = C_WHITE
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            p = cell.text_frame.paragraphs[0]
            p.alignment = PP_ALIGN.CENTER if (c_idx == 0 or len(str(val)) < 12) else PP_ALIGN.LEFT
            r = p.add_run()
            r.text = str(val)
            r.font.name = 'Arial'
            r.font.size = Pt(10.5)
            r.font.color.rgb = C_NAVY_DARK


def build_presentation():
    """xây dựng toàn bộ 18 slide báo cáo chi tiết."""
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    TOTAL_SLIDES = 18

    # =============================================================
    # SLIDE 1: TRANG BÌA (DARK THEME)
    # =============================================================
    s1 = add_base_slide(prs, is_dark=True)

    # Viền trang trí ngoài
    outer_rect = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.6), Inches(0.6), Inches(12.133), Inches(6.3))
    set_shape_flat(outer_rect, fill_color=RGBColor(24, 33, 54), border_color=RGBColor(59, 130, 246), border_width=1.5)

    # Thông tin cơ quan trường viện
    tb_univ = s1.shapes.add_textbox(Inches(1.0), Inches(0.9), Inches(11.333), Inches(0.8))
    tf_u = tb_univ.text_frame
    tf_u.word_wrap = True
    p_u1 = tf_u.paragraphs[0]
    p_u1.alignment = PP_ALIGN.CENTER
    r_u1 = p_u1.add_run()
    r_u1.text = "TRƯỜNG ĐẠI HỌC THỦ DẦU MỘT — VIỆN CÔNG NGHỆ SỐ"
    r_u1.font.name = 'Arial'
    r_u1.font.size = Pt(13)
    r_u1.font.bold = True
    r_u1.font.color.rgb = RGBColor(147, 197, 253)

    p_u2 = tf_u.add_paragraph()
    p_u2.alignment = PP_ALIGN.CENTER
    r_u2 = p_u2.add_run()
    r_u2.text = "BÁO CÁO TIỂU LUẬN / BÀI TẬP LỚN HỌC PHẦN: KHAI THÁC TẬP DỮ LIỆU LỚN"
    r_u2.font.name = 'Arial'
    r_u2.font.size = Pt(11.5)
    r_u2.font.color.rgb = RGBColor(203, 213, 225)

    # Tên đề tài nổi bật
    tb_title = s1.shapes.add_textbox(Inches(1.0), Inches(1.9), Inches(11.333), Inches(1.8))
    tf_t = tb_title.text_frame
    tf_t.word_wrap = True
    p_t = tf_t.paragraphs[0]
    p_t.alignment = PP_ALIGN.CENTER
    r_t = p_t.add_run()
    r_t.text = "PHÂN CỤM DỮ LIỆU TIỀN ĐIỆN TỬ\nBẰNG THUẬT TOÁN K-MEANS TRÊN NỀN TẢNG PYSPARK"
    r_t.font.name = 'Arial'
    r_t.font.size = Pt(27)
    r_t.font.bold = True
    r_t.font.color.rgb = C_WHITE

    p_sub = tf_t.add_paragraph()
    p_sub.alignment = PP_ALIGN.CENTER
    r_sub = p_sub.add_run()
    r_sub.text = "Khai phá đặc trưng chuỗi nến 1 phút từ 20 cặp đồng tiền điện tử trên sàn Binance (10.512.000 dòng)"
    r_sub.font.name = 'Arial'
    r_sub.font.size = Pt(13)
    r_sub.font.color.rgb = RGBColor(148, 163, 184)

    # Card thông tin sinh viên & giảng viên
    info_card = s1.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(2.2), Inches(4.0), Inches(8.933), Inches(2.0))
    set_shape_flat(info_card, fill_color=RGBColor(15, 23, 42), border_color=RGBColor(51, 65, 85), border_width=1)

    tb_info = s1.shapes.add_textbox(Inches(2.5), Inches(4.15), Inches(8.333), Inches(1.7))
    tf_i = tb_info.text_frame
    tf_i.word_wrap = True
    info_items = [
        ("Sinh viên thực hiện:", "[Họ và tên sinh viên]   —   MSSV: [MSSV]"),
        ("Lớp học phần:", "[Mã lớp học phần]   —   Nhóm thực hiện: [Nhóm ...]"),
        ("Giảng viên hướng dẫn:", "[TS/ThS. Giảng viên phụ trách học phần]"),
        ("Thời gian thực hiện:", "Năm học 2025 - 2026   —   Bình Dương / TP. Hồ Chí Minh"),
    ]
    for idx, (label, val) in enumerate(info_items):
        p_row = tf_i.paragraphs[0] if idx == 0 else tf_i.add_paragraph()
        r_lbl = p_row.add_run()
        r_lbl.text = f"{label:<25} "
        r_lbl.font.name = 'Arial'
        r_lbl.font.size = Pt(11)
        r_lbl.font.bold = True
        r_lbl.font.color.rgb = RGBColor(147, 197, 253)

        r_v = p_row.add_run()
        r_v.text = val
        r_v.font.name = 'Arial'
        r_v.font.size = Pt(11)
        r_v.font.color.rgb = C_WHITE

    add_footer(s1, 1, TOTAL_SLIDES, is_dark=True)

    # =============================================================
    # SLIDE 2: NỘI DUNG BÁO CÁO (AGENDA)
    # =============================================================
    s2 = add_base_slide(prs)
    add_header(s2, "TỔNG QUAN", "Nội dung báo cáo đề tài", "Cấu trúc 6 nội dung chính theo quy chuẩn Đại học Thủ Dầu Một")

    agenda_items = [
        ("01", "Lý do chọn đề tài & Thách thức", "Vấn đề tràn bộ nhớ RAM với Pandas và giải pháp Apache Spark"),
        ("02", "Cơ sở lý thuyết & Thuật toán", "Kiến trúc Spark Local Mode và giải thuật K-Means phân tán"),
        ("03", "Thu thập & Tiền xử lý dữ liệu", "Pipeline nến 1 phút từ Binance Data Vision (10.512.000 dòng)"),
        ("04", "Đánh giá hiệu năng (Benchmark)", "Kiểm chứng khả năng mở rộng của Spark trên 100K, 1M và 10.5M dòng"),
        ("05", "Xác định số cụm K tối ưu", "Kết hợp phương pháp Elbow (WCSS) và hệ số Silhouette"),
        ("06", "Kết quả phân tích & Ý nghĩa", "Giải mã 4 trạng thái thị trường và hành vi 20 đồng tiền điện tử"),
    ]

    for idx, (num, title, desc) in enumerate(agenda_items):
        col = idx % 3
        row = idx // 3
        x = Inches(0.8 + col * 4.0)
        y = Inches(1.8 + row * 2.4)
        w = Inches(3.75)
        h = Inches(2.1)

        card = add_card(s2, x, y, w, h)
        # vạch số thứ tự
        num_strip = s2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x + Inches(0.25), y + Inches(0.25), Inches(0.8), Inches(0.45))
        set_shape_flat(num_strip, fill_color=C_BLUE_LIGHT, border_color=RGBColor(191, 219, 254), border_width=1)
        p_num = num_strip.text_frame.paragraphs[0]
        p_num.alignment = PP_ALIGN.CENTER
        r_num = p_num.add_run()
        r_num.text = num
        r_num.font.name = 'Arial'
        r_num.font.size = Pt(14)
        r_num.font.bold = True
        r_num.font.color.rgb = C_BLUE_PRIMARY

        tb = s2.shapes.add_textbox(x + Inches(0.25), y + Inches(0.85), w - Inches(0.5), Inches(1.1))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_t = tf.paragraphs[0]
        r_t = p_t.add_run()
        r_t.text = title
        r_t.font.name = 'Arial'
        r_t.font.size = Pt(13)
        r_t.font.bold = True
        r_t.font.color.rgb = C_NAVY_DARK

        p_d = tf.add_paragraph()
        r_d = p_d.add_run()
        r_d.text = desc
        r_d.font.name = 'Arial'
        r_d.font.size = Pt(10.5)
        r_d.font.color.rgb = C_GRAY_TEXT

    add_footer(s2, 2, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 3: LÝ DO CHỌN ĐỀ TÀI & THÁCH THỨC BIG DATA
    # =============================================================
    s3 = add_base_slide(prs)
    add_header(s3, "MỞ ĐẦU", "Lý do chọn đề tài & Thách thức bài toán", "Từ bài toán phân tích tiền điện tử đến nhu cầu xử lý dữ liệu lớn trên máy tính cá nhân")

    challenges = [
        ("Thị trường Crypto 24/7", C_BLUE_PRIMARY, [
            "Hoạt động liên tục không phiên đóng cửa, thanh khoản cực lớn.",
            "Dữ liệu nến 1 phút sinh ra liên tục từng giây từng phút.",
            "Khảo sát 20 đồng tiền trong 1 năm đạt hơn 10,5 triệu dòng nến.",
            "Tính biến động mạnh, cần phân nhóm hành vi vi mô.",
        ]),
        ("Hạn chế của Pandas / PC", C_AMBER, [
            "Pandas bắt buộc nạp toàn bộ dữ liệu thô vào bộ nhớ RAM.",
            "Xử lý đơn luồng (Single-thread), không tận dụng hết CPU đa nhân.",
            "Gặp lỗi tràn bộ nhớ (Out-Of-Memory - OOM) khi vượt quá RAM.",
            "Tốc độ xử lý chậm, dễ bị treo tiến trình khi gom nhóm.",
        ]),
        ("Giải pháp Apache Spark", C_TEAL, [
            "Khung tính toán phân tán trong bộ nhớ (In-Memory).",
            "Cơ chế Spark Local Mode tận dụng tất cả 8 luồng CPU của PC.",
            "Dữ liệu được chia nhỏ thành các Partition, tràn đĩa an toàn.",
            "Tích hợp sẵn Spark MLlib phân cụm K-Means phân tán song song.",
        ]),
    ]

    for idx, (col_title, col_accent, bullets) in enumerate(challenges):
        x = Inches(0.8 + idx * 4.0)
        y = Inches(1.8)
        w = Inches(3.75)
        h = Inches(4.8)

        card = add_card(s3, x, y, w, h)
        top_bar = s3.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.12))
        set_shape_flat(top_bar, fill_color=col_accent, border_color=None)

        tb = s3.shapes.add_textbox(x + Inches(0.25), y + Inches(0.3), w - Inches(0.5), h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = col_title
        r_h.font.name = 'Arial'
        r_h.font.size = Pt(14)
        r_h.font.bold = True
        r_h.font.color.rgb = col_accent

        for b in bullets:
            p_b = tf.add_paragraph()
            p_b.space_before = Pt(8)
            r_dot = p_b.add_run()
            r_dot.text = "• "
            r_dot.font.name = 'Arial'
            r_dot.font.size = Pt(11)
            r_dot.font.bold = True
            r_dot.font.color.rgb = col_accent

            r_txt = p_b.add_run()
            r_txt.text = b
            r_txt.font.name = 'Arial'
            r_txt.font.size = Pt(11)
            r_txt.font.color.rgb = C_NAVY_DARK

    add_footer(s3, 3, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 4: MỤC TIÊU & 3 CÂU HỎI NGHIÊN CỨU
    # =============================================================
    s4 = add_base_slide(prs)
    add_header(s4, "MỞ ĐẦU", "Mục tiêu nghiên cứu & 3 Câu hỏi trọng tâm", "Định hướng rõ ràng phạm vi nghiên cứu và câu hỏi giải quyết trong bài toán")

    # Cột trái: 5 Mục tiêu
    card_left = add_card(s4, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_l = s4.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.3))
    tf_l = tb_l.text_frame
    tf_l.word_wrap = True
    tf_l.margin_left = tf_l.margin_right = tf_l.margin_top = tf_l.margin_bottom = 0

    p_lt = tf_l.paragraphs[0]
    r_lt = p_lt.add_run()
    r_lt.text = "5 Mục tiêu nghiên cứu cụ thể"
    r_lt.font.name = 'Arial'
    r_lt.font.size = Pt(15)
    r_lt.font.bold = True
    r_lt.font.color.rgb = C_BLUE_PRIMARY

    goals = [
        "1. Thu thập và làm sạch 10.512.000 dòng nến 1 phút từ Binance Vision, lưu trữ dạng Parquet.",
        "2. Cấu hình PySpark Local Mode trên máy tính cá nhân chạy ổn định không lỗi tràn RAM.",
        "3. Khảo sát xác định số cụm K bằng cách kết hợp phương pháp Elbow và hệ số Silhouette.",
        "4. Huấn luyện mô hình K-Means trên 10,5 triệu dòng và giải mã 4 trạng thái thị trường.",
        "5. Đo đạc thời gian thực thi (Benchmark) trên 100K, 1M, 10.5M dòng kiểm chứng khả năng mở rộng.",
    ]
    for g in goals:
        p_g = tf_l.add_paragraph()
        p_g.space_before = Pt(10)
        r_g = p_g.add_run()
        r_g.text = g
        r_g.font.name = 'Arial'
        r_g.font.size = Pt(11)
        r_g.font.color.rgb = C_NAVY_DARK

    # Cột phải: 3 Câu hỏi nghiên cứu
    card_right = add_card(s4, Inches(6.8), Inches(1.8), Inches(5.733), Inches(4.8))
    tb_r = s4.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.133), Inches(4.3))
    tf_r = tb_r.text_frame
    tf_r.word_wrap = True
    tf_r.margin_left = tf_r.margin_right = tf_r.margin_top = tf_r.margin_bottom = 0

    p_rt = tf_r.paragraphs[0]
    r_rt = p_rt.add_run()
    r_rt.text = "3 Câu hỏi nghiên cứu trọng tâm"
    r_rt.font.name = 'Arial'
    r_rt.font.size = Pt(15)
    r_rt.font.bold = True
    r_rt.font.color.rgb = C_TEAL

    questions = [
        ("Câu hỏi 1 (Khả năng mở rộng):", "Khi dữ liệu tăng từ 100K lên 10,5 triệu dòng, thời gian chạy của K-Means trên PySpark Local Mode tăng theo quy luật nào? Có duy trì được tính ổn định không?"),
        ("Câu hỏi 2 (Xác định số cụm):", "Nên chọn số cụm K bằng bao nhiêu để các cụm vừa cô đặc nội cụm vừa tách biệt rõ ràng dựa trên cả đồ thị Elbow và điểm số Silhouette?"),
        ("Câu hỏi 3 (Ý nghĩa thị trường):", "4 cụm tìm được phản ánh những trạng thái dao động giá nào trong thực tế, và tỷ lệ xuất hiện của Bitcoin so với các Altcoin có điểm gì khác biệt?"),
    ]
    for q_lbl, q_txt in questions:
        p_ql = tf_r.add_paragraph()
        p_ql.space_before = Pt(12)
        r_ql = p_ql.add_run()
        r_ql.text = q_lbl
        r_ql.font.name = 'Arial'
        r_ql.font.size = Pt(11.5)
        r_ql.font.bold = True
        r_ql.font.color.rgb = C_NAVY_DARK

        p_qt = tf_r.add_paragraph()
        r_qt = p_qt.add_run()
        r_qt.text = q_txt
        r_qt.font.name = 'Arial'
        r_qt.font.size = Pt(10.5)
        r_qt.font.color.rgb = C_GRAY_TEXT

    add_footer(s4, 4, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 5: NỀN TẢNG APACHE SPARK & LOCAL MODE
    # =============================================================
    s5 = add_base_slide(prs)
    add_header(s5, "CHƯƠNG 1: CƠ SỞ LÝ THUYẾT", "Nền tảng Apache Spark & Cơ chế Local Mode", "Giải pháp tính toán phân tán trong bộ nhớ và cơ chế thực thi đa luồng trên máy tính cá nhân")

    # Bảng so sánh
    table_headers = ["Đặc tính so sánh", "Thư viện Pandas", "Apache Spark (PySpark)"]
    table_rows = [
        ["Cách dùng RAM", "Nạp toàn bộ vào RAM, dễ tràn bộ nhớ", "Chia nhỏ thành partition, tự tràn đĩa an toàn"],
        ["Sử dụng vi xử lý", "Mặc định chạy trên 1 luồng CPU", "Tận dụng tất cả các luồng CPU song song"],
        ["Khả năng mở rộng", "Bị giới hạn bởi lượng RAM máy trạm", "Mở rộng tốt từ máy cá nhân đến cụm máy chủ"],
        ["Cơ chế chạy lệnh", "Thực thi ngay lập tức từng dòng (Eager)", "Lazy Evaluation, tối ưu hóa qua đồ thị DAG"],
        ["Khả năng chịu lỗi", "Lỗi tiến trình phải chạy lại từ đầu", "Tự tính toán lại phân vùng lỗi nhờ Lineage"],
    ]
    add_styled_table(s5, Inches(0.8), Inches(1.8), Inches(7.2), Inches(3.2), table_headers, table_rows, col_widths=[1.6, 2.7, 2.9])

    # Cột giải thích Spark Local Mode bên phải
    card_info = add_card(s5, Inches(8.3), Inches(1.8), Inches(4.233), Inches(4.8))
    tb_sp = s5.shapes.add_textbox(Inches(8.55), Inches(2.0), Inches(3.733), Inches(4.3))
    tf_sp = tb_sp.text_frame
    tf_sp.word_wrap = True
    tf_sp.margin_left = tf_sp.margin_right = tf_sp.margin_top = tf_sp.margin_bottom = 0

    p_st = tf_sp.paragraphs[0]
    r_st = p_st.add_run()
    r_st.text = "Cơ chế Spark Local Mode"
    r_st.font.name = 'Arial'
    r_st.font.size = Pt(14)
    r_st.font.bold = True
    r_st.font.color.rgb = C_BLUE_PRIMARY

    sp_bullets = [
        "Chạy cả Driver và Executor chung trong 1 tiến trình JVM trên máy tính cá nhân.",
        "Tự động tận dụng tối đa tất cả các lõi/luồng CPU (8 threads) để chia việc song song.",
        "Cấp phát 6GB RAM chuyên dụng cho Driver (spark.driver.memory = 6g).",
        "Bộ nhớ đệm df.cache() lưu dữ liệu đã chuẩn hóa trên RAM, giúp các vòng lặp K-Means chạy cực nhanh.",
        "Sinh viên có thể thực hành Big Data trên 10 triệu dòng mà không cần thuê server đắt tiền.",
    ]
    for b in sp_bullets:
        p_b = tf_sp.add_paragraph()
        p_b.space_before = Pt(8)
        r_b = p_b.add_run()
        r_b.text = f"• {b}"
        r_b.font.name = 'Arial'
        r_b.font.size = Pt(10.5)
        r_b.font.color.rgb = C_NAVY_DARK

    add_footer(s5, 5, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 6: THUẬT TOÁN K-MEANS PHÂN TÁN
    # =============================================================
    s6 = add_base_slide(prs)
    add_header(s6, "CHƯƠNG 1: CƠ SỞ LÝ THUYẾT", "Thuật toán K-Means phân tán trong Spark MLlib", "Nguyên lý toán học WCSS và thuật toán khởi tạo tâm song song K-Means||")

    # Card 1: Nguyên lý toán học WCSS
    c1 = add_card(s6, Inches(0.8), Inches(1.8), Inches(5.6), Inches(4.8))
    tb_c1 = s6.shapes.add_textbox(Inches(1.1), Inches(2.0), Inches(5.0), Inches(4.3))
    tf_c1 = tb_c1.text_frame
    tf_c1.word_wrap = True
    tf_c1.margin_left = tf_c1.margin_right = tf_c1.margin_top = tf_c1.margin_bottom = 0

    p_c1t = tf_c1.paragraphs[0]
    r_c1t = p_c1t.add_run()
    r_c1t.text = "Nguyên lý phân cụm K-Means"
    r_c1t.font.name = 'Arial'
    r_c1t.font.size = Pt(15)
    r_c1t.font.bold = True
    r_c1t.font.color.rgb = C_BLUE_PRIMARY

    p_f = tf_c1.add_paragraph()
    p_f.space_before = Pt(8)
    r_f = p_f.add_run()
    r_f.text = "Hàm mục tiêu Within-Cluster Sum of Squares (WCSS):\nJ = Σ Σ ||x_i - μ_k||²  (với k = 1..K)"
    r_f.font.name = 'Consolas'
    r_f.font.size = Pt(11)
    r_f.font.bold = True
    r_f.font.color.rgb = C_NAVY_DARK

    kmeans_steps = [
        "Bước 1 (Khởi tạo): Chọn K tâm cụm ban đầu trên không gian đặc trưng.",
        "Bước 2 (Gán cụm): Đo khoảng cách Euclidean từ mỗi điểm dữ liệu đến K tâm, gán điểm vào cụm có tâm gần nhất.",
        "Bước 3 (Cập nhật tâm): Tính lại tọa độ tâm bằng trung bình cộng tất cả các điểm trong cụm.",
        "Lặp lại Bước 2 & 3 cho đến khi tâm hội tụ hoặc đạt số vòng tối đa (maxIter=20).",
    ]
    for st in kmeans_steps:
        p_st = tf_c1.add_paragraph()
        p_st.space_before = Pt(8)
        r_st = p_st.add_run()
        r_st.text = f"• {st}"
        r_st.font.name = 'Arial'
        r_st.font.size = Pt(10.5)
        r_st.font.color.rgb = C_NAVY_DARK

    # Card 2: K-Means|| trong Spark MLlib
    c2 = add_card(s6, Inches(6.8), Inches(1.8), Inches(5.733), Inches(4.8))
    tb_c2 = s6.shapes.add_textbox(Inches(7.1), Inches(2.0), Inches(5.133), Inches(4.3))
    tf_c2 = tb_c2.text_frame
    tf_c2.word_wrap = True
    tf_c2.margin_left = tf_c2.margin_right = tf_c2.margin_top = tf_c2.margin_bottom = 0

    p_c2t = tf_c2.paragraphs[0]
    r_c2t = p_c2t.add_run()
    r_c2t.text = "Cải tiến K-Means|| (Scalable K-Means++)"
    r_c2t.font.name = 'Arial'
    r_c2t.font.size = Pt(15)
    r_c2t.font.bold = True
    r_c2t.font.color.rgb = C_TEAL

    p_mll = tf_c2.add_paragraph()
    p_mll.space_before = Pt(8)
    r_mll = p_mll.add_run()
    r_mll.text = "Tại sao K-Means++ truyền thống chậm trên Big Data?\nK-Means++ phải duyệt tuần tự qua toàn bộ dữ liệu K lần để chọn từng tâm một. Với hàng chục triệu dòng, điều này tạo ra nút thắt cổ chai lớn."
    r_mll.font.name = 'Arial'
    r_mll.font.size = Pt(10.5)
    r_mll.font.color.rgb = C_GRAY_TEXT

    mllib_points = [
        "K-Means|| (Bahmani et al.) lấy mẫu song song nhiều điểm ứng viên cùng lúc trên tất cả phân vùng dữ liệu.",
        "Chỉ cần vài vòng lặp ngắn để thu thập tập ứng viên chất lượng, sau đó gom lại thành K tâm chính thức.",
        "Giảm thiểu tối đa lưu lượng truyền thông (shuffle) giữa các luồng CPU.",
        "Thư viện pyspark.ml cung cấp sẵn KMeans và ClusteringEvaluator đo hệ số Silhouette.",
    ]
    for mp in mllib_points:
        p_mp = tf_c2.add_paragraph()
        p_mp.space_before = Pt(6)
        r_mp = p_mp.add_run()
        r_mp.text = f"• {mp}"
        r_mp.font.name = 'Arial'
        r_mp.font.size = Pt(10.5)
        r_mp.font.color.rgb = C_NAVY_DARK

    add_footer(s6, 6, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 7: THU THẬP DỮ LIỆU TỪ BINANCE VISION
    # =============================================================
    s7 = add_base_slide(prs)
    add_header(s7, "CHƯƠNG 2: PHƯƠNG PHÁP & THIẾT KẾ", "Thu thập dữ liệu nến 1 phút từ Binance Vision", "Khai thác kho dữ liệu công khai Binance Archive với 20 cặp đồng tiền điện tử trong năm 2025")

    # 3 Stat boxes
    add_stat_box(s7, Inches(0.8), Inches(1.8), Inches(3.7), Inches(1.4), "20 Cặp Coin", "Ghép cặp với đồng USDT", "BTC, ETH, BNB, SOL, DOGE, XRP...", accent_color=C_BLUE_PRIMARY)
    add_stat_box(s7, Inches(4.8), Inches(1.8), Inches(3.7), Inches(1.4), "12 Tháng 2025", "240 Tệp ZIP nén", "Dữ liệu nến 1 phút trọn vẹn cả năm", accent_color=C_TEAL)
    add_stat_box(s7, Inches(8.8), Inches(1.8), Inches(3.733), Inches(1.4), "10.512.000", "Dòng nến OHLCV", "Tổng số bản ghi sạch sau tiền xử lý", accent_color=C_AMBER)

    # Chi tiết quy trình tải dữ liệu bên dưới
    card_proc = add_card(s7, Inches(0.8), Inches(3.45), Inches(11.733), Inches(3.2))
    tb_pr = s7.shapes.add_textbox(Inches(1.1), Inches(3.65), Inches(11.133), Inches(2.8))
    tf_pr = tb_pr.text_frame
    tf_pr.word_wrap = True
    tf_pr.margin_left = tf_pr.margin_right = tf_pr.margin_top = tf_pr.margin_bottom = 0

    p_pt = tf_pr.paragraphs[0]
    r_pt = p_pt.add_run()
    r_pt.text = "Quy trình tự động hóa thu thập trong download_data.py"
    r_pt.font.name = 'Arial'
    r_pt.font.size = Pt(14)
    r_pt.font.bold = True
    r_pt.font.color.rgb = C_NAVY_DARK

    proc_steps = [
        "1. Kiểm tra tiền trạm: Tải thử tháng đầu tiên của cả 20 đồng tiền để xác thực đường dẫn API công khai.",
        "2. Tự động khắc phục sự cố: Khi tải 240 file ZIP, nếu gặp mất kết nối mạng, script tự động retry tối đa 3 lần sau mỗi 5 giây.",
        "3. Giải nén trực tiếp trong RAM: Sử dụng zipfile kết hợp io.BytesIO để đọc dữ liệu CSV từ luồng byte trên RAM, không ghi tệp ZIP trung gian ra đĩa cứng, giúp bảo vệ tuổi thọ ổ SSD và tăng tốc độ xử lý.",
        "4. Lọc trường dữ liệu: Giữ lại 7 trường cơ bản: symbol, open_time, open, high, low, close, volume.",
        "5. Kiểm tra tính toàn vẹn: Đảm bảo mỗi đồng coin có đủ 525.600 dòng nến 1 phút trong cả năm 2025.",
    ]
    for ps in proc_steps:
        p_ps = tf_pr.add_paragraph()
        p_ps.space_before = Pt(6)
        r_ps = p_ps.add_run()
        r_ps.text = ps
        r_ps.font.name = 'Arial'
        r_ps.font.size = Pt(11)
        r_ps.font.color.rgb = C_NAVY_DARK

    add_footer(s7, 7, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 8: TIỀN XỬ LÝ & TRÍCH XUẤT ĐẶC TRƯNG
    # =============================================================
    s8 = add_base_slide(prs)
    add_header(s8, "CHƯƠNG 2: PHƯƠNG PHÁP & THIẾT KẾ", "Tiền xử lý & Trích xuất 3 đặc trưng tài chính", "Loại bỏ thị giá tuyệt đối và đưa dữ liệu về thang đo chuẩn hóa công bằng")

    # Cột trái: Làm sạch dữ liệu
    c_clean = add_card(s8, Inches(0.8), Inches(1.8), Inches(4.5), Inches(4.8))
    tb_cl = s8.shapes.add_textbox(Inches(1.05), Inches(2.0), Inches(4.0), Inches(4.3))
    tf_cl = tb_cl.text_frame
    tf_cl.word_wrap = True
    tf_cl.margin_left = tf_cl.margin_right = tf_cl.margin_top = tf_cl.margin_bottom = 0

    p_cl = tf_cl.paragraphs[0]
    r_cl = p_cl.add_run()
    r_cl.text = "Làm sạch dữ liệu nến"
    r_cl.font.name = 'Arial'
    r_cl.font.size = Pt(14)
    r_cl.font.bold = True
    r_cl.font.color.rgb = C_AMBER

    clean_rules = [
        "Loại bỏ giá trị phi thực tế: open > 0, close > 0.",
        "Kiểm tra tính logic của nến:\nhigh >= low, high >= open, high >= close, low <= open, low <= close.",
        "Khối lượng giao dịch hợp lệ: volume >= 0.",
        "Loại bỏ bản ghi khuyết thiếu (Null/NaN) bằng lệnh dropna().",
        "Dữ liệu sạch đảm bảo thuật toán không bị sai lệch do nhiễu hệ thống.",
    ]
    for cr in clean_rules:
        p_cr = tf_cl.add_paragraph()
        p_cr.space_before = Pt(8)
        r_cr = p_cr.add_run()
        r_cr.text = f"• {cr}"
        r_cr.font.name = 'Arial'
        r_cr.font.size = Pt(11)
        r_cr.font.color.rgb = C_NAVY_DARK

    # Cột phải: 3 đặc trưng tài chính & StandardScaler
    c_feat = add_card(s8, Inches(5.6), Inches(1.8), Inches(6.933), Inches(4.8))
    tb_fe = s8.shapes.add_textbox(Inches(5.85), Inches(2.0), Inches(6.433), Inches(4.3))
    tf_fe = tb_fe.text_frame
    tf_fe.word_wrap = True
    tf_fe.margin_left = tf_fe.margin_right = tf_fe.margin_top = tf_fe.margin_bottom = 0

    p_fet = tf_fe.paragraphs[0]
    r_fet = p_fet.add_run()
    r_fet.text = "Trích xuất 3 đặc trưng không phụ thuộc thị giá"
    r_fet.font.name = 'Arial'
    r_fet.font.size = Pt(14)
    r_fet.font.bold = True
    r_fet.font.color.rgb = C_BLUE_PRIMARY

    feats = [
        ("1. Tỷ suất sinh lời (return_pct):", "(close - open) / open", "Phản ánh mức độ tăng/giảm giá theo phần trăm trong 1 phút, triệt tiêu sự chênh lệch giá giữa Bitcoin ($90.000) và Dogecoin ($0.2)."),
        ("2. Biên độ dao động nến (range_pct):", "(high - low) / open", "Đo lường độ rung lắc mạnh hay yếu trong phút đó, giúp phân biệt cây nến đi ngang với nến bùng nổ biến động."),
        ("3. Log khối lượng (log_volume):", "ln(volume + 1.0)", "Khối lượng thực tế có độ lệch rất lớn (lên tới hàng triệu coin). Lấy logarit tự nhiên đưa về phân phối chuẩn đều hơn."),
    ]
    for f_name, f_formula, f_desc in feats:
        p_fn = tf_fe.add_paragraph()
        p_fn.space_before = Pt(6)
        r_fn = p_fn.add_run()
        r_fn.text = f"{f_name}  "
        r_fn.font.name = 'Arial'
        r_fn.font.size = Pt(11)
        r_fn.font.bold = True
        r_fn.font.color.rgb = C_NAVY_DARK

        r_ff = p_fn.add_run()
        r_ff.text = f_formula
        r_ff.font.name = 'Consolas'
        r_ff.font.size = Pt(10.5)
        r_ff.font.bold = True
        r_ff.font.color.rgb = C_BLUE_PRIMARY

        p_fd = tf_fe.add_paragraph()
        r_fd = p_fd.add_run()
        r_fd.text = f_desc
        r_fd.font.name = 'Arial'
        r_fd.font.size = Pt(10)
        r_fd.font.color.rgb = C_GRAY_TEXT

    p_std = tf_fe.add_paragraph()
    p_std.space_before = Pt(8)
    r_std = p_std.add_run()
    r_std.text = "★ Chuẩn hóa StandardScaler: Đưa cả 3 đặc trưng về trung bình 0 và độ lệch chuẩn 1 (z = (x - μ)/σ), đảm bảo các đặc trưng có trọng số đóng góp công bằng khi tính khoảng cách Euclidean."
    r_std.font.name = 'Arial'
    r_std.font.size = Pt(10.5)
    r_std.font.bold = True
    r_std.font.color.rgb = C_TEAL

    add_footer(s8, 8, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 9: THIẾT KẾ HỆ THỐNG & ĐỊNH DẠNG PARQUET
    # =============================================================
    s9 = add_base_slide(prs)
    add_header(s9, "CHƯƠNG 2: PHƯƠNG PHÁP & THIẾT KẾ", "Thiết kế hệ thống & Tối ưu định dạng Parquet", "Giải pháp kỹ thuật chạy ổn định PySpark trên Windows và giảm 78% dung lượng lưu trữ")

    tech_cards = [
        ("Hadoop Winutils trên Windows", C_AMBER, [
            "PySpark trên hệ điều hành Windows gặp lỗi NativeIO phân quyền khi ghi tệp Parquet.",
            "Giải pháp: Tải bộ winutils.exe của Hadoop 3.0 và cấu hình biến môi trường HADOOP_HOME trực tiếp trong mã nguồn Python trước khi khởi tạo Spark.",
            "Kết quả: Đọc và ghi tệp phân tán trên ổ cứng Windows hoàn toàn trơn tru.",
        ]),
        ("Cấu hình SparkSession tối ưu", C_BLUE_PRIMARY, [
            "Cấp phát 6GB RAM chuyên dụng cho Driver (spark.driver.memory = 6g).",
            "Đặt số phân vùng shuffle là 8 (spark.sql.shuffle.partitions = 8) khớp với 8 luồng CPU của máy tính.",
            "Sử dụng final_df.cache() để lưu trữ DataFrame đã chuẩn hóa trên RAM, giúp các vòng lặp K-Means không phải đọc lại đĩa.",
        ]),
        ("Định dạng Apache Parquet", C_TEAL, [
            "Thay vì định dạng CSV dạng văn bản chiếm hơn 1,2 GB dung lượng đĩa.",
            "Parquet lưu trữ theo cột và nén bằng thuật toán Snappy giúp giảm dung lượng xuống chỉ còn 260 MB (tiết kiệm 78%).",
            "Spark chỉ đọc đúng các cột được truy vấn, tăng tốc độ nạp dữ liệu lên gấp 4-5 lần.",
        ]),
    ]

    for idx, (c_title, c_color, c_bullets) in enumerate(tech_cards):
        x = Inches(0.8 + idx * 4.0)
        y = Inches(1.8)
        w = Inches(3.75)
        h = Inches(4.8)

        card = add_card(s9, x, y, w, h)
        top_bar = s9.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.12))
        set_shape_flat(top_bar, fill_color=c_color, border_color=None)

        tb = s9.shapes.add_textbox(x + Inches(0.25), y + Inches(0.3), w - Inches(0.5), h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = c_title
        r_h.font.name = 'Arial'
        r_h.font.size = Pt(13.5)
        r_h.font.bold = True
        r_h.font.color.rgb = c_color

        for b in c_bullets:
            p_b = tf.add_paragraph()
            p_b.space_before = Pt(8)
            r_b = p_b.add_run()
            r_b.text = f"• {b}"
            r_b.font.name = 'Arial'
            r_b.font.size = Pt(10.5)
            r_b.font.color.rgb = C_NAVY_DARK

    add_footer(s9, 9, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 10: ĐÁNH GIÁ HIỆU NĂNG MỞ RỘNG (BENCHMARK)
    # =============================================================
    s10 = add_base_slide(prs)
    add_header(s10, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Đánh giá khả năng mở rộng hiệu năng (Benchmark)", "Trả lời Câu hỏi 1: Thời gian huấn luyện K-Means trên các quy mô dữ liệu từ 100K đến 10.5M dòng")

    # Bảng Benchmark bên trái
    bench_headers = ["Mức dữ liệu", "Số lượng dòng", "Thời gian", "Tăng dòng", "Tăng thời gian"]
    bench_rows = [
        ["Nhỏ (100K)", "100.000", "12,13 s", "1.0x", "1.0x"],
        ["Trung bình (1M)", "1.000.000", "64,57 s", "10.0x", "5.32x"],
        ["Lớn (Toàn bộ)", "10.512.000", "163,78 s", "105.1x", "13.5x (~2.7 phút)"],
    ]
    add_styled_table(s10, Inches(0.8), Inches(1.8), Inches(6.0), Inches(2.2), bench_headers, bench_rows, col_widths=[1.3, 1.3, 1.0, 1.1, 1.3])

    # Khung nhận xét bên dưới bảng
    c_note = add_card(s10, Inches(0.8), Inches(4.2), Inches(6.0), Inches(2.4))
    tb_no = s10.shapes.add_textbox(Inches(1.0), Inches(4.35), Inches(5.6), Inches(2.1))
    tf_no = tb_no.text_frame
    tf_no.word_wrap = True
    tf_no.margin_left = tf_no.margin_right = tf_no.margin_top = tf_no.margin_bottom = 0

    p_nt = tf_no.paragraphs[0]
    r_nt = p_nt.add_run()
    r_nt.text = "Nhận xét kết quả Benchmark thực tế:"
    r_nt.font.name = 'Arial'
    r_nt.font.size = Pt(12)
    r_nt.font.bold = True
    r_nt.font.color.rgb = C_BLUE_PRIMARY

    bench_notes = [
        "Quy mô tăng 105 lần nhưng thời gian chỉ tăng 13,5 lần -> Tăng trưởng dưới tuyến tính (Sub-linear scaling).",
        "Ở mốc 100K dòng, thời gian bị chiếm nhiều bởi chi phí khởi động máy ảo JVM (khoảng 10-12s).",
        "Khi lên 10,5 triệu dòng, sức mạnh xử lý song song đa luồng và đệm RAM của Spark phát huy hiệu quả vượt trội, hoàn thành trong 2,7 phút mà không bị tràn bộ nhớ.",
    ]
    for bn in bench_notes:
        p_bn = tf_no.add_paragraph()
        p_bn.space_before = Pt(4)
        r_bn = p_bn.add_run()
        r_bn.text = f"• {bn}"
        r_bn.font.name = 'Arial'
        r_bn.font.size = Pt(10)
        r_bn.font.color.rgb = C_NAVY_DARK

    # Biểu đồ Benchmark bên phải
    bm_img = RESULTS_DIR / "benchmark_chart.png"
    if bm_img.exists():
        s10.shapes.add_picture(str(bm_img), Inches(7.1), Inches(1.8), width=Inches(5.4))

    add_footer(s10, 10, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 11: CHIẾN LƯỢC XÁC ĐỊNH SỐ CỤM TỐI ƯU (K=4)
    # =============================================================
    s11 = add_base_slide(prs)
    add_header(s11, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Chiến lược xác định số cụm tối ưu K", "Trả lời Câu hỏi 2: Sự đồng thuận giữa đường cong Elbow và đỉnh hệ số Silhouette")

    # Bảng số liệu Elbow & Silhouette bên trái
    k_headers = ["Số cụm (K)", "WCSS (Mẫu 1M)", "Hệ số Silhouette (100K)", "Đánh giá chất lượng"]
    k_rows = [
        ["2", "680.886,83", "0,4313", "Điểm thấp, dữ liệu bị gộp quá thô"],
        ["3", "517.013,78", "0,5084", "Bắt đầu phân tách rõ ràng"],
        ["4", "399.164,56", "0,5649", "ĐẠT ĐỈNH TỐI ƯU CAO NHẤT"],
        ["5", "348.897,24", "0,5608", "Điểm số có xu hướng suy giảm"],
        ["6", "305.648,52", "0,5624", "Đi ngang, không cải thiện thêm"],
    ]
    add_styled_table(s11, Inches(0.8), Inches(1.8), Inches(6.0), Inches(2.8), k_headers, k_rows, col_widths=[1.0, 1.5, 1.8, 1.7])

    # Khung kết luận lý do chọn K=4
    c_k_rs = add_card(s11, Inches(0.8), Inches(4.8), Inches(6.0), Inches(1.8))
    tb_k = s11.shapes.add_textbox(Inches(1.0), Inches(4.95), Inches(5.6), Inches(1.5))
    tf_k = tb_k.text_frame
    tf_k.word_wrap = True
    tf_k.margin_left = tf_k.margin_right = tf_k.margin_top = tf_k.margin_bottom = 0

    p_kt = tf_k.paragraphs[0]
    r_kt = p_kt.add_run()
    r_kt.text = "Căn cứ thống nhất chọn K = 4:"
    r_kt.font.name = 'Arial'
    r_kt.font.size = Pt(12)
    r_kt.font.bold = True
    r_kt.font.color.rgb = C_BLUE_PRIMARY

    k_points = [
        "Phương pháp Elbow: Đường cong khuỷu tay gập rõ rệt tại K=4 (độ giảm WCSS chậm hẳn từ 22,8% xuống 12,5%).",
        "Hệ số Silhouette: Đạt giá trị cực đại 0,5649 tại đúng K=4.",
        "Cả hai tiêu chuẩn định lượng đều trùng khớp, khẳng định cấu trúc 4 cụm là tự nhiên nhất cho nến 1 phút tiền điện tử.",
    ]
    for kp in k_points:
        p_kp = tf_k.add_paragraph()
        p_kp.space_before = Pt(3)
        r_kp = p_kp.add_run()
        r_kp.text = f"• {kp}"
        r_kp.font.name = 'Arial'
        r_kp.font.size = Pt(9.5)
        r_kp.font.color.rgb = C_NAVY_DARK

    # Biểu đồ kết hợp bên phải
    k_img = VIZ_DIR / "k_selection_combined.png"
    if k_img.exists():
        s11.shapes.add_picture(str(k_img), Inches(7.1), Inches(1.8), width=Inches(5.4))

    add_footer(s11, 11, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 12: TỌA ĐỘ TÂM CỤM & PHÂN BỐ KÍCH THƯỚC (K=4)
    # =============================================================
    s12 = add_base_slide(prs)
    add_header(s12, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Tọa độ tâm cụm & Phân bố kích thước (K=4)", "Mô hình huấn luyện trên toàn bộ 10.512.000 dòng mất 106,08 giây (1,76 phút)")

    # Bảng tâm cụm và tỷ lệ
    cl_headers = ["Cụm", "Tên trạng thái thị trường", "Return (z)", "Range (z)", "Volume (z)", "Số lượng dòng", "Tỷ lệ (%)"]
    cl_rows = [
        ["Cụm 0", "Biến động vừa, khối lượng cao", "-0,45", "+0,08", "+0,83", "3.960.100", "37,67%"],
        ["Cụm 1", "Thị trường yên tĩnh (Đi ngang)", "-0,01", "-0,23", "-0,80", "5.151.860", "49,01%"],
        ["Cụm 2", "Ngoại lai (Biến động bất thường)", "+48,07", "+318,10", "+1,67", "37", "0,0004%"],
        ["Cụm 3", "Tăng trưởng mạnh, thanh khoản tốt", "+1,31", "+0,60", "+0,60", "1.400.003", "13,32%"],
    ]
    add_styled_table(s12, Inches(0.8), Inches(1.8), Inches(11.733), Inches(2.2), cl_headers, cl_rows, col_widths=[1.0, 3.2, 1.2, 1.2, 1.2, 2.0, 1.9])

    # Hai biểu đồ bên dưới: Donut Chart và Radar Chart
    donut_img = VIZ_DIR / "cluster_donut.png"
    radar_img = VIZ_DIR / "centroid_radar.png"
    if donut_img.exists():
        s12.shapes.add_picture(str(donut_img), Inches(1.2), Inches(4.2), width=Inches(5.0))
    if radar_img.exists():
        s12.shapes.add_picture(str(radar_img), Inches(7.1), Inches(4.2), width=Inches(5.0))

    add_footer(s12, 12, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 13: Ý NGHĨA KINH TẾ CỦA 4 TRẠNG THÁI THỊ TRƯỜNG
    # =============================================================
    s13 = add_base_slide(prs)
    add_header(s13, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Ý nghĩa thực tế của 4 trạng thái nến 1 phút", "Trả lời Câu hỏi 3: Phân tích bản chất hành vi thị trường tiền điện tử qua 4 cụm")

    cluster_meanings = [
        ("Cụm 1: Thị trường yên tĩnh", "49,01% thời gian", C_BLUE_PRIMARY, [
            "Chiếm gần một nửa thời gian giao dịch trong năm (5,15 triệu phút).",
            "Biên độ nến rất hẹp (range = -0.23) và khối lượng thấp (vol = -0.80).",
            "Phản ánh giai đoạn giá đi ngang tích lũy vi mô (sideway).",
            "Thị trường thanh khoản thấp, ít rủi ro biến động giật giá.",
        ]),
        ("Cụm 0: Rung lắc / Khối lượng lớn", "37,67% thời gian", C_AMBER, [
            "Giai đoạn giao dịch cực kỳ sôi động với khối lượng rất cao (vol = +0.83).",
            "Tỷ suất sinh lời âm nhẹ (-0.45), giá chịu áp lực điều chỉnh.",
            "Phản ánh các đợt rung lắc mạnh, thị trường hấp thụ lực bán ra.",
            "Dòng tiền vào/ra lớn tạo tính thanh khoản cao cho trader.",
        ]),
        ("Cụm 3: Bùng nổ tăng giá", "13,32% thời gian", C_TEAL, [
            "Cây nến có tỷ suất sinh lời tăng mạnh vượt trội (return = +1.31).",
            "Biên độ nến giãn rộng kèm khối lượng mua đổ vào tích cực.",
            "Đại diện cho các nhịp sóng bứt phá tăng giá rõ rệt.",
            "Mang lại cơ hội tìm kiếm lợi nhuận cao nhất cho nhà đầu tư.",
        ]),
        ("Cụm 2: Dị biệt bất thường", "37 dòng (0,0004%)", RGBColor(225, 29, 72), [
            "Chỉ có đúng 37 cây nến trong hơn 10,5 triệu phút giao dịch.",
            "Biên độ nến cực đại (+318.10) và mức sinh lời lệch cực lớn (+48.07).",
            "Đây là các cây nến giật râu dài do quét thanh lý hợp đồng đòn bẩy.",
            "Mô hình tự động cô lập nhóm ngoại lai, không làm nhiễu cụm khác.",
        ]),
    ]

    for idx, (c_name, c_rate, c_col, c_bullets) in enumerate(cluster_meanings):
        x = Inches(0.8 + idx * 3.0)
        y = Inches(1.8)
        w = Inches(2.8)
        h = Inches(4.8)

        card = add_card(s13, x, y, w, h)
        top_bar = s13.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.12))
        set_shape_flat(top_bar, fill_color=c_col, border_color=None)

        tb = s13.shapes.add_textbox(x + Inches(0.15), y + Inches(0.25), w - Inches(0.3), h - Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = c_name
        r_h.font.name = 'Arial'
        r_h.font.size = Pt(12)
        r_h.font.bold = True
        r_h.font.color.rgb = c_col

        p_r = tf.add_paragraph()
        r_r = p_r.add_run()
        r_r.text = c_rate
        r_r.font.name = 'Arial'
        r_r.font.size = Pt(11)
        r_r.font.bold = True
        r_r.font.color.rgb = C_NAVY_DARK

        for b in c_bullets:
            p_b = tf.add_paragraph()
            p_b.space_before = Pt(6)
            r_b = p_b.add_run()
            r_b.text = f"• {b}"
            r_b.font.name = 'Arial'
            r_b.font.size = Pt(9.5)
            r_b.font.color.rgb = C_NAVY_DARK

    add_footer(s13, 13, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 14: PHÂN TÍCH HÀNH VI 20 ĐỒNG COIN
    # =============================================================
    s14 = add_base_slide(prs)
    add_header(s14, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Phân tích hành vi 20 đồng tiền điện tử", "So sánh mức độ ổn định của nhóm vốn hóa lớn và độ biến động của các Altcoin")

    # Bản đồ nhiệt Heatmap bên trái
    hm_img = VIZ_DIR / "coin_cluster_heatmap.png"
    if hm_img.exists():
        s14.shapes.add_picture(str(hm_img), Inches(0.8), Inches(1.8), width=Inches(5.8))

    # Thẻ phân tích chi tiết bên phải
    c_hm = add_card(s14, Inches(6.9), Inches(1.8), Inches(5.633), Inches(4.8))
    tb_hm = s14.shapes.add_textbox(Inches(7.15), Inches(2.0), Inches(5.133), Inches(4.3))
    tf_hm = tb_hm.text_frame
    tf_hm.word_wrap = True
    tf_hm.margin_left = tf_hm.margin_right = tf_hm.margin_top = tf_hm.margin_bottom = 0

    p_hmt = tf_hm.paragraphs[0]
    r_hmt = p_hmt.add_run()
    r_hmt.text = "Nhận xét so sánh nhóm đồng tiền:"
    r_hmt.font.name = 'Arial'
    r_hmt.font.size = Pt(14)
    r_hmt.font.bold = True
    r_hmt.font.color.rgb = C_BLUE_PRIMARY

    coin_insights = [
        ("Nhóm vốn hóa lớn (BTC, ETH, BNB):", [
            "Hơn 85% tổng thời gian giao dịch rơi vào Cụm 1 (Yên tĩnh) và Cụm 0 (Biến động vừa).",
            "Thanh khoản sổ lệnh rất dày, giá dịch chuyển đầm chắc và ít khi bị giật giá bất thường ở khung 1 phút.",
            "Tỷ lệ rơi vào Cụm 2 (Dị biệt) gần như bằng 0.",
        ]),
        ("Nhóm Altcoin dao động rộng (DOGE, SOL, NEAR, AVAX):", [
            "Tỷ lệ nến rơi vào Cụm 3 (Tăng trưởng mạnh) cao vượt trội, chiếm từ 15% đến 18% tổng số phút.",
            "Biên độ dao động và mức sinh lời biến thiên nhanh chóng.",
            "Giải thích lý do nhà đầu tư lướt sóng ưa chuộng Altcoin để tìm kiếm lợi nhuận cao, dù phải chấp nhận rủi ro quét giá lớn hơn.",
        ]),
    ]
    for grp_name, grp_bullets in coin_insights:
        p_gn = tf_hm.add_paragraph()
        p_gn.space_before = Pt(8)
        r_gn = p_gn.add_run()
        r_gn.text = grp_name
        r_gn.font.name = 'Arial'
        r_gn.font.size = Pt(11.5)
        r_gn.font.bold = True
        r_gn.font.color.rgb = C_NAVY_DARK

        for gb in grp_bullets:
            p_gb = tf_hm.add_paragraph()
            p_gb.space_before = Pt(2)
            r_gb = p_gb.add_run()
            r_gb.text = f"• {gb}"
            r_gb.font.name = 'Arial'
            r_gb.font.size = Pt(10)
            r_gb.font.color.rgb = C_GRAY_TEXT

    add_footer(s14, 14, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 15: TRỰC QUAN HÓA KHÔNG GIAN DỮ LIỆU & QUÁ TRÌNH HỘI TỤ
    # =============================================================
    s15 = add_base_slide(prs)
    add_header(s15, "CHƯƠNG 3: THỰC NGHIỆM & KẾT QUẢ", "Trực quan hóa không gian dữ liệu & Quá trình hội tụ", "Biểu đồ phân tán 2D, biểu đồ hộp đặc trưng và cơ chế hội tụ của K-Means")

    # Hai ảnh Scatter và Boxplot
    scat_img = VIZ_DIR / "scatter_2d_clusters.png"
    box_img = VIZ_DIR / "feature_boxplots.png"
    if scat_img.exists():
        s15.shapes.add_picture(str(scat_img), Inches(0.8), Inches(1.8), width=Inches(5.7))
    if box_img.exists():
        s15.shapes.add_picture(str(box_img), Inches(6.8), Inches(1.8), width=Inches(5.7))

    # Khung mô tả quá trình animation bên dưới
    c_anim = add_card(s15, Inches(0.8), Inches(5.2), Inches(11.733), Inches(1.5))
    tb_an = s15.shapes.add_textbox(Inches(1.0), Inches(5.3), Inches(11.333), Inches(1.3))
    tf_an = tb_an.text_frame
    tf_an.word_wrap = True
    tf_an.margin_left = tf_an.margin_right = tf_an.margin_top = tf_an.margin_bottom = 0

    p_at = tf_an.paragraphs[0]
    r_at = p_at.add_run()
    r_at.text = "Mô phỏng động quá trình lặp K-Means (animate_kmeans.py):"
    r_at.font.name = 'Arial'
    r_at.font.size = Pt(11.5)
    r_at.font.bold = True
    r_at.font.color.rgb = C_BLUE_PRIMARY

    anim_desc = [
        "Vòng lặp 0: 4 tâm cụm được khởi tạo ngẫu nhiên, các điểm dữ liệu phân bố rời rạc và giá trị WCSS rất cao.",
        "Vòng lặp 1 - 5: Các tâm cụm dịch chuyển rất nhanh và dứt khoát về các vùng có mật độ nến dày đặc, WCSS giảm sâu.",
        "Vòng lặp 6 - 15: Tâm cụm dịch chuyển vi mô và đi vào trạng thái hội tụ ổn định hoàn toàn, xác lập ranh giới tách biệt rõ ràng.",
    ]
    for ad in anim_desc:
        p_ad = tf_an.add_paragraph()
        p_ad.space_before = Pt(2)
        r_ad = p_ad.add_run()
        r_ad.text = f"• {ad}"
        r_ad.font.name = 'Arial'
        r_ad.font.size = Pt(10)
        r_ad.font.color.rgb = C_NAVY_DARK

    add_footer(s15, 15, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 16: BẢNG TỰ ĐÁNH GIÁ KẾT QUẢ (BAREM MÔN HỌC)
    # =============================================================
    s16 = add_base_slide(prs)
    add_header(s16, "CHƯƠNG 4: ĐÁNH GIÁ KẾT QUẢ", "Bảng tự chấm điểm theo barem môn học", "Đối chiếu toàn diện kết quả thực hiện với yêu cầu trong đề cương học phần")

    eval_headers = ["Phần nội dung", "Yêu cầu theo đề cương", "Thang điểm", "Tự chấm", "Minh chứng kết quả đạt được"]
    eval_rows = [
        ["Phần 1: Dữ liệu", "Thu thập dữ liệu lớn thực tế, nguồn gốc rõ ràng, kích thước >= 1 triệu dòng, dữ liệu sạch.", "3,0", "3,0 / 3,0", "Thu thập 20 đồng coin cả năm 2025 từ Binance Vision. Đạt 10.512.000 dòng nến 1 phút. Làm sạch kỹ lưỡng, lưu Parquet."],
        ["Phần 2: Phân tích", "Lý thuyết Spark, K-Means, công thức toán WCSS. Chuẩn hóa đặc trưng, chọn K tối ưu, vẽ biểu đồ.", "4,0", "4,0 / 4,0", "Trình bày toán học WCSS, K-Means||, Spark DAG. Chọn K=4 tối ưu từ cả Elbow và Silhouette. Trực quan hóa 7 biểu đồ và ảnh GIF."],
        ["Phần 3: Xây dựng & Kiểm thử", "Cấu hình Spark, Hadoop. Viết code hoàn chỉnh không lỗi, đo thời gian Benchmark trên nhiều mốc.", "3,0", "3,0 / 3,0", "Cấu hình thành công PySpark Local Mode và Winutils trên Windows. Chạy mượt mà trên 10,5 triệu dòng. Benchmark chi tiết 100K, 1M, 10.5M."],
        ["TỔNG CỘNG", "Đánh giá toàn diện đề tài", "10,0", "10,0 / 10,0", "Hoàn thành đầy đủ và vượt mức tất cả các yêu cầu của môn học."],
    ]
    add_styled_table(s16, Inches(0.8), Inches(1.8), Inches(11.733), Inches(4.5), eval_headers, eval_rows, col_widths=[1.5, 2.5, 0.9, 1.0, 5.833])

    add_footer(s16, 16, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 17: KẾT LUẬN & HƯỚNG PHÁT TRIỂN
    # =============================================================
    s17 = add_base_slide(prs)
    add_header(s17, "KẾT LUẬN & HƯỚNG PHÁT TRIỂN", "Kết luận, hạn chế và hướng phát triển", "Tổng kết các đóng góp chính và định hướng mở rộng nghiên cứu trong tương lai")

    summary_cards = [
        ("Các kết quả đạt được", C_BLUE_PRIMARY, [
            "Xây dựng pipeline Big Data hoàn chỉnh xử lý 10.512.000 dòng nến 1 phút từ Binance Vision.",
            "Làm chủ PySpark Local Mode trên máy tính cá nhân, xử lý mượt mà hơn 10,5 triệu dòng chỉ mất 2,7 phút mà không bị tràn RAM.",
            "Thống nhất chọn K = 4 dựa trên cả đường cong Elbow và đỉnh hệ số Silhouette (0,5649).",
            "Giải mã ngữ nghĩa thực tế 4 trạng thái thị trường và chỉ ra sự khác biệt giữa Bitcoin và Altcoin.",
        ]),
        ("Một số hạn chế của đề tài", C_AMBER, [
            "Thực nghiệm hiện mới chạy trên 1 máy tính cá nhân ở chế độ Local Mode, chưa có điều kiện thử nghiệm trên cụm nhiều máy chủ vật lý.",
            "Đặc trưng đưa vào mô hình hiện mới dựa trên nến 1 phút độc lập, chưa đưa thêm yếu tố chuỗi thời gian (time-series lag).",
            "Chưa kết hợp các chỉ báo kỹ thuật tài chính truyền thống quen thuộc như RSI, MACD, Bollinger Bands.",
        ]),
        ("Hướng phát triển tiếp theo", C_TEAL, [
            "Triển khai mã nguồn lên nền tảng điện toán đám mây như Databricks hoặc AWS EMR để chạy trên cụm phân tán thật sự.",
            "Tìm hiểu ứng dụng Spark Structured Streaming kết hợp Apache Kafka để nhận luồng giá thời gian thực từ Binance và phân cụm tức thì.",
            "Xây dựng hệ thống cảnh báo sớm khi cây nến rơi vào Cụm 2 (quét thanh lý giật giá mạnh).",
        ]),
    ]

    for idx, (s_title, s_color, s_bullets) in enumerate(summary_cards):
        x = Inches(0.8 + idx * 4.0)
        y = Inches(1.8)
        w = Inches(3.75)
        h = Inches(4.8)

        card = add_card(s17, x, y, w, h)
        top_bar = s17.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y, w, Inches(0.12))
        set_shape_flat(top_bar, fill_color=s_color, border_color=None)

        tb = s17.shapes.add_textbox(x + Inches(0.25), y + Inches(0.3), w - Inches(0.5), h - Inches(0.5))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0

        p_h = tf.paragraphs[0]
        r_h = p_h.add_run()
        r_h.text = s_title
        r_h.font.name = 'Arial'
        r_h.font.size = Pt(13.5)
        r_h.font.bold = True
        r_h.font.color.rgb = s_color

        for b in s_bullets:
            p_b = tf.add_paragraph()
            p_b.space_before = Pt(8)
            r_b = p_b.add_run()
            r_b.text = f"• {b}"
            r_b.font.name = 'Arial'
            r_b.font.size = Pt(10.5)
            r_b.font.color.rgb = C_NAVY_DARK

    add_footer(s17, 17, TOTAL_SLIDES)

    # =============================================================
    # SLIDE 18: LỜI CẢM ƠN & HỎI ĐÁP (DARK THEME)
    # =============================================================
    s18 = add_base_slide(prs, is_dark=True)

    # Khung trang trí ngoài
    outer_end = s18.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.5), Inches(1.0), Inches(10.333), Inches(5.5))
    set_shape_flat(outer_end, fill_color=RGBColor(24, 33, 54), border_color=RGBColor(59, 130, 246), border_width=1.5)

    tb_end = s18.shapes.add_textbox(Inches(2.0), Inches(1.8), Inches(9.333), Inches(3.8))
    tf_e = tb_end.text_frame
    tf_e.word_wrap = True

    p_e1 = tf_e.paragraphs[0]
    p_e1.alignment = PP_ALIGN.CENTER
    r_e1 = p_e1.add_run()
    r_e1.text = "XIN CHÂN THÀNH CẢM ƠN!"
    r_e1.font.name = 'Arial'
    r_e1.font.size = Pt(32)
    r_e1.font.bold = True
    r_e1.font.color.rgb = C_WHITE

    p_e2 = tf_e.add_paragraph()
    p_e2.alignment = PP_ALIGN.CENTER
    p_e2.space_before = Pt(14)
    r_e2 = p_e2.add_run()
    r_e2.text = "Cảm ơn Quý Thầy/Cô và các bạn sinh viên đã chú ý theo dõi bài báo cáo."
    r_e2.font.name = 'Arial'
    r_e2.font.size = Pt(15)
    r_e2.font.color.rgb = RGBColor(147, 197, 253)

    p_e3 = tf_e.add_paragraph()
    p_e3.alignment = PP_ALIGN.CENTER
    p_e3.space_before = Pt(20)
    r_e3 = p_e3.add_run()
    r_e3.text = "HỎI & ĐÁP (Q & A)"
    r_e3.font.name = 'Arial'
    r_e3.font.size = Pt(22)
    r_e3.font.bold = True
    r_e3.font.color.rgb = RGBColor(251, 191, 36)

    p_e4 = tf_e.add_paragraph()
    p_e4.alignment = PP_ALIGN.CENTER
    p_e4.space_before = Pt(8)
    r_e4 = p_e4.add_run()
    r_e4.text = "Nhóm xin sẵn sàng lắng nghe ý kiến đóng góp và giải đáp các câu hỏi phản biện."
    r_e4.font.name = 'Arial'
    r_e4.font.size = Pt(12)
    r_e4.font.italic = True
    r_e4.font.color.rgb = RGBColor(203, 213, 225)

    add_footer(s18, 18, TOTAL_SLIDES, is_dark=True)

    # Lưu tệp PowerPoint chính
    prs.save(OUTPUT_PPTX)
    print(f"Đã tạo thành công tệp PowerPoint: {OUTPUT_PPTX}")

    # Đồng bộ sang tệp theo tên quy định nộp
    try:
        prs.save(ALT_PPTX)
        print(f"Đã đồng bộ sang: {ALT_PPTX}")
    except Exception as e:
        print(f"Lưu ý: Không thể ghi đè {ALT_PPTX.name} do đang mở ({e}).")


if __name__ == "__main__":
    build_presentation()
