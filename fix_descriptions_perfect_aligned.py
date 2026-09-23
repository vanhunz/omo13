# -*- coding: utf-8 -*-
import docx
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

def set_font_tnr(run, size_pt=13, bold=False, italic=False, color_rgb=(0,0,0)):
    run.font.name = 'Times New Roman'
    run.font.size = Pt(size_pt)
    run.bold = bold
    run.italic = italic
    if color_rgb:
        run.font.color.rgb = RGBColor(*color_rgb)
    rPr = run._r.get_or_add_rPr()
    rFonts = parse_xml(f'<w:rFonts {nsdecls("w")} w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman" w:eastAsia="Times New Roman"/>')
    rPr.append(rFonts)

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{fill_hex}"/>')
    tcPr.append(shd)

def set_table_borders(table, color="D3D3D3", sz="4", val="single"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        f'<w:top w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:bottom w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideH w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:insideV w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:left w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'<w:right w:val="{val}" w:sz="{sz}" w:space="0" w:color="{color}"/>'
        f'</w:tblBorders>'
    )
    tblPr.append(borders)

def format_cell(cell, text, bold=False, italic=False, size_pt=10, color_rgb=(0,0,0), align=WD_ALIGN_PARAGRAPH.LEFT):
    cell.text = ""
    p = cell.paragraphs[0]
    p.alignment = align
    p.paragraph_format.space_before = Pt(2)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.line_spacing = 1.15
    run = p.add_run(text)
    set_font_tnr(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)

def add_p_after(para, text, style='Normal', bold=False, italic=False, size_pt=13, color_rgb=(0,0,0), align=WD_ALIGN_PARAGRAPH.LEFT, space_before=2, space_after=4):
    new_p = OxmlElement('w:p')
    para._p.addnext(new_p)
    new_para = docx.text.paragraph.Paragraph(new_p, para._parent)
    try:
        new_para.style = style
    except:
        new_para.style = 'Normal'
    new_para.alignment = align
    new_para.paragraph_format.space_before = Pt(space_before)
    new_para.paragraph_format.space_after = Pt(space_after)
    new_para.paragraph_format.line_spacing = 1.35
    if text:
        run = new_para.add_run(text)
        set_font_tnr(run, size_pt=size_pt, bold=bold, italic=italic, color_rgb=color_rgb)
    return new_para

def apply_global_times_new_roman(doc):
    for p in doc.paragraphs:
        p.paragraph_format.line_spacing = 1.35
        for r in p.runs:
            set_font_tnr(r, size_pt=r.font.size.pt if r.font.size else 13, bold=r.bold, italic=r.italic, color_rgb=(0,0,0) if not r.font.color or not r.font.color.rgb else None)
    
    for t in doc.tables:
        set_table_borders(t)
        for r in t.rows:
            for c in r.cells:
                for p in c.paragraphs:
                    p.paragraph_format.line_spacing = 1.15
                    for run in p.runs:
                        set_font_tnr(run, size_pt=run.font.size.pt if run.font.size else 10, bold=run.bold, italic=run.italic)

def fix_all_descriptions_and_align_with_source():
    print("Loading base document DoAnNhom13_IS425_OMO.docx...")
    doc = docx.Document('DoAnNhom13_IS425_OMO.docx')
    print(f"Loaded document with {len(doc.paragraphs)} paragraphs and {len(doc.tables)} tables.")

    # 1. Standardize margins
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)

    tables = list(doc.tables)

    # -------------------------------------------------------------
    # 2. UPDATE TABLE 1: EXACT 3 CATEGORIES & 15 PRODUCTS FROM js/products.js
    # -------------------------------------------------------------
    if len(tables) > 1:
        t1 = tables[1]
        print("Populating Table 1 (Bảng 1.2) with EXACT 3 categories and 15 products from js/products.js...")
        
        while len(t1.rows) < 4:
            t1.add_row()
        while len(t1.rows) > 4:
            t1._tbl.remove(t1.rows[-1]._tr)

        headers_1 = ["Phân Nhóm Kích Thước (Danh Mục)", "Khoảng Giá & Chi Tiết 15 Sản Phẩm Trong Source Code", "Đặc Tính Chất Liệu & Tiêu Chuẩn Gia Công"]
        for c_idx in range(min(len(headers_1), len(t1.rows[0].cells))):
            format_cell(t1.rows[0].cells[c_idx], headers_1[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t1.rows[0].cells[c_idx], "4A6FA5")

        prod_exact_3cat = [
            (
                "🌟 Phân nhóm 1: Gấu nhỏ (5cm - 15cm)\n(Mã danh mục: gau-nho)", 
                "Khoảng giá: 45.000đ - 95.000đ (4 sản phẩm)\n\n"
                "• sp-01: Móc Khóa Gấu Bông Keychain Xinh\n"
                "  - Giá: 45.000đ (Giá gốc: 60.000đ) | Kích thước: 10cm - 12cm | Tag: Hot Trend | Đánh giá: 5.0★ (48 reviews)\n"
                "  - Mô tả: Bộ sưu tập móc khóa gấu bông nhỏ xinh với đủ tạo hình biểu cảm siêu ngộ nghĩnh và đáng yêu. Thích hợp treo balo, túi xách, chìa khóa xe hoặc làm quà tặng bạn bè.\n\n"
                "• sp-02: Hội Bạn Thân Opanchu Usagi\n"
                "  - Giá: 85.000đ (Giá gốc: 110.000đ) | Kích thước: 12cm - 15cm | Tag: Bán Chạy | Đánh giá: 5.0★ (62 reviews)\n"
                "  - Mô tả: Hội bạn thân Opanchu Usagi cực hot với những biểu cảm 'dở khóc dở cười' đặc trưng! Chất liệu siêu êm tay, nhỏ gọn dễ dàng mang theo bên mình.\n\n"
                "• sp-03: Bộ Ba Gấu We Bare Bears Nhỏ\n"
                "  - Giá: 95.000đ (Giá gốc: 120.000đ) | Kích thước: 15cm | Tag: Mới Về | Đánh giá: 4.9★ (35 reviews)\n"
                "  - Mô tả: Bộ ba gấu xám Grizzly, gấu trúc Panda và gấu trắng Ice Bear đã cập bến Ô MƠ! Kích thước mini cực đáng yêu, để bàn học hay góc làm việc rất chill.\n\n"
                "• sp-04: Móc Khóa Sumikko Gurashi\n"
                "  - Giá: 55.000đ (Giá gốc: 75.000đ) | Kích thước: 8cm - 10cm | Tag: Yêu Thích | Đánh giá: 4.8★ (29 reviews)\n"
                "  - Mô tả: Gia đình góc nhỏ Sumikko đáng yêu xỉu, biểu cảm ngây thơ giúp xoa dịu tâm trạng mỗi khi bạn ngắm nhìn.",
                "• Vỏ bọc: Vải nhung tuyết co giãn 4 chiều mềm mịn, không xơ cứng, không rụng lông khi ma sát.\n"
                "• Ruột bông: Bông gòn vi sợi PP cao cấp siêu êm, độ đàn hồi cao.\n"
                "• Phụ kiện: Móc khóa kim loại mạ tĩnh điện chống rỉ sét, móc cài chắc chắn chịu lực tốt."
            ),
            (
                "🌟 Phân nhóm 2: Gấu vừa (20cm - 30cm)\n(Mã danh mục: gau-vua)", 
                "Khoảng giá: 175.000đ - 210.000đ (6 sản phẩm)\n\n"
                "• sp-05: Gấu Bông Kirby Tròn Ủm Hồng\n"
                "  - Giá: 200.000đ (Giá gốc: 240.000đ) | Kích thước: 30cm | Tag: Best Seller | Đánh giá: 5.0★ (112 reviews)\n"
                "  - Mô tả: Chú Kirby hồng tròn ủm siêu đáng yêu với kích thước 30cm cực thích hợp để ôm trọn vào lòng khi ngủ hoặc tựa lưng khi học tập.\n\n"
                "• sp-06: Minions - Chú Bé Bob Đáng Yêu\n"
                "  - Giá: 185.000đ (Giá gốc: 220.000đ) | Kích thước: 25cm | Tag: Nổi Bật | Đánh giá: 4.9★ (84 reviews)\n"
                "  - Mô tả: Thiết kế ngộ nghĩnh, chất liệu cao cấp an toàn, sẵn sàng mang đến tiếng cười và sự vui tươi cho căn phòng của bạn.\n\n"
                "• sp-07: Cún Tai Dài Cinnamoroll Trắng Muốt\n"
                "  - Giá: 210.000đ (Giá gốc: 250.000đ) | Kích thước: 30cm | Tag: Sanrio Hot | Đánh giá: 5.0★ (95 reviews)\n"
                "  - Mô tả: Chú cún tai dài Cinnamoroll trắng muốt, êm ái như một đám mây nhỏ với kích thước 30cm. Sắc màu nhẹ nhàng, kiểu dáng đáng yêu chuẩn phong cách Nhật.\n\n"
                "• sp-08: Winnie the Pooh Áo Đỏ Ôm Mật\n"
                "  - Giá: 195.000đ (Giá gốc: 230.000đ) | Kích thước: 28cm | Tag: Disney Classic | Đánh giá: 5.0★ (130 reviews)\n"
                "  - Mô tả: Chú gấu Pooh mê mật ong truyền thống với chiếc áo đỏ quen thuộc và chất lông xù mềm mại, ấm áp.\n\n"
                "• sp-09: Gudetama - Trứng Lười Dễ Thương\n"
                "  - Giá: 175.000đ (Giá gốc: 210.000đ) | Kích thước: 25cm | Tag: Hài Hước | Đánh giá: 4.8★ (47 reviews)\n"
                "  - Mô tả: Thánh lười Gudetama với biểu cảm thảnh thơi vô cùng giải trí! Giúp giải tỏa căng thẳng sau một ngày dài làm việc.\n\n"
                "• sp-10: Cá Ngố Hangyodon Độc Lạ\n"
                "  - Giá: 190.000đ (Giá gốc: 230.000đ) | Kích thước: 28cm | Tag: Độc Lạ | Đánh giá: 4.9★ (53 reviews)\n"
                "  - Mô tả: Chú cá ngố Hangyodon độc lạ với đôi mắt tròn xoe và chiếc miệng cá ngộ nghĩnh, biểu cảm cực kỳ đáng yêu.",
                "• Vỏ bọc: Vỏ nhung pha lê cao cấp kết hợp lông thỏ mềm mại, lông xù xoắn dịu nhẹ, an toàn tuyệt đối cho làn da nhạy cảm.\n"
                "• Ruột bông: Bông gòn bi 7D tinh khiết 100% nhập khẩu, khả năng phục hồi form dáng tuyệt đối, chống xẹp lún sau thời gian dài sử dụng.\n"
                "• May thêu: Đường chỉ may đôi giấu mũi tinh xảo, mắt mũi thêu thủ công sắc nét không phai màu."
            ),
            (
                "🌟 Phân nhóm 3: Gấu to (30cm trở lên)\n(Mã danh mục: gau-to)", 
                "Khoảng giá: 280.000đ - 490.000đ (5 sản phẩm)\n\n"
                "• sp-11: My Melody Nơ Hồng Xinh Xắn\n"
                "  - Giá: 290.000đ (Giá gốc: 350.000đ) | Kích thước: 45cm | Tag: Siêu Hot | Đánh giá: 5.0★ (156 reviews)\n"
                "  - Mô tả: Bạn thỏ hồng My Melody dịu dàng đáng yêu, kích thước 45cm ôm cực đã tay. Món quà hoàn hảo để dành tặng bạn gái hay em nhỏ.\n\n"
                "• sp-12: Chuột Mickey Mouse Cổ Điển\n"
                "  - Giá: 280.000đ (Giá gốc: 330.000đ) | Kích thước: 40cm | Tag: Disney Classic | Đánh giá: 4.9★ (78 reviews)\n"
                "  - Mô tả: Chú chuột Mickey kinh điển của nhà Disney với dáng ngồi gọn gàng, khuôn mặt rạng rỡ, chất lượng hoàn thiện tuyệt hảo.\n\n"
                "• sp-13: Cáo Tuyết Kitsune Nhật Bản\n"
                "  - Giá: 320.000đ (Giá gốc: 380.000đ) | Kích thước: 50cm | Tag: Độc Quyền | Đánh giá: 5.0★ (89 reviews)\n"
                "  - Mô tả: Bé cáo tuyết Kitsune mang nét đẹp thần thoại Nhật Bản với bộ lông trắng mướt mềm mại, đôi tai hồng đáng yêu.\n\n"
                "• sp-14: Cún Shiba Inu Béo Tròn\n"
                "  - Giá: 310.000đ (Giá gốc: 370.000đ) | Kích thước: 50cm | Tag: Yêu Thích | Đánh giá: 5.0★ (142 reviews)\n"
                "  - Mô tả: Bé cún Shiba Inu má phúng phính, thân hình mập mạp tròn trịa, cảm giác ôm vào lòng cực kỳ ấm áp và bình yên.\n\n"
                "• sp-15: Gấu Bông Khổng Lồ Teddy ÔMƠ Vỗ Về\n"
                "  - Giá: 490.000đ (Giá gốc: 590.000đ) | Kích thước: 80cm - 100cm | Tag: Khổng Lồ (Signature) | Đánh giá: 5.0★ (210 reviews)\n"
                "  - Mô tả: Chú gấu Teddy khổng lồ signature của ÔMƠ! Thay bạn trao cái ôm siết dịu dàng và ấm áp nhất cho người thương.",
                "• Vỏ bọc: Vải lông thỏ siêu mịn cao cấp kết hợp nỉ lông cừu nhân tạo ấm áp, áo len dệt thủ công có thể tháo rời giặt giũ dễ dàng.\n"
                "• Ruột bông: Bông gòn 3D tinh khiết kháng khuẩn đạt chuẩn kiểm định an toàn dệt may, nhồi căng tròn tạo cảm giác ôm đầm tay.\n"
                "• Đóng gói: Hỗ trợ nén hút chân không chuyên nghiệp, đóng hộp Kraft Vintage cao cấp kèm thiệp viết tay theo yêu cầu."
            )
        ]

        for r_idx, (cat_name, cat_desc, cat_mat) in enumerate(prod_exact_3cat, start=1):
            if r_idx < len(t1.rows):
                format_cell(t1.rows[r_idx].cells[0], cat_name, bold=True, size_pt=9.5)
                format_cell(t1.rows[r_idx].cells[1], cat_desc, size_pt=9.0)
                if len(t1.rows[r_idx].cells) > 2:
                    format_cell(t1.rows[r_idx].cells[2], cat_mat, size_pt=9.0)

    # -------------------------------------------------------------
    # 3. FIX INCONSISTENT PARAGRAPHS & ALIGN ALL DESCRIPTIONS
    # -------------------------------------------------------------
    print("Fixing all narrative descriptions in paragraphs...")

    tech_spec_text = (
        "KIẾN TRÚC VÀ CÔNG NGHỆ TRIỂN KHAI TOÀN DIỆN TRÊN HỆ THỐNG WEBSITE ÔMƠ:\n\n"
        "1. Nền Tảng Công Nghệ Phía Client (Frontend Architecture):\n"
        "• Cấu trúc ngữ nghĩa chuẩn HTML5: Sử dụng đầy đủ các thẻ ngữ nghĩa (<header>, <nav>, <main>, <section>, <article>, <aside>, <footer>) giúp tối ưu hóa khả năng lập chỉ mục của công cụ tìm kiếm Googlebot.\n"
        "• Hệ thống giao diện CSS3 Vanilla thuần túy (assets/css/style.css – ~69KB): Thiết kế theo kiến trúc CSS BEM, tích hợp CSS Variables bảng màu Vintage Pastel đồng bộ, hệ thống bố cục lưới CSS Grid & Flexbox linh hoạt, hiệu ứng chuyển động mượt mà (keyframes float, pulse, toast-in, reveal-on-scroll) không phụ thuộc framework nặng nề, tối ưu điểm Core Web Vitals.\n"
        "• Xử lý logic Vanilla JavaScript (js/products.js, js/cart.js):\n"
        "  - Quản lý trạng thái giỏ hàng thời gian thực (Cart State Management) qua HTML5 LocalStorage API (khóa omo_cart_items_v1), tự động tính toán tổng phụ, phí vận chuyển (miễn phí đơn từ 300.000 VNĐ; phí ship 30.000 VNĐ cho đơn dưới 300.000 VNĐ) và lưu vết phiên mua sắm.\n"
        "  - Hệ thống mã khuyến mãi (Coupon Engine): Hỗ trợ các mã giảm giá thực tế trong mã nguồn gồm OMOYEUTHUONG (giảm 10% tổng đơn), OMOCHU (giảm 10% tổng đơn), FREESHIP (giảm 5% tổng đơn).\n"
        "  - Danh mục yêu thích (Wishlist Storage): Quản lý qua khóa omo_wishlist_items_v1 cho phép khách hàng lưu sản phẩm quan tâm qua nút Floating FAB.\n"
        "  - Trợ lý ảo Chatbot AI 'Trợ lý Mơ' (omo_chat_messages_v1): Tích hợp giao diện chat nổi góc dưới màn hình, tự động nhận diện từ khóa (giá, ship, đổi trả, tư vấn quà) và phản hồi tức thì 24/7.\n"
        "  - Cổng thanh toán quét mã VietQR MB Bank & Ví MoMo API: Tự động trích xuất số tiền, mã đơn hàng (#OMO-XXXXX) và tạo mã QR chuyển khoản trực tiếp hỗ trợ thanh toán 1 chạm.\n\n"
        "2. Nền Tảng Máy Chủ & Môi Trường Triển Khai (Backend & Infrastructure):\n"
        "• Web Server Node.js (server.js): Máy chủ HTTP nhẹ phục vụ các tệp tĩnh với MIME Types đầy đủ, chạy trên cổng PORT 3000 phục vụ phát triển cục bộ và kiểm thử.\n"
        "• Nền tảng Đám mây Vercel Hosting (vercel.json): Triển khai toàn bộ mã nguồn lên Vercel tại địa chỉ chính thức https://omo13.vercel.app/ với chứng chỉ SSL tự động gia hạn, CDN Edge Network phân tán tốc độ cao và uptime 99.99%."
    )

    pages_deep_text = (
        "PHÂN TÍCH CHI TIẾT TỪNG PHÂN HỆ VÀ GIAO DIỆN 8 TRANG CHÍNH TRÊN HỆ THỐNG:\n\n"
        "📄 1. Trang Chủ (index.html – URL: /index.html):\n"
        "• Thẻ Title chuẩn SEO: 'ÔMƠ – Ôm trọn vỗ về, giữ trọn giấc mơ | Tiệm Gấu Bông Đà Nẵng' (61 ký tự).\n"
        "• Meta Description: 'Tiệm gấu ÔMƠ chuyên cung cấp các mẫu gấu bông cao cấp, êm ái, an toàn từ móc khóa xinh xắn đến gấu khổng lồ vỗ về giấc mơ của bạn.' (130 ký tự).\n"
        "• Header & Navigation: Logo ÔMƠ nhận diện thương hiệu, menu điều hướng 8 trang trung tâm với hiệu ứng hover pastel, bên phải là thanh tìm kiếm và biểu tượng Giỏ hàng có badge số lượng real-time.\n"
        "• Hero Section: Banner chính phong cách Vintage Pastel hiển thị hình ảnh gấu bông cùng slogan 'Ôm một chú gấu, giữ một giấc mơ' và nút CTA 'KHÁM PHÁ CỬA HÀNG'.\n"
        "• Section 'Sản Phẩm Được Yêu Thích': Trưng bày các sản phẩm bán chạy (Kirby hồng, Opanchu Usagi, Cinnamoroll) kèm tag giá ưu đãi và nút 'Thêm Vào Giỏ'.\n"
        "• Section 'Vì Sao Nên Chọn Chúng Tớ?': Nêu bật 4 cam kết vàng (Gòn PP 100% tinh khiết kháng khuẩn, Vải nhung tuyết mịn êm, Đóng gói hộp quà Vintage thủ công, Giao hỏa tốc 1-2h tại Đà Nẵng).\n"
        "• Section 'Khoảnh Khắc Đáng Yêu Cùng Gấu Bông': Khối hình ảnh khách hàng feedback thực tế gia tăng độ tin cậy E-E-A-T.\n"
        "• Section 'Tin Tức & Cẩm Nang Gấu Bông': Trích dẫn 3 bài viết nổi bật hỗ trợ điều hướng SEO Silo.\n"
        "• Footer 4 Cột: Đồng bộ thông tin thực thể Entity, Google Maps, chính sách và liên kết mạng xã hội.\n\n"
        "📄 2. Trang Cửa Hàng (cua-hang.html – URL: /cua-hang.html):\n"
        "• Thẻ Title: 'Cửa Hàng Gấu Bông ÔMƠ – Đầy Đủ Mẫu Mã Xinh Xắn & Cao Cấp' (56 ký tự).\n"
        "• Meta Description: 'Khám phá toàn bộ bộ sưu tập gấu bông ÔMƠ: Móc khóa gấu nhỏ (5-15cm), gấu vừa ôm (20-30cm) và gấu bông khổng lồ mềm mại (30cm+).' (127 ký tự).\n"
        "• Bộ Lọc Đa Tiêu Chí (Multi-Filter): 3 tab phân nhóm kích thước (Tất Cả Mẫu, Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+) kết hợp ô tìm kiếm real-time và bộ sắp xếp giá (Thấp đến Cao, Cao đến Thấp, Đánh giá cao nhất).\n"
        "• Product Grid: Hiển thị trọn vẹn 15 sản phẩm dạng lưới 3-4 cột với ảnh chất lượng cao, tên chuẩn SEO, nhãn Hot/Sale, rating 5 sao và nút CTA mua hàng.\n"
        "• Quick View Modal: Cửa sổ popup xem nhanh thông số chất liệu, kích thước, ảnh phóng to, bộ chọn số lượng và nút 'Thêm Vào Giỏ Hàng'.\n\n"
        "📄 3. Trang Giới Thiệu (gioi-thieu.html – URL: /gioi-thieu.html):\n"
        "• Thẻ Title: 'Giới Thiệu Về ÔMƠ – Ôm một chú gấu, giữ một giấc mơ' (51 ký tự).\n"
        "• Meta Description: 'Khám phá câu chuyện hình thành và sứ mệnh lan tỏa những cái ôm ấm áp của tiệm gấu ÔMƠ Đà Nẵng.' (94 ký tự).\n"
        "• Storytelling: Kể câu chuyện hình thành thương hiệu ÔMƠ và sứ mệnh lan tỏa những cái ôm xoa dịu cảm xúc cho giới trẻ.\n"
        "• Hệ Giá Trị Cốt Lõi 5T và Giới thiệu đội ngũ 6 thành viên sáng lập.\n"
        "• Quy trình đóng gói quà tặng thủ công với hộp giấy Kraft và thiệp hoa khô viết tay theo yêu cầu.\n\n"
        "📄 4. Trang Tin Tức & Cẩm Nang (tin-tuc.html – URL: /tin-tuc.html):\n"
        "• Thẻ Title: 'Tin Tức & Cẩm Nang Gấu Bông – ÔMƠ Shop' (38 ký tự).\n"
        "• Meta Description: 'Khám phá các mẹo chọn quà tặng gấu bông, cẩm nang vệ sinh và bảo quản gấu bông luôn thơm mềm tại nhà.' (101 ký tự).\n"
        "• Bài viết 1: '5 Cách Chọn Gấu Bông Làm Quà Tặng Phù Hợp Cho Người Thương' (Tác giả: ÔMƠ Team, Ngày đăng: 15/04/2026, Chuyên mục: Mẹo Chọn Quà).\n"
        "• Bài viết 2: 'Cách Vệ Sinh Và Bảo Quản Gấu Bông Đúng Cách Tại Nhà Luôn Thơm Tho' (Tác giả: ÔMƠ Care, Ngày đăng: 10/04/2026, Chuyên mục: Chăm Sóc Gấu Bông).\n"
        "• Bài viết 3: 'Gấu Bông Có Ý Nghĩa Gì Khi Làm Quà Tặng Người Yêu & Bạn Bè?' (Tác giả: ÔMƠ Story, Ngày đăng: 02/04/2026, Chuyên mục: Ý Nghĩa Yêu Thương).\n"
        "• Tích hợp đầy đủ Mục lục tự động (TOC), hình ảnh minh họa chân thực có thẻ Alt và liên kết nội bộ trỏ về danh mục sản phẩm.\n\n"
        "📄 5. Trang Giỏ Hàng (gio-hang.html – URL: /gio-hang.html):\n"
        "• Thẻ Title: 'Giỏ Hàng Của Bạn – ÔMƠ Shop' (27 ký tự).\n"
        "• Meta Description: 'Xem và quản lý các sản phẩm gấu bông trong giỏ hàng của bạn tại ÔMƠ Shop.' (73 ký tự).\n"
        "• Bảng danh sách sản phẩm: Hiển thị thumbnail, tên sản phẩm, đơn giá, bộ tăng giảm số lượng + -, nút xóa X, tổng tiền từng món.\n"
        "• Thanh tiến trình miễn phí vận chuyển (Freeship Progress Bar): Thông báo miễn phí vận chuyển cho đơn từ 300.000 VNĐ (nhắc nhở khách mua thêm số tiền còn thiếu).\n"
        "• Ô nhập mã giảm giá (Coupon Box): Áp dụng ngay voucher OMOYEUTHUONG (-10%), OMOCHU (-10%) hoặc FREESHIP (-5%).\n\n"
        "📄 6. Trang Thanh Toán (thanh-toan.html – URL: /thanh-toan.html):\n"
        "• Thẻ Title: 'Thanh Toán Đơn Hàng – ÔMƠ Shop' (30 ký tự).\n"
        "• Meta Description: 'Thanh toán an toàn, tiện lợi qua Chuyển khoản QR ngân hàng, Ví MoMo hoặc Tiền mặt khi nhận hàng (COD).' (102 ký tự).\n"
        "• Form thu thập thông tin người nhận: Họ tên, Số điện thoại, Địa chỉ nhận hàng chi tiết, Ghi chú viết thiệp tặng.\n"
        "• Tùy chọn 3 phương thức thanh toán: (1) Quét mã VietQR chuyển khoản MB Bank tự động (STK: 022456789999 - TIỆM GẤU ÔMƠ); (2) Ví điện tử MoMo; (3) Thanh toán tiền mặt khi nhận hàng (COD).\n"
        "• Popup xác nhận đặt hàng thành công (#OMO-XXXXX) và lưu thông tin đơn hàng vào LocalStorage.\n\n"
        "📄 7. Trang Chính Sách (chinh-sach.html – URL: /chinh-sach.html):\n"
        "• Thẻ Title: 'Chính Sách Mua Hàng & Đổi Trả – ÔMƠ Shop' (40 ký tự).\n"
        "• Meta Description: 'Chính sách mua hàng, giao nhận hỏa tốc và quy định đổi trả trong vòng 3 ngày tại tiệm gấu ÔMƠ Đà Nẵng.' (102 ký tự).\n"
        "• Chính sách đổi trả 1-1 trong 3 ngày nếu phát hiện lỗi đường may hoặc lỗi nhà sản xuất.\n"
        "• Chính sách giao hàng hỏa tốc trong 1-2 giờ nội thành Đà Nẵng và giao nhanh COD toàn quốc 2-4 ngày.\n"
        "• Chính sách bảo mật thông tin khách hàng tuyệt đối và cam kết chất lượng gòn tinh khiết 100%.\n\n"
        "📄 8. Trang Liên Hệ (lien-he.html – URL: /lien-he.html):\n"
        "• Thẻ Title: 'Liên Hệ Tiệm Gấu ÔMƠ – 255 Hà Huy Tập, Đà Nẵng' (46 ký tự).\n"
        "• Meta Description: 'Liên hệ tiệm gấu bông ÔMƠ tại Đà Nẵng. Hotline: 222-456-789. Địa chỉ: 255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng.' (112 ký tự).\n"
        "• Bản đồ Google Maps tương tác nhúng trực quan định vị cửa hàng tại 255 Hà Huy Tập, Thanh Khê, Đà Nẵng.\n"
        "• Thông tin liên hệ: Hotline: 022456789 / 222-456-789 | Email: omoshop@gmail.com | Địa chỉ: 255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng.\n"
        "• Form gửi tin nhắn tư vấn và liên kết trực tiếp tới Fanpage Facebook, Instagram và kênh Zalo OA."
    )

    for p in doc.paragraphs:
        txt = p.text.strip()
        
        # Update tech spec
        if "KIẾN TRÚC VÀ CÔNG NGHỆ TRIỂN KHAI TOÀN DIỆN" in txt:
            p.text = tech_spec_text
            for r in p.runs:
                set_font_tnr(r, size_pt=13)

        # Update 8 pages deep breakdown
        if "PHÂN TÍCH CHI TIẾT TỪNG PHÂN HỆ VÀ GIAO DIỆN 8 TRANG CHÍNH" in txt:
            p.text = pages_deep_text
            for r in p.runs:
                set_font_tnr(r, size_pt=13)

        # Fix wrong address occurrences in paragraphs
        if "137 Nguyễn Thị Thập" in txt:
            p.text = txt.replace("137 Nguyễn Thị Thập, Phường Hòa Minh, Quận Liên Chiểu, TP. Đà Nẵng", "255 Hà Huy Tập, Phường Hòa Khê, Quận Thanh Khê, TP. Đà Nẵng").replace("137 Nguyễn Thị Thập, Liên Chiểu, Đà Nẵng", "255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng").replace("137 Nguyễn Thị Thập", "255 Hà Huy Tập")
            for r in p.runs:
                set_font_tnr(r, size_pt=13)

        # Fix wrong category mentions in Chapter 1 / Chapter 2
        if "4 phân nhóm" in txt:
            p.text = txt.replace("4 phân nhóm", "3 phân nhóm kích thước chính (Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+)")
            for r in p.runs:
                set_font_tnr(r, size_pt=13)

        # Fix 4P price paragraph
        if "chiến lược định giá thâm nhập thị trường" in txt.lower():
            p.text = (
                "ÔMƠ áp dụng chiến lược định giá thâm nhập thị trường (Penetration Pricing) kết hợp định giá theo giá trị cảm nhận (Value-based Pricing): "
                "Mức giá dao động từ 45.000 VNĐ (Móc khóa gấu nhỏ 5-15cm) đến 490.000 VNĐ (Gấu bông khổng lồ Teddy ÔMƠ 80-100cm), "
                "cực kỳ cạnh tranh và phù hợp với túi tiền của học sinh, sinh viên và người đi làm trẻ tuổi. "
                "Bên cạnh đó, tiệm triển khai các gói định giá thông minh: Combo Quà Tặng (Gấu bông + Hộp quà Kraft + Thiệp hoa khô viết tay), "
                "chính sách Miễn phí vận chuyển (Freeship) cho đơn hàng từ 300.000 VNĐ (phí ship tiêu chuẩn 30.000 VNĐ cho đơn dưới 300.000 VNĐ) "
                "và hệ thống mã giảm giá ưu đãi (OMOYEUTHUONG giảm 10%, OMOCHU giảm 10%, FREESHIP giảm 5%)."
            )
            for r in p.runs:
                set_font_tnr(r, size_pt=13)

    # -------------------------------------------------------------
    # 4. FIX TABLES WITH INACCURATE DESCRIPTIONS
    # -------------------------------------------------------------
    print("Fixing table descriptions across the document...")

    # Table 17 (Bố cục giỏ hàng)
    if len(tables) > 17:
        t17 = tables[17]
        for r in t17.rows:
            for c in r.cells:
                if "250k" in c.text or "WELCOMEOMO" in c.text:
                    c.text = c.text.replace("250k", "300.000đ (phí ship tiêu chuẩn 30.000đ)").replace("WELCOMEOMO", "OMOYEUTHUONG (-10%), OMOCHU (-10%), FREESHIP (-5%)")
                    format_cell(c, c.text, size_pt=10)

    # Table 26 (Plugins)
    if len(tables) > 26:
        t26 = tables[26]
        for r in t26.rows:
            for c in r.cells:
                if "WELCOMEOMO" in c.text or "250k" in c.text:
                    c.text = c.text.replace("WELCOMEOMO", "OMOYEUTHUONG (10%), OMOCHU (10%), FREESHIP (5%)").replace("250k", "300.000đ")
                    format_cell(c, c.text, size_pt=10)

    # Table 29 (Entity khai báo)
    if len(tables) > 29:
        t29 = tables[29]
        for r in t29.rows:
            for c in r.cells:
                if "137 Nguyễn Thị Thập" in c.text:
                    c.text = c.text.replace("137 Nguyễn Thị Thập, Phường Hòa Minh, Quận Liên Chiểu, TP. Đà Nẵng", "255 Hà Huy Tập, Phường Hòa Khê, Quận Thanh Khê, TP. Đà Nẵng").replace("137 Nguyễn Thị Thập", "255 Hà Huy Tập")
                    format_cell(c, c.text, size_pt=10)

    # -------------------------------------------------------------
    # 5. APPLY GLOBAL TIMES NEW ROMAN & SAVE
    # -------------------------------------------------------------
    print("Applying global Times New Roman styling to all paragraphs, headings and tables...")
    apply_global_times_new_roman(doc)

    out1 = 'DoAnNhom13_IS425_OMO.docx'
    out2 = 'BAO_CAO_DO_AN_IS425_NHOM13_100_TRANG.docx'
    out3 = 'NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx'

    print(f"Saving final report to {out1}...")
    doc.save(out1)
    doc.save(out2)
    doc.save(out3)
    print("All files saved successfully!")

if __name__ == '__main__':
    fix_all_descriptions_and_align_with_source()
