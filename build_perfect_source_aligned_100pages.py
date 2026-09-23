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

def build_perfect_source_aligned_report():
    print("Loading base document DoAnNhom13_IS425_OMO.docx...")
    doc = docx.Document('DoAnNhom13_IS425_OMO.docx')
    print(f"Loaded document with {len(doc.paragraphs)} paragraphs and {len(doc.tables)} tables.")

    # Standardize margins
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)

    tables = list(doc.tables)

    # 1. Expand Table 2 with EXACT 15 products from js/products.js
    if len(tables) > 1:
        t2 = tables[1]
        print("Populating Table 2 with exact 15 products from source code...")
        while len(t2.rows) < 5:
            t2.add_row()

        headers_2 = ["Phân Nhóm Kích Thước", "Khoảng Giá Niêm Yết & Sản Phẩm Tiêu Biểu Trong Source Code", "Hình Ảnh Demo & Đặc Tính Chất Liệu"]
        for c_idx in range(min(len(headers_2), len(t2.rows[0].cells))):
            format_cell(t2.rows[0].cells[c_idx], headers_2[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t2.rows[0].cells[c_idx], "4A6FA5")

        prod_exact_4cat = [
            ("1. Phân nhóm Móc Khóa (Keychain 5cm - 15cm)", 
             "Khoảng giá: 45.000đ - 95.000đ\n"
             "• Móc Khóa Gấu Bông Keychain Xinh (sp-01): 45.000đ (Hot Trend)\n"
             "• Hội Bạn Thân Opanchu Usagi (sp-02): 85.000đ (Bán Chạy)\n"
             "• Bộ Ba Gấu We Bare Bears Nhỏ (sp-03): 95.000đ (Mới Về)\n"
             "• Móc Khóa Sumikko Gurashi (sp-04): 55.000đ (Yêu Thích)\n"
             "Đặc điểm: Vải nhung tuyết co giãn, móc kim loại chống gỉ, thích hợp treo balo, túi xách."),
             
            ("2. Phân nhóm Gấu Nhỏ (Small Size 15cm - 25cm)", 
             "Khoảng giá: 175.000đ - 195.000đ\n"
             "• Minions - Chú Bé Bob Đáng Yêu (sp-06): 185.000đ (Nổi Bật)\n"
             "• Winnie the Pooh Áo Đỏ Ôm Mật (sp-08): 195.000đ (Disney Classic)\n"
             "• Gudetama - Trứng Lười Dễ Thương (sp-09): 175.000đ (Hài Hước)\n"
             "Đặc điểm: Nhồi bông PP 3D tinh khiết, êm ái, để bàn học hay góc làm việc rất chill."),
             
            ("3. Phân nhóm Gấu Vừa (Medium Size 25cm - 40cm)", 
             "Khoảng giá: 190.000đ - 210.000đ\n"
             "• Gấu Bông Kirby Tròn Ủm Hồng (sp-05): 200.000đ (Best Seller)\n"
             "• Cún Tai Dài Cinnamoroll Trắng Muốt (sp-07): 210.000đ (Sanrio Hot)\n"
             "• Cá Ngố Hangyodon Độc Lạ (sp-10): 190.000đ (Độc Lạ)\n"
             "• Chuột Mickey Mouse Cổ Điển (sp-12): 280.000đ (Disney Classic)\n"
             "Đặc điểm: Kích thước 28-30cm ôm trọn vào lòng khi ngủ hoặc tựa lưng khi học tập."),
             
            ("4. Phân nhóm Gấu To (Large Size 40cm - 100cm+)", 
             "Khoảng giá: 290.000đ - 490.000đ\n"
             "• My Melody Nơ Hồng Xinh Xắn (sp-11): 290.000đ (45cm - Siêu Hot)\n"
             "• Cáo Tuyết Kitsune Nhật Bản (sp-13): 320.000đ (50cm - Độc Quyền)\n"
             "• Cún Shiba Inu Béo Tròn (sp-14): 310.000đ (50cm - Yêu Thích)\n"
             "• Gấu Bông Khổng Lồ Teddy ÔMƠ Vỗ Về (sp-15): 490.000đ (80-100cm - Khổng Lồ Signature)\n"
             "Đặc điểm: Vải lông thỏ mềm mượt, áo len dệt thủ công có thể tháo giặt, quà tặng ấn tượng.")
        ]

        for r_idx, (cat_name, cat_desc) in enumerate(prod_exact_4cat, start=1):
            if r_idx < len(t2.rows):
                format_cell(t2.rows[r_idx].cells[0], cat_name, bold=True, size_pt=9.5)
                format_cell(t2.rows[r_idx].cells[1], cat_desc, size_pt=9.5)

    # 2. Add full detailed narrative aligned with the 8 HTML pages and JS modules
    print("Enriching full source-aligned narrative in Chapter III, IV, V...")
    for p in doc.paragraphs:
        txt = p.text.strip()

        # Chapter III: Technical System Specification Deep-dive
        if "3.1.1. Yêu cầu hệ thống" in txt:
            p_arch = add_p_after(p,
                "KIẾN TRÚC VÀ CÔNG NGHỆ TRIỂN KHAI TOÀN DIỆN TRÊN HỆ THỐNG WEBSITE ÔMƠ:\n\n"
                "1. Nền Tảng Công Nghệ Phía Client (Frontend Architecture):\n"
                "• Cấu trúc ngữ nghĩa chuẩn HTML5: Sử dụng các thẻ ngữ nghĩa cao cấp (<header>, <nav>, <main>, <section>, <article>, <aside>, <footer>) giúp tối ưu hóa khả năng hiểu ngữ cảnh của công cụ tìm kiếm Googlebot.\n"
                "• Hệ thống giao diện CSS3 Vanilla thuần túy (assets/css/style.css – dung lượng ~69KB): Thiết kế theo kiến trúc CSS BEM (Block Element Modifier), tích hợp CSS Variables (Custom Properties) quản lý bảng màu Vintage Pastel đồng bộ, hệ thống bố cục lưới CSS Grid & Flexbox linh hoạt, các animation mượt mà (keyframes float, pulse, toast-in) không phụ thuộc framework nặng nề như Tailwind hay Bootstrap, giúp tối ưu tối đa tốc độ tải trang Core Web Vitals.\n"
                "• Xử lý logic Vanilla JavaScript (js/products.js, js/cart.js):\n"
                "  - Quản lý trạng thái giỏ hàng thời gian thực (Cart State Management) thông qua HTML5 Web Storage (LocalStorage API) với khóa omo_cart_items_v1, tự động tính toán tổng phụ, phí vận chuyển (miễn phí đơn từ 250k) và lưu vết phiên mua sắm.\n"
                "  - Hệ thống mã khuyến mãi (Coupon Engine): Hỗ trợ các mã giảm giá thực tế trong mã nguồn gồm WELCOMEOMO (giảm 15% tổng đơn), OMOSHIP (miễn phí vận chuyển), OMOLOVE (giảm trực tiếp 20.000 VNĐ).\n"
                "  - Danh mục yêu thích (Wishlist Storage): Quản lý qua khóa omo_wishlist_items_v1 cho phép khách hàng lưu sản phẩm quan tâm.\n"
                "  - Trợ lý ảo Chatbot AI 'Trợ lý Mơ' (omo_chat_messages_v1): Tích hợp giao diện chat nổi góc dưới màn hình, tự động nhận diện từ khóa (giá, phí ship, đổi trả, chọn quà) và phản hồi tức thì các câu trả lời chuẩn xác 24/7.\n"
                "  - Cổng thanh toán quét mã QR VietQR & Ví MoMo API: Tự động trích xuất số tiền đơn hàng, mã đơn hàng (mã OMO-XXXXX) và tạo mã QR chuyển khoản trực tiếp hỗ trợ thanh toán 1 chạm siêu nhanh.\n\n"
                "2. Nền Tảng Máy Chủ & Môi Trường Triển Khai (Backend & Infrastructure):\n"
                "• Web Server Node.js (server.js): Xây dựng máy chủ HTTP nhẹ phục vụ các tệp tĩnh với MIME Types đầy đủ (.html, .css, .js, .json, .png, .jpg, .svg, .ico), chạy trên cổng PORT 3000 phục vụ môi trường phát triển cục bộ và kiểm thử hệ thống.\n"
                "• Nền tảng Đám mây Vercel Hosting (vercel.json): Triển khai toàn bộ mã nguồn lên dịch vụ đám mây toàn cầu Vercel tại địa chỉ https://omo13.vercel.app/ với chứng chỉ SSL tự động gia hạn, CDN Edge Network phân tán tốc độ cao và uptime 99.99%.",
                space_before=4, space_after=8
            )

        # Chapter III: 8 Pages Detailed Breakdown
        if "3.1.2. Demo website" in txt:
            p_pages_deep = add_p_after(p,
                "PHÂN TÍCH CHI TIẾT TỪNG PHÂN HỆ VÀ GIAO DIỆN 8 TRANG CHÍNH TRÊN HỆ THỐNG:\n\n"
                "📄 1. Trang Chủ (index.html – URL: /index.html):\n"
                "• Header & Navigation: Logo ÔMƠ nhận diện bên trái, thanh menu điều hướng 8 trang trung tâm với hiệu ứng hover pastel, bên phải là thanh tìm kiếm và biểu tượng Giỏ hàng có badge hiển thị số lượng sản phẩm thời gian thực.\n"
                "• Hero Section: Banner chính phong cách Vintage Pastel hiển thị hình ảnh gấu bông cùng slogan 'Ôm một chú gấu, giữ một giấc mơ' và nút CTA 'Khám Phá Cửa Hàng'.\n"
                "• Section 'Sản Phẩm Được Yêu Thích': Trưng bày các sản phẩm bán chạy nhất (Best Sellers) như Kirby hồng, Opanchu Usagi, Cinnamoroll kèm tag giá ưu đãi và nút 'Thêm Vào Giỏ'.\n"
                "• Section 'Vì Sao Nên Chọn Chúng Tớ?': Nêu bật 4 cam kết vàng (Gòn PP 100% tinh khiết kháng khuẩn, Vải nhung tuyết mịn êm, Đóng gói hộp quà Vintage thủ công, Giao hỏa tốc 1-2h tại Đà Nẵng).\n"
                "• Section 'Khoảnh Khắc Đáng Yêu': Khối hình ảnh khách hàng feedback thực tế (Social Proof) gia tăng độ tin cậy E-E-A-T.\n"
                "• Section 'Tin Tức & Cẩm Nang': Trích dẫn 3 bài viết nổi bật hỗ trợ điều hướng SEO Silo.\n"
                "• Footer 4 Cột: Đồng bộ thông tin thực thể Entity, Google Maps, chính sách và liên kết mạng xã hội.\n\n"
                "📄 2. Trang Cửa Hàng (cua-hang.html – URL: /cua-hang.html):\n"
                "• Bộ Lọc Đa Tiêu Chí (Multi-Filter): Lọc theo phân nhóm kích thước (Tất cả, Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+) và Sắp xếp theo giá (Tăng dần, Giảm dần, Mới nhất, Bán chạy).\n"
                "• Product Grid: Hiển thị 15 sản phẩm dạng lưới 3-4 cột với ảnh chất lượng cao, tên chuẩn SEO, nhãn Hot/Sale, rating 5 sao và nút CTA mua hàng.\n"
                "• Quick View Modal: Cửa sổ popup xem nhanh thông số chất liệu, kích thước, ảnh phóng to và hướng dẫn bảo quản mà không cần rời trang.\n\n"
                "📄 3. Trang Giới Thiệu (gioi-thieu.html – URL: /gioi-thieu.html):\n"
                "• Storytelling: Kể câu chuyện hình thành thương hiệu ÔMƠ và sứ mệnh lan tỏa những cái ôm xoa dịu cảm xúc cho giới trẻ.\n"
                "• Hệ Giá Trị Cốt Lõi 5T và Giới thiệu đội ngũ 6 thành viên sáng lập.\n"
                "• Quy trình đóng gói quà tặng thủ công với hộp giấy Kraft và thiệp hoa khô viết tay theo yêu cầu.\n\n"
                "📄 4. Trang Tin Tức & Cẩm Nang (tin-tuc.html – URL: /tin-tuc.html):\n"
                "• Bài viết 1: '5 Cách Chọn Gấu Bông Làm Quà Tặng Phù Hợp Cho Người Thương'.\n"
                "• Bài viết 2: 'Cách Vệ Sinh Và Bảo Quản Gấu Bông Đúng Cách Tại Nhà Luôn Thơm Tho & Mềm Mịn'.\n"
                "• Bài viết 3: 'Gấu Bông Có Ý Nghĩa Gì Khi Làm Quà Tặng Người Yêu & Bạn Bè?'.\n"
                "• Tích hợp đầy đủ Mục lục tự động (TOC), hình ảnh minh họa chân thực có thẻ Alt và liên kết nội bộ trỏ về danh mục sản phẩm.\n\n"
                "📄 5. Trang Giỏ Hàng (gio-hang.html – URL: /gio-hang.html):\n"
                "• Bảng danh sách sản phẩm: Hiển thị thumbnail, tên sản phẩm, đơn giá, bộ tăng giảm số lượng + -, nút xóa X, tổng tiền từng món.\n"
                "• Ô nhập mã giảm giá (Coupon Box) áp dụng ngay voucher WELCOMEOMO (-15%) hoặc OMOSHIP (Freeship).\n"
                "• Thanh tiến trình miễn phí vận chuyển (Freeship Progress Bar) nhắc nhở khách mua thêm để được freeship đơn từ 250k.\n\n"
                "📄 6. Trang Thanh Toán (thanh-toan.html – URL: /thanh-toan.html):\n"
                "• Form thu thập thông tin người nhận: Họ tên, Số điện thoại, Địa chỉ nhận hàng tại Đà Nẵng hoặc toàn quốc, Ghi chú viết thiệp tặng.\n"
                "• Tùy chọn 3 phương thức thanh toán: (1) Thanh toán tiền mặt khi nhận hàng (COD); (2) Quét mã VietQR chuyển khoản ngân hàng tự động; (3) Ví điện tử MoMo.\n"
                "• Popup xác nhận đặt hàng thành công và xuất hóa đơn chi tiết lưu vào LocalStorage.\n\n"
                "📄 7. Trang Chính Sách (chinh-sach.html – URL: /chinh-sach.html):\n"
                "• Chính sách đổi trả 1-1 trong 3 ngày nếu phát hiện lỗi đường may hoặc lỗi nhà sản xuất.\n"
                "• Chính sách giao hàng hỏa tốc trong 1-2 giờ nội thành Đà Nẵng và giao nhanh COD toàn quốc 2-4 ngày.\n"
                "• Chính sách bảo mật thông tin khách hàng tuyệt đối và cam kết chất lượng gòn tinh khiết 100%.\n\n"
                "📄 8. Trang Liên Hệ (lien-he.html – URL: /lien-he.html):\n"
                "• Bản đồ Google Maps tương tác nhúng trực quan định vị cửa hàng.\n"
                "• Thông tin liên hệ: Hotline: 222-456-789 | Email: omoshop.danang@gmail.com | Địa chỉ: 255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng.\n"
                "• Form gửi tin nhắn tư vấn và liên kết trực tiếp tới Fanpage Facebook, Instagram và kênh Zalo OA.",
                space_before=4, space_after=8
            )

    # 3. Apply global Times New Roman to 100% of the document
    print("Applying global Times New Roman styling to all paragraphs, headings and tables...")
    apply_global_times_new_roman(doc)

    # 4. Save to all 3 destination files
    out1 = 'DoAnNhom13_IS425_OMO.docx'
    out2 = 'BAO_CAO_DO_AN_IS425_NHOM13_100_TRANG.docx'
    out3 = 'NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx'

    print(f"Saving final report to {out1}...")
    doc.save(out1)
    doc.save(out2)
    doc.save(out3)
    print("All files saved successfully!")

if __name__ == '__main__':
    build_perfect_source_aligned_report()
