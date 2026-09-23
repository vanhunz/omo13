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

def perform_complete_body_and_spec_update():
    print("Loading NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx...")
    doc = docx.Document('NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx')
    print(f"Loaded document with {len(doc.paragraphs)} paragraphs and {len(doc.tables)} tables.")

    # 1. Standardize margins
    for s in doc.sections:
        s.top_margin = Cm(2.0)
        s.bottom_margin = Cm(2.0)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.0)

    tables = list(doc.tables)

    # -------------------------------------------------------------
    # 2. UPDATE PARAGRAPHS: TECHNICAL SPECIFICATIONS & SYSTEM REQS
    # -------------------------------------------------------------
    print("Updating system requirements, platform, and full architecture narrative...")

    for i, p in enumerate(doc.paragraphs):
        txt = p.text.strip()

        # Update KPI paragraph if mentions Google Sites
        if "trên google sites" in txt.lower():
            p.text = txt.replace("trên Google Sites", "trên nền tảng Vercel Cloud (omo13.vercel.app)")
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Update Platform
        if txt.startswith("Nền tảng:"):
            p.text = (
                "Nền tảng & Kiến trúc công nghệ: Hệ thống website thương mại điện tử ÔMƠ SHOP được xây dựng bằng kiến trúc Frontend chuẩn W3C: "
                "Cấu trúc Semantic HTML5, hệ thống giao diện CSS3 Vanilla thuần túy theo quy chuẩn BEM (Block Element Modifier) tích hợp CSS Variables đồng bộ bảng màu Vintage Pastel, "
                "kết hợp logic xử lý Vanilla JavaScript ES6+ (quản lý giỏ hàng thời gian thực, Chatbot AI Trợ lý Mơ, cổng thanh toán VietQR & MoMo). "
                "Mã nguồn được biên dịch và triển khai trên dịch vụ Đám mây toàn cầu Vercel Cloud Platform tại địa chỉ chính thức https://omo13.vercel.app/."
            )
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Update Hosting
        if txt.startswith("Hosting:"):
            p.text = "Hosting & Hạ tầng mạng: Nền tảng Đám mây Vercel Hosting (CDN Edge Network phân tán tốc độ cao, tự động cấp phát và gia hạn chứng chỉ bảo mật SSL/HTTPS 256-bit, Uptime 99.99%)."
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Update Domain
        if txt.startswith("Domain:"):
            p.text = "Domain: https://omo13.vercel.app (Tên miền chính thức của dự án trên hệ thống đám mây Vercel)."
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Update Tools
        if txt.startswith("Công cụ hỗ trợ:"):
            p.text = "Công cụ phát triển & quản trị: Visual Studio Code, Git/GitHub, Node.js (Static Server kiểm thử cục bộ), Vercel CLI, Figma (Thiết kế Wireframe & UI/UX), Google Search Console, Google Analytics 4, Screaming Frog SEO Spider, SEOquake."
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Update System Specification (Đặc tả hệ thống)
        if "* đặc tả hệ thống" in txt.lower() or txt == "* Đặc tả hệ thống":
            p.text = (
                "* Đặc tả toàn diện kiến trúc các phân hệ giao diện trên website ÔMƠ:\n\n"
                "1. Thanh Tiện Ích Đầu Trang (Top Bar):\n"
                "• Phía bên trái: Hiển thị thông tin địa chỉ thực thể cửa hàng (📍 255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng) và khung giờ mở cửa (⏰ 08:00 – 22:00 từ Thứ 2 đến Chủ Nhật).\n"
                "• Phía bên phải: Hiển thị cam kết dịch vụ vận chuyển (🚚 Giao hỏa tốc 1-2h tại Đà Nẵng) và số điện thoại đường dây nóng (📞 Hotline: 222-456-789 / 022456789) có gắn liên kết `tel:` hỗ trợ bấm gọi ngay.\n\n"
                "2. Thanh Điều Hướng Chính Cố Định (Main Sticky Header):\n"
                "• Khối nhận diện thương hiệu (Brand Logo): Bên trái đặt logo hình chú gấu ÔMƠ kèm tên thương hiệu 'ÔMƠ' (thẻ H1) và khẩu hiệu slogan 'Ôm một chú gấu, giữ một giấc mơ'.\n"
                "• Menu điều hướng trung tâm (Navigation Menu): Gồm đầy đủ 8 liên kết điều hướng mượt mà đến 8 trang chức năng chính: Trang chủ (index.html), Giới thiệu (gioi-thieu.html), Cửa hàng (cua-hang.html), Tin tức (tin-tuc.html), Giỏ hàng (gio-hang.html), Thanh toán (thanh-toan.html), Chính sách (chinh-sach.html), Liên hệ (lien-he.html).\n"
                "• Cụm công cụ tương tác (Header Actions): Phía bên phải tích hợp nút liên kết Fanpage Facebook, nút tìm kiếm nhanh 🔍, biểu tượng Giỏ hàng 🛒 có gắn huy hiệu (Cart Badge) tự động nhảy số lượng sản phẩm thời gian thực khi khách bấm mua hàng, và nút Menu di động ☰ (Mobile Hamburger Menu) co giãn linh hoạt trên smartphone.\n\n"
                "3. Thân Trang & Các Phân Hệ Cốt Lõi (Body & Core Page Sections):\n"
                "• Phân hệ Trang Chủ (index.html): Bố cục tuần tự qua 7 khối nội dung trực quan: (1) Hero Banner Pastel với slogan, 2 nút CTA và 3 chỉ số uy tín (100% bông tinh khiết, 500+ mẫu gấu, 2.000+ khách hàng); (2) Khối 'Câu Chuyện Của Chúng Mình' giải mã ý nghĩa ÔM và MƠ kèm 4 cam kết vàng; (3) Khối 'Sản Phẩm Được Yêu Thích' hiển thị 8 mẫu gấu Best Sellers có rating 5 sao và nút Quick View; (4) Khối 'Vì Sao Nên Chọn Chúng Tớ?' với 4 thẻ giá trị; (5) Khối thông điệp trích dẫn 'Gửi gắm yêu thương trong từng chi tiết'; (6) Khối 'Khoảnh Khắc Đáng Yêu Cùng Gấu Bông' trưng bày 4 hình ảnh feedback khách hàng Instagram; (7) Khối 'Tin Tức & Cẩm Nang' trích dẫn 3 bài viết blog chuyên sâu; (8) Khối Hotline Callout Banner hỗ trợ tư vấn tức thì.\n"
                "• Phân hệ Trang Cửa Hàng (cua-hang.html): Bộ lọc 3 tab kích thước (Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+), ô tìm kiếm gấu bông thời gian thực, bộ lọc sắp xếp giá bán/đánh giá, lưới 15 sản phẩm chuẩn SEO và cửa sổ popup Quick View xem nhanh chi tiết sản phẩm.\n"
                "• Phân hệ Trang Giới Thiệu (gioi-thieu.html): Storytelling thương hiệu, Hệ giá trị cốt lõi 5T, thông tin 6 thành viên sáng lập, quy trình đóng hộp quà Kraft vintage và thiệp hoa khô viết tay.\n"
                "• Phân hệ Trang Tin Tức (tin-tuc.html): 3 bài viết cẩm nang chuẩn SEO (Mẹo chọn quà, Hướng dẫn bảo quản, Ý nghĩa gấu bông) tích hợp mục lục tự động TOC và hệ thống liên kết nội bộ Silo.\n"
                "• Phân hệ Trang Giỏ Hàng (gio-hang.html): Bảng tóm tắt danh sách sản phẩm (ảnh, tên, đơn giá, bộ đếm số lượng +/-), thanh tiến trình thông minh nhắc nhở miễn phí vận chuyển cho đơn từ 300.000đ, ô nhập mã voucher ưu đãi (OMOYEUTHUONG 10%, OMOCHU 10%, FREESHIP 5%) và thẻ tóm tắt đơn hàng chuyển tiếp thanh toán.\n"
                "• Phân hệ Trang Thanh Toán (thanh-toan.html): Form thu thập địa chỉ nhận hàng và ghi chú viết thiệp, tích hợp 3 phương thức thanh toán linh hoạt (Quét mã VietQR MB Bank STK 022456789999, Ví MoMo, COD khi nhận hàng), popup xác nhận đặt hàng thành công #OMO-XXXXX và lưu trữ LocalStorage.\n"
                "• Phân hệ Trang Chính Sách (chinh-sach.html): 4 văn bản quy định minh bạch về mua hàng, giao hàng hỏa tốc 1-2h tại Đà Nẵng / COD 2-4 ngày toàn quốc, đổi trả 1-1 trong 3 ngày và bảo mật dữ liệu.\n"
                "• Phân hệ Trang Liên Hệ (lien-he.html): Bản đồ tương tác Google Maps nhúng trực quan tại 255 Hà Huy Tập, form gửi tin nhắn và liên kết kết nối đa kênh Facebook/Instagram/TikTok/Zalo.\n\n"
                "4. Chân Trang Đồng Bộ (Standard 4-Column Site Footer):\n"
                "• Cột 1 (Thương hiệu ÔMƠ): Giới thiệu ngắn sứ mệnh mang lại cái ôm vỗ về và 4 biểu tượng mạng xã hội (Facebook, Instagram, TikTok, Zalo).\n"
                "• Cột 2 (Về ÔMƠ): Danh mục menu liên kết nội bộ nhanh đến Trang chủ, Giới thiệu, Cửa hàng, Tin tức, Liên hệ.\n"
                "• Cột 3 (Chính sách & Hỗ trợ): Liên kết chính sách mua hàng, quy định đổi trả trong 3 ngày, giỏ hàng và hướng dẫn thanh toán.\n"
                "• Cột 4 (Thông tin liên hệ thực thể): Hiển thị địa chỉ 255 Hà Huy Tập, Q. Thanh Khê, TP. Đà Nẵng; Hotline / Zalo: 022456789 (222-456-789); Email: omoshop@gmail.com.\n"
                "• Dòng bản quyền dưới cùng: '© 2026 ÔMƠ Shop – Ôm trọn vỗ về, giữ trọn giấc mơ. All rights reserved.'\n\n"
                "5. Lớp Tương Tác Trải Nghiệm Người Dùng (Interactive Layer & AI Chatbot):\n"
                "• Thanh cuộn tiến trình (Scroll Progress Bar) hiển thị tỷ lệ đọc trang tại đỉnh màn hình.\n"
                "• Trợ lý ảo Chatbot AI 'Trợ lý Mơ' góc dưới màn hình: Tự động phản hồi thông minh theo từ khóa (giá, ship, đổi trả, tư vấn quà) kèm các nút gợi ý Quick Replies.\n"
                "• Danh sách sản phẩm yêu thích (Wishlist FAB): Lưu trữ các mẫu gấu bông yêu thích qua LocalStorage.\n"
                "• Cụm nút nổi (Floating Actions): Nút gọi Hotline khẩn cấp, nút chat Zalo và nút Back-to-top cuộn mượt mà lên đầu trang.\n"
                "• Thông báo Toast Notifications: Hiển thị thông báo nổi sinh động khi thêm giỏ hàng, lưu yêu thích hoặc nhập mã khuyến mãi."
            )
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Clean demo blog text (replace any non-source characters with source-aligned characters)
        if "TRÍCH XUẤT NỘI DUNG TOÀN VĂN BÀI VIẾT DEMO" in txt:
            p.text = (
                "TRÍCH XUẤT NỘI DUNG TOÀN VĂN BÀI VIẾT DEMO CHUẨN SEO (ĐỘ DÀI 1.500 TỪ):\n"
                "Tiêu đề H1: 5 Cách Chọn Gấu Bông Làm Quà Tặng Phù Hợp Cho Từng Dịp & Đối Tượng Ý Nghĩa – ÔMƠ Shop\n\n"
                "Đoạn Mở Bài (Sapo): Gấu bông từ lâu đã trở thành một trong những món quà tặng ý nghĩa và được yêu thích nhất mọi thời đại. Không chỉ mang vẻ ngoài dễ thương, một chú thú nhồi bông mềm mại còn là 'sứ giả cảm xúc' gửi gắm sự quan tâm chân thành và cái ôm ấm áp đến người nhận. Tuy nhiên, làm thế nào để chọn được một chú gấu bông vừa vặn với sở thích, đúng dịp và đảm bảo chất lượng cao cấp? Hãy cùng Tiệm gấu bông ÔMƠ khám phá ngay 5 cách chọn gấu bông làm quà tặng chuẩn không cần chỉnh trong bài viết dưới đây!\n\n"
                "H2: 1. Vì Sao Gấu Bông Luôn Là Món Quà Tặng Được Yêu Thích Hàng Đầu?\n"
                "Theo các nghiên cứu tâm lý học hành vi, việc ôm một chú gấu bông mềm mại có khả năng kích thích não bộ sản sinh hormone Oxytocin (hormone hạnh phúc), giúp giảm nhịp tim, xoa dịu cảm giác cô đơn và giảm căng thẳng (stress) hiệu quả. Dù ở bất kỳ độ tuổi nào, từ trẻ nhỏ, học sinh sinh viên cho đến người đi làm, một chú thú bông êm ái luôn mang lại cảm giác được vỗ về và an tâm.\n\n"
                "H2: 2. Tiêu Chí Chọn Gấu Bông Chất Lượng Cao Không Thể Bỏ Qua\n"
                "• Chất liệu vỏ bọc bên ngoài: Nên ưu tiên vải nhung tuyết co giãn 4 chiều mịn màng, sờ mát tay, không bị xù lông hoặc rụng lông khi giặt.\n"
                "• Chất liệu bông gòn bên trong: Tuyệt đối tránh các loại bông tạp chất, gòn phế phẩm giá rẻ. Hãy chọn gấu nhồi 100% bông gòn trắng tinh khiết PP 3D có độ đàn hồi cao, không bị xẹp lún sau thời gian dài sử dụng.\n"
                "• Đường may và chi tiết: Các đường kim mũi chỉ phải đều đặn, chắc chắn; các chi tiết mắt, mũi được đính kết an toàn, không dễ bị rơi rớt gây nguy hiểm cho trẻ nhỏ.\n\n"
                "H2: 3. 5 Cách Chọn Gấu Bông Phù Hợp Từng Dịp Và Đối Tượng\n"
                "H3: 3.1. Chọn gấu bông tặng người yêu dịp Valentine, Ngày kỷ niệm\n"
                "Đối với bạn gái, những chú gấu Teddy cỡ lớn (từ 50cm đến 1m - Gấu Teddy ÔMƠ Khổng Lồ) mang phong cách Vintage hoặc các mẫu thỏ hồng My Melody nơ xinh xắn luôn là lựa chọn số 1. Món quà thể hiện mong muốn được ở bên cạnh che chở và sưởi ấm cho người yêu mỗi ngày.\n"
                "H3: 3.2. Chọn gấu bông tặng bạn bè dịp Sinh nhật, Tốt nghiệp\n"
                "Với bạn thân, các mẫu thú bông hot-trend mang tính hài hước, độc lạ như Hội bạn thân Opanchu Usagi, Minions chú bé Bob, Cá ngố Hangyodon hay Trứng lười Gudetama size 20 - 30cm sẽ mang lại tiếng cười và niềm vui bất ngờ.\n"
                "H3: 3.3. Chọn móc khóa gấu bông xinh xắn làm quà nhỏ bất ngờ\n"
                "Những chiếc móc khóa thú bông mini (size 5 - 15cm) như Móc khóa Keychain Xinh, Sumikko Gurashi, We Bare Bears có giá mềm (45k - 95k) rất thích hợp để tặng bạn bè cùng lớp treo balo, túi xách đi học như một vật kỷ niệm đáng yêu.\n"
                "H3: 3.4. Chọn gấu bông chữa lành (Healing Toys) cho bản thân hoặc người đang stress\n"
                "Dòng thú bông nhung mềm tông màu Pastel như Kirby tròn ủm hồng, Cún tai dài Cinnamoroll trắng muốt hay Cún Shiba Inu béo tròn giúp xoa dịu những áp lực tinh thần sau ngày dài làm việc.\n"
                "H3: 3.5. Chọn gấu bông cho trẻ nhỏ (An toàn, kháng khuẩn)\n"
                "Cần chọn các mẫu thú bông kinh điển như Gấu Pooh áo đỏ, Chuột Mickey Mouse cổ điển với mắt mũi thêu thủ công sắc nét, ruột bông tinh khiết kháng khuẩn 100% để đảm bảo an toàn tuyệt đối cho bé khi chơi và ôm ngủ.\n\n"
                "H2: 4. Dịch Vụ Đóng Gói Quà Tặng Tinh Tế Tại Tiệm Gấu Bông ÔMƠ\n"
                "Khi mua gấu bông tại ÔMƠ Shop (https://omo13.vercel.app/), bạn sẽ được trải nghiệm dịch vụ quà tặng trọn gói: Đóng hộp quà Vintage Kraft thắt nơ ruy băng thủ công, tặng kèm thiệp hoa khô viết tay theo lời chúc bạn yêu cầu, và dịch vụ giao hàng hỏa tốc trong 1-2 giờ tại nội thành Đà Nẵng.\n\n"
                "H2: 5. Lời Kết\n"
                "Một chú gấu bông xinh xắn được lựa chọn tỉ mỉ chính là món quà tinh thần vô giá thay bạn gửi trao những yêu thương chân thành nhất. Hãy ghé ngay Cửa hàng Tiệm gấu bông ÔMƠ để chọn cho mình và người thương một 'người bạn nhỏ' ngọt ngào nhất nhé!"
            )
            for r in p.runs: set_font_tnr(r, size_pt=13)

        # Clear redundant separate short Header/Footer list paragraphs right after * Đặc tả hệ thống if present
        if txt.startswith("Header: Logo và tên thương hiệu") or txt.startswith("Footer: Tên thương hiệu ÔMƠ SHOP"):
            p.text = ""

    # -------------------------------------------------------------
    # 3. UPDATE TABLE 11: BẢNG 3.1 ĐẶC TẢ HỆ THỐNG WEBSITE
    # -------------------------------------------------------------
    if len(tables) > 11:
        t11 = tables[11]
        print("Updating Table 11 (Bảng 3.1. Đặc tả hệ thống website)...")
        headers_11 = ["Trang / Phân Hệ", "Yêu Cầu Giao Diện & Thành Phần Body Chi Tiết Trong Source Code", "Mục Đích Trải Nghiệm (UX) & Mục Tiêu SEO"]
        for c_idx in range(min(len(headers_11), len(t11.rows[0].cells))):
            format_cell(t11.rows[0].cells[c_idx], headers_11[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t11.rows[0].cells[c_idx], "4A6FA5")

        t11_data = [
            (
                "Trang Chủ\n(index.html)",
                "• Top Bar: Địa chỉ 255 Hà Huy Tập, giờ mở cửa 8h-22h, giao hỏa tốc 1-2h, Hotline 222-456-789.\n"
                "• Header: Logo ÔMƠ + Slogan 'Ôm một chú gấu, giữ một giấc mơ', menu 8 trang, badge giỏ hàng nhảy số real-time, nút mobile toggle.\n"
                "• Body (Thân trang):\n"
                "  (1) Hero Section: Banner pastel, nút 'Khám Phá Cửa Hàng', 3 chỉ số uy tín (100% bông tinh khiết, 500+ mẫu, 2000+ khách hàng);\n"
                "  (2) Story Summary: Ý nghĩa ÔM & MƠ, 4 cam kết vàng;\n"
                "  (3) Best Sellers Grid: 8 mẫu gấu tiêu biểu (Kirby, Opanchu, Bob, Cinnamoroll, Pooh, Gudetama, Hangyodon, My Melody) kèm tag giá & Quick View;\n"
                "  (4) Why Choose Us: 4 thẻ giá trị cốt lõi (Đa dạng mẫu mã, Gửi trao yêu thương, Nâng niu giấc mơ, Hỗ trợ tận tâm);\n"
                "  (5) Quote Box: Thông điệp 'Gửi gắm yêu thương trong từng chi tiết';\n"
                "  (6) Social Feed: Lưới 4 ảnh feedback Instagram thực tế;\n"
                "  (7) Blog Grid: 3 bài viết cẩm nang nổi bật;\n"
                "  (8) Hotline Banner: Kêu gọi hành động tư vấn quà tặng.\n"
                "• Footer 4 cột + Floating Actions (Chatbot AI Trợ lý Mơ, Wishlist FAB, Hotline, Back-to-top).",
                "• Trải nghiệm (UX): Tạo ấn tượng đầu tiên ngọt ngào, dẫn dắt khách hàng qua từng điểm chạm cảm xúc và kích thích chuyển đổi xem danh mục sản phẩm.\n"
                "• SEO: Khai báo thẻ H1 thương hiệu, tối ưu thẻ Title (61 ký tự), Meta Description (130 ký tự), phân bổ bộ từ khóa chính và liên kết Silo."
            ),
            (
                "Trang Giới Thiệu\n(gioi-thieu.html)",
                "• Top Bar & Header đồng bộ toàn hệ thống.\n"
                "• Hero Page: Tiêu đề 'Về ÔMƠ Shop 🧸' kèm đường dẫn Breadcrumb.\n"
                "• Body:\n"
                "  - Khối 'Câu Chuyện Của ÔMƠ': Storytelling về sứ mệnh lan tỏa những cái ôm xoa dịu cảm xúc cho giới trẻ;\n"
                "  - Khối 'Sứ Mệnh Của ÔMƠ': Hệ giá trị cốt lõi 5T (Tâm - Tinh - Thật - Tận - Tình);\n"
                "  - Khối 'Đội Ngũ Sáng Lập': Giới thiệu 6 thành viên sáng lập nhóm 13;\n"
                "  - Khối 'Quy Trình Đóng Gói Thủ Công': Hộp quà Kraft Vintage và thiệp hoa khô viết tay theo yêu cầu;\n"
                "  - Khối Cam kết chất lượng bông gòn tinh khiết 100%.\n"
                "• Footer 4 cột & Chatbot AI.",
                "• Trải nghiệm (UX): Xây dựng niềm tin sâu sắc, truyền tải câu chuyện cảm xúc và sự chỉn chu trong từng gói quà tặng.\n"
                "• SEO: Gia tăng tín hiệu thực thể E-E-A-T (Chuyên môn, Trải nghiệm, Thẩm quyền, Đáng tin cậy) và chứa từ khóa thương hiệu ÔMƠ."
            ),
            (
                "Trang Cửa Hàng\n(cua-hang.html)",
                "• Top Bar & Header đồng bộ.\n"
                "• Hero Page: Tiêu đề 'Cửa Hàng ÔMƠ 🛍️' và Breadcrumb.\n"
                "• Body:\n"
                "  - Bộ lọc đa tiêu chí (Multi-Filter): 3 tab phân nhóm kích thước (Tất Cả, Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+);\n"
                "  - Ô tìm kiếm gấu bông thời gian thực kết hợp bộ chọn sắp xếp giá (Thấp đến Cao, Cao đến Thấp, Đánh giá cao nhất);\n"
                "  - Bộ đếm sản phẩm hiển thị real-time;\n"
                "  - Product Grid: 15 sản phẩm chuẩn SEO dạng lưới 3-4 cột với ảnh chất lượng cao, nhãn Hot/Trend/Sale, rating 5 sao, giá niêm yết và nút CTA;\n"
                "  - Quick View Modal: Cửa sổ popup xem nhanh thông số chất liệu, kích thước, ảnh phóng to, bộ chọn số lượng và nút Thêm vào giỏ;\n"
                "  - Hotline Callout Banner tư vấn kích thước riêng.\n"
                "• Footer 4 cột & Widget tương tác.",
                "• Trải nghiệm (UX): Tìm kiếm và lọc sản phẩm cực nhanh dưới 5 giây, xem chi tiết không cần tải lại trang, thao tác mua hàng mượt mà.\n"
                "• SEO: Tối ưu bộ từ khóa danh mục ('gấu nhỏ', 'gấu vừa', 'gấu khổng lồ', 'móc khóa gấu bông') và dữ liệu có cấu trúc Product Schema."
            ),
            (
                "Trang Tin Tức\n(tin-tuc.html)",
                "• Top Bar & Header đồng bộ.\n"
                "• Hero Page: Tiêu đề 'Tin Tức & Cẩm Nang 📰' và Breadcrumb.\n"
                "• Body:\n"
                "  - Khối Bài viết nổi bật (Featured Post);\n"
                "  - Lưới bài viết chia theo 3 bài blog chuyên sâu: (1) 5 Cách chọn gấu bông làm quà tặng; (2) Hướng dẫn vệ sinh và bảo quản gấu bông; (3) Ý nghĩa tặng gấu bông cho người yêu & bạn bè;\n"
                "  - Định dạng bài viết: Tích hợp đầy đủ Mục lục tự động (TOC), hình ảnh minh họa có thẻ Alt, danh sách liệt kê và liên kết nội bộ trỏ về Cửa hàng.\n"
                "• Footer 4 cột & Trợ lý ảo Chatbot.",
                "• Trải nghiệm (UX): Đọc nội dung hữu ích, dễ quét thông tin nhờ bố cục chia đoạn ngắn và mục lục trực quan.\n"
                "• SEO: Khai thác bộ từ khóa KGR, cung cấp nội dung chất lượng cao (Helpful Content) và phân bổ internal links theo mô hình Silo."
            ),
            (
                "Trang Giỏ Hàng\n(gio-hang.html)",
                "• Top Bar & Header đồng bộ (Badge giỏ hàng nhảy số real-time).\n"
                "• Hero Page: Tiêu đề 'Giỏ Hàng Của Bạn 🛒' và Breadcrumb.\n"
                "• Body Layout 2 Cột:\n"
                "  - Thanh thông báo tiến trình Freeship (Free Shipping Notice Bar) nhắc nhở mua thêm để được miễn phí ship đơn từ 300.000đ;\n"
                "  - Cột trái: Bảng danh sách sản phẩm đã chọn (Ảnh thumbnail, Tên, Size, Đơn giá, Bộ tăng giảm số lượng +/-, Nút xóa 🗑️, Thành tiền) và nút tiếp tục mua sắm;\n"
                "  - Cột phải: Thẻ Tóm tắt đơn hàng (Tạm tính, Phí ship 30k/0đ, Dòng mã giảm giá xanh lá), Ô nhập mã ưu đãi (OMOYEUTHUONG 10%, OMOCHU 10%, FREESHIP 5%), Tổng thanh toán và nút 'Tiến Hành Thanh Toán 💕';\n"
                "  - Trạng thái giỏ hàng trống (Empty State) hiển thị icon 🧸💔 và nút quay lại cửa hàng.\n"
                "• Footer 4 cột.",
                "• Trải nghiệm (UX): Luồng thanh toán không ma sát (Frictionless UX), cập nhật giá tiền tức thì qua LocalStorage không cần reload trang.\n"
                "• SEO: Đảm bảo luồng chuyển đổi mượt mà, tối ưu thẻ Title và Meta Description chuẩn SEO."
            ),
            (
                "Trang Thanh Toán\n(thanh-toan.html)",
                "• Top Bar & Header đồng bộ.\n"
                "• Hero Page: Tiêu đề 'Thanh Toán Đơn Hàng 💳' và Breadcrumb.\n"
                "• Dải thông báo nhắc nhở kiểm tra thông tin nhận hàng.\n"
                "• Body Layout 2 Cột:\n"
                "  - Cột trái: (1) Form thông tin người nhận (Họ tên, SĐT, Địa chỉ chi tiết, Ghi chú viết thiệp); (2) Lựa chọn 3 phương thức thanh toán: Quét mã VietQR MB Bank STK 022456789999, Ví MoMo, Thanh toán tiền mặt COD; (3) Khung hiển thị mã VietQR động kèm thông tin chuyển khoản có nút Sao chép (Copy) tiện lợi;\n"
                "  - Cột phải: Thẻ tóm tắt đơn hàng (Danh sách món hàng preview, Mã đơn #OMO-XXXXX, Tổng thanh toán) và nút 'Xác Nhận Đặt Hàng Ngay 💕';\n"
                "  - Modal Popup đặt hàng thành công: Xuất hóa đơn chi tiết và lưu trữ đơn hàng vào LocalStorage.\n"
                "• Footer 4 cột.",
                "• Trải nghiệm (UX): Tối giản quy trình 1 trang (One-Page Checkout), quét mã QR tự động điền số tiền và nội dung, hạn chế sai sót chuyển khoản.\n"
                "• SEO: Đạt chuẩn trải nghiệm người dùng Core Web Vitals, bảo mật giao dịch an toàn."
            ),
            (
                "Trang Chính Sách\n(chinh-sach.html)",
                "• Top Bar & Header đồng bộ.\n"
                "• Hero Page: Tiêu đề 'Chính Sách ÔMƠ 📜' và Breadcrumb.\n"
                "• Body:\n"
                "  - Khối Chính sách mua hàng & Quy định đặt hàng trực tuyến;\n"
                "  - Khối Chính sách giao hàng hỏa tốc trong 1-2h tại Đà Nẵng và giao nhanh COD toàn quốc 2-4 ngày;\n"
                "  - Khối Chính sách đổi trả 1-1 trong 3 ngày nếu phát hiện lỗi đường may hoặc lỗi nhà sản xuất;\n"
                "  - Khối Chính sách bảo mật thông tin khách hàng tuyệt đối.\n"
                "• Footer 4 cột & Chatbot AI.",
                "• Trải nghiệm (UX): Minh bạch hóa thông tin, giúp khách hàng yên tâm tối đa trước và sau khi đặt hàng.\n"
                "• SEO: Bổ sung các trang chính sách tiêu chuẩn nâng cao điểm chất lượng Entity E-E-A-T trên Google Search."
            ),
            (
                "Trang Liên Hệ\n(lien-he.html)",
                "• Top Bar & Header đồng bộ.\n"
                "• Hero Page: Tiêu đề 'Liên Hệ Với Chúng Tớ 💌' và Breadcrumb.\n"
                "• Body:\n"
                "  - Khung bản đồ Google Maps tương tác nhúng trực quan định vị cửa hàng tại 255 Hà Huy Tập, Thanh Khê, Đà Nẵng;\n"
                "  - Khối thông tin liên hệ chính thức: Địa chỉ, Hotline 222-456-789 / 022456789, Email omoshop@gmail.com, Giờ hoạt động 8h-22h;\n"
                "  - Form gửi tin nhắn tư vấn và phản hồi trực tuyến;\n"
                "  - Cụm nút liên kết trực tiếp tới Fanpage Facebook, Instagram, TikTok và kênh Zalo OA.\n"
                "• Footer 4 cột & Trợ lý ảo Chatbot.",
                "• Trải nghiệm (UX): Đa kênh tương tác, hỗ trợ khách hàng tìm đường đến cửa hàng và gửi yêu cầu tư vấn nhanh chóng.\n"
                "• SEO: Đồng bộ dữ liệu Local SEO (NAP: Name - Address - Phone) trùng khớp 100% với Google Business Profile và Schema Organization."
            )
        ]

        for r_idx, (page_name, page_body, page_ux_seo) in enumerate(t11_data, start=1):
            if r_idx < len(t11.rows):
                format_cell(t11.rows[r_idx].cells[0], page_name, bold=True, size_pt=9.5)
                format_cell(t11.rows[r_idx].cells[1], page_body, size_pt=9.0)
                if len(t11.rows[r_idx].cells) > 2:
                    format_cell(t11.rows[r_idx].cells[2], page_ux_seo, size_pt=9.0)

    # -------------------------------------------------------------
    # 4. UPDATE TABLE 12: BẢNG 3.2 PHÂN TẦNG WEBSITE
    # -------------------------------------------------------------
    if len(tables) > 12:
        t12 = tables[12]
        print("Updating Table 12 (Bảng 3.2. Phân tầng website)...")
        headers_12 = ["Tầng Kiến Trúc", "Trang / Phân Nhóm Danh Mục / Sản Phẩm", "Đường Dẫn Truy Cập (URL) & Đặc Điểm Kỹ Thuật"]
        for c_idx in range(min(len(headers_12), len(t12.rows[0].cells))):
            format_cell(t12.rows[0].cells[c_idx], headers_12[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t12.rows[0].cells[c_idx], "4A6FA5")

        t12_data = [
            ("Tầng 1 (Homepage)", "Trang chủ", "https://omo13.vercel.app/index.html"),
            ("Tầng 2 (Core Pages)", "Giới thiệu", "https://omo13.vercel.app/gioi-thieu.html"),
            ("Tầng 2 (Core Pages)", "Cửa hàng", "https://omo13.vercel.app/cua-hang.html"),
            ("Tầng 2 (Core Pages)", "Tin tức & Cẩm nang", "https://omo13.vercel.app/tin-tuc.html"),
            ("Tầng 2 (Core Pages)", "Giỏ hàng", "https://omo13.vercel.app/gio-hang.html"),
            ("Tầng 2 (Core Pages)", "Thanh toán", "https://omo13.vercel.app/thanh-toan.html"),
            ("Tầng 2 (Core Pages)", "Chính sách mua hàng & Đổi trả", "https://omo13.vercel.app/chinh-sach.html"),
            ("Tầng 2 (Core Pages)", "Liên hệ & Bản đồ", "https://omo13.vercel.app/lien-he.html"),
            ("Tầng 3 (Categories)", "Phân nhóm Gấu nhỏ (5cm - 15cm)", "Nằm trong trang Cửa hàng (Bộ lọc Tab: data-cat='gau-nho')"),
            ("Tầng 3 (Categories)", "Phân nhóm Gấu vừa (20cm - 30cm)", "Nằm trong trang Cửa hàng (Bộ lọc Tab: data-cat='gau-vua')"),
            ("Tầng 3 (Categories)", "Phân nhóm Gấu to (30cm trở lên)", "Nằm trong trang Cửa hàng (Bộ lọc Tab: data-cat='gau-to')"),
            ("Tầng 3 (Blog Posts)", "Bài viết 1: 5 Cách Chọn Gấu Bông Làm Quà Tặng", "Nằm trong trang Tin tức (Mẹo Chọn Quà - Tác giả: ÔMƠ Team)"),
            ("Tầng 3 (Blog Posts)", "Bài viết 2: Cách Vệ Sinh & Bảo Quản Gấu Bông", "Nằm trong trang Tin tức (Chăm Sóc Gấu Bông - Tác giả: ÔMƠ Care)"),
            ("Tầng 3 (Blog Posts)", "Bài viết 3: Ý Nghĩa Của Gấu Bông Khi Làm Quà Tặng", "Nằm trong trang Tin tức (Ý Nghĩa Yêu Thương - Tác giả: ÔMƠ Story)"),
            ("Tầng 4 (Products Demo)", "sp-01: Móc Khóa Gấu Bông Keychain Xinh", "Nằm trong danh mục Gấu nhỏ (Giá: 45.000đ - Hot Trend)"),
            ("Tầng 4 (Products Demo)", "sp-02: Hội Bạn Thân Opanchu Usagi", "Nằm trong danh mục Gấu nhỏ (Giá: 85.000đ - Bán Chạy)"),
            ("Tầng 4 (Products Demo)", "sp-05: Gấu Bông Kirby Tròn Ủm Hồng", "Nằm trong danh mục Gấu vừa (Giá: 200.000đ - Best Seller)"),
            ("Tầng 4 (Products Demo)", "sp-07: Cún Tai Dài Cinnamoroll Trắng Muốt", "Nằm trong danh mục Gấu vừa (Giá: 210.000đ - Sanrio Hot)"),
            ("Tầng 4 (Products Demo)", "sp-11: My Melody Nơ Hồng Xinh Xắn", "Nằm trong danh mục Gấu to (Giá: 290.000đ - Siêu Hot)"),
            ("Tầng 4 (Products Demo)", "sp-15: Gấu Bông Khổng Lồ Teddy ÔMƠ Vỗ Về", "Nằm trong danh mục Gấu to (Giá: 490.000đ - Khổng Lồ Signature)")
        ]

        for r_idx, (layer_name, page_item, url_item) in enumerate(t12_data, start=1):
            if r_idx < len(t12.rows):
                format_cell(t12.rows[r_idx].cells[0], layer_name, bold=True, size_pt=9.5)
                format_cell(t12.rows[r_idx].cells[1], page_item, size_pt=9.0)
                if len(t12.rows[r_idx].cells) > 2:
                    format_cell(t12.rows[r_idx].cells[2], url_item, size_pt=9.0)

    # -------------------------------------------------------------
    # 5. UPDATE TABLE 22: TỐI ƯU UI TRANG CHỦ (HEADER, FOOTER, BODY)
    # -------------------------------------------------------------
    if len(tables) > 22:
        t22 = tables[22]
        print("Updating Table 22 (Tối ưu UI trang chủ: Header, Footer, Body)...")
        headers_22 = ["Phân Vùng Giao Diện", "Giải Pháp Tối Ưu UI Chi Tiết Khớp Mã Nguồn Thực Tế"]
        format_cell(t22.rows[0].cells[0], headers_22[0], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
        format_cell(t22.rows[0].cells[1], headers_22[1], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
        set_cell_background(t22.rows[0].cells[0], "4A6FA5")
        set_cell_background(t22.rows[0].cells[1], "4A6FA5")

        t22_data = [
            (
                "Header &\nTop Bar",
                "• Top Bar: Cố định nền hồng phấn nhẹ trên cùng hiển thị thực thể địa chỉ tại 255 Hà Huy Tập, giờ mở cửa 8h-22h, cam kết giao hỏa tốc 1-2h và Hotline 222-456-789.\n"
                "• Sticky Header: Thiết kế theo hiệu ứng bán trong suốt (Glassmorphism), tự động thêm shadow khi cuộn trang; bên trái là Logo hình chú gấu ÔMƠ và Slogan 'Ôm một chú gấu, giữ một giấc mơ'; ở giữa là menu điều hướng 8 trang với hiệu ứng hover pastel mềm mại; bên phải là thanh tìm kiếm 🔍, nút Facebook và biểu tượng Giỏ hàng 🛒 có badge hiển thị số lượng sản phẩm real-time."
            ),
            (
                "Footer 4 Cột\n& Widgets",
                "• Footer 4 Cột: Cột 1 giới thiệu thương hiệu ÔMƠ SHOP và 4 icon mạng xã hội (Facebook, Instagram, TikTok, Zalo); Cột 2 liên kết nhanh các trang chính; Cột 3 hiển thị chính sách đổi trả 3 ngày và thanh toán; Cột 4 công bố thông tin liên hệ chính thức tại Đà Nẵng (255 Hà Huy Tập, Hotline 022456789, Email omoshop@gmail.com).\n"
                "• Cụm Widget Nổi: Scroll Progress Bar đỉnh trang, Chatbot AI Trợ lý Mơ góc phải dưới, Wishlist FAB '♥ Đã lưu', Hotline/Zalo/Back-to-top."
            ),
            (
                "Body (Thân trang\nTrang Chủ)",
                "Body trang chủ (`index.html`) được cấu trúc tuần tự và chặt chẽ theo tâm lý học mua sắm:\n"
                "1. Hero Section: Banner chính hiển thị huy hiệu '🧸 Tiệm Gấu Bông Xinh Xắn & Đáng Yêu', tiêu đề lớn 'ÔMƠ SHOP - Ôm một chú gấu, giữ một giấc mơ', 2 nút CTA ('KHÁM PHÁ CỬA HÀNG', 'Tìm hiểu về ÔMƠ'), 3 số liệu uy tín (100% bông tinh khiết, 500+ mẫu, 2.000+ khách hàng) cùng hình ảnh gấu bông có 2 nhãn động lơ lửng ('✨ Mềm mịn như mây', '🎁 Quà tặng vỗ về');\n"
                "2. Story Summary Section: Trình bày câu chuyện 'ÔMƠ – Ôm trọn vỗ về, giữ trọn giấc mơ' kèm 4 biểu tượng cam kết vàng (Vải nhung mịn, An toàn da nhạy cảm, Gói quà kèm thiệp, Ship hỏa tốc);\n"
                "3. Best Sellers Section: Trưng bày 8 mẫu gấu bán chạy nhất trong source code (Kirby hồng, Opanchu Usagi, Bob Minions, Cinnamoroll, Pooh áo đỏ, Gudetama, Cá ngố Hangyodon, My Melody) với ảnh sắc nét, rating 5 sao, giá niêm yết, giá gốc, nhãn tag giảm giá và nút Quick View / Thêm giỏ;\n"
                "4. Why Choose Us Section: 4 thẻ giá trị cốt lõi (Đa Dạng Mẫu Mã, Gửi Trao Yêu Thương, Nâng Niu Từng Giấc Mơ, Hỗ Trợ Tận Tâm);\n"
                "5. Quote & Craftsmanship Section: Khối trích dẫn 'Gửi gắm yêu thương trong từng chi tiết' trên nền gradient kem hồng cam kết gia công tỉ mỉ;\n"
                "6. Social Feed Section: Lưới 4 khung ảnh feedback thực tế của khách hàng kết nối Instagram/TikTok (#sanrio #disney #feedback);\n"
                "7. News & Blog Section: 3 bài viết cẩm nang chọn quà và chăm sóc gấu bông hữu ích;\n"
                "8. Hotline Callout Banner: Banner kêu gọi hành động với số điện thoại 222-456-789 và nút chat tư vấn trực tiếp."
            )
        ]

        for r_idx, (part_name, part_desc) in enumerate(t22_data, start=1):
            if r_idx < len(t22.rows):
                format_cell(t22.rows[r_idx].cells[0], part_name, bold=True, size_pt=9.5)
                format_cell(t22.rows[r_idx].cells[1], part_desc, size_pt=9.0)

    # -------------------------------------------------------------
    # 6. UPDATE TABLE 23: TỐI ƯU UI TRANG DANH MỤC SẢN PHẨM
    # -------------------------------------------------------------
    if len(tables) > 23:
        t23 = tables[23]
        print("Updating Table 23 (Tối ưu UI trang danh mục sản phẩm)...")
        headers_23 = ["Hạng Mục Tối Ưu UI", "Nội Dung Triển Khai Trong Mã Nguồn Thực Tế"]
        for c_idx in range(min(len(headers_23), len(t23.rows[0].cells))):
            format_cell(t23.rows[0].cells[c_idx], headers_23[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t23.rows[0].cells[c_idx], "4A6FA5")

        t23_data = [
            ("1. Tổng quan giao diện", "Header, Top Bar và Footer đồng bộ phong cách Vintage Pastel. Nền trắng sáng thoáng đãng (#FAF6F0 / #FFFBFB) kết hợp border bo tròn mềm mại 16px tạo cảm giác ngọt ngào, dễ chịu."),
            ("2. Bộ lọc kích thước & Tìm kiếm", "Tích hợp thanh điều khiển thông minh: 3 Tabs phân loại theo kích thước (Gấu nhỏ 5-15cm, Gấu vừa 20-30cm, Gấu to 30cm+) kết hợp ô tìm kiếm từ khóa real-time và bộ sắp xếp theo giá (Thấp đến Cao, Cao đến Thấp, Đánh giá cao nhất)."),
            ("3. Lưới sản phẩm (Product Grid)", "Hiển thị trọn vẹn 15 sản phẩm dạng lưới 3-4 cột đồng đều; mỗi thẻ sản phẩm (Product Card) gồm: ảnh chụp thực tế rõ nét, tag giảm giá / Best Seller / Hot Trend, nút lưu yêu thích (Wishlist Heart Toggle), tên sản phẩm chuẩn SEO, rating 5 sao, giá niêm yết & giá gốc, nút 'Thêm Giỏ' và nút xem nhanh 'Quick View'."),
            ("4. Cửa sổ Quick View Modal", "Popup xem nhanh chi tiết sản phẩm ngay trên trang: hiển thị ảnh phóng to, thông số chất liệu vải/bông, kích thước chuẩn, mô tả công dụng, bộ tăng giảm số lượng và nút thêm giỏ hàng tiện lợi."),
            ("5. Trạng thái phản hồi & Hotline", "Tích hợp bộ đếm sản phẩm phù hợp thời gian thực; màn hình Empty State thông minh khi không tìm thấy kết quả; khối Hotline Banner chân trang sẵn sàng hỗ trợ khách tìm mẫu gấu theo kích thước riêng.")
        ]

        for r_idx, (cat_item, cat_desc) in enumerate(t23_data, start=1):
            if r_idx < len(t23.rows):
                format_cell(t23.rows[r_idx].cells[0], cat_item, bold=True, size_pt=9.5)
                format_cell(t23.rows[r_idx].cells[1], cat_desc, size_pt=9.0)

    # -------------------------------------------------------------
    # 7. UPDATE TABLE 24: TỐI ƯU UI TRANG THANH TOÁN
    # -------------------------------------------------------------
    if len(tables) > 24:
        t24 = tables[24]
        print("Updating Table 24 (Tối ưu UI trang thanh toán)...")
        headers_24 = ["Phân Vùng Giao Diện", "Giải Pháp Tối Ưu UI & Bảo Mật Giao Dịch"]
        for c_idx in range(min(len(headers_24), len(t24.rows[0].cells))):
            format_cell(t24.rows[0].cells[c_idx], headers_24[c_idx], bold=True, size_pt=10, color_rgb=(255,255,255), align=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_background(t24.rows[0].cells[c_idx], "4A6FA5")

        t24_data = [
            (
                "Cột Trái: Thông Tin & Phương Thức",
                "• Dải thông báo màu hồng phấn nhắc nhở khách hàng kiểm tra kỹ thông tin nhận hàng trước khi thanh toán.\n"
                "• Form giao hàng tinh gọn: Họ tên, Số điện thoại nhận hàng, Địa chỉ chi tiết, Ghi chú cho tiệm gấu ÔMƠ (viết thiệp tặng quà theo yêu cầu).\n"
                "• Tùy chọn 3 phương thức thanh toán linh hoạt: (1) Chuyển khoản VietQR MB Bank STK 022456789999 (Chủ TK: TIỆM GẤU ÔMƠ); (2) Ví điện tử MoMo; (3) Thanh toán tiền mặt khi nhận hàng (COD).\n"
                "• Box mã VietQR thông minh: Tự động trích xuất chính xác số tiền đơn hàng và cú pháp chuyển khoản 'OMOSHOP [Mã đơn]' kèm nút Sao chép (Copy) 1 chạm."
            ),
            (
                "Cột Phải: Tóm Tắt & Xác Nhận",
                "• Thẻ tóm tắt đơn hàng (Order Summary): Preview danh sách sản phẩm đã chọn, hiển thị Mã đơn hàng định dạng #OMO-XXXXX, Giá trị đơn hàng, Phí vận chuyển (miễn phí từ 300k), Phí giao dịch 0đ và Tổng thanh toán nổi bật.\n"
                "• Nút CTA 'Xác Nhận Đặt Hàng Ngay 💕' kích hoạt Modal Đặt hàng thành công, hiển thị đầy đủ thông tin người nhận và lưu trữ đơn hàng vào LocalStorage."
            )
        ]

        while len(t24.rows) < 3:
            t24.add_row()

        for r_idx, (part_name, part_desc) in enumerate(t24_data, start=1):
            if r_idx < len(t24.rows):
                format_cell(t24.rows[r_idx].cells[0], part_name, bold=True, size_pt=9.5)
                format_cell(t24.rows[r_idx].cells[1], part_desc, size_pt=9.0)

    # -------------------------------------------------------------
    # 8. UPDATE TABLE 25: TỐI ƯU UX CHI TIẾT
    # -------------------------------------------------------------
    if len(tables) > 25:
        t25 = tables[25]
        print("Updating Table 25 (Tối ưu UX chi tiết)...")
        # Row 8: Giỏ hàng
        if len(t25.rows) > 8:
            format_cell(t25.rows[8].cells[1], "Bố cục 2 cột trực quan: Cột trái quản lý danh sách sản phẩm (ảnh, tên, đơn giá, bộ đếm số lượng +/- và nút xóa), thanh tiến trình freeship thông minh cho đơn từ 300k; Cột phải tóm tắt đơn hàng và ô nhập voucher giảm giá (OMOYEUTHUONG 10%, OMOCHU 10%, FREESHIP 5%).", size_pt=9.0)
            format_cell(t25.rows[8].cells[2], "Tối ưu luồng mua sắm không ma sát, lưu trạng thái giỏ hàng real-time qua LocalStorage không bị mất khi tải lại trang.", size_pt=9.0)
        # Row 9: Tin tức
        if len(t25.rows) > 9:
            format_cell(t25.rows[9].cells[1], "Bố cục bài viết chuyên sâu: Featured post banner nổi bật, danh sách 3 bài viết blog có hình ảnh sắc nét, mục lục tự động TOC, định dạng in đậm/nghiêng giúp đọc quét nhanh và liên kết nội bộ Silo trỏ về cửa hàng.", size_pt=9.0)
            format_cell(t25.rows[9].cells[2], "Giữ chân người đọc lâu hơn (On-site time > 2 phút), cung cấp thông tin hữu ích và gia tăng tín hiệu SEO.", size_pt=9.0)
        # Row 10: Thanh toán
        if len(t25.rows) > 10:
            format_cell(t25.rows[10].cells[1], "Quy trình thanh toán 1 trang (One-Page Checkout): Form thông tin ngắn gọn, 3 phương thức thanh toán trực quan (VietQR MB Bank STK 022456789999, MoMo, COD), mã QR động tự điền số tiền và modal thông báo đặt hàng thành công.", size_pt=9.0)
            format_cell(t25.rows[10].cells[2], "Giảm tỷ lệ bỏ giỏ hàng xuống dưới 15%, hoàn tất thanh toán siêu tốc chỉ trong 30 giây.", size_pt=9.0)

    # -------------------------------------------------------------
    # 9. UPDATE TABLE 27: SỬA LIÊN KẾT GOOGLE SITES SANG VERCEL
    # -------------------------------------------------------------
    if len(tables) > 27:
        t27 = tables[27]
        print("Updating Table 27 (Sửa liên kết Onpage sang Vercel)...")
        for r in t27.rows:
            for c in r.cells:
                if "sites.google.com" in c.text or "google.com/view" in c.text:
                    c.text = (
                        c.text.replace("https://sites.google.com/view/omo13/tin-tuc/5-cach-chon-gau-bong-lam-qua-tang", "https://omo13.vercel.app/tin-tuc.html")
                              .replace("https://sites.google.com/view/omo13/trang-chủ", "https://omo13.vercel.app/index.html")
                              .replace("https://sites.google.com/view/omo13/cửa-hàng", "https://omo13.vercel.app/cua-hang.html")
                              .replace("https://sites.google.com/view/omo13/giỏ-hàng", "https://omo13.vercel.app/gio-hang.html")
                              .replace("https://sites.google.com/view/omo13/chính-sách", "https://omo13.vercel.app/chinh-sach.html")
                    )
                    format_cell(c, c.text, size_pt=9.0)

    # -------------------------------------------------------------
    # 10. APPLY GLOBAL TIMES NEW ROMAN & SAVE
    # -------------------------------------------------------------
    print("Applying global Times New Roman styling to all paragraphs, headings and tables...")
    apply_global_times_new_roman(doc)

    out_main = 'NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx'
    print(f"Saving main requested report to {out_main}...")
    doc.save(out_main)
    print(f"Saved {out_main} successfully!")

    # Attempt to save additional mirrors
    for alt_out in ['DoAnNhom13_IS425_OMO.docx', 'BAO_CAO_DO_AN_IS425_NHOM13_100_TRANG.docx']:
        try:
            doc.save(alt_out)
            print(f"Saved mirror {alt_out} successfully!")
        except Exception as e:
            print(f"Notice: Could not save mirror {alt_out} (file might be opened in Word): {e}")

if __name__ == '__main__':
    perform_complete_body_and_spec_update()
