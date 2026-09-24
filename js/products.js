// Danh sách sản phẩm tiệm gấu ÔMƠ
const PRODUCTS = [
  // --- GẤU NHỎ (5cm - 15cm) ---
  {
    id: 'sp-01',
    name: 'Móc Khóa Gấu Bông Keychain Xinh',
    category: 'gau-nho',
    categoryName: 'Gấu nhỏ (5-15cm)',
    price: 45000,
    originalPrice: 60000,
    size: '10cm - 12cm',
    image: 'assets/images/img_24.png',
    badge: 'Hot Trend',
    rating: 5,
    reviewsCount: 48,
    description: 'Bộ sưu tập móc khóa gấu bông nhỏ xinh với đủ tạo hình biểu cảm siêu ngộ nghĩnh và đáng yêu. Thích hợp treo balo, túi xách, chìa khóa xe hoặc làm quà tặng bạn bè.',
    material: 'Bông gòn PP cao cấp mềm mịn, vải nhung tuyết không rụng lông.',
    inStock: true
  },
  {
    id: 'sp-02',
    name: 'Hội Bạn Thân Opanchu Usagi',
    category: 'gau-nho',
    categoryName: 'Gấu nhỏ (5-15cm)',
    price: 85000,
    originalPrice: 110000,
    size: '12cm - 15cm',
    image: 'assets/images/img_25.png',
    badge: 'Bán Chạy',
    rating: 5,
    reviewsCount: 62,
    description: 'Hội bạn thân Opanchu Usagi cực hot với những biểu cảm "dở khóc dở cười" đặc trưng! Chất liệu siêu êm tay, nhỏ gọn dễ dàng mang theo bên mình.',
    material: 'Chất liệu vải nhung mềm mại, bông nhồi 100% tinh khiết.',
    inStock: true
  },
  {
    id: 'sp-03',
    name: 'Bộ Ba Gấu We Bare Bears Nhỏ',
    category: 'gau-nho',
    categoryName: 'Gấu nhỏ (5-15cm)',
    price: 95000,
    originalPrice: 120000,
    size: '15cm',
    image: 'assets/images/img_26.png',
    badge: 'Mới Về',
    rating: 4.9,
    reviewsCount: 35,
    description: 'Bộ ba gấu xám Grizzly, gấu trúc Panda và gấu trắng Ice Bear đã cập bến Ô MƠ! Kích thước mini cực đáng yêu, để bàn học hay góc làm việc rất chill.',
    material: 'Vải co giãn 4 chiều mềm mịn, bông vi sợi đàn hồi tốt.',
    inStock: true
  },
  {
    id: 'sp-04',
    name: 'Móc Khóa Sumikko Gurashi',
    category: 'gau-nho',
    categoryName: 'Gấu nhỏ (5-15cm)',
    price: 55000,
    originalPrice: 75000,
    size: '8cm - 10cm',
    image: 'assets/images/img_28.png',
    badge: 'Yêu Thích',
    rating: 4.8,
    reviewsCount: 29,
    description: 'Gia đình góc nhỏ Sumikko đáng yêu xỉu, biểu cảm ngây thơ giúp xoa dịu tâm trạng mỗi khi bạn ngắm nhìn.',
    material: 'Vải lông nhung ngắn mềm mịn, móc kim loại chống gỉ.',
    inStock: true
  },

  // --- GẤU VỪA (20cm - 30cm) ---
  {
    id: 'sp-05',
    name: 'Gấu Bông Kirby Tròn Ủm Hồng',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 200000,
    originalPrice: 240000,
    size: '30cm',
    image: 'assets/images/img_30.png',
    badge: 'Best Seller',
    rating: 5,
    reviewsCount: 112,
    description: 'Chú Kirby hồng tròn ủm siêu đáng yêu với kích thước 30cm cực thích hợp để ôm trọn vào lòng khi ngủ hoặc tựa lưng khi học tập.',
    material: 'Vỏ nhung pha lê cao cấp, bông gòn bi 7D êm ái chống xẹp lún.',
    inStock: true
  },
  {
    id: 'sp-06',
    name: 'Minions - Chú Bé Bob Đáng Yêu',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 185000,
    originalPrice: 220000,
    size: '25cm',
    image: 'assets/images/img_3.png',
    badge: 'Nổi Bật',
    rating: 4.9,
    reviewsCount: 84,
    description: 'Thiết kế ngộ nghĩnh, chất liệu cao cấp an toàn, sẵn sàng mang đến tiếng cười và sự vui tươi cho căn phòng của bạn.',
    material: 'Vải nỉ nhung cao cấp, mắt thêu tinh xảo không phai màu.',
    inStock: true
  },
  {
    id: 'sp-07',
    name: 'Cún Tai Dài Cinnamoroll Trắng Muốt',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 210000,
    originalPrice: 250000,
    size: '30cm',
    image: 'assets/images/img_36.png',
    badge: 'Sanrio Hot',
    rating: 5,
    reviewsCount: 95,
    description: 'Chú cún tai dài Cinnamoroll trắng muốt, êm ái như một đám mây nhỏ với kích thước 30cm. Sắc màu nhẹ nhàng, kiểu dáng đáng yêu chuẩn phong cách Nhật.',
    material: 'Lông thỏ mịn màng không rụng lông, bông gòn tinh khiết 100%.',
    inStock: true
  },
  {
    id: 'sp-08',
    name: 'Winnie the Pooh Áo Đỏ Ôm Mật',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 195000,
    originalPrice: 230000,
    size: '28cm',
    image: 'assets/images/img_2.png',
    badge: 'Disney Classic',
    rating: 5,
    reviewsCount: 130,
    description: 'Chú gấu Pooh mê mật ong truyền thống với chiếc áo đỏ quen thuộc và chất lông xù mềm mại, ấm áp.',
    material: 'Bông sạch cao cấp, lông xù xoắn mềm mại an toàn cho trẻ em.',
    inStock: true
  },
  {
    id: 'sp-09',
    name: 'Gudetama - Trứng Lười Dễ Thương',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 175000,
    originalPrice: 210000,
    size: '25cm',
    image: 'assets/images/img_37.png',
    badge: 'Hài Hước',
    rating: 4.8,
    reviewsCount: 47,
    description: 'Thánh lười Gudetama với biểu cảm thảnh thơi vô cùng giải trí! Giúp giải tỏa căng thẳng sau một ngày dài làm việc.',
    material: 'Vải thun 4 chiều siêu mịn đàn hồi, bông gòn êm như mây.',
    inStock: true
  },
  {
    id: 'sp-10',
    name: 'Cá Ngố Hangyodon Độc Lạ',
    category: 'gau-vua',
    categoryName: 'Gấu vừa (20-30cm)',
    price: 190000,
    originalPrice: 230000,
    size: '28cm',
    image: 'assets/images/img_38.png',
    badge: 'Độc Lạ',
    rating: 4.9,
    reviewsCount: 53,
    description: 'Chú cá ngố Hangyodon độc lạ với đôi mắt tròn xoe và chiếc miệng cá ngộ nghĩnh, biểu cảm cực kỳ đáng yêu.',
    material: 'Vải nhung tuyết mịn màng, may đường kim chắc chắn.',
    inStock: true
  },

  // --- GẤU TO (30cm trở lên) ---
  {
    id: 'sp-11',
    name: 'My Melody Nơ Hồng Xinh Xắn',
    category: 'gau-to',
    categoryName: 'Gấu to (30cm trở lên)',
    price: 290000,
    originalPrice: 350000,
    size: '45cm',
    image: 'assets/images/img_4.png',
    badge: 'Siêu Hot',
    rating: 5,
    reviewsCount: 156,
    description: 'Bạn thỏ hồng My Melody dịu dàng đáng yêu, kích thước 45cm ôm cực đã tay. Món quà hoàn hảo để dành tặng bạn gái hay em nhỏ.',
    material: 'Lông thỏ siêu mịn cao cấp, nhồi bông gòn trắng đàn hồi không xẹp.',
    inStock: true
  },
  {
    id: 'sp-12',
    name: 'Chuột Mickey Mouse Cổ Điển',
    category: 'gau-to',
    categoryName: 'Gấu to (30cm trở lên)',
    price: 280000,
    originalPrice: 330000,
    size: '40cm',
    image: 'assets/images/img_40.png',
    badge: 'Disney Classic',
    rating: 4.9,
    reviewsCount: 78,
    description: 'Chú chuột Mickey kinh điển của nhà Disney với dáng ngồi gọn gàng, khuôn mặt rạng rỡ, chất lượng hoàn thiện tuyệt hảo.',
    material: 'Vải nhung mịn kết hợp cotton cao cấp, bông tinh khiết an toàn.',
    inStock: true
  },
  {
    id: 'sp-13',
    name: 'Cáo Tuyết Kitsune Nhật Bản',
    category: 'gau-to',
    categoryName: 'Gấu to (30cm trở lên)',
    price: 320000,
    originalPrice: 380000,
    size: '50cm',
    image: 'assets/images/img_17.png',
    badge: 'Độc Quyền',
    rating: 5,
    reviewsCount: 89,
    description: 'Bé cáo tuyết Kitsune mang nét đẹp thần thoại Nhật Bản với bộ lông trắng mướt mềm mại, đôi tai hồng đáng yêu.',
    material: 'Vải lông cừu nhân tạo cao cấp, ruột bông gòn 3D êm ái.',
    inStock: true
  },
  {
    id: 'sp-14',
    name: 'Cún Shiba Inu Béo Tròn',
    category: 'gau-to',
    categoryName: 'Gấu to (30cm trở lên)',
    price: 310000,
    originalPrice: 370000,
    size: '50cm',
    image: 'assets/images/img_18.png',
    badge: 'Yêu Thích',
    rating: 5,
    reviewsCount: 142,
    description: 'Bé cún Shiba Inu má phúng phính, thân hình mập mạp tròn trịa, cảm giác ôm vào lòng cực kỳ ấm áp và bình yên.',
    material: 'Chất vải nhung co giãn 4 chiều mềm mịn, bông nhồi căng phồng.',
    inStock: true
  },
  {
    id: 'sp-15',
    name: 'Gấu Bông Khổng Lồ Teddy ÔMƠ Vỗ Về',
    category: 'gau-to',
    categoryName: 'Gấu to (30cm trở lên)',
    price: 490000,
    originalPrice: 590000,
    size: '80cm - 100cm',
    image: 'assets/images/img_19.png',
    badge: 'Khổng Lồ',
    rating: 5,
    reviewsCount: 210,
    description: 'Chú gấu Teddy khổng lồ signature của ÔMƠ! Thay bạn trao cái ôm siết dịu dàng và ấm áp nhất cho người thương.',
    material: 'Bông gòn bi trắng 100%, áo len dệt thủ công có thể tháo giặt tiện lợi.',
    inStock: true
  }
];

// Tin tức / Blog
const BLOG_POSTS = [
  {
    id: 'post-01',
    title: '5 Cách Chọn Gấu Bông Làm Quà Tặng Phù Hợp Cho Người Thương',
    date: '15 Tháng 04, 2026',
    author: 'ÔMƠ Team',
    category: 'Mẹo Chọn Quà',
    image: 'assets/images/img_43.png',
    url: 'tin-tuc/5-cach-chon-gau-bong-lam-qua-tang/',
    summary: 'Để món quà thêm phần trọn vẹn, việc chọn đúng chú gấu bông là rất quan trọng. Cùng ÔMƠ khám phá 5 tiêu chí: sở thích, độ tuổi, kích thước, chất liệu và dịp tặng quà...',
    content: `
      <p>Giữa nhịp sống hối hả, đôi khi những món quà đắt tiền lại không mang sức mạnh chữa lành bằng một chú gấu bông mềm mại. Gấu bông không chỉ là món đồ chơi, mà còn là "người bạn" đại diện cho sự hiện diện của bạn bên cạnh người thương. Dù là dịp sinh nhật, kỷ niệm hay chỉ đơn giản là một ngày bình thường muốn tạo bất ngờ, một chú gấu bông xinh xắn từ ÔMƠ chắc chắn sẽ là cầu nối tuyệt vời để bạn gửi gắm những yêu thương chưa ngỏ.</p>
      <h3>5 Tiêu chí quan trọng khi chọn gấu bông:</h3>
      <ul>
        <li><strong>1. Dựa vào sở thích:</strong> Hãy để ý xem người nhận thích động vật nào (mèo, cún, thỏ) hay nhân vật hoạt hình gì (Kirby, Winnie the Pooh, Sanrio, Minions...).</li>
        <li><strong>2. Độ tuổi:</strong> Trẻ nhỏ thích hợp với gấu nhỏ, chất liệu kháng khuẩn, mắt thêu an toàn; người lớn thường thích gấu ôm ngủ hoặc gấu Teddy cỡ lớn.</li>
        <li><strong>3. Kích thước:</strong> Chọn size phù hợp với không gian phòng (để bàn học, góc sofa hay ôm trên giường ngủ).</li>
        <li><strong>4. Chất liệu bông:</strong> Ưu tiên vỏ nhung mịn, bông nhồi 100% tinh khiết để đảm bảo an toàn tuyệt đối cho làn da, không gây dị ứng hay rụng lông.</li>
        <li><strong>5. Dịp tặng quà:</strong> Lễ tình nhân có thể tặng gấu ôm tim, tốt nghiệp thì tặng gấu mặc áo cử nhân, sinh nhật tặng gấu ôm bánh kem.</li>
      </ul>
    `
  },
  {
    id: 'post-02',
    title: 'Cách Vệ Sinh Và Bảo Quản Gấu Bông Đúng Cách Tại Nhà Luôn Thơm Tho',
    date: '10 Tháng 04, 2026',
    author: 'ÔMƠ Care',
    category: 'Chăm Sóc Gấu Bông',
    image: 'assets/images/img_44.png',
    summary: 'Giữ cho những "người bạn nhỏ" luôn sạch sẽ, thơm tho là điều rất quan trọng. Tùy vào kích thước của gấu bông, chúng ta sẽ có cách vệ sinh riêng biệt để tránh làm hỏng form dáng...',
    content: `
      <p>Giữ cho những "người bạn nhỏ" luôn sạch sẽ, thơm tho là điều rất quan trọng để bảo vệ sức khỏe hô hấp và giữ gấu luôn bền đẹp như ngày đầu mới mua. Tùy vào kích thước của gấu bông, chúng ta có cách vệ sinh riêng biệt:</p>
      <h3>1. Vệ sinh đối với gấu bông cỡ nhỏ (dưới 30cm)</h3>
      <p>Với những bé gấu nhỏ gọn, bạn có thể tiết kiệm thời gian bằng cách giặt trực tiếp bằng máy giặt:</p>
      <ul>
        <li>Cho gấu vào túi giặt chuyên dụng để bảo vệ lớp lông và chi tiết mắt, mũi không bị cọ xát.</li>
        <li>Chọn chế độ giặt nhẹ dịu (Delicate/Gentle) với nước lạnh hoặc nước ấm dưới 30°C.</li>
        <li>Sử dụng nước giặt dịu nhẹ, hạn chế nước xả quá nồng.</li>
        <li>Phơi ở nơi thoáng gió, có nắng nhẹ tự nhiên để gấu khô thấu bên trong, chống ẩm mốc.</li>
      </ul>
      <h3>2. Vệ sinh đối với gấu bông cỡ lớn (trên 40cm - 1m)</h3>
      <p>Những chú gấu Teddy khổng lồ không thể nhét vừa máy giặt thì sao? Đừng lo lắng:</p>
      <ul>
        <li>Tìm đường chỉ may (thường ở sau lưng gấu), khéo léo tháo một đoạn nhỏ và lấy toàn bộ ruột bông ra ngoài.</li>
        <li>Đem phần ruột bông phơi nắng to để khử khuẩn và làm tơi xốp bông.</li>
        <li>Giặt riêng phần vỏ vải như quần áo thông thường, ngâm xả thơm tho.</li>
        <li>Khi vỏ khô hoàn toàn, nhồi bông lại căng tròn và dùng kim khâu giấu chỉ khép lại. Chú gấu sẽ xinh và thơm phức như mới!</li>
      </ul>
    `
  },
  {
    id: 'post-03',
    title: 'Gấu Bông Có Ý Nghĩa Gì Khi Làm Quà Tặng Người Yêu & Bạn Bè?',
    date: '02 Tháng 04, 2026',
    author: 'ÔMƠ Story',
    category: 'Ý Nghĩa Yêu Thương',
    image: 'assets/images/img_48.png',
    summary: 'Tặng gấu bông mang ý nghĩa về sự bảo bọc, chở che và khao khát được đồng hành. "Dù không có mặt ở đây, chú gấu này sẽ thay mình lắng nghe và ôm ấp bạn những lúc mỏi mệt"...',
    content: `
      <p>Bạn có biết vì sao gấu bông luôn nằm trong danh sách những món quà được yêu thích nhất mọi thời đại không? Tặng gấu bông mang ý nghĩa biểu tượng sâu sắc:</p>
      <ul>
        <li><strong>Biểu tượng của sự chở che:</strong> Khi bạn tặng ai đó một chú gấu, bạn đang ngầm gửi gắm thông điệp: <em>"Dù mình không ở ngay bên cạnh, chú gấu này sẽ thay mình vỗ về, lắng nghe và ở bên bạn mỗi đêm."</em></li>
        <li><strong>Sự ấm áp và gắn kết:</strong> Một cái ôm siết chú gấu êm mềm giúp sản sinh hormone thư giãn, xua tan căng thẳng sau ngày dài học tập hay làm việc.</li>
        <li><strong>Lưu giữ kỷ niệm bền lâu:</strong> Khác với hoa tươi dễ tàn hay quà ăn uống nhanh hết, gấu bông sẽ ở bên người nhận suốt nhiều năm, trở thành một chứng nhân tình cảm đầy trân quý.</li>
      </ul>
    `
  }
];

// Xuất ra window để mọi script dùng chung
if (typeof window !== 'undefined') {
  window.PRODUCTS = PRODUCTS;
  window.BLOG_POSTS = BLOG_POSTS;
}
