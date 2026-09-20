/**
 * ÔMƠ SHOP - CART & INTERACTION LOGIC
 */

const CART_STORAGE_KEY = 'omo_cart_items_v1';
const VOUCHER_STORAGE_KEY = 'omo_active_voucher_v1';
const WISHLIST_STORAGE_KEY = 'omo_wishlist_items_v1';
const CHAT_STORAGE_KEY = 'omo_chat_messages_v1';

function getChatMessages() {
  try {
    const data = localStorage.getItem(CHAT_STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch (e) {
    return [];
  }
}

function saveChatMessages(messages) {
  localStorage.setItem(CHAT_STORAGE_KEY, JSON.stringify(messages.slice(-30)));
}

function addChatMessage(text, sender = 'bot') {
  const messages = getChatMessages();
  messages.push({ text, sender, time: Date.now() });
  saveChatMessages(messages);
  renderChatMessages();
}

function getChatReply(message) {
  const text = message.toLowerCase();
  if (text.includes('giá') || text.includes('rẻ')) return 'ÔMƠ có nhiều mẫu từ 45.000 VND. Bạn muốn gấu nhỏ, gấu vừa hay gấu to?';
  if (text.includes('ship') || text.includes('giao')) return 'Đà Nẵng có giao hỏa tốc 1-2 giờ. Các tỉnh khác tụi mình gửi nhanh trong 2-4 ngày nhé!';
  if (text.includes('đổi') || text.includes('trả')) return 'Bạn có thể đổi trả trong 3 ngày nếu sản phẩm còn nguyên trạng. Xem chi tiết tại trang Chính sách nha.';
  if (text.includes('tư vấn') || text.includes('quà')) return 'Bạn tặng ai và ngân sách khoảng bao nhiêu? Mình sẽ gợi ý một bé gấu thật hợp gu!';
  if (text.includes('hello') || text.includes('hi') || text.includes('chào')) return 'Chào bạn, mình là Mơ đây! Mình có thể tư vấn quà, giá, size và giao hàng cho bạn.';
  return 'Mình đã ghi nhận rồi nha. Bạn thử hỏi về giá, giao hàng, đổi trả hoặc gợi ý quà để mình hỗ trợ nhanh hơn nhé!';
}

function renderChatMessages() {
  const list = document.querySelector('.omo-chat-messages');
  if (!list) return;
  list.innerHTML = '';
  getChatMessages().forEach(message => {
    const item = document.createElement('div');
    item.className = `omo-chat-message ${message.sender}`;
    item.textContent = message.text;
    list.appendChild(item);
  });
  list.scrollTop = list.scrollHeight;
}

function submitChatMessage(text) {
  const cleanText = text.trim();
  if (!cleanText) return;
  addChatMessage(cleanText, 'user');
  window.setTimeout(() => addChatMessage(getChatReply(cleanText)), 450);
}

function setupChatWidget() {
  if (document.querySelector('.omo-chat')) return;
  const chat = document.createElement('aside');
  chat.className = 'omo-chat';
  chat.innerHTML = `
    <button class="omo-chat-launcher" type="button" aria-label="Mở chat với ÔMƠ" aria-expanded="false">
      <span class="omo-chat-launcher-icon">💬</span><span class="omo-chat-launcher-label">Chat với Mơ</span>
    </button>
    <div class="omo-chat-panel" aria-hidden="true">
      <div class="omo-chat-head">
        <div><strong>Trợ lý Mơ</strong><small><i></i> Đang online</small></div>
        <button class="omo-chat-close" type="button" aria-label="Đóng chat">×</button>
      </div>
      <div class="omo-chat-messages"></div>
      <div class="omo-chat-quick-replies">
        <button type="button">Giá bao nhiêu?</button><button type="button">Giao hàng</button><button type="button">Gợi ý quà</button>
      </div>
      <form class="omo-chat-form">
        <input type="text" placeholder="Nhắn Mơ một câu nhé..." aria-label="Tin nhắn chat">
        <button type="submit" aria-label="Gửi tin nhắn">➤</button>
      </form>
    </div>
  `;
  document.body.appendChild(chat);
  const launcher = chat.querySelector('.omo-chat-launcher');
  const panel = chat.querySelector('.omo-chat-panel');
  const close = chat.querySelector('.omo-chat-close');
  const form = chat.querySelector('.omo-chat-form');
  const input = form.querySelector('input');
  const openChat = () => {
    chat.classList.add('is-open');
    launcher.setAttribute('aria-expanded', 'true');
    panel.setAttribute('aria-hidden', 'false');
    input.focus();
  };
  const closeChat = () => {
    chat.classList.remove('is-open');
    launcher.setAttribute('aria-expanded', 'false');
    panel.setAttribute('aria-hidden', 'true');
  };
  launcher.addEventListener('click', () => chat.classList.contains('is-open') ? closeChat() : openChat());
  close.addEventListener('click', closeChat);
  form.addEventListener('submit', event => {
    event.preventDefault();
    submitChatMessage(input.value);
    input.value = '';
  });
  chat.querySelectorAll('.omo-chat-quick-replies button').forEach(button => {
    button.addEventListener('click', () => submitChatMessage(button.textContent));
  });
  if (!getChatMessages().length) {
    addChatMessage('Chào bạn, mình là Mơ! Bạn cần tìm một chiếc ôm hay một món quà thật xinh hôm nay?');
  } else {
    renderChatMessages();
  }
}

function getWishlist() {
  try {
    const data = localStorage.getItem(WISHLIST_STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch (e) {
    return [];
  }
}

function toggleWishlist(productId) {
  const wishlist = getWishlist();
  const isLoved = wishlist.includes(productId);
  const nextWishlist = isLoved
    ? wishlist.filter(id => id !== productId)
    : [...wishlist, productId];

  localStorage.setItem(WISHLIST_STORAGE_KEY, JSON.stringify(nextWishlist));
  updateWishlistUI();

  const product = (window.PRODUCTS || []).find(item => item.id === productId);
  showToast(isLoved ? `Đã bỏ "${product ? product.name : 'sản phẩm'}" khỏi yêu thích.` : `Đã lưu "${product ? product.name : 'sản phẩm'}" vào yêu thích!`, isLoved ? '♡' : '♥');
}

function updateWishlistUI() {
  const wishlist = getWishlist();
  document.querySelectorAll('.wishlist-btn').forEach(button => {
    const isLoved = wishlist.includes(button.dataset.productId);
    button.classList.toggle('is-loved', isLoved);
    button.setAttribute('aria-pressed', String(isLoved));
    button.setAttribute('aria-label', isLoved ? 'Bỏ khỏi yêu thích' : 'Thêm vào yêu thích');
    button.textContent = isLoved ? '♥' : '♡';
  });
  document.querySelectorAll('.wishlist-count').forEach(count => {
    count.textContent = wishlist.length;
  });
}

function setupInteractionLayer() {
  const progress = document.createElement('div');
  progress.className = 'scroll-progress';
  document.body.appendChild(progress);

  const wishlistFab = document.createElement('button');
  wishlistFab.className = 'wishlist-fab';
  wishlistFab.type = 'button';
  wishlistFab.innerHTML = '<span>♥</span> Đã lưu <span class="wishlist-count">0</span>';
  wishlistFab.addEventListener('click', () => {
    const count = getWishlist().length;
    showToast(count ? `Bạn đang lưu ${count} sản phẩm yêu thích.` : 'Bạn chưa lưu sản phẩm nào. Hãy bấm trái tim để lưu nhé!', '♥');
  });
  document.body.appendChild(wishlistFab);

  const revealObserver = new IntersectionObserver(entries => {
    entries.forEach(entry => {
      if (entry.isIntersecting) {
        entry.target.classList.add('is-visible');
        revealObserver.unobserve(entry.target);
      }
    });
  }, { threshold: 0.12 });

  const decorateRevealElements = () => {
    document.querySelectorAll('.section-padding, .product-card, .why-card, .blog-card, .story-summary-box, .hotline-banner, .site-footer > .container').forEach(element => {
      if (!element.classList.contains('reveal-on-scroll')) {
        element.classList.add('reveal-on-scroll');
        revealObserver.observe(element);
      }
    });
    updateWishlistUI();
  };

  decorateRevealElements();
  const contentObserver = new MutationObserver(decorateRevealElements);
  document.querySelectorAll('#homeProductGrid, #shopProductGrid, #homeBlogGrid').forEach(grid => {
    contentObserver.observe(grid, { childList: true });
  });

  const header = document.querySelector('.main-header');
  const updateScrollState = () => {
    const pageHeight = document.documentElement.scrollHeight - window.innerHeight;
    const progressValue = pageHeight > 0 ? (window.scrollY / pageHeight) * 100 : 0;
    progress.style.width = `${progressValue}%`;
    if (header) header.classList.toggle('is-scrolled', window.scrollY > 12);
  };
  window.addEventListener('scroll', updateScrollState, { passive: true });
  updateScrollState();
  setupChatWidget();
}

// Lấy danh sách giỏ hàng
function getCart() {
  try {
    const data = localStorage.getItem(CART_STORAGE_KEY);
    return data ? JSON.parse(data) : [];
  } catch (e) {
    return [];
  }
}

// Lưu giỏ hàng
function saveCart(cart) {
  localStorage.setItem(CART_STORAGE_KEY, JSON.stringify(cart));
  updateCartBadge();
}

// Cập nhật số lượng hiển thị trên icon giỏ hàng ở header
function updateCartBadge() {
  const cart = getCart();
  const totalCount = cart.reduce((sum, item) => sum + item.quantity, 0);
  const badges = document.querySelectorAll('.cart-badge');
  badges.forEach(b => {
    b.textContent = totalCount;
    b.style.display = totalCount > 0 ? 'flex' : 'none';
  });
}

// Thêm vào giỏ hàng
function addToCart(productId, quantity = 1, options = {}) {
  const product = (window.PRODUCTS || []).find(p => p.id === productId);
  if (!product) return;

  const cart = getCart();
  const existingIndex = cart.findIndex(item => item.id === productId);

  if (existingIndex > -1) {
    cart[existingIndex].quantity += quantity;
  } else {
    cart.push({
      id: product.id,
      name: product.name,
      price: product.price,
      originalPrice: product.originalPrice,
      image: product.image,
      size: options.size || product.size,
      quantity: quantity
    });
  }

  saveCart(cart);
  showToast(`Đã thêm "${product.name}" vào giỏ hàng! 💖`);
}

// Cập nhật số lượng
function updateCartQuantity(productId, newQty) {
  let cart = getCart();
  if (newQty <= 0) {
    cart = cart.filter(item => item.id !== productId);
  } else {
    const item = cart.find(i => i.id === productId);
    if (item) {
      item.quantity = newQty;
    }
  }
  saveCart(cart);
  if (typeof renderCartPage === 'function') {
    renderCartPage();
  }
  if (typeof renderCheckoutPage === 'function') {
    renderCheckoutPage();
  }
}

// Xóa sản phẩm khỏi giỏ
function removeFromCart(productId) {
  let cart = getCart();
  const item = cart.find(i => i.id === productId);
  cart = cart.filter(i => i.id !== productId);
  saveCart(cart);
  if (item) {
    showToast(`Đã xóa "${item.name}" khỏi giỏ hàng.`);
  }
  if (typeof renderCartPage === 'function') {
    renderCartPage();
  }
}

// Làm trống giỏ hàng
function clearCart() {
  localStorage.removeItem(CART_STORAGE_KEY);
  localStorage.removeItem(VOUCHER_STORAGE_KEY);
  updateCartBadge();
  if (typeof renderCartPage === 'function') {
    renderCartPage();
  }
}

// Định dạng tiền tệ VND
function formatVND(amount) {
  return new Intl.NumberFormat('vi-VN', { style: 'currency', currency: 'VND' }).format(amount).replace('₫', 'VND');
}

// Hiển thị Toast thông báo
function showToast(message, icon = '🌸') {
  let container = document.querySelector('.toast-container');
  if (!container) {
    container = document.createElement('div');
    container.className = 'toast-container';
    document.body.appendChild(container);
  }

  const toast = document.createElement('div');
  toast.className = 'toast';
  toast.innerHTML = `<span class="toast-icon">${icon}</span> <span>${message}</span>`;
  container.appendChild(toast);

  // Trigger animation
  setTimeout(() => toast.classList.add('show'), 10);

  // Remove after 3.5s
  setTimeout(() => {
    toast.classList.remove('show');
    setTimeout(() => toast.remove(), 400);
  }, 3500);
}

// Quick View Modal
function openQuickView(productId) {
  const product = (window.PRODUCTS || []).find(p => p.id === productId);
  if (!product) return;

  let modalBackdrop = document.querySelector('#quickViewModal');
  if (!modalBackdrop) {
    modalBackdrop = document.createElement('div');
    modalBackdrop.id = 'quickViewModal';
    modalBackdrop.className = 'modal-backdrop';
    document.body.appendChild(modalBackdrop);
  }

  modalBackdrop.innerHTML = `
    <div class="modal-card">
      <button class="modal-close-btn" onclick="closeQuickView()">&times;</button>
      <div class="modal-product-grid">
        <div class="modal-img-wrap">
          <img src="${product.image}" alt="${product.name}" class="modal-img">
        </div>
        <div class="modal-details">
          <span class="product-category-tag">${product.categoryName}</span>
          <h2 style="font-family: var(--font-heading); font-size: 1.5rem; margin: 8px 0; color: var(--text-main);">${product.name}</h2>
          <div class="product-rating" style="margin-bottom: 12px;">
            ⭐⭐⭐⭐⭐ <span class="rating-count">(${product.reviewsCount} đánh giá yêu thích)</span>
          </div>
          <div class="product-price-row" style="margin-bottom: 16px;">
            <span class="current-price" style="font-size: 1.5rem;">${formatVND(product.price)}</span>
            <span class="original-price" style="font-size: 1rem;">${formatVND(product.originalPrice)}</span>
          </div>
          <p style="color: var(--text-muted); font-size: 0.95rem; line-height: 1.7; margin-bottom: 16px;">
            ${product.description}
          </p>
          <div style="background: var(--soft-pink-bg); padding: 12px 16px; border-radius: var(--radius-sm); margin-bottom: 20px; font-size: 0.88rem;">
            <strong>✨ Chất liệu:</strong> ${product.material}<br>
            <strong>📏 Kích thước:</strong> ${product.size}
          </div>
          <div style="display: flex; gap: 12px; align-items: center; margin-bottom: 20px;">
            <span style="font-weight: 600; font-size: 0.9rem;">Số lượng:</span>
            <div class="qty-stepper">
              <button class="qty-btn" onclick="adjustModalQty(-1)">-</button>
              <input type="number" id="modalQtyInput" class="qty-input" value="1" min="1" max="99">
              <button class="qty-btn" onclick="adjustModalQty(1)">+</button>
            </div>
          </div>
          <div style="display: flex; gap: 12px; flex-wrap: wrap;">
            <button class="btn btn-primary" onclick="handleAddFromModal('${product.id}')" style="flex-grow: 1;">
              Thêm Vào Giỏ Hàng 💕
            </button>
            <a href="cua-hang.html" class="btn btn-outline">Xem Thêm</a>
          </div>
        </div>
      </div>
    </div>
  `;

  modalBackdrop.classList.add('show');
}

function closeQuickView() {
  const modalBackdrop = document.querySelector('#quickViewModal');
  if (modalBackdrop) {
    modalBackdrop.classList.remove('show');
  }
}

function adjustModalQty(delta) {
  const input = document.querySelector('#modalQtyInput');
  if (input) {
    let val = parseInt(input.value) || 1;
    val = Math.max(1, val + delta);
    input.value = val;
  }
}

function handleAddFromModal(productId) {
  const input = document.querySelector('#modalQtyInput');
  const qty = input ? parseInt(input.value) || 1 : 1;
  addToCart(productId, qty);
  closeQuickView();
}

// Khởi chạy khi load trang
document.addEventListener('DOMContentLoaded', () => {
  updateCartBadge();
  setupInteractionLayer();

  // Mobile menu toggle
  const toggleBtn = document.querySelector('.mobile-menu-toggle');
  const navMenu = document.querySelector('.nav-menu');
  if (toggleBtn && navMenu) {
    toggleBtn.setAttribute('aria-expanded', 'false');
    toggleBtn.addEventListener('click', () => {
      const isOpen = navMenu.classList.toggle('active');
      toggleBtn.setAttribute('aria-expanded', String(isOpen));
    });

    navMenu.querySelectorAll('a').forEach(link => {
      link.addEventListener('click', () => {
        navMenu.classList.remove('active');
        toggleBtn.setAttribute('aria-expanded', 'false');
      });
    });

    document.addEventListener('click', event => {
      if (!navMenu.contains(event.target) && !toggleBtn.contains(event.target)) {
        navMenu.classList.remove('active');
        toggleBtn.setAttribute('aria-expanded', 'false');
      }
    });
  }

  // Back to top
  const backTopBtn = document.querySelector('.floating-btn.back-top');
  if (backTopBtn) {
    window.addEventListener('scroll', () => {
      if (window.scrollY > 300) {
        backTopBtn.classList.add('show');
      } else {
        backTopBtn.classList.remove('show');
      }
    });
    backTopBtn.addEventListener('click', () => {
      window.scrollTo({ top: 0, behavior: 'smooth' });
    });
  }
});
