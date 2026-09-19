/**
 * ÔMƠ SHOP - CART & INTERACTION LOGIC
 */

const CART_STORAGE_KEY = 'omo_cart_items_v1';
const VOUCHER_STORAGE_KEY = 'omo_active_voucher_v1';

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

  // Mobile menu toggle
  const toggleBtn = document.querySelector('.mobile-menu-toggle');
  const navMenu = document.querySelector('.nav-menu');
  if (toggleBtn && navMenu) {
    toggleBtn.addEventListener('click', () => {
      navMenu.classList.toggle('active');
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
