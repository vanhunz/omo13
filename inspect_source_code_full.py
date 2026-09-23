import sys
import json
import re
import glob

sys.stdout.reconfigure(encoding='utf-8')

# 1. Read js/products.js
with open('js/products.js', 'r', encoding='utf-8') as f:
    js_content = f.read()

# Extract products array
print("=== PRODUCTS FROM js/products.js ===")
products = []
matches = re.findall(r'\{\s*id:\s*(\d+),\s*name:\s*[\'"](.*?)[\'"],\s*category:\s*[\'"](.*?)[\'"],\s*price:\s*(\d+),\s*originalPrice:\s*(\d+),\s*image:\s*[\'"](.*?)[\'"],\s*tag:\s*[\'"](.*?)[\'"],\s*rating:\s*([\d\.]+),\s*sold:\s*(\d+),\s*description:\s*[\'"](.*?)[\'"]', js_content, re.DOTALL)

print(f"Found {len(matches)} products:")
for p in matches:
    p_id, name, cat, price, orig_price, img, tag, rating, sold, desc = p
    print(f"- ID {p_id}: {name} | Cat: {cat} | Price: {int(price):,}đ (Orig: {int(orig_price):,}đ) | Tag: {tag} | Sold: {sold} | Desc: {desc[:60]}...")

# 2. Read all HTML pages
print("\n=== ALL HTML PAGES DETAILS ===")
html_files = glob.glob('*.html')
for hf in sorted(html_files):
    with open(hf, 'r', encoding='utf-8') as f:
        content = f.read()
        title_m = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE)
        meta_desc = re.search(r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']', content, re.IGNORECASE)
        h1_m = re.findall(r'<h1.*?>(.*?)</h1>', content, re.IGNORECASE | re.DOTALL)
        h2_m = re.findall(r'<h2.*?>(.*?)</h2>', content, re.IGNORECASE | re.DOTALL)
        
        title = title_m.group(1).strip() if title_m else "N/A"
        desc = meta_desc.group(1).strip() if meta_desc else "N/A"
        h1_clean = [re.sub(r'<.*?>', '', x).strip() for x in h1_m]
        h2_clean = [re.sub(r'<.*?>', '', x).strip() for x in h2_m]
        
        print(f"\n--- FILE: {hf} ---")
        print(f"Title ({len(title)} chars): {title}")
        print(f"Meta ({len(desc)} chars): {desc}")
        print(f"H1: {h1_clean}")
        print(f"H2 ({len(h2_clean)} tags): {h2_clean}")

# 3. Read cart.js features
print("\n=== JS/CART.JS FEATURES ===")
with open('js/cart.js', 'r', encoding='utf-8') as f:
    cart_js = f.read()
    coupons = re.findall(r'[\'"]([A-Z0-9]+)[\'"]\s*:\s*\{.*?discount:\s*([\d\.]+)', cart_js, re.DOTALL)
    print(f"Coupons in cart.js: {coupons}")
