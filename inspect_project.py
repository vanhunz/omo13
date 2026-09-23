import sys
import os
import json
import docx

def analyze_all():
    output_lines = []
    
    # 1. Crawled JSON analysis
    with open('all_pages_crawled.json', 'r', encoding='utf-8') as f:
        crawled = json.load(f)
    
    output_lines.append("=== CRAWLED PAGES SUMMARY ===")
    for page_name, page_data in crawled.items():
        output_lines.append(f"Page: {page_name}")
        output_lines.append(f"  URL: {page_data.get('url', '')}")
        output_lines.append(f"  Title: {page_data.get('title', '')}")
        output_lines.append(f"  Images count: {len(page_data.get('images', []))}")
        text_content = page_data.get('text', '')
        output_lines.append(f"  Text length: {len(text_content)} chars")
        output_lines.append(f"  Text preview: {text_content[:200]}...")
        output_lines.append("-" * 40)

    # 2. HTML Files analysis
    output_lines.append("\n=== HTML FILES IN PROJECT ===")
    html_files = [f for f in os.listdir('.') if f.endswith('.html')]
    for hf in html_files:
        with open(hf, 'r', encoding='utf-8') as f:
            content = f.read()
        output_lines.append(f"File: {hf} ({len(content)} bytes)")
        # Extract title, meta description, h1, h2
        import re
        title_match = re.search(r'<title>(.*?)</title>', content, re.IGNORECASE)
        meta_desc = re.search(r'<meta\s+name=["\']description["\']\s+content=["\'](.*?)["\']', content, re.IGNORECASE)
        h1s = re.findall(r'<h1[^>]*>(.*?)</h1>', content, re.IGNORECASE | re.DOTALL)
        h2s = re.findall(r'<h2[^>]*>(.*?)</h2>', content, re.IGNORECASE | re.DOTALL)
        output_lines.append(f"  Title: {title_match.group(1) if title_match else 'None'}")
        output_lines.append(f"  Meta Desc: {meta_desc.group(1) if meta_desc else 'None'}")
        output_lines.append(f"  H1s ({len(h1s)}): {[re.sub(r'<[^>]+>', '', h).strip() for h in h1s]}")
        output_lines.append(f"  H2s ({len(h2s)}): {[re.sub(r'<[^>]+>', '', h).strip() for h in h2s[:5]]}")
        output_lines.append("-" * 40)

    # 3. Products in JS
    output_lines.append("\n=== PRODUCTS IN js/products.js ===")
    with open('js/products.js', 'r', encoding='utf-8') as f:
        js_text = f.read()
    # Extract PRODUCTS array or analyze lines
    output_lines.append(f"Length of js/products.js: {len(js_text)}")
    # Count products
    prod_names = re.findall(r"name:\s*['\"]([^'\"]+)['\"]", js_text)
    prod_cats = re.findall(r"categoryName:\s*['\"]([^'\"]+)['\"]", js_text)
    prod_prices = re.findall(r"price:\s*([0-9]+)", js_text)
    output_lines.append(f"Total products found: {len(prod_names)}")
    for i in range(len(prod_names)):
        cat = prod_cats[i] if i < len(prod_cats) else ''
        pr = prod_prices[i] if i < len(prod_prices) else ''
        output_lines.append(f"  {i+1}. {prod_names[i]} | Danh mục: {cat} | Giá: {pr}đ")

    # 4. News / Articles in tin-tuc.html
    output_lines.append("\n=== NEWS / ARTICLES IN tin-tuc.html ===")
    with open('tin-tuc.html', 'r', encoding='utf-8') as f:
        news_content = f.read()
    news_titles = re.findall(r'<h3[^>]*class=["\'][^"\']*article-title[^"\']*["\'][^>]*>(.*?)</h3>', news_content, re.IGNORECASE | re.DOTALL)
    if not news_titles:
        news_titles = re.findall(r'<h3[^>]*>(.*?)</h3>', news_content, re.IGNORECASE | re.DOTALL)
    output_lines.append(f"Total news/articles found in tin-tuc.html: {len(news_titles)}")
    for i, nt in enumerate(news_titles):
        clean_nt = re.sub(r'<[^>]+>', '', nt).strip()
        output_lines.append(f"  {i+1}. {clean_nt}")

    with open('project_analysis_summary.txt', 'w', encoding='utf-8') as f:
        f.write('\n'.join(output_lines))
    print("Saved project_analysis_summary.txt successfully.")

if __name__ == '__main__':
    analyze_all()
