# -*- coding: utf-8 -*-
import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document('NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx')

print(f"=== VERIFICATION OF NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx ===")
print(f"Paragraphs: {len(doc.paragraphs)}, Tables: {len(doc.tables)}")

# 1. Check for any forbidden/outdated words
forbidden = ['google sites', 'google.com/view', 'capybara', 'lotso', '137 nguyễn thị thập', 'welcomeomo']
errors_found = 0

for i, p in enumerate(doc.paragraphs):
    for f in forbidden:
        if f in p.text.lower():
            print(f"[ERROR in P{i}] Found '{f}': {p.text[:100]}...")
            errors_found += 1

for t_idx, t in enumerate(doc.tables):
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            for f in forbidden:
                if f in c.text.lower():
                    print(f"[ERROR in Table {t_idx} R{r_idx} C{c_idx}] Found '{f}': {c.text[:100]}...")
                    errors_found += 1

if errors_found == 0:
    print("SUCCESS: 0 outdated/incorrect terms found in the entire document!")
else:
    print(f"WARNING: Found {errors_found} potential issues.")

# 2. Print key sections to verify content
print("\n--- SAMPLE CHECK: Section 3.1.1 Đặc tả hệ thống ---")
for p in doc.paragraphs:
    if "* Đặc tả toàn diện kiến trúc" in p.text:
        print(p.text[:400] + "...\n")

print("--- SAMPLE CHECK: Table 22 Row 3 (Body Trang chủ) ---")
if len(doc.tables) > 22:
    print(doc.tables[22].rows[2].cells[1].text[:300] + "...\n")

print("--- SAMPLE CHECK: Table 11 Row 1 (Trang chủ Body & SEO) ---")
if len(doc.tables) > 11:
    print(doc.tables[11].rows[1].cells[1].text[:300] + "...\n")

print("Verification complete!")
