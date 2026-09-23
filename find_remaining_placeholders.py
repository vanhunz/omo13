import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document('DoAnNhom13_IS425_OMO.docx')

print("--- REMAINING PLACEHOLDERS IN PARAGRAPHS ---")
for i, p in enumerate(doc.paragraphs):
    t = p.text.strip()
    if any(k in t.lower() for k in ['chèn ảnh', 'chụp ảnh', 'chụp phần', 'mô tả ở đây']):
        print(f"P[{i}]: \"{t}\"")
