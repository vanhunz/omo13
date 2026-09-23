import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document('DoAnNhom13_IS425_OMO.docx')

print('--- CHECKING ALL TABLES FOR PLACEHOLDER OR EMPTY TEXT ---')
for i, t in enumerate(doc.tables):
    for r_idx, r in enumerate(t.rows):
        for c_idx, c in enumerate(r.cells):
            txt = c.text.strip()
            imgs = len(c._tc.xpath('.//w:drawing') + c._tc.xpath('.//w:pict'))
            if any(p in txt.lower() for p in ['viết nội dung', 'chèn ảnh', 'mô tả ở đây', '...']) or (txt == '' and imgs == 0):
                print(f'Table {i+1} [Row {r_idx}, Col {c_idx}]: "{txt[:80]}" (imgs: {imgs})')
