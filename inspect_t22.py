# -*- coding: utf-8 -*-
import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')
doc = docx.Document('NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx')

t22 = doc.tables[22]
print(f"Table 22 has {len(t22.rows)} rows:")
for i, r in enumerate(t22.rows):
    print(f"Row {i}: Col0='{r.cells[0].text[:30]}' | Col1='{r.cells[1].text[:80]}...'")
