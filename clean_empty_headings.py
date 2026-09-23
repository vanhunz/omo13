import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

doc = docx.Document('DoAnNhom13_IS425_OMO.docx')

# Remove empty heading paragraphs or redundant empty paragraphs
paras_to_remove = []
for p in doc.paragraphs:
    # If paragraph has no text and no drawings
    drawings = p._p.xpath('.//w:drawing') + p._p.xpath('.//w:pict')
    if not p.text.strip() and len(drawings) == 0:
        # Check if it is not a required blank line
        if p.style.name.startswith('Heading'):
            paras_to_remove.append(p)

print(f"Removing {len(paras_to_remove)} empty heading paragraphs...")
for p in paras_to_remove:
    p._element.getparent().remove(p._element)

doc.save('DoAnNhom13_IS425_OMO.docx')
doc.save('NHOM13_IS425_BAO_CAO_DO_AN_HOAN_CHINH.docx')
print("Saved clean document!")
