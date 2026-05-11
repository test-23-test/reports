import sys, io, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
from docx import Document

doc = Document('CU_finalprojectreport_rewritten.docx')

full = ' '.join(doc.paragraphs[i].text for i in range(85, 491) if doc.paragraphs[i].text.strip())

# Check for ANY remaining AI-detectable Unicode artifacts
checks = {
    "Smart left double quote": len(re.findall(r'\u201c', full)),
    "Smart right double quote": len(re.findall(r'\u201d', full)),
    "Smart left single quote": len(re.findall(r'\u2018', full)),
    "Smart right single quote": len(re.findall(r'\u2019', full)),
    "Em-dash": len(re.findall(r'\u2014', full)),
    "En-dash": len(re.findall(r'\u2013', full)),
    "Math minus": len(re.findall(r'\u2212', full)),
    "Ellipsis char": len(re.findall(r'\u2026', full)),
    "Non-breaking space": len(re.findall(r'\u00a0', full)),
    "Thin space": len(re.findall(r'\u2009', full)),
    "Zero-width space": len(re.findall(r'\u200b', full)),
    "Zero-width joiner": len(re.findall(r'\u200d', full)),
    "BOM": len(re.findall(r'\ufeff', full)),
    "Word joiner": len(re.findall(r'\u2060', full)),
    "Bullet": len(re.findall(r'[\u2022\u25e6\u00b7]', full)),
}

print("=== UNICODE ARTIFACT SCAN (Chapters 1-8) ===")
all_clean = True
for name, count in checks.items():
    status = "CLEAN" if count == 0 else f"FOUND {count}"
    if count > 0:
        all_clean = False
    print(f"  {name}: {status}")

print()
if all_clean:
    print("ALL AI-DETECTABLE UNICODE ARTIFACTS: REMOVED")
else:
    print("WARNING: Some artifacts remain")

# Show what IS in the text now (should be plain ASCII-like)
straight_apos = full.count("'")
straight_dq = full.count('"')
hyphens = full.count("-")
questions = full.count("?")
print()
print("=== CLEAN TEXT MARKERS ===")
print(f"  Straight apostrophes: {straight_apos}")
print(f"  Straight double quotes: {straight_dq}")
print(f"  Plain hyphens: {hyphens}")
print(f"  Question marks: {questions}")

# Sample output
print()
print("=== SAMPLE CLEANED TEXT ===")
for idx in [85, 95, 190, 434, 462]:
    print(f"P{idx}: {doc.paragraphs[idx].text[:130]}...")
    print()

# Structure check
orig = Document('CU_finalprojectreport-1.docx')
print(f"Structure: {len(doc.paragraphs)} paras, {len(doc.tables)} tables")
o_imgs = len([r for r in orig.part.rels.values() if 'image' in r.reltype])
r_imgs = len([r for r in doc.part.rels.values() if 'image' in r.reltype])
print(f"Images: {r_imgs} (orig: {o_imgs})")
for idx, label in [(6,"Declaration"),(19,"Ack"),(492,"Refs"),(518,"Appendix")]:
    ok = orig.paragraphs[idx].text == doc.paragraphs[idx].text
    print(f"{label}: {'INTACT' if ok else 'CHANGED'}")
