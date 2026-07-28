from docx import Document
b = Document(r"E:\OC DOCS\Comprendre_Optimum_Control_backup.docx")
for i in range(33, 37):
    p = b.paragraphs[i]
    t = p.text.strip()[:80] if p.text.strip() else "(vide)"
    s = p.style.name if p.style else ""
    print(i, s, t)
