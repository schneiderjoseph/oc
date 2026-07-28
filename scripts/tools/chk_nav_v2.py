from docx import Document
v = Document(r"E:\OC DOCS\Comprendre_Optimum_Control_V2.docx")
lines = []
for i in range(33, 46):
    p = v.paragraphs[i]
    t = p.text.strip()[:75] if p.text.strip() else "(vide)"
    s = p.style.name if p.style else ""
    lines.append(f"{i}|{s}|{t}")
open(r"E:\OC DOCS\nav_section.txt", "w", encoding="utf-8").write("\n".join(lines))
