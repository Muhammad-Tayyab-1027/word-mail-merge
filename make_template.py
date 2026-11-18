"""Creates templates/interview_invite.docx, the sample letter template. Run once; edit the .docx in Word afterwards."""
from pathlib import Path

from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt

doc = Document()
style = doc.styles["Normal"]
style.font.name = "Arial"
style.font.size = Pt(11)

head = doc.sections[0].header.paragraphs[0]
head.text = "Example Services Ltd | 12 Sample Street, Dublin 2"
head.alignment = WD_ALIGN_PARAGRAPH.RIGHT

doc.add_paragraph("{{letter_date}}")
doc.add_paragraph("{{first_name}} {{last_name}}\n{{address_line}}\n{{city}}")
doc.add_paragraph("Dear {{first_name}},")
title = doc.add_paragraph()
title.add_run("Interview for the position of {{role}}").bold = True
doc.add_paragraph(
    "Thank you for applying for the {{role}} role. We are pleased to invite you to an interview "
    "on {{interview_date}} at {{interview_time}}.")
table = doc.add_table(rows=3, cols=2)
table.style = "Table Grid"
for i, (label, value) in enumerate([("Location", "{{location}}"), ("Interviewer", "{{interviewer}}"),
                                    ("Reference", "{{reference}}")]):
    table.cell(i, 0).text = label
    table.cell(i, 1).text = value
doc.add_paragraph()
doc.add_paragraph("Please reply to this letter or email recruitment@example.ie to confirm you can attend. "
                  "Bring photo ID with you on the day.")
doc.add_paragraph("Kind regards,\n\nRecruitment Team")
doc.sections[0].footer.paragraphs[0].text = "Sample letter for a portfolio project. Company and people are made up."

Path("templates").mkdir(exist_ok=True)
doc.save("templates/interview_invite.docx")
print("saved templates/interview_invite.docx")
