from pathlib import Path

from docx import Document

from mail_merge import clean_row, fill_paragraph, format_date, merge, missing_fields, safe_filename, template_fields


def make_template(path):
    doc = Document()
    p = doc.add_paragraph()
    # Word often splits text into several runs; the placeholder here spans three of them.
    p.add_run("Dear {{fir")
    p.add_run("st_na")
    p.add_run("me}},")
    doc.add_paragraph("Your interview is on {{interview_date}}.")
    table = doc.add_table(rows=1, cols=1)
    table.cell(0, 0).text = "Ref: {{reference}}"
    doc.sections[0].footer.paragraphs[0].text = "Role: {{role}}"
    doc.save(path)


def test_format_date():
    assert format_date("2026-11-24") == "Tuesday 24 November 2026"
    assert format_date("24/11/2026") == "Tuesday 24 November 2026"
    assert format_date("next week") == "next week"


def test_clean_row():
    row = clean_row({" First_Name ": "  aoife ", "last_name": "BYRNE", "interview_date": "2026-11-24", "city": "dublin"})
    assert row == {"first_name": "Aoife", "last_name": "Byrne",
                   "interview_date": "Tuesday 24 November 2026", "city": "dublin"}


def test_mixed_case_names_are_kept():
    assert clean_row({"last_name": "O'Neill"})["last_name"] == "O'Neill"


def test_missing_fields():
    assert missing_fields(["a", "b"], {"a": "x", "b": ""}) == ["b"]


def test_safe_filename():
    assert safe_filename("002_O'Neill") == "002_O_Neill"


def test_fill_paragraph_across_runs():
    doc = Document()
    p = doc.add_paragraph()
    p.add_run("Hello {{fir")
    p.add_run("st_name}}!")
    fill_paragraph(p, {"first_name": "Aoife"})
    assert p.text == "Hello Aoife!"


def test_template_fields(tmp_path):
    t = tmp_path / "t.docx"
    make_template(t)
    assert template_fields(t) == ["first_name", "interview_date", "reference", "role"]


def test_merge_creates_and_skips(tmp_path):
    t = tmp_path / "t.docx"
    make_template(t)
    rows = [
        {"first_name": "aoife", "last_name": "Byrne", "interview_date": "2026-11-24", "reference": "R1", "role": "Admin"},
        {"first_name": "Liam", "last_name": "Doyle", "interview_date": "", "reference": "R2", "role": "Admin"},
    ]
    log = merge(t, rows, tmp_path / "out")
    assert [e["status"] for e in log] == ["Created", "Skipped"]
    assert "interview_date" in log[1]["detail"]
    letter = Document(tmp_path / "out" / log[0]["file"])
    text = "\n".join(p.text for p in letter.paragraphs)
    assert "Dear Aoife," in text
    assert "Tuesday 24 November 2026" in text
    assert "{{" not in text
    assert letter.tables[0].cell(0, 0).text == "Ref: R1"
    assert letter.sections[0].footer.paragraphs[0].text == "Role: Admin"


def test_duplicate_surnames_get_unique_files(tmp_path):
    t = tmp_path / "t.docx"
    make_template(t)
    row = {"first_name": "A", "last_name": "Byrne", "interview_date": "2026-11-24", "reference": "R", "role": "X"}
    log = merge(t, [row, dict(row)], tmp_path / "out")
    assert len({e["file"] for e in log}) == 2
