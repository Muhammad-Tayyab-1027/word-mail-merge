#!/usr/bin/env python3
"""Create one personalised Word letter per row of a CSV file.

Usage:
    python mail_merge.py templates/interview_invite.docx sample_data/candidates.csv --out output

The template uses placeholders such as {{first_name}}. Each placeholder must match a CSV column.
Rows with missing values are skipped and listed in merge_log.csv, so nothing half-filled is sent.
"""
import argparse
import csv
import re
from datetime import datetime
from pathlib import Path

from docx import Document

PLACEHOLDER = re.compile(r"\{\{\s*([a-z0-9_]+)\s*\}\}")
DATE_INPUT_FORMATS = ("%Y-%m-%d", "%d/%m/%Y", "%d-%m-%Y")


def format_date(value):
    """Turn 2026-11-24 or 24/11/2026 into 'Tuesday 24 November 2026'. Other text is returned unchanged."""
    for fmt in DATE_INPUT_FORMATS:
        try:
            d = datetime.strptime(value.strip(), fmt)
            return f"{d:%A} {d.day} {d:%B %Y}"
        except ValueError:
            continue
    return value


def clean_row(row):
    """Trim spaces, tidy name capitals and format any column whose name ends in _date."""
    cleaned = {}
    for key, value in row.items():
        key = (key or "").strip().lower()
        value = " ".join((value or "").split())
        if key.endswith("name") and value.isupper() or key.endswith("name") and value.islower():
            value = value.title()
        if key.endswith("_date") and value:
            value = format_date(value)
        cleaned[key] = value
    return cleaned


def all_paragraphs(doc):
    """Every paragraph in the body, in tables, and in headers and footers."""
    yield from doc.paragraphs
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                yield from cell.paragraphs
    for section in doc.sections:
        for part in (section.header, section.footer):
            yield from part.paragraphs


def template_fields(path):
    """The placeholder names used anywhere in the template."""
    doc = Document(path)
    return sorted({m.group(1) for p in all_paragraphs(doc) for m in PLACEHOLDER.finditer(p.text)})


def fill_paragraph(paragraph, values):
    """Replace placeholders in a paragraph, even when Word has split one across several runs."""
    text = paragraph.text
    if "{{" not in text:
        return
    new_text = PLACEHOLDER.sub(lambda m: values[m.group(1)], text)
    runs = paragraph.runs
    runs[0].text = new_text  # keeps the first run's formatting
    for run in runs[1:]:
        run.text = ""


def missing_fields(fields, row):
    return [f for f in fields if not row.get(f)]


def safe_filename(text):
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_") or "letter"


def merge(template, rows, out_dir, name_field="last_name"):
    """Write one .docx per complete row. Returns a list of log entries."""
    fields = template_fields(template)
    out_dir.mkdir(parents=True, exist_ok=True)
    log, used = [], set()
    for number, raw in enumerate(rows, 2):  # row 1 of the CSV is the header
        row = clean_row(raw)
        missing = missing_fields(fields, row)
        if missing:
            log.append({"row": number, "file": "", "status": "Skipped", "detail": "Missing: " + ", ".join(missing)})
            continue
        doc = Document(template)
        for paragraph in all_paragraphs(doc):
            fill_paragraph(paragraph, row)
        base = safe_filename(f"{number:03d}_{row.get(name_field, '')}")
        name = base
        while name in used:
            name += "_x"
        used.add(name)
        doc.save(out_dir / f"{name}.docx")
        log.append({"row": number, "file": f"{name}.docx", "status": "Created", "detail": ""})
    return log


def main():
    parser = argparse.ArgumentParser(description="Create personalised Word letters from a CSV file.")
    parser.add_argument("template", type=Path, help="Word template (.docx) with {{placeholders}}")
    parser.add_argument("data", type=Path, help="CSV file, one row per letter")
    parser.add_argument("--out", type=Path, default=Path("output"), help="folder for the letters")
    args = parser.parse_args()

    with args.data.open(newline="", encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    fields = template_fields(args.template)
    columns = {c.strip().lower() for c in rows[0]} if rows else set()
    unknown = [f for f in fields if f not in columns]
    if unknown:
        raise SystemExit(f"The template uses fields the CSV does not have: {', '.join(unknown)}")

    log = merge(args.template, rows, args.out)
    with (args.out / "merge_log.csv").open("w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=["row", "file", "status", "detail"])
        writer.writeheader()
        writer.writerows(log)
    created = sum(1 for e in log if e["status"] == "Created")
    print(f"Fields in template: {', '.join(fields)}")
    print(f"Rows read:          {len(rows)}")
    print(f"Letters created:    {created}")
    print(f"Rows skipped:       {len(rows) - created}  (see merge_log.csv)")


if __name__ == "__main__":
    main()
