# Word Mail Merge

A Python tool that creates one personalised Word letter for each row of a spreadsheet. Office and HR administrators do this every week for interview invitations, appointment letters and reminders.

**All names, addresses and the company are made up.**

## What it does

- Reads a Word template with placeholders such as `{{first_name}}` and `{{interview_date}}`
- Fills placeholders in body text, tables, headers and footers, keeping the template's formatting
- Tidies data as it goes: extra spaces, names typed in all capitals, and dates such as `24/11/2026`, which become "Tuesday 24 November 2026"
- Skips any row with a missing value instead of sending a half-filled letter, and logs why in `merge_log.csv`
- Stops before starting if the template uses a field the spreadsheet does not have

## Run it

```
pip install -r requirements.txt
python mail_merge.py templates/interview_invite.docx sample_data/candidates.csv --out output
```

With the sample data:

```
Rows read:          7
Letters created:    5
Rows skipped:       2  (see merge_log.csv)
```

| Row | Result |
|---|---|
| 2, 3, 4, 6, 8 | Letter created |
| 5 | Skipped: missing address |
| 7 | Skipped: missing interview date |

The sample letters are in the [output](output) folder.

## Use your own letter

1. Write the letter in Word and type placeholders where the details go, such as `{{first_name}}`.
2. Save a CSV with a column for each placeholder.
3. Run the command above with your files.

Any column ending in `_date` is formatted as a full date.

## Tests

```
pip install pytest
pytest
```

The 9 tests cover date formatting, data tidying, placeholders split across Word's text runs, tables, footers, skipped rows and duplicate file names.
