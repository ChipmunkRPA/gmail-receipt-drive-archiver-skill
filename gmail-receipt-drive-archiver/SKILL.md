---
name: gmail-receipt-drive-archiver
description: Use when a user wants unread Gmail receipt, invoice, purchase, or order-confirmation emails rendered to PDF, uploaded to a Google Drive folder named with the current date plus "processed invoice", and then marked read.
---

# Gmail Receipt Drive Archiver

## When to use

Use this skill when the user wants a Gmail cleanup/archive pass that:

- finds unread receipt, invoice, purchase, or order-confirmation emails
- prints the rendered email itself to PDF rather than summarizing the content
- uploads those PDFs into a Google Drive folder named for the current date
- marks only the processed emails as read

## Workflow

1. Confirm `gws` auth works for Gmail and Drive.
2. (Optional) Preinstall Node dependencies from the skill directory:

```powershell
npm install
```

   The Python script now auto-installs missing Node deps (`playwright-core`) when needed.

3. Run the archiver script:

```powershell
python scripts/process_unread_receipts.py
```

4. Read the JSON summary from stdout and report:
   - the Drive folder name and id
   - how many emails were processed
   - whether the processed emails were marked read

## Behavior

- The script creates or reuses a Drive folder named `M.D.YYYY processed invoice` using the current local date.
- If that folder already exists, the script appends new PDFs to the existing folder contents (no clearing/deletion).
- Matching is based on Gmail search plus subject/snippet checks and attachment signals (image/PDF attachments + commerce hints, and receipt-like attachment filenames).
- Forwarded receipt/order emails (for example, `Fwd:` Uber Eats order receipts) are also treated as candidates when forward markers and commerce/receipt hints are present.
- Official filing/payment receipts (for example, USPTO trademark filing receipts with serial number and amount paid) are treated as candidates based on sender/subject/snippet clues.
- The PDF is produced from rendered email HTML so inline images and image-based totals can survive into the PDF.
- Non-inline image attachments (for example, a PNG/JPG scanned receipt) are appended into the rendered PDF under an attachments section.
- Local artifacts are written under the current working directory:
  - `output/pdf/`
  - `tmp/email_html/`

## Files

- `scripts/process_unread_receipts.py`: Gmail and Drive orchestration through `gws`
- `scripts/render_email_html_to_pdf.js`: Playwright-based HTML-to-PDF renderer using local Chrome

## Prerequisites

- `gws` authenticated for Gmail and Drive
- `python`
- `node` and `npm`
- Google Chrome installed

If Chrome is not on `PATH`, set `CHROME_PATH`.
If `gws` is not on `PATH`, set `GWS_EXE`.
If the CLI config is not in the default location, set `GOOGLE_WORKSPACE_CLI_CONFIG_DIR`.
