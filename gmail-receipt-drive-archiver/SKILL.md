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
2. From the skill directory, ensure Node dependencies are installed:

```powershell
npm install
```

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
- If that folder already exists, it clears the folder before uploading the new PDFs for that run.
- Matching is based on Gmail search plus subject/snippet checks only. It does not extract amounts or try to semantically summarize the email body.
- The PDF is produced from rendered email HTML so inline images and image-based totals can survive into the PDF.
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
