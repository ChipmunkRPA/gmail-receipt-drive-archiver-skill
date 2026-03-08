# Gmail Receipt Drive Archiver Skill

This repository contains a Codex skill that finds unread Gmail receipt and purchase emails, prints the rendered email to PDF, uploads the PDFs into a Google Drive folder named with the current date plus `processed invoice`, and marks the processed emails as read.

## Repo Layout

- `gmail-receipt-drive-archiver/`: installable skill folder
- `gmail-receipt-drive-archiver/scripts/process_unread_receipts.py`: Gmail and Drive workflow
- `gmail-receipt-drive-archiver/scripts/render_email_html_to_pdf.js`: Playwright-based renderer

## Prerequisites

- `gws` authenticated for Gmail and Drive
- Python
- Node.js and npm
- Google Chrome

## Install Dependencies

```powershell
cd gmail-receipt-drive-archiver
npm install
```

## Run

```powershell
cd gmail-receipt-drive-archiver
python scripts/process_unread_receipts.py
```

The script creates or reuses a Drive folder named like `3.8.2026 processed invoice`, clears that folder for the current run, uploads the new PDFs, and marks only the processed emails as read.

## Environment Overrides

- `GWS_EXE`: explicit path to `gws`
- `GOOGLE_WORKSPACE_CLI_CONFIG_DIR`: explicit `gws` config directory
- `CHROME_PATH`: explicit Chrome executable path

## License

MIT. See [LICENSE](./LICENSE).
