import base64
import html
import json
import os
import re
import shutil
import subprocess
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.message import Message
from email.parser import BytesParser
from pathlib import Path


SKILL_DIR = Path(__file__).resolve().parent.parent
GWS_EXE = os.environ.get("GWS_EXE") or shutil.which("gws") or r"C:\Codex\tools\gws\gws.exe"
GWS_CONFIG_DIR = os.environ.get("GOOGLE_WORKSPACE_CLI_CONFIG_DIR") or r"C:\Codex\.config\gws"
OUTPUT_DIR = Path.cwd() / "output" / "pdf"
HTML_DIR = Path.cwd() / "tmp" / "email_html"
SEARCH_QUERY = 'is:unread (receipt OR invoice OR purchase OR "order confirmation" OR "order number" OR "payment receipt" OR "payment verification" OR "business receipt" OR "purchase is complete" OR "status paid")'

RECEIPT_PATTERNS = (
    r"\breceipt\b",
    r"\binvoice\b",
    r"\bpurchase\b",
    r"\border confirmation\b",
    r"\border number\b",
    r"\bpayment receipt\b",
    r"\bpayment verification\b",
    r"\bbusiness receipt\b",
    r"\bpurchase is complete\b",
    r"\bstatus paid\b",
)


@dataclass
class Candidate:
    message_id: str
    subject: str
    snippet: str


def folder_name_for_today() -> str:
    now = datetime.now()
    return f"{now.month}.{now.day}.{now.year} processed invoice"


def gws(*args: str, params: dict | None = None, json_body: dict | None = None, upload: str | None = None) -> dict:
    command = [GWS_EXE, *args]
    if params is not None:
        command.extend(["--params", json.dumps(params, separators=(",", ":"))])
    if json_body is not None:
        command.extend(["--json", json.dumps(json_body, separators=(",", ":"))])
    if upload is not None:
        command.extend(["--upload", upload])

    env = os.environ.copy()
    env["GOOGLE_WORKSPACE_CLI_CONFIG_DIR"] = GWS_CONFIG_DIR

    result = subprocess.run(command, capture_output=True, env=env, check=False)
    stdout = result.stdout.decode("utf-8", errors="replace")
    stderr = result.stderr.decode("utf-8", errors="replace")
    if result.returncode != 0:
        raise RuntimeError(f"gws failed: {' '.join(command)}\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}")
    return json.loads(stdout) if stdout.strip() else {}


def normalize_text(value: str | None) -> str:
    if not value:
        return ""
    value = value.replace("\xa0", " ")
    value = re.sub(r"\s+", " ", value)
    return value.strip()


def decode_raw_message(raw_value: str) -> Message:
    padding = "=" * ((4 - len(raw_value) % 4) % 4)
    raw_bytes = base64.urlsafe_b64decode(raw_value + padding)
    return BytesParser(policy=policy.default).parsebytes(raw_bytes)


def safe_filename(subject: str, message_id: str, suffix: str) -> str:
    safe = re.sub(r"[^A-Za-z0-9._ -]", "", subject).strip()
    safe = re.sub(r"\s+", " ", safe)
    safe = safe[:80].rstrip(" .") or f"message-{message_id}"
    return f"{safe}-{message_id}.{suffix}"


def ensure_dirs() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    HTML_DIR.mkdir(parents=True, exist_ok=True)


def clear_local_artifacts() -> None:
    for directory in (OUTPUT_DIR, HTML_DIR):
        directory.mkdir(parents=True, exist_ok=True)
        for path in directory.glob("*"):
            if path.is_file():
                path.unlink()


def ensure_drive_folder(folder_name: str) -> str:
    response = gws(
        "drive",
        "files",
        "list",
        params={
            "q": f"mimeType = 'application/vnd.google-apps.folder' and trashed = false and name = '{folder_name}'",
            "fields": "files(id,name)",
            "pageSize": 10,
        },
    )
    files = response.get("files", [])
    if files:
        return files[0]["id"]

    created = gws(
        "drive",
        "files",
        "create",
        json_body={
            "name": folder_name,
            "mimeType": "application/vnd.google-apps.folder",
        },
    )
    return created["id"]


def clear_drive_folder(folder_id: str) -> None:
    response = gws(
        "drive",
        "files",
        "list",
        params={
            "q": f"trashed = false and '{folder_id}' in parents",
            "fields": "files(id,name)",
            "pageSize": 200,
        },
    )
    for item in response.get("files", []):
        gws("drive", "files", "delete", params={"fileId": item["id"]})


def extract_header(message_json: dict, header_name: str) -> str:
    headers = message_json.get("payload", {}).get("headers", [])
    for header in headers:
        if header.get("name", "").lower() == header_name.lower():
            return normalize_text(header.get("value", ""))
    return ""


def is_receipt_candidate(candidate: Candidate) -> bool:
    haystack = f"{candidate.subject}\n{candidate.snippet}".lower()
    return any(re.search(pattern, haystack) for pattern in RECEIPT_PATTERNS)


def get_candidates() -> list[Candidate]:
    response = gws(
        "gmail",
        "users",
        "messages",
        "list",
        params={"userId": "me", "q": SEARCH_QUERY, "maxResults": 100},
    )
    candidates: list[Candidate] = []
    for item in response.get("messages", []):
        detail = gws(
            "gmail",
            "users",
            "messages",
            "get",
            params={"userId": "me", "id": item["id"], "format": "metadata"},
        )
        candidates.append(
            Candidate(
                message_id=item["id"],
                subject=extract_header(detail, "Subject"),
                snippet=normalize_text(detail.get("snippet", "")),
            )
        )
    return [candidate for candidate in candidates if is_receipt_candidate(candidate)]


def build_cid_map(message: Message) -> dict[str, str]:
    cid_map: dict[str, str] = {}
    for part in message.walk():
        if part.is_multipart():
            continue
        content_id = part.get("Content-ID")
        if not content_id:
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        cid = content_id.strip().strip("<>")
        mime_type = part.get_content_type()
        encoded = base64.b64encode(payload).decode("ascii")
        cid_map[cid] = f"data:{mime_type};base64,{encoded}"
    return cid_map


def first_html_part(message: Message) -> str | None:
    for part in message.walk():
        if part.is_multipart():
            continue
        if part.get_content_type() == "text/html":
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return None


def first_text_part(message: Message) -> str | None:
    for part in message.walk():
        if part.is_multipart():
            continue
        if part.get_content_type() == "text/plain":
            payload = part.get_payload(decode=True) or b""
            charset = part.get_content_charset() or "utf-8"
            return payload.decode(charset, errors="replace")
    return None


def inject_cid_images(html_body: str, cid_map: dict[str, str]) -> str:
    updated = html_body
    for cid, data_uri in cid_map.items():
        updated = re.sub(
            rf"cid:{re.escape(cid)}(?=[\"'> ])",
            data_uri,
            updated,
            flags=re.IGNORECASE,
        )
    return updated


def wrap_email_html(subject: str, sender: str, date_value: str, body_html: str) -> str:
    return f"""<!DOCTYPE html>
<html>
<head>
  <meta charset="utf-8">
  <title>{html.escape(subject or "Email")}</title>
  <style>
    body {{
      margin: 0;
      padding: 24px 28px;
      font-family: Arial, sans-serif;
      background: #ffffff;
      color: #111827;
    }}
    .meta {{
      margin-bottom: 24px;
      border-bottom: 1px solid #e5e7eb;
      padding-bottom: 12px;
    }}
    .meta h1 {{
      margin: 0 0 12px;
      font-size: 24px;
      line-height: 1.2;
      word-break: break-word;
    }}
    .meta p {{
      margin: 4px 0;
      font-size: 13px;
    }}
    img {{
      max-width: 100%;
      height: auto;
    }}
  </style>
</head>
<body>
  <div class="meta">
    <h1>{html.escape(subject or "(No subject)")}</h1>
    <p><strong>From:</strong> {html.escape(sender or "(Unknown sender)")}</p>
    <p><strong>Date:</strong> {html.escape(date_value or "")}</p>
  </div>
  <div class="email-body">{body_html}</div>
</body>
</html>
"""


def plain_to_html(text: str) -> str:
    return f"<pre>{html.escape(text)}</pre>"


def write_renderable_html(message_id: str, subject: str, message: Message) -> Path:
    sender = normalize_text(message.get("From", ""))
    date_value = normalize_text(message.get("Date", ""))
    html_body = first_html_part(message)
    if html_body is None:
        plain_body = first_text_part(message) or ""
        html_body = plain_to_html(plain_body)
    else:
        html_body = inject_cid_images(html_body, build_cid_map(message))
    wrapped = wrap_email_html(subject, sender, date_value, html_body)
    html_path = HTML_DIR / safe_filename(subject, message_id, "html")
    html_path.write_text(wrapped, encoding="utf-8")
    return html_path


def chrome_path() -> str:
    browser = os.environ.get("CHROME_PATH") or shutil.which("chrome") or shutil.which("chrome.exe")
    if browser:
        return browser
    fallback = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
    if Path(fallback).exists():
        return fallback
    raise FileNotFoundError("Chrome not found. Set CHROME_PATH.")


def print_html_to_pdf(html_path: Path, pdf_path: Path) -> None:
    renderer = Path(__file__).with_name("render_email_html_to_pdf.js")
    command = [
        "node",
        str(renderer),
        str(html_path.resolve()),
        str(pdf_path.resolve()),
        chrome_path(),
    ]
    result = subprocess.run(command, capture_output=True, text=True, check=False, timeout=120)
    if result.returncode != 0:
        raise RuntimeError(f"PDF render failed for {html_path}\nSTDOUT:\n{result.stdout}\nSTDERR:\n{result.stderr}")


def upload_pdf(pdf_path: Path, folder_id: str) -> str:
    uploaded = gws(
        "drive",
        "files",
        "create",
        json_body={
            "name": pdf_path.name,
            "parents": [folder_id],
            "mimeType": "application/pdf",
        },
        upload=str(pdf_path),
    )
    return uploaded["id"]


def mark_read(message_id: str) -> None:
    gws(
        "gmail",
        "users",
        "messages",
        "modify",
        params={"userId": "me", "id": message_id},
        json_body={"removeLabelIds": ["UNREAD"]},
    )


def process_message(candidate: Candidate, folder_id: str) -> dict:
    response = gws(
        "gmail",
        "users",
        "messages",
        "get",
        params={"userId": "me", "id": candidate.message_id, "format": "raw"},
    )
    parsed = decode_raw_message(response["raw"])
    subject = normalize_text(parsed.get("Subject", "")) or candidate.subject or "(No subject)"
    html_path = write_renderable_html(candidate.message_id, subject, parsed)
    pdf_path = OUTPUT_DIR / safe_filename(subject, candidate.message_id, "pdf")
    print_html_to_pdf(html_path, pdf_path)
    drive_file_id = upload_pdf(pdf_path, folder_id)
    mark_read(candidate.message_id)
    return {
        "messageId": candidate.message_id,
        "subject": subject,
        "pdfPath": str(pdf_path),
        "driveFileId": drive_file_id,
    }


def main() -> None:
    ensure_dirs()
    clear_local_artifacts()
    folder_name = folder_name_for_today()
    folder_id = ensure_drive_folder(folder_name)
    clear_drive_folder(folder_id)
    processed = [process_message(candidate, folder_id) for candidate in get_candidates()]
    print(json.dumps({
        "folderId": folder_id,
        "folderName": folder_name,
        "processedCount": len(processed),
        "processed": processed,
    }, indent=2))


if __name__ == "__main__":
    main()
