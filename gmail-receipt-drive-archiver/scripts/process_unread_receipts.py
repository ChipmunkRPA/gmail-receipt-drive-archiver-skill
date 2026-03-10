import base64
import html
import json
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass
from datetime import datetime
from email import policy
from email.message import Message
from email.parser import BytesParser
from pathlib import Path
from time import sleep


SKILL_DIR = Path(__file__).resolve().parent.parent
GWS_EXE = os.environ.get("GWS_EXE") or shutil.which("gws") or r"C:\Codex\tools\gws\gws.exe"
GWS_CONFIG_DIR = os.environ.get("GOOGLE_WORKSPACE_CLI_CONFIG_DIR") or r"C:\Codex\.config\gws"
OUTPUT_DIR = Path.cwd() / "output" / "pdf"
HTML_DIR = Path.cwd() / "tmp" / "email_html"
NODE_MODULES_DIR = SKILL_DIR / "node_modules"
PLAYWRIGHT_CORE_DIR = NODE_MODULES_DIR / "playwright-core"
SEARCH_QUERY = (
    'is:unread ((receipt OR invoice OR purchase OR "order confirmation" OR "order number" OR '
    '"payment receipt" OR "payment verification" OR "business receipt" OR "purchase is complete" OR '
    '"status paid" OR "food delivery" OR takeout OR restaurant OR fantuan OR "uber eats" OR doordash OR grubhub OR seamless OR '
    '"filing receipt" OR uspto OR "trademark application" OR "serial number" OR "amount paid") OR has:attachment)'
)

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
    r"\bfood delivery\b",
    r"\btakeout\b",
    r"\brestaurant\b",
    r"\bfantuan\b",
    r"\buber eats\b",
    r"\bdoordash\b",
    r"\bgrubhub\b",
    r"\bseamless\b",
)

COMMERCE_HINT_PATTERNS = (
    r"\border\b",
    r"\bdelivery\b",
    r"\bfood\b",
    r"\brestaurant\b",
    r"\bmeal\b",
    r"\btotal\b",
    r"\bpaid\b",
    r"\bcharge\b",
    r"\buber eats\b",
    r"\bdoordash\b",
    r"\bgrubhub\b",
    r"\bseamless\b",
)

FORWARD_PREFIX_PATTERNS = (
    r"^\s*fwd:",
    r"^\s*fw:",
)

FORWARDED_RECEIPT_HINT_PATTERNS = (
    r"\border\b",
    r"\breceipt\b",
    r"\buber eats\b",
    r"\bdoordash\b",
    r"\bgrubhub\b",
    r"\bfantuan\b",
    r"\bfrom:\s*uber receipts\b",
)

GOV_RECEIPT_CORE_PATTERNS = (
    r"\buspto\b",
    r"\btrademark\b",
    r"\bservice mark\b",
    r"\bapplication\b",
    r"\bserial number\b",
    r"\bdocket number\b",
    r"\bfiling date\b",
)

GOV_RECEIPT_PROOF_PATTERNS = (
    r"\bfiling receipt\b",
    r"\bamount paid\b",
    r"\bpaid\b",
    r"\bfee\b",
    r"\bserial number\b",
    r"\bdocket number\b",
)

ATTACHMENT_NAME_PATTERNS = (
    r"\breceipt\b",
    r"\binvoice\b",
    r"\border\b",
    r"\bbill\b",
    r"\bstatement\b",
    r"\bpayment\b",
    r"\bscan\b",
    r"\bimg\b",
    r"\bphoto\b",
)

ATTACHMENT_EXTENSIONS = (".png", ".jpg", ".jpeg", ".webp", ".heic", ".pdf")


@dataclass
class Candidate:
    message_id: str
    subject: str
    snippet: str
    sender: str
    attachments: list[dict[str, str]]


def folder_name_for_today() -> str:
    now = datetime.now()
    return f"{now.month}.{now.day}.{now.year} processed invoice"


def resolve_gws_exe() -> str:
    candidates = [
        os.environ.get("GWS_EXE"),
        shutil.which("gws"),
        shutil.which("gws.exe"),
        r"C:\Codex\tools\gws\gws.exe",
        r"C:\Codex\tools\gws\gws.EXE",
    ]
    for candidate in candidates:
        if candidate and Path(candidate).exists():
            return str(Path(candidate))
    raise FileNotFoundError(
        "Could not find gws executable. Set GWS_EXE or add gws to PATH."
    )


def run_command(
    command: list[str],
    env: dict[str, str] | None = None,
    timeout: int | None = None,
    cwd: Path | None = None,
) -> subprocess.CompletedProcess:
    return subprocess.run(
        command,
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        env=env,
        cwd=str(cwd) if cwd else None,
        check=False,
        timeout=timeout,
    )


def ensure_playwright_core() -> None:
    if PLAYWRIGHT_CORE_DIR.exists():
        return

    npm_exe = shutil.which("npm.cmd") or shutil.which("npm")
    if not npm_exe:
        raise FileNotFoundError("npm was not found. Install Node.js/npm or add npm to PATH.")

    install = run_command(
        [npm_exe, "install", "--no-audit", "--no-fund"],
        env=os.environ.copy(),
        timeout=300,
        cwd=SKILL_DIR,
    )
    if install.returncode != 0:
        raise RuntimeError(
            "Failed to install Node dependencies.\n"
            f"Command: {' '.join([npm_exe, 'install', '--no-audit', '--no-fund'])}\n"
            f"STDOUT:\n{install.stdout}\nSTDERR:\n{install.stderr}"
        )
    if not PLAYWRIGHT_CORE_DIR.exists():
        raise RuntimeError("npm install completed, but playwright-core is still missing.")


def gws(*args: str, params: dict | None = None, json_body: dict | None = None, upload: str | None = None) -> dict:
    command = [resolve_gws_exe(), *args]
    if params is not None:
        command.extend(["--params", json.dumps(params, separators=(",", ":"))])
    if json_body is not None:
        command.extend(["--json", json.dumps(json_body, separators=(",", ":"))])
    if upload is not None:
        command.extend(["--upload", upload])

    env = os.environ.copy()
    env["GOOGLE_WORKSPACE_CLI_CONFIG_DIR"] = GWS_CONFIG_DIR

    attempts = 3
    for attempt in range(1, attempts + 1):
        result = run_command(command, env=env)
        stdout = result.stdout
        stderr = result.stderr
        if result.returncode == 0:
            return json.loads(stdout) if stdout.strip() else {}

        # Retry transient transport/discovery failures.
        retryable = ("discoveryError" in stdout) or ("tcp connect error" in stdout.lower())
        if retryable and attempt < attempts:
            sleep(attempt)
            continue

        raise RuntimeError(f"gws failed: {' '.join(command)}\nSTDOUT:\n{stdout}\nSTDERR:\n{stderr}")
    raise RuntimeError("Unreachable gws retry state.")


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


def extract_header(message_json: dict, header_name: str) -> str:
    headers = message_json.get("payload", {}).get("headers", [])
    for header in headers:
        if header.get("name", "").lower() == header_name.lower():
            return normalize_text(header.get("value", ""))
    return ""


def collect_attachment_metadata(payload: dict | None) -> list[dict[str, str]]:
    attachments: list[dict[str, str]] = []
    if not payload:
        return attachments

    def walk(part: dict) -> None:
        filename = normalize_text(part.get("filename", ""))
        mime_type = normalize_text(part.get("mimeType", "")).lower()
        body = part.get("body", {}) or {}
        has_attachment_id = bool(body.get("attachmentId"))
        if filename or has_attachment_id:
            attachments.append({"filename": filename, "mimeType": mime_type})
        for child in part.get("parts", []) or []:
            walk(child)

    walk(payload)
    return attachments


def is_document_attachment(meta: dict[str, str]) -> bool:
    filename = meta.get("filename", "").lower()
    mime_type = meta.get("mimeType", "").lower()
    if mime_type.startswith("image/") or mime_type == "application/pdf":
        return True
    return any(filename.endswith(ext) for ext in ATTACHMENT_EXTENSIONS)


def attachment_name_looks_receipt(meta: dict[str, str]) -> bool:
    filename = meta.get("filename", "").lower()
    return any(re.search(pattern, filename) for pattern in ATTACHMENT_NAME_PATTERNS)


def is_receipt_candidate(candidate: Candidate) -> bool:
    haystack = f"{candidate.subject}\n{candidate.snippet}\n{candidate.sender}".lower()
    text_match = any(re.search(pattern, haystack) for pattern in RECEIPT_PATTERNS)
    if text_match:
        return True

    subject = candidate.subject.lower()
    is_forward = any(re.search(pattern, subject) for pattern in FORWARD_PREFIX_PATTERNS)
    if is_forward and any(re.search(pattern, haystack) for pattern in FORWARDED_RECEIPT_HINT_PATTERNS):
        return True

    if any(attachment_name_looks_receipt(meta) for meta in candidate.attachments):
        return True

    has_gov_core = any(re.search(pattern, haystack) for pattern in GOV_RECEIPT_CORE_PATTERNS)
    has_gov_proof = any(re.search(pattern, haystack) for pattern in GOV_RECEIPT_PROOF_PATTERNS)
    if has_gov_core and has_gov_proof:
        return True

    has_document_attachment = any(is_document_attachment(meta) for meta in candidate.attachments)
    has_commerce_hint = any(re.search(pattern, haystack) for pattern in COMMERCE_HINT_PATTERNS)
    return has_document_attachment and has_commerce_hint


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
                sender=extract_header(detail, "From"),
                attachments=collect_attachment_metadata(detail.get("payload")),
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


def wrap_email_html(subject: str, sender: str, date_value: str, body_html: str, attachments_html: str) -> str:
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
    .attachments {{
      margin-top: 28px;
      border-top: 1px solid #e5e7eb;
      padding-top: 18px;
    }}
    .attachments h2 {{
      margin: 0 0 12px;
      font-size: 16px;
    }}
    figure {{
      margin: 0 0 18px;
    }}
    figcaption {{
      margin-top: 6px;
      font-size: 12px;
      color: #4b5563;
    }}
    img {{
      max-width: 100%;
      height: auto;
      border: 1px solid #e5e7eb;
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
  {attachments_html}
</body>
</html>
"""


def plain_to_html(text: str) -> str:
    return f"<pre>{html.escape(text)}</pre>"


def extract_attachment_images_html(message: Message) -> str:
    figures: list[str] = []
    for part in message.walk():
        if part.is_multipart():
            continue
        if part.get_content_disposition() != "attachment":
            continue
        mime_type = part.get_content_type().lower()
        if not mime_type.startswith("image/"):
            continue
        payload = part.get_payload(decode=True)
        if not payload:
            continue
        filename = normalize_text(part.get_filename() or "")
        encoded = base64.b64encode(payload).decode("ascii")
        data_uri = f"data:{mime_type};base64,{encoded}"
        label = filename or mime_type
        figures.append(
            "<figure>"
            f"<img src=\"{data_uri}\" alt=\"{html.escape(label)}\" />"
            f"<figcaption>{html.escape(label)}</figcaption>"
            "</figure>"
        )
    if not figures:
        return ""
    return "<section class=\"attachments\"><h2>Attached receipt images</h2>" + "".join(figures) + "</section>"


def write_renderable_html(message_id: str, subject: str, message: Message) -> Path:
    sender = normalize_text(message.get("From", ""))
    date_value = normalize_text(message.get("Date", ""))
    html_body = first_html_part(message)
    if html_body is None:
        plain_body = first_text_part(message) or ""
        html_body = plain_to_html(plain_body)
    else:
        html_body = inject_cid_images(html_body, build_cid_map(message))
    attachments_html = extract_attachment_images_html(message)
    wrapped = wrap_email_html(subject, sender, date_value, html_body, attachments_html)
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
    ensure_playwright_core()
    command = [
        "node",
        str(renderer),
        str(html_path.resolve()),
        str(pdf_path.resolve()),
        chrome_path(),
    ]
    result = run_command(command, timeout=120)
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
    ensure_playwright_core()
    folder_name = folder_name_for_today()
    folder_id = ensure_drive_folder(folder_name)
    processed = [process_message(candidate, folder_id) for candidate in get_candidates()]
    print(json.dumps({
        "folderId": folder_id,
        "folderName": folder_name,
        "processedCount": len(processed),
        "processed": processed,
    }, indent=2))


if __name__ == "__main__":
    try:
        main()
    except Exception as error:
        print(json.dumps({"error": str(error)}))
        sys.exit(1)
