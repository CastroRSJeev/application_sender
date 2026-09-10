"""
Job Poster -> Application Bot
------------------------------
Flow:
  1. User uploads a hiring poster image.
  2. NVIDIA NIM (meta/llama-3.2-11b-vision-instruct) extracts structured
     HR/job info from the image as strict JSON.
  3. We validate the extracted email.
  4. We check the Excel log for duplicates (don't re-apply to the same
     contact/job).
  5. Unless dry_run=true, we email the resume to the HR contact.
  6. We log the attempt (sent / skipped / failed) into an Excel workbook.

Run:
  uvicorn main:app --reload --port 8000

Env vars (see .env.example):
  NVIDIA_API_KEY, SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASS,
  SENDER_NAME, RESUME_PATH, EXCEL_PATH
"""

import base64
import json
import os
import re
import smtplib
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path
from typing import Optional

import requests
from dotenv import load_dotenv
from fastapi import FastAPI, File, HTTPException, UploadFile
from fastapi.responses import FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from openpyxl import Workbook, load_workbook
from pydantic import BaseModel, EmailStr, ValidationError

load_dotenv()

NVIDIA_API_KEY = os.environ["NVIDIA_API_KEY"]
NVIDIA_MODEL = "meta/llama-3.2-11b-vision-instruct"
NVIDIA_URL = "https://integrate.api.nvidia.com/v1/chat/completions"

SMTP_HOST = os.environ["SMTP_HOST"]
SMTP_PORT = int(os.environ.get("SMTP_PORT", 587))
SMTP_USER = os.environ["SMTP_USER"]
SMTP_PASS = os.environ["SMTP_PASS"]
SENDER_NAME = os.environ.get("SENDER_NAME", SMTP_USER)

RESUME_PATH = Path(os.environ.get("RESUME_PATH", "resume.pdf"))
EXCEL_PATH = Path(os.environ.get("EXCEL_PATH", "applications.xlsx"))

EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

# ---------------------------------------------------------------------------
# System prompt for the vision model
# ---------------------------------------------------------------------------
SYSTEM_PROMPT = """You are an information-extraction engine. You are shown an
image of a job/hiring poster (a flyer, LinkedIn screenshot, WhatsApp forward,
or similar). Your ONLY job is to extract recruiting contact details and
return them as strict JSON. You do not chat, explain, or add commentary.

Extract these fields exactly, using null when a field is not present in the
image (never invent or guess a value):

{
  "company_name": string or null,
  "job_title": string or null,
  "hr_name": string or null,
  "hr_email": string or null,
  "hr_phone": string or null,
  "location": string or null,
  "application_deadline": string or null,
  "other_instructions": string or null   // e.g. "apply only via portal", "WhatsApp only"
}

Rules:
- Output ONLY the JSON object. No markdown fences, no preamble, no trailing text.
- hr_email must be a literal email address exactly as printed in the image.
  If only a phone number or a web form URL is given, leave hr_email null and
  put whatever contact method exists into "other_instructions".
- If multiple emails appear, prefer the one explicitly labeled HR/recruiter/
  hiring/contact over a generic info@/support@ address.
- If the image is not a job poster at all, return all fields as null.
- Never fabricate a company name, email, or phone number that is not
  visibly printed in the image.
"""

app = FastAPI(title="Job Poster Application Bot")
app.mount("/static", StaticFiles(directory="."), name="static")


@app.get("/")
async def index():
    return FileResponse("index.html")


# ---------------------------------------------------------------------------
# Schema for the extracted data
# ---------------------------------------------------------------------------
class ExtractedPoster(BaseModel):
    company_name: Optional[str] = None
    job_title: Optional[str] = None
    hr_name: Optional[str] = None
    hr_email: Optional[str] = None
    hr_phone: Optional[str] = None
    location: Optional[str] = None
    application_deadline: Optional[str] = None
    other_instructions: Optional[str] = None


# ---------------------------------------------------------------------------
# Step 1: call NVIDIA NIM vision model
# ---------------------------------------------------------------------------
def extract_poster_info(image_bytes: bytes, mime: str) -> ExtractedPoster:
    b64 = base64.b64encode(image_bytes).decode()
    data_uri = f"data:{mime};base64,{b64}"

    payload = {
        "model": NVIDIA_MODEL,
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "image_url", "image_url": {"url": data_uri}},
                    {
                        "type": "text",
                        "text": "Extract the fields as instructed. Return JSON only.",
                    },
                ],
            },
        ],
        "max_tokens": 512,
        "temperature": 0.0,  # deterministic extraction, not creative
    }
    headers = {
        "Authorization": f"Bearer {NVIDIA_API_KEY}",
        "Accept": "application/json",
    }

    resp = requests.post(NVIDIA_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    raw_text = resp.json()["choices"][0]["message"]["content"].strip()

    # Model occasionally wraps JSON in ```json fences despite instructions - strip them.
    raw_text = re.sub(r"^```(json)?|```$", "", raw_text.strip(), flags=re.MULTILINE).strip()

    try:
        data = json.loads(raw_text)
    except json.JSONDecodeError as e:
        raise HTTPException(
            status_code=502,
            detail=f"Model did not return valid JSON: {raw_text[:300]}",
        ) from e

    try:
        return ExtractedPoster(**data)
    except ValidationError as e:
        raise HTTPException(status_code=502, detail=f"Malformed extraction: {e}") from e


# ---------------------------------------------------------------------------
# Step 2: excel log helpers
# ---------------------------------------------------------------------------
EXCEL_HEADERS = [
    "timestamp", "company_name", "job_title", "hr_name", "hr_email",
    "hr_phone", "location", "status", "notes",
]


def _open_or_create_workbook() -> Workbook:
    if EXCEL_PATH.exists():
        return load_workbook(EXCEL_PATH)
    wb = Workbook()
    ws = wb.active
    ws.title = "applications"
    ws.append(EXCEL_HEADERS)
    return wb


def already_applied(hr_email: Optional[str], job_title: Optional[str]) -> bool:
    if not hr_email or not EXCEL_PATH.exists():
        return False
    wb = _open_or_create_workbook()
    ws = wb["applications"]
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row[4] == hr_email and row[2] == job_title and row[7] == "sent":
            return True
    return False


def log_application(info: ExtractedPoster, status: str, notes: str = "") -> None:
    wb = _open_or_create_workbook()
    ws = wb["applications"]
    ws.append([
        datetime.now().isoformat(timespec="seconds"),
        info.company_name, info.job_title, info.hr_name, info.hr_email,
        info.hr_phone, info.location, status, notes,
    ])
    wb.save(EXCEL_PATH)


# ---------------------------------------------------------------------------
# Step 3: email sending
# ---------------------------------------------------------------------------
def send_application_email(info: ExtractedPoster) -> None:
    if not RESUME_PATH.exists():
        raise HTTPException(status_code=500, detail=f"Resume not found at {RESUME_PATH}")

    msg = EmailMessage()
    msg["Subject"] = f"Application for {info.job_title or 'the advertised position'}"
    msg["From"] = f"{SENDER_NAME} <{SMTP_USER}>"
    msg["To"] = info.hr_email

    greeting = f"Dear {info.hr_name}," if info.hr_name else "Dear Hiring Manager,"
    body = f"""{greeting}

I came across the opening for {info.job_title or "the role advertised"} at
{info.company_name or "your company"} and would like to apply. Please find
my resume attached.

I'd welcome the chance to discuss how I can contribute to the team.

Best regards,
{SENDER_NAME}
"""
    msg.set_content(body)

    with open(RESUME_PATH, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="octet-stream",
            filename=RESUME_PATH.name,
        )

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.send_message(msg)


# ---------------------------------------------------------------------------
# API
# ---------------------------------------------------------------------------
@app.post("/process-poster")
async def process_poster(
    file: UploadFile = File(...),
    dry_run: bool = False,
):
    """
    Upload a hiring poster image. Extracts HR contact info, optionally
    sends your resume, and logs the attempt to Excel.

    dry_run=true -> extract & log only, do NOT send an email. Use this
    first to sanity-check what the model read off the poster.
    """
    if file.content_type not in ("image/jpeg", "image/png", "image/webp"):
        raise HTTPException(status_code=400, detail="Upload a JPEG, PNG, or WEBP image.")

    image_bytes = await file.read()
    info = extract_poster_info(image_bytes, file.content_type)

    if not info.company_name and not info.job_title and not info.hr_email:
        log_application(info, status="skipped", notes="Not recognized as a job poster")
        return JSONResponse(
            status_code=422,
            content={"extracted": info.model_dump(), "status": "skipped",
                     "reason": "Doesn't look like a job poster."},
        )

    if not info.hr_email or not EMAIL_RE.match(info.hr_email):
        log_application(info, status="skipped", notes="No valid HR email found")
        return {
            "extracted": info.model_dump(),
            "status": "skipped",
            "reason": "No valid HR email extracted. Apply manually using other_instructions.",
        }

    if already_applied(info.hr_email, info.job_title):
        return {
            "extracted": info.model_dump(),
            "status": "skipped",
            "reason": "Already applied to this contact for this role.",
        }

    if dry_run:
        return {"extracted": info.model_dump(), "status": "dry_run_only"}

    try:
        send_application_email(info)
    except Exception as e:
        log_application(info, status="failed", notes=str(e))
        raise HTTPException(status_code=500, detail=f"Email send failed: {e}") from e

    log_application(info, status="sent")
    return {"extracted": info.model_dump(), "status": "sent"}


@app.get("/applications")
def get_applications():
    if not EXCEL_PATH.exists():
        return {"rows": []}
    wb = _open_or_create_workbook()
    ws = wb["applications"]
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        rows.append(dict(zip(EXCEL_HEADERS, row)))
    return {"rows": rows}


@app.get("/applications/download")
def download_applications():
    if not EXCEL_PATH.exists():
        raise HTTPException(status_code=404, detail="No applications logged yet.")
    return FileResponse(
        path=str(EXCEL_PATH),
        filename="applications.xlsx",
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/health")
def health():
    return {"ok": True}
