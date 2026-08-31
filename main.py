from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from docx import Document
from io import BytesIO
from pathlib import Path
from datetime import datetime
import re
import os
import json
import base64
import urllib.request
import urllib.error

app = FastAPI(
    title="Entry Pass API",
    version="1.0.0",
    description="خدمة مستقلة لإنشاء استمارة سمة الدخول بصيغة Word"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=False,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "entry_pass_template.docx"
HOTEL_TEMPLATE_PATH = BASE_DIR / "hotel_booking_template.docx"
GEMMA_MODEL = "gemma-4-26b-a4b-it"
GEMMA_API_KEY = os.getenv("GEMMA_API_KEY", "").strip()



class EntryPassRequest(BaseModel):
    nationality: str = Field(default="الباكستانية")
    count: int = Field(ge=1, le=500)
    first_name: str
    last_name: str
    entry_port: str = ""
    hotel: str = ""
    arrival_date: str = ""
    marketing_company: str = ""
    telegram_destination: str = ""
    document_date: str = ""


class HotelBookingRequest(BaseModel):
    company: str = ""
    first_name: str = ""
    last_name: str = ""
    count: int = Field(ge=1, le=500)
    hotel: str = ""

class GemmaReadRequest(BaseModel):
    image_base64: str
    mime_type: str = "image/jpeg"



def _clean(value: str) -> str:
    value = (value or "").strip()
    value = re.sub(r"\s+", " ", value)
    return value


def _replace_in_paragraph(paragraph, replacements: dict[str, str]) -> None:
    # القالب الذي نولده يحفظ كل placeholder ضمن Run واحد.
    # هذا المسار أيضاً يعالج الحالات التي ينقسم بها النص بين عدة Runs.
    full_text = "".join(run.text for run in paragraph.runs)
    if not full_text:
        return

    new_text = full_text
    changed = False
    for key, value in replacements.items():
        if key in new_text:
            new_text = new_text.replace(key, value)
            changed = True

    if not changed:
        return

    if paragraph.runs:
        paragraph.runs[0].text = new_text
        for run in paragraph.runs[1:]:
            run.text = ""
    else:
        paragraph.add_run(new_text)


def _replace_everywhere(doc: Document, replacements: dict[str, str]) -> None:
    for p in doc.paragraphs:
        _replace_in_paragraph(p, replacements)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    _replace_in_paragraph(p, replacements)

    for section in doc.sections:
        for p in section.header.paragraphs:
            _replace_in_paragraph(p, replacements)
        for p in section.footer.paragraphs:
            _replace_in_paragraph(p, replacements)


@app.get("/")
def root():
    return {
        "ok": True,
        "service": "entry-pass-api",
        "message": "سيرفر سمة الدخول يعمل"
    }


@app.get("/health")
def health():
    return {
        "ok": TEMPLATE_PATH.exists() and HOTEL_TEMPLATE_PATH.exists(),
        "entry_template": TEMPLATE_PATH.name,
        "entry_template_exists": TEMPLATE_PATH.exists(),
        "hotel_template": HOTEL_TEMPLATE_PATH.name,
        "hotel_template_exists": HOTEL_TEMPLATE_PATH.exists(),
        "gemma_model": GEMMA_MODEL,
        "gemma_configured": bool(GEMMA_API_KEY),
    }


@app.post("/fill-entry-pass")
def fill_entry_pass(data: EntryPassRequest):
    if not TEMPLATE_PATH.exists():
        return {
            "ok": False,
            "error": "قالب Word غير موجود على السيرفر"
        }

    document_date = _clean(data.document_date)
    if not document_date:
        document_date = datetime.now().strftime("%d / %m / %Y")

    replacements = {
        "{{NATIONALITY}}": _clean(data.nationality),
        "{{COUNT}}": str(data.count),
        "{{FIRST_NAME}}": _clean(data.first_name),
        "{{LAST_NAME}}": _clean(data.last_name),
        "{{ENTRY_PORT}}": _clean(data.entry_port),
        "{{HOTEL}}": _clean(data.hotel),
        "{{ARRIVAL_DATE}}": _clean(data.arrival_date),
        "{{MARKETING_COMPANY}}": _clean(data.marketing_company),
        "{{TELEGRAM_DESTINATION}}": _clean(data.telegram_destination),
        "{{DOCUMENT_DATE}}": document_date,
    }

    doc = Document(str(TEMPLATE_PATH))
    _replace_everywhere(doc, replacements)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    filename = f"entry-pass-{datetime.now().strftime('%Y%m%d-%H%M%S')}.docx"

    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Entry-Pass-Count": str(data.count),
        },
    )


@app.post("/fill-hotel-booking")
def fill_hotel_booking(data: HotelBookingRequest):
    if not HOTEL_TEMPLATE_PATH.exists():
        return {
            "ok": False,
            "error": "قالب الحجز الفندقي غير موجود على السيرفر"
        }

    replacements = {
        "{{COMPANY}}": _clean(data.company),
        "{{FIRST_NAME}}": _clean(data.first_name),
        "{{LAST_NAME}}": _clean(data.last_name),
        "{{COUNT}}": str(data.count),
        "{{HOTEL}}": _clean(data.hotel),
    }

    doc = Document(str(HOTEL_TEMPLATE_PATH))
    _replace_everywhere(doc, replacements)

    output = BytesIO()
    doc.save(output)
    output.seek(0)

    filename = f"hotel-booking-{datetime.now().strftime('%Y%m%d-%H%M%S')}.docx"
    return StreamingResponse(
        output,
        media_type=(
            "application/vnd.openxmlformats-officedocument."
            "wordprocessingml.document"
        ),
        headers={
            "Content-Disposition": f'attachment; filename="{filename}"',
            "X-Hotel-Booking-Count": str(data.count),
        },
    )


GEMMA_PROMPT = r"""
You are a passport data extraction engine. Analyze ONLY the passport identity page image.
Return ONLY one valid JSON object, with no markdown and no explanation.

Required keys:
{
  "given_name_en": null,
  "father_name_en": null,
  "surname_en": null,
  "passport_number": null,
  "nationality": null,
  "residence_country": null,
  "birth_date": null,
  "issue_date": null,
  "expiry_date": null,
  "sex": null,
  "mrz_line1": null,
  "mrz_line2": null,
  "confidence": 0.0
}

Rules:
- Never guess unreadable values; use null.
- Dates must be YYYY-MM-DD.
- nationality should be the 3-letter ICAO code when available, e.g. IRQ, PAK, IND.
- sex must be M, F, or X when readable.
- Copy MRZ lines exactly as visible, using < fillers.
- Compare printed fields with MRZ before returning the result.
- Prefer checksum-consistent MRZ for passport number, nationality, birth date, expiry date and sex.
- For Pakistani passports, Father/Husband Name may be printed separately. Put that value in father_name_en.
- Preserve the holder's printed English names; do not translate names here.
- confidence must be a number from 0.0 to 1.0 reflecting overall readability and agreement.
"""


def _extract_json_object(text: str) -> dict:
    raw = (text or "").strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*", "", raw, flags=re.I)
        raw = re.sub(r"\s*```$", "", raw)
    try:
        obj = json.loads(raw)
        if isinstance(obj, dict):
            return obj
    except Exception:
        pass

    start = raw.find("{")
    end = raw.rfind("}")
    if start >= 0 and end > start:
        obj = json.loads(raw[start:end + 1])
        if isinstance(obj, dict):
            return obj
    raise ValueError("Gemma returned invalid JSON")


@app.post("/read-gemma")
def read_gemma(data: GemmaReadRequest):
    if not GEMMA_API_KEY:
        return {
            "ok": False,
            "error": "GEMMA_API_KEY is not configured on Render"
        }

    try:
        image_bytes = base64.b64decode(data.image_base64, validate=True)
    except Exception:
        return {"ok": False, "error": "Invalid base64 image"}

    if not image_bytes:
        return {"ok": False, "error": "Empty image"}

    if len(image_bytes) > 15 * 1024 * 1024:
        return {"ok": False, "error": "Image is too large"}

    mime_type = (data.mime_type or "image/jpeg").strip().lower()
    if mime_type not in {"image/jpeg", "image/jpg", "image/png", "image/webp"}:
        mime_type = "image/jpeg"

    payload = {
        "contents": [{
            "role": "user",
            "parts": [
                {"text": GEMMA_PROMPT},
                {
                    "inline_data": {
                        "mime_type": mime_type,
                        "data": base64.b64encode(image_bytes).decode("ascii")
                    }
                }
            ]
        }],
        "generationConfig": {
            "temperature": 0.0,
            "thinkingConfig": {"thinkingLevel": "minimal"}
        }
    }

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models/"
        f"{GEMMA_MODEL}:generateContent"
    )
    request = urllib.request.Request(
        url,
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Content-Type": "application/json",
            "x-goog-api-key": GEMMA_API_KEY,
        },
        method="POST",
    )

    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            raw_response = response.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        return {
            "ok": False,
            "error": f"Gemma HTTP {e.code}",
            "detail": detail[:1500],
        }
    except Exception as e:
        return {"ok": False, "error": f"Gemma connection error: {e}"}

    try:
        api_json = json.loads(raw_response)
        parts = (
            api_json.get("candidates", [{}])[0]
            .get("content", {})
            .get("parts", [])
        )
        text = "\n".join(
            str(part.get("text", ""))
            for part in parts
            if isinstance(part, dict) and part.get("text")
        ).strip()
        result = _extract_json_object(text)
    except Exception as e:
        return {
            "ok": False,
            "error": f"Could not parse Gemma response: {e}",
            "raw": raw_response[:1500],
        }

    return {
        "ok": True,
        "model": GEMMA_MODEL,
        "data": result,
    }
