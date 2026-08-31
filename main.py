from fastapi import FastAPI
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from docx import Document
from io import BytesIO
from pathlib import Path
from datetime import datetime
import re

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
