from fastapi import FastAPI
from fastapi.responses import StreamingResponse, JSONResponse
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from docx import Document
from io import BytesIO
from pathlib import Path
from datetime import datetime
import re

app = FastAPI(title="Entry Pass API", version="2.0.0")
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_credentials=False, allow_methods=["*"], allow_headers=["*"])
BASE_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = BASE_DIR / "entry_pass_template.docx"

class EntryPassRequest(BaseModel):
    nationality: str = ""
    count: int = Field(ge=1, le=500)
    first_name: str = ""
    last_name: str = ""
    entry_port: str = ""
    hotel: str = ""
    arrival_date: str = ""
    marketing_company: str = ""
    telegram_destination: str = ""
    document_date: str = ""

def clean(v):
    return re.sub(r"\s+", " ", (v or "").strip())

def replace_para(p, repl):
    full = "".join(r.text for r in p.runs)
    if not full:
        return
    new = full
    for k, v in repl.items():
        new = new.replace(k, v)
    if new != full:
        if p.runs:
            p.runs[0].text = new
            for r in p.runs[1:]:
                r.text = ""
        else:
            p.add_run(new)

def replace_all(doc, repl):
    for p in doc.paragraphs:
        replace_para(p, repl)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    replace_para(p, repl)
    for sec in doc.sections:
        for p in sec.header.paragraphs:
            replace_para(p, repl)
        for p in sec.footer.paragraphs:
            replace_para(p, repl)

@app.get("/")
def root():
    return {"ok": True, "service": "entry-pass-api", "version": "2.0.0"}

@app.get("/health")
def health():
    return {"ok": TEMPLATE_PATH.exists(), "template": TEMPLATE_PATH.name}

@app.post("/fill-entry-pass")
def fill_entry_pass(data: EntryPassRequest):
    if not TEMPLATE_PATH.exists():
        return JSONResponse(status_code=500, content={"detail": "Word template not found"})

    document_date = clean(data.document_date) or datetime.now().strftime("%d / %m / %Y")
    repl = {
        "{{NATIONALITY}}": clean(data.nationality),
        "{{COUNT}}": str(data.count),
        "{{FIRST_NAME}}": clean(data.first_name),
        "{{LAST_NAME}}": clean(data.last_name),
        "{{ENTRY_PORT}}": clean(data.entry_port),
        "{{HOTEL}}": clean(data.hotel),
        "{{ARRIVAL_DATE}}": clean(data.arrival_date),
        "{{MARKETING_COMPANY}}": clean(data.marketing_company),
        "{{TELEGRAM_DESTINATION}}": clean(data.telegram_destination),
        "{{DOCUMENT_DATE}}": document_date,
    }

    doc = Document(str(TEMPLATE_PATH))
    replace_all(doc, repl)
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    filename = f"entry-pass-{datetime.now().strftime('%Y%m%d-%H%M%S')}.docx"
    return StreamingResponse(output, media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document", headers={"Content-Disposition": f'attachment; filename="{filename}"'})
