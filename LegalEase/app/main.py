from fastapi import FastAPI, Request, Form
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from pathlib import Path
from datetime import datetime
from docx import Document
from docx.shared import Inches, Pt
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
import re
import uuid
import textwrap

BASE = Path(__file__).resolve().parent.parent
GENERATED = BASE / "generated"
GENERATED.mkdir(exist_ok=True)

app = FastAPI(title="LegalEase", version="1.0.0")
app.mount("/static", StaticFiles(directory=str(BASE / "app" / "static")), name="static")
templates = Jinja2Templates(directory=str(BASE / "app" / "templates"))

TEMPLATES = {
    "employment": {
        "name": "Employment Contract",
        "description": "A professional agreement between an employer and employee.",
        "fields": [
            ("employer", "Employer / Company", "Acme Technologies Pvt. Ltd."),
            ("employee", "Employee Name", "Alex Johnson"),
            ("job_title", "Job Title", "Software Engineer"),
            ("effective_date", "Effective Date", "2026-10-01"),
            ("salary", "Annual Salary", "₹8,00,000"),
            ("work_location", "Work Location", "Chennai, Tamil Nadu"),
            ("notice_period", "Notice Period", "30 days"),
            ("probation", "Probation Period", "3 months"),
        ],
    },
    "lease": {
        "name": "Lease Agreement",
        "description": "A structured residential or commercial lease draft.",
        "fields": [
            ("landlord", "Landlord", "Priya Kumar"),
            ("tenant", "Tenant", "Arun Kumar"),
            ("property", "Property Address", "12 Anna Nagar, Chennai, Tamil Nadu"),
            ("effective_date", "Lease Start Date", "2026-10-01"),
            ("end_date", "Lease End Date", "2027-09-30"),
            ("rent", "Monthly Rent", "₹25,000"),
            ("deposit", "Security Deposit", "₹75,000"),
            ("notice_period", "Notice Period", "30 days"),
        ],
    },
    "nda": {
        "name": "Non-Disclosure Agreement",
        "description": "A mutual confidentiality agreement for sensitive information.",
        "fields": [
            ("disclosing_party", "Disclosing Party", "Acme Technologies Pvt. Ltd."),
            ("receiving_party", "Receiving Party", "Beta Innovations Pvt. Ltd."),
            ("effective_date", "Effective Date", "2026-10-01"),
            ("purpose", "Purpose", "Evaluation of a software partnership"),
            ("duration", "Confidentiality Period", "2 years"),
            ("jurisdiction", "Governing Jurisdiction", "Tamil Nadu, India"),
        ],
    },
    "custom": {
        "name": "Custom Legal Document",
        "description": "Start with a flexible professional structure.",
        "fields": [
            ("party_a", "First Party", "Party A"),
            ("party_b", "Second Party", "Party B"),
            ("effective_date", "Effective Date", "2026-10-01"),
            ("purpose", "Purpose / Subject", "Business arrangement"),
            ("key_terms", "Key Terms", "Payment, confidentiality, termination"),
        ],
    },
}

class GenerateRequest(BaseModel):
    doc_type: str
    values: dict
    brand_name: str = "LegalEase"
    logo_url: str = ""
    font: str = "Helvetica"
    title: str = ""
    extra_terms: str = ""

def clean(s):
    return re.sub(r"\s+", " ", str(s or "")).strip()

def build_document(req: GenerateRequest):
    t = TEMPLATES.get(req.doc_type, TEMPLATES["custom"])
    v = {k: clean(val) for k, val in req.values.items()}
    title = clean(req.title) or t["name"]
    brand = clean(req.brand_name) or "LegalEase"
    date = datetime.now().strftime("%d %B %Y")

    sections = []
    if req.doc_type == "employment":
        sections = [
            ("1. Appointment and Role", f"{v.get('employer','The Employer')} appoints {v.get('employee','the Employee')} as {v.get('job_title','the Employee’s role')}, effective {v.get('effective_date',date)}. The Employee will perform duties reasonably associated with the role and other lawful duties assigned by the Employer."),
            ("2. Compensation", f"The Employee will receive an annual salary of {v.get('salary','the agreed salary')}, subject to applicable deductions and the Employer's normal payroll procedures."),
            ("3. Work Location", f"The primary work location will be {v.get('work_location','the agreed location')}, subject to reasonable business requirements and applicable law."),
            ("4. Probation", f"The initial probation period will be {v.get('probation','the agreed probation period')}. Any continuation or confirmation will be handled according to the Employer's policies and applicable law."),
            ("5. Confidentiality", "The Employee must protect confidential business, technical, financial, customer and operational information obtained through employment and use it only for legitimate work purposes."),
            ("6. Termination and Notice", f"Either party may terminate employment subject to applicable law and the agreed notice period of {v.get('notice_period','the applicable notice period')}."),
            ("7. General", "This document records the principal terms supplied by the parties. Mandatory employment protections and local law prevail where applicable."),
        ]
        parties = f"Employer: {v.get('employer')} | Employee: {v.get('employee')}"
        terms = [("Effective Date", v.get("effective_date")), ("Salary", v.get("salary")), ("Notice", v.get("notice_period")), ("Probation", v.get("probation"))]
    elif req.doc_type == "lease":
        sections = [
            ("1. Premises", f"The Landlord, {v.get('landlord')}, leases to the Tenant, {v.get('tenant')}, the property located at {v.get('property')}."),
            ("2. Term", f"The lease begins on {v.get('effective_date')} and ends on {v.get('end_date')}, unless ended earlier in accordance with this agreement or applicable law."),
            ("3. Rent", f"The Tenant will pay monthly rent of {v.get('rent')}. Payment is due according to the parties' agreed payment schedule."),
            ("4. Security Deposit", f"The Tenant will provide a security deposit of {v.get('deposit')}, subject to lawful deductions and return requirements."),
            ("5. Use and Care", "The premises will be used only for lawful purposes. The Tenant will keep the premises reasonably clean and promptly report material damage or maintenance issues."),
            ("6. Termination", f"Termination and notice will be handled in accordance with the agreed notice period of {v.get('notice_period')} and applicable law."),
            ("7. General", "Mandatory local tenancy protections and applicable law prevail over conflicting provisions in this draft."),
        ]
        parties = f"Landlord: {v.get('landlord')} | Tenant: {v.get('tenant')}"
        terms = [("Start Date", v.get("effective_date")), ("End Date", v.get("end_date")), ("Monthly Rent", v.get("rent")), ("Deposit", v.get("deposit"))]
    elif req.doc_type == "nda":
        sections = [
            ("1. Purpose", f"{v.get('disclosing_party')} and {v.get('receiving_party')} enter this agreement for {v.get('purpose')}."),
            ("2. Confidential Information", "Confidential Information means non-public business, technical, commercial, financial or other information disclosed in connection with the stated purpose, whether marked confidential or reasonably understood to be confidential."),
            ("3. Obligations", "The Receiving Party will use Confidential Information only for the stated purpose, restrict access to people who need it, and take reasonable measures to prevent unauthorized disclosure."),
            ("4. Exclusions", "Information that is public without breach, already lawfully known, independently developed without use of Confidential Information, or lawfully received from a third party without confidentiality restrictions is generally excluded."),
            ("5. Duration", f"Confidentiality obligations will continue for {v.get('duration')} unless a longer period is required by applicable law or agreed in writing."),
            ("6. Governing Law", f"The parties intend the agreement to be interpreted under the laws applicable in {v.get('jurisdiction')}, subject to mandatory legal requirements."),
            ("7. General", "This draft should be reviewed for transaction-specific definitions, remedies, data protection obligations and local enforceability."),
        ]
        parties = f"Disclosing Party: {v.get('disclosing_party')} | Receiving Party: {v.get('receiving_party')}"
        terms = [("Effective Date", v.get("effective_date")), ("Purpose", v.get("purpose")), ("Duration", v.get("duration")), ("Jurisdiction", v.get("jurisdiction"))]
    else:
        sections = [
            ("1. Parties", f"{v.get('party_a')} and {v.get('party_b')} agree to the terms described in this document."),
            ("2. Effective Date", f"This document is effective from {v.get('effective_date')}."),
            ("3. Purpose", f"The purpose of this document is {v.get('purpose')}."),
            ("4. Key Terms", f"The principal terms supplied by the user are: {v.get('key_terms')}."),
            ("5. Confidentiality", "Each party should protect non-public information received in connection with the arrangement and use it only for the agreed purpose."),
            ("6. Termination", "The parties should specify termination events, notice requirements, outstanding obligations and any post-termination provisions."),
            ("7. General", "This is a generated draft and should be reviewed and adapted to the transaction and applicable law."),
        ]
        parties = f"First Party: {v.get('party_a')} | Second Party: {v.get('party_b')}"
        terms = [("Effective Date", v.get("effective_date")), ("Purpose", v.get("purpose")), ("Key Terms", v.get("key_terms"))]

    if clean(req.extra_terms):
        sections.append(("8. Additional Terms", clean(req.extra_terms)))
        terms.append(("Additional Terms", clean(req.extra_terms)))

    text_parts = [title, brand, date, parties, ""]
    for heading, body in sections:
        text_parts += [heading, body, ""]
    text_parts += ["Acknowledgement", "The parties should review this draft carefully, complete any missing provisions, and obtain qualified legal advice where appropriate."]
    return {"title": title, "brand": brand, "date": date, "parties": parties, "sections": sections, "terms": terms, "font": req.font}

@app.get("/", response_class=HTMLResponse)
async def home(request: Request):
    return templates.TemplateResponse("index.html", {"request": request, "templates": TEMPLATES})

@app.get("/api/templates")
async def get_templates():
    return TEMPLATES

@app.post("/api/generate")
async def generate(req: GenerateRequest):
    data = build_document(req)
    token = uuid.uuid4().hex
    txt_path = GENERATED / f"{token}.txt"
    txt_path.write_text("\n\n".join([data["title"], data["parties"]] + [f"{h}\n{b}" for h,b in data["sections"]]), encoding="utf-8")
    return {"id": token, "document": data, "download": {"txt": f"/api/download/{token}/txt", "docx": f"/api/download/{token}/docx", "pdf": f"/api/download/{token}/pdf"}}

def get_data(token):
    p = GENERATED / f"{token}.txt"
    if not p.exists():
        return None
    return p.read_text(encoding="utf-8")

@app.get("/api/download/{token}/txt")
async def download_txt(token: str):
    p = GENERATED / f"{token}.txt"
    if not p.exists(): return JSONResponse({"error":"Document not found"}, status_code=404)
    return FileResponse(p, filename="LegalEase_Document.txt", media_type="text/plain")

@app.get("/api/download/{token}/docx")
async def download_docx(token: str):
    p = GENERATED / f"{token}.txt"
    if not p.exists(): return JSONResponse({"error":"Document not found"}, status_code=404)
    raw = p.read_text(encoding="utf-8")
    lines = raw.splitlines()
    doc = Document()
    doc.add_heading(lines[0] if lines else "LegalEase Document", 0)
    doc.add_paragraph("Generated by LegalEase")
    for line in lines[1:]:
        if not line.strip(): continue
        if re.match(r"^\d+\.", line) or line in ("Acknowledgement",):
            doc.add_heading(line, level=1)
        else:
            doc.add_paragraph(line)
    out = GENERATED / f"{token}.docx"
    doc.save(out)
    return FileResponse(out, filename="LegalEase_Document.docx", media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

@app.get("/api/download/{token}/pdf")
async def download_pdf(token: str):
    p = GENERATED / f"{token}.txt"
    if not p.exists(): return JSONResponse({"error":"Document not found"}, status_code=404)
    raw = p.read_text(encoding="utf-8")
    out = GENERATED / f"{token}.pdf"
    c = canvas.Canvas(str(out), pagesize=A4)
    width, height = A4
    x, y = 20*mm, height - 20*mm
    c.setFont("Helvetica-Bold", 17)
    first = raw.splitlines()[0] if raw.splitlines() else "LegalEase Document"
    c.drawString(x, y, first[:80])
    y -= 10*mm
    c.setFont("Helvetica", 9)
    c.drawString(x, y, "Generated by LegalEase")
    y -= 10*mm
    c.setFont("Helvetica", 10)
    for para in raw.split("\n"):
        if not para.strip():
            y -= 4*mm
            continue
        bold = bool(re.match(r"^\d+\.", para)) or para.strip() == "Acknowledgement"
        c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
        for line in textwrap.wrap(para, 95) or [""]:
            if y < 18*mm:
                c.showPage()
                y = height - 20*mm
                c.setFont("Helvetica-Bold" if bold else "Helvetica", 10)
            c.drawString(x, y, line)
            y -= 5*mm
    c.save()
    return FileResponse(out, filename="LegalEase_Document.pdf", media_type="application/pdf")
