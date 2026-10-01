# LegalEase

LegalEase is a full-stack legal document drafting demo built with FastAPI and a modern HTML/CSS/JavaScript frontend.

## Features
- Employment Contract, Lease Agreement, NDA and Custom Document templates
- Dynamic party/date/term inputs
- Editable document preview
- Automatic key-term table
- Custom company/brand name, logo URL and font
- Export to PDF, DOCX and TXT
- Responsive dashboard UI
- Local-only document generation; no external AI API key required
- Optional AI provider hook can be added in `app/main.py`

> Legal disclaimer: LegalEase generates document drafts for informational/productivity purposes and is not a substitute for advice from a qualified lawyer.

## Run
Python 3.10+ recommended.

```bash
python -m venv .venv
# Windows:
.venv\Scripts\activate
# macOS/Linux:
source .venv/bin/activate

pip install -r requirements.txt
uvicorn app.main:app --reload
```

Open http://127.0.0.1:8000

## Project structure
- `app/main.py` - FastAPI backend and document generation
- `app/templates/index.html` - UI
- `app/static/style.css` - styling
- `app/static/app.js` - frontend logic
- `generated/` - generated exports
