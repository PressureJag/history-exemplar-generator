# GCSE History Edexcel Exemplar Answer Generator

A Flask web app that generates grade-accurate exemplar answers for GCSE History (Edexcel) using Claude AI. Upload your question paper, mark scheme, and examiner insights report — and get a fully annotated Word document with student-style answers at grades 1, 3, 5, 7, and 9.

## What it does

1. **Upload** your three source documents (PDF or Word): question paper, mark scheme, and examiner insights report.
2. **Configure** the question reference and which grades to generate (individual grades or all five).
3. **Generate** — the app calls Claude to:
   - Analyse the mark scheme and extract grade descriptors, boundaries, and examiner insights.
   - Write grade-accurate exemplar answers, each with an authentic student voice calibrated to the mark scheme language.
4. **Download** a formatted `.docx` with:
   - A cover page (question, AO breakdown, colour key)
   - Per-grade sections with the full clean answer, segment-by-segment examiner annotations, AO colour coding, "why not full marks" notes, and a mark justification table.

## Setup

**Requirements:** Python 3.11+, [Claude Code CLI](https://claude.ai/code) installed and authenticated (the app calls `claude -p` as a subprocess).

```bash
git clone https://github.com/PressureJag/History-Examiners-Exemplar.git
cd History-Examiners-Exemplar
pip install -r requirements.txt
python app.py
```

Open `http://localhost:5002` in your browser.

## Project structure

```
app.py                  # Flask routes and SSE streaming
exemplar_builder.py     # File parsing, Claude calls, Word doc builder
templates/              # Jinja2 HTML templates
static/                 # Logo and CSS assets
requirements.txt
```

## Dependencies

| Package | Purpose |
|---|---|
| `flask` | Web framework |
| `python-docx` | Word document generation |
| `pdfplumber` | PDF text extraction |

Claude is accessed via the Claude Code CLI (`claude -p --model opus`) — no direct API key needed if you're already authenticated.

## Notes

- Session uploads are stored in `uploads/` and outputs in `output/` (both gitignored).
- The app streams progress updates to the browser via Server-Sent Events while Claude is working.
- Generation time is typically 60–120 seconds depending on the number of grades requested.
