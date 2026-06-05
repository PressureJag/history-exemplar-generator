import os
import uuid
from pathlib import Path

from flask import (
    Flask, render_template, request, redirect,
    url_for, session, send_file, flash, Response, stream_with_context,
)

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", os.urandom(24))

UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "output"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {".pdf", ".docx"}


def _save_upload(field_name: str, dest_dir: Path):
    f = request.files.get(field_name)
    if not f or not f.filename:
        return None
    ext = Path(f.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return None
    dest = dest_dir / f"{field_name}{ext}"
    f.save(dest)
    return str(dest)


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    session_id = uuid.uuid4().hex[:10]
    session_dir = UPLOAD_DIR / session_id
    session_dir.mkdir(exist_ok=True)

    files = {}
    labels = {
        "question_paper":    "Question Paper",
        "mark_scheme":       "Mark Scheme",
        "examiner_insights": "Examiner Insights",
    }

    for field, label in labels.items():
        path = _save_upload(field, session_dir)
        if not path:
            flash(f'Please upload a valid PDF or Word document for "{label}".')
            return redirect(url_for("index"))
        files[field] = path

    session["session_id"] = session_id
    session["files"] = files
    return redirect(url_for("configure"))


@app.route("/configure")
def configure():
    if "files" not in session:
        return redirect(url_for("index"))
    return render_template("configure.html")


@app.route("/generate", methods=["POST"])
def generate():
    if "files" not in session:
        return Response("data: ERROR:Session expired. Please re-upload your files.\n\n",
                        mimetype="text/event-stream")

    question_ref = request.form.get("question_ref", "").strip()
    grades_raw   = request.form.getlist("grades")

    if not question_ref or not grades_raw:
        return Response("data: ERROR:Missing question reference or grade selection.\n\n",
                        mimetype="text/event-stream")

    grades = [1, 3, 5, 7, 9] if "all" in grades_raw else sorted(int(g) for g in grades_raw)

    files        = dict(session["files"])
    session_id   = session["session_id"]
    grade_label  = "All_Grades" if len(grades) == 5 else "_".join(f"G{g}" for g in grades)
    filename     = f"Exemplar_{grade_label}_{session_id}.docx"
    output_path  = OUTPUT_DIR / filename

    def stream():
        import traceback
        from exemplar_builder import parse_file, analyze_documents, generate_exemplars, build_word_doc

        def log(msg):
            with open("/tmp/exemplar_debug.log", "a") as f:
                import datetime
                f.write(f"[{datetime.datetime.now().strftime('%H:%M:%S')}] {msg}\n")

        log(f"=== NEW REQUEST === question_ref={question_ref!r} grades={grades}")

        # Stage 1 — parse uploaded files
        yield "data: Reading uploaded documents…\n\n"
        try:
            qp_text = parse_file(files["question_paper"])
            ms_text = parse_file(files["mark_scheme"])
            ei_text = parse_file(files["examiner_insights"])
            log(f"Files parsed OK: qp={len(qp_text)} ms={len(ms_text)} ei={len(ei_text)}")
        except Exception as exc:
            log(f"File parse error: {exc}")
            yield f"data: ERROR:Could not read files — {exc}\n\n"
            return

        # Stage 2 — deep mark scheme analysis (Opus)
        yield "data: Analysing mark scheme and examiner report…\n\n"
        try:
            log("Calling analyze_documents...")
            analysis = analyze_documents(qp_text, ms_text, ei_text, question_ref)
            log(f"Analysis OK: {list(analysis.keys())}")
        except RuntimeError as exc:
            log(f"Analysis error: {exc}\n{traceback.format_exc()}")
            if "overloaded" in str(exc).lower():
                yield "data: Analysing mark scheme and examiner report… (API busy, retrying)\n\n"
            yield f"data: ERROR:{exc}\n\n"
            return

        # Stage 3 — generate grade-accurate exemplars (Opus)
        grade_str = ", ".join(f"Grade {g}" for g in grades)
        yield f"data: Generating exemplar answers ({grade_str})…\n\n"
        try:
            exemplar_data = generate_exemplars(analysis, grades)
        except Exception as exc:
            yield f"data: ERROR:Generation failed — {exc}\n\n"
            return

        # Stage 4 — build Word document
        yield "data: Building annotated Word document…\n\n"
        try:
            build_word_doc(exemplar_data, output_path)
        except Exception as exc:
            yield f"data: ERROR:Document build failed — {exc}\n\n"
            return

        session["last_download"] = filename
        yield f"data: DONE:{filename}\n\n"

    return Response(stream_with_context(stream()), mimetype="text/event-stream")


@app.route("/result")
def result():
    filename = request.args.get("file") or session.get("last_download")
    if not filename:
        return redirect(url_for("index"))
    # Validate filename is a real output file (prevent path traversal)
    safe = (OUTPUT_DIR / filename).resolve()
    if not str(safe).startswith(str(OUTPUT_DIR.resolve())):
        return redirect(url_for("index"))
    return render_template("result.html", filename=filename)


@app.route("/download/<filename>")
def download(filename):
    path = OUTPUT_DIR / filename
    if not path.exists():
        flash("File not found.")
        return redirect(url_for("index"))
    return send_file(
        path,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


if __name__ == "__main__":
    app.run(debug=True, port=5002)
