import io
import json
import os
import zipfile

from dotenv import load_dotenv

load_dotenv()

from flask import Flask, request, render_template_string, send_file, send_from_directory, redirect, url_for
from werkzeug.utils import secure_filename

from .extraction import (
    parse_pdf_pages,
    parse_pdf_page_range_detailed,
    extract_profile,
    parse_reference_document,
    compose_document_text,
)
from .schema import TENDER_PROFILE_FIELDS, profile_to_json, profile_from_json
from .registry import (
    resolve_values,
    load_company_profile,
    STORE,
    build_document_selection,
)
from .filler import fill_template
from .generator import build_verbatim_document, build_ai_drafted_document, slugify

APP_DIR = os.path.dirname(os.path.dirname(__file__))
UPLOAD_DIR = os.path.join(APP_DIR, "uploads")
OUTPUT_DIR = os.path.join(APP_DIR, "output")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

app = Flask(__name__)

BASE_CSS = """
body { font-family: -apple-system, Segoe UI, sans-serif; width: 92%; max-width: 1600px; margin: 40px auto; color: #e8e8e8; background: #000; }
h1 { font-size: 22px; } h2 { font-size: 17px; margin-top: 28px; } h3 { font-size: 14px; margin-top: 20px; color: #bbb; }
.field { margin-bottom: 14px; }
.field label { display: block; font-size: 12px; color: #aaa; margin-bottom: 3px; }
.field input[type=text] { width: 100%; padding: 7px 9px; border: 1px solid #444; border-radius: 5px; font-size: 14px; box-sizing: border-box; background: #1a1a1a; color: #e8e8e8; }
.low { background: #4d4320; color: #e8e8e8; }
.missing { background: #4a2222; color: #e8e8e8; }
.badge { font-size: 10px; padding: 1px 6px; border-radius: 8px; margin-left: 6px; }
.badge.low { background: #ffe08a; color: #1a1a1a; }
.badge.missing { background: #ffb3b3; color: #1a1a1a; }
.badge.high { background: #c8f7c5; color: #1a1a1a; }
.tmpl { border: 1px solid #333; border-radius: 6px; padding: 10px 14px; margin-bottom: 8px; }
.tmpl.generated { border-left: 3px solid #ffb300; background: #1a1608; }
button, input[type=submit] { background: #1F4E79; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-size: 14px; cursor: pointer; }
.upload-box { border: 2px dashed #555; border-radius: 10px; padding: 40px; text-align: center; }
small.hint { color: #999; }
.uploads-panel { background: #161b22; border-radius: 6px; padding: 10px 14px; margin-bottom: 18px; font-size: 13px; }
.uploads-panel a { color: #6ea8d8; text-decoration: none; margin-right: 16px; }
.uploads-panel a:hover { text-decoration: underline; }
.file-list { list-style: none; padding: 0; margin: 8px 0 0; font-size: 13px; }
.file-list li { padding: 3px 0; color: #ccc; }
"""

UPLOAD_PAGE = """
<html><head><title>Airox Autoform</title><style>{{css}}</style></head>
<body>
<h1>Airox Nigen — Tender Document Autofill</h1>
<p><small class="hint">Upload the incoming company PDF (RFP). It will be scanned once for the facts every
certificate needs and for the list of documents it requires, then you'll review everything before anything
is generated.</small></p>
<form action="/extract" method="post" enctype="multipart/form-data">
  <div class="upload-box">
    <input type="file" name="pdf" accept="application/pdf" required>
    <br><br>
    <input type="submit" value="Scan PDF">
  </div>
  <div class="field" style="margin-top:18px;">
    <label>Optional — your own reference document(s) (a previously filled certificate, net worth /
      financial capacity statement, etc.) — select more than one at once if you need to. Used only to
      fill facts the tender itself doesn't state; the tender's own wording always wins when both
      mention the same fact.</label>
    <input type="file" id="reference_doc" name="reference_doc" accept="application/pdf,.docx" multiple>
    <ul id="reference_doc_list" class="file-list"></ul>
  </div>
</form>
<script>
document.getElementById('reference_doc').addEventListener('change', function () {
  var list = document.getElementById('reference_doc_list');
  list.innerHTML = '';
  for (var i = 0; i < this.files.length; i++) {
    var li = document.createElement('li');
    li.textContent = '📎 ' + this.files[i].name;
    list.appendChild(li);
  }
});
</script>
</body></html>
"""

REVIEW_PAGE = """
<html><head><title>Review — Airox Autoform</title><style>{{css}}</style></head>
<body>
<h1>Review extracted facts</h1>

<div class="uploads-panel">
  <b>Uploaded for this run:</b><br>
  <a href="/uploads/{{tender_upload}}" target="_blank">📄 {{tender_upload}}</a>
  {% for r in reference_uploads %}<a href="/uploads/{{r}}" target="_blank">📎 {{r}}</a>{% endfor %}
  <br><small class="hint">PDFs open in a new tab; Word files download — open them in Word (browsers can't preview .docx).</small>
</div>

<p><small class="hint">Anything highlighted was low-confidence or not found — fix it before generating.
This is the only place you need to correct facts; every document below is filled from these same fields.</small></p>
<form action="/generate" method="post">
<input type="hidden" name="profile_json" value='{{profile_json}}'>
<input type="hidden" name="tender_upload" value="{{tender_upload}}">
<input type="hidden" name="gen_docs_json" value='{{gen_docs_json}}'>
{% for key, desc in fields %}
  {% set fv = profile[key] %}
  <div class="field">
    <label>{{desc}}
      {% if not fv.value %}<span class="badge missing">missing</span>
      {% elif fv.confidence == 'low' %}<span class="badge low">low confidence</span>
      {% else %}<span class="badge high">ok</span>{% endif %}
      {% if fv.source_page %}<small class="hint">(source: page {{fv.source_page}})</small>{% endif %}
    </label>
    <input type="text" name="ov_{{key}}" value="{{fv.value or ''}}"
      class="{{ 'missing' if not fv.value else ('low' if fv.confidence=='low' else '') }}">
  </div>
{% endfor %}

<h2>Select documents to generate</h2>
<p><small class="hint">Every item below was found inside <strong>this specific tender's</strong> own checklist — it changes with each tender you upload, it's never a fixed list. All are selected by default; uncheck anything you don't need.</small></p>
{% if document_selection %}
{% for d in document_selection %}
  <div class="tmpl {{ '' if d.filename else 'generated' }}">
    {% if d.filename %}
    <label><input type="checkbox" name="tmpl" value="{{d.filename}}" checked> <b>{{d.label}}</b></label>
    {% if d.requirement_name and d.requirement_name != d.label %}<br><small class="hint">Tender calls this: "{{d.requirement_name}}"{% if d.source_page %} (page {{d.source_page}}){% endif %}</small>{% endif %}
    {% if d.is_company_data %}<br><small class="hint">⚠ Mostly your own company data (not from the client PDF) — figures here are placeholders, fill/verify by hand.</small>{% endif %}
    {% else %}
    <label><input type="checkbox" name="tmpl" value="gen:{{d.gen_id}}" checked> <b>{{d.requirement_name}}</b></label>
    {% if d.mode == 'ai_compose' %}
    <br><small class="hint">🤖 No ready-made format for this exists anywhere in the tender — Airox's AI will
    draft it from scratch under your letterhead. Review carefully, and check any [FILL: ...] items, before
    signing/submission.</small>
    {% else %}
    <br><small class="hint">⚠ No template for this — will be generated by reproducing the tender's own text
    verbatim (pages {{d.source_page}}{% if d.end_page and d.end_page != d.source_page %}–{{d.end_page}}{% endif %})
    under your letterhead, highlighted for review.</small>
    {% endif %}
    {% endif %}
  </div>
{% endfor %}
{% else %}
<p><small class="hint">⚠ Nothing could be detected from this tender's own text — no document/checklist section
was found, or extraction couldn't run (check your API key). Nothing is offered here rather than guessing.</small></p>
{% endif %}

<br>
<input type="submit" value="Generate Word documents">
</form>
</body></html>
"""


@app.route("/")
def index():
    return render_template_string(UPLOAD_PAGE, css=BASE_CSS)


@app.route("/favicon.ico")
def favicon():
    return send_from_directory(os.path.join(APP_DIR, "static"), "airox_logo.jpg", mimetype="image/jpeg")


@app.route("/uploads/<path:filename>")
def view_upload(filename):
    """Read-only view of a file uploaded this session (tender PDF or
    reference doc) so you can double-check exactly what was used, without
    exposing anything outside uploads/ -- send_from_directory refuses any
    path that resolves outside UPLOAD_DIR."""
    return send_from_directory(UPLOAD_DIR, filename)


def _save_upload(file_storage, default_name: str) -> str:
    """Save an uploaded file under a sanitized, collision-free name (so two
    reference documents that happen to share a filename don't overwrite
    each other -- each stays independently viewable afterward). Returns
    the name it was saved under."""
    base = secure_filename(file_storage.filename) or default_name
    name, ext = os.path.splitext(base)
    candidate = base
    i = 1
    while os.path.exists(os.path.join(UPLOAD_DIR, candidate)):
        i += 1
        candidate = f"{name}_{i}{ext}"
    file_storage.save(os.path.join(UPLOAD_DIR, candidate))
    return candidate


@app.route("/extract", methods=["POST"])
def extract():
    f = request.files["pdf"]
    tender_upload = _save_upload(f, "tender.pdf")

    reference_uploads = []
    reference_parts = []
    for ref_file in request.files.getlist("reference_doc"):
        if not ref_file or not ref_file.filename:
            continue
        saved_name = _save_upload(ref_file, "reference")
        reference_uploads.append(saved_name)
        try:
            text = parse_reference_document(os.path.join(UPLOAD_DIR, saved_name))
        except ValueError:
            continue
        if text.strip():
            reference_parts.append(f"--- REFERENCE FILE: {saved_name} ---\n{text}")

    reference_text = "\n\n".join(reference_parts) if reference_parts else None

    pages = parse_pdf_pages(os.path.join(UPLOAD_DIR, tender_upload))
    profile, required_documents = extract_profile(pages, reference_text=reference_text)

    document_selection = build_document_selection(required_documents)
    gen_docs = [
        {
            "gen_id": d["gen_id"],
            "name": d["requirement_name"],
            "source_page": d["source_page"],
            "end_page": d["end_page"],
            "mode": d["mode"],
        }
        for d in document_selection
        if d["gen_id"]
    ]

    return render_template_string(
        REVIEW_PAGE,
        css=BASE_CSS,
        fields=TENDER_PROFILE_FIELDS,
        profile=profile,
        profile_json=json.dumps(profile_to_json(profile)).replace("'", "&#39;"),
        document_selection=document_selection,
        gen_docs_json=json.dumps(gen_docs).replace("'", "&#39;"),
        tender_upload=tender_upload,
        reference_uploads=reference_uploads,
    )


@app.route("/generate", methods=["POST"])
def generate():
    profile = profile_from_json(json.loads(request.form["profile_json"]))
    overrides = {}
    for key, _ in TENDER_PROFILE_FIELDS:
        v = request.form.get(f"ov_{key}", "").strip()
        if v:
            overrides[key] = v

    selected = request.form.getlist("tmpl")
    values = resolve_values(profile, overrides=overrides)

    gen_docs = {d["gen_id"]: d for d in json.loads(request.form.get("gen_docs_json") or "[]")}
    tender_upload = request.form.get("tender_upload")

    # Parse only the specific pages each verbatim doc needs (not the whole
    # tender -- for a 150+ page tender a full re-parse alone took ~70s and
    # was the main cause of a slow "Generate" click) and pull genuine tables
    # alongside the plain text, so build_verbatim_document can render a
    # tender page's checklist/form tables as real Word tables instead of
    # flattening them into scrambled paragraph text. A handful of pages with
    # table detection still finishes in well under a second.
    tender_path = os.path.join(UPLOAD_DIR, tender_upload) if tender_upload else None
    ranged_pages_cache = {}

    def _pages_for_range(start: int, end: int):
        key = (start, end)
        if key not in ranged_pages_cache:
            ranged_pages_cache[key] = parse_pdf_page_range_detailed(tender_path, start, end)
        return ranged_pages_cache[key]

    # Facts to hand an AI-drafted document (see below) -- only the ones
    # actually known, never the literal "[FILL: ...]" placeholder text, so
    # the drafting prompt isn't confused into treating that string as a
    # real fact.
    known_tender_facts = {
        desc: values.get(key)
        for key, desc in TENDER_PROFILE_FIELDS
        if values.get(key) and not values[key].startswith("[FILL:")
    }

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        for value in selected:
            if value.startswith("gen:"):
                doc_info = gen_docs.get(value[len("gen:"):])
                if not doc_info:
                    continue
                out_name = slugify(doc_info["name"])
                out_path = os.path.join(OUTPUT_DIR, out_name)
                company = load_company_profile()
                if doc_info.get("mode") == "ai_compose":
                    # Nothing in the tender to copy from at all -- the AI
                    # drafts the body itself; a failed API call still
                    # produces a branded, signable file rather than losing
                    # every other document in this same ZIP.
                    try:
                        body_text = compose_document_text(doc_info["name"], known_tender_facts, company)
                    except Exception as exc:
                        body_text = (
                            f"[FILL: this document's content -- AI drafting failed ({exc}), "
                            "draft this one by hand]"
                        )
                    build_ai_drafted_document(
                        doc_info["name"], body_text, out_path,
                        company, values["signing_date"], values["signing_place"],
                    )
                else:
                    start = doc_info["source_page"] or 1
                    end = doc_info["end_page"] or start
                    page_entries = _pages_for_range(start, end)
                    build_verbatim_document(
                        doc_info["name"], page_entries, out_path,
                        company, values["signing_date"], values["signing_place"],
                    )
                zf.write(out_path, arcname=out_name)
            else:
                template_path = os.path.join(STORE, value)
                out_path = os.path.join(OUTPUT_DIR, value)
                fill_template(template_path, values, out_path)
                zf.write(out_path, arcname=value)
    buf.seek(0)
    return send_file(buf, as_attachment=True, download_name="airox_tender_documents.zip",
                      mimetype="application/zip")


if __name__ == "__main__":
    app.run(debug=True, port=5000)
