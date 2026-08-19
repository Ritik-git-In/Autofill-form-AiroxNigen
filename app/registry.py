import datetime
import json
import os
import re

from .filler import scan_placeholders

HERE = os.path.dirname(os.path.dirname(__file__))
STORE = os.path.join(HERE, "templates_store")
COMPANY_PATH = os.path.join(HERE, "config", "company_profile.json")

TEMPLATE_LABELS = {
    "anti_collusion_certificate.docx": "Anti-Collusion Certificate",
    "anti_blacklisting_affidavit.docx": "Anti-Blacklisting Affidavit",
    "statement_of_legal_capacity.docx": "Statement of Legal Capacity",
    "restrictions_on_sourcing.docx": "Restrictions on Sourcing of Equipment",
    "certification_not_availing_subsidy.docx": "Certification for Not Availing Subsidy",
    "declaration_of_shareholding.docx": "Declaration of Shareholding Pattern",
    "form_of_tender.docx": "Form of Tender",
    "power_of_attorney.docx": "Power of Attorney for Signing of Bid",
    "site_details.docx": "Site Details",
    "financial_capacity.docx": "Financial Capacity",
    "tender_capacity.docx": "Tender Capacity",
    "no_deviation_certificate.docx": "No Deviation Certificate",
}

# Templates whose content is mostly/entirely about AIROX's own fixed
# financial/execution history, not the incoming client PDF. Flagged in the
# UI so you don't expect the scanner to fill these from the RFP.
COMPANY_DATA_TEMPLATES = {"financial_capacity.docx", "tender_capacity.docx"}


def load_company_profile() -> dict:
    with open(COMPANY_PATH) as f:
        return json.load(f)


_STOPWORDS = {
    "of", "the", "for", "and", "to", "a", "an", "by", "on", "in", "from",
    "is", "are", "as", "this", "that", "with", "or", "at", "be",
}


def _keywords(text: str) -> set[str]:
    words = re.findall(r"[a-zA-Z]+", text.lower())
    # Crude singular/plural fold ("restrictions" -> "restriction") so a
    # tender's plural phrasing still lines up with a singular label, or
    # vice versa.
    normed = (w[:-1] if w.endswith("s") and len(w) > 4 else w for w in words)
    return {w for w in normed if w not in _STOPWORDS and len(w) > 2}


def match_template_for_requirement(name: str) -> str | None:
    """Try to match a document name as stated in a tender's own checklist
    (e.g. 'Annex VI - Anti-Collusion Certificate') to one of our templates
    by keyword overlap against TEMPLATE_LABELS, since the tender's wording
    rarely matches our filename/label exactly. Returns the templates_store
    filename, or None if nothing matches well enough to trust.

    Deliberately conservative: a single shared word ('capacity', 'details')
    between two otherwise-different document names is not enough -- that's
    how 'Technical Capacity' would wrongly pair with our Financial Capacity
    template. We want most of the template's own significant words present,
    or a majority overlap plus at least two shared words."""
    req_kw = _keywords(name)
    if not req_kw:
        return None
    best_file, best_score, best_overlap = None, 0.0, 0
    for fname, label in TEMPLATE_LABELS.items():
        label_kw = _keywords(label)
        if not label_kw:
            continue
        overlap = req_kw & label_kw
        if not overlap:
            continue
        score = len(overlap) / len(label_kw)
        if score > best_score:
            best_score, best_file, best_overlap = score, fname, len(overlap)
    if best_file and (best_score >= 0.75 or (best_score >= 0.5 and best_overlap >= 2)):
        return best_file
    return None


def build_document_selection(required_documents: list[dict]) -> list[dict]:
    """Turn the tender's own detected document requirements into the
    review screen's document list: each requirement matched against our
    template library where possible, and left unmatched (but still listed,
    per the tender's own wording) where we don't have a template yet.
    Unmatched entries carry a page range + a stable gen_id so the caller
    can reproduce that document verbatim from the tender PDF instead."""
    out = []
    for req in required_documents:
        name = (req.get("name") or "").strip()
        if not name:
            continue
        fname = match_template_for_requirement(name)
        out.append({
            "requirement_name": name,
            "source_page": req.get("source_page"),
            "end_page": req.get("end_page"),
            "filename": fname,
            "gen_id": None,
            "label": TEMPLATE_LABELS.get(fname, name) if fname else name,
            "placeholders": scan_placeholders(os.path.join(STORE, fname)) if fname else [],
            "is_company_data": fname in COMPANY_DATA_TEMPLATES if fname else False,
        })

    # Drop duplicate requirements that refer to the same underlying document.
    # Tenders sometimes list one physical form under two different names --
    # e.g. "Contractor Quality Assurance Plan" and its own form number
    # "FORM No. 11.20(DQM) F-09" both pointing at the same page -- which
    # would otherwise generate two identical copies of that page range.
    # Only applied to unmatched (no-template) entries; two genuinely
    # different bundled documents starting on the exact same tender page
    # isn't realistic, so an exact source_page match is a safe signal.
    seen_start_pages = set()
    deduped = []
    for d in out:
        if not d["filename"] and d["source_page"]:
            if d["source_page"] in seen_start_pages:
                continue
            seen_start_pages.add(d["source_page"])
        deduped.append(d)
    out = deduped

    # Fill in a missing end_page for unmatched entries: the page just before
    # the next detected document starts, since the model isn't always sure
    # exactly where one bundled document ends and the next begins.
    unmatched = [d for d in out if not d["filename"] and d["source_page"]]
    unmatched.sort(key=lambda d: d["source_page"])
    for i, d in enumerate(unmatched):
        if d["end_page"]:
            continue
        if i + 1 < len(unmatched) and unmatched[i + 1]["source_page"]:
            d["end_page"] = max(d["source_page"], unmatched[i + 1]["source_page"] - 1)
        else:
            d["end_page"] = d["source_page"]

    for i, d in enumerate(d for d in out if not d["filename"]):
        d["gen_id"] = f"gen{i}"

    return out


def resolve_values(tender_profile: dict, company: dict | None = None, overrides: dict | None = None) -> dict:
    """Flatten Tender Profile FieldValues + Company Profile into the plain
    str->str dict the filler needs, applying '[FILL: ...]' for anything
    missing. `overrides` lets the review-screen edits win over the raw
    extraction."""
    company = company or load_company_profile()
    overrides = overrides or {}

    values = {}
    for key, fv in tender_profile.items():
        if key in overrides and overrides[key]:
            values[key] = overrides[key]
        else:
            values[key] = fv.or_placeholder(key)

    # company master-data fields available to any template
    values["company_name"] = company["company_name"]
    values["gstin"] = company["gstin"]

    if overrides.get("signing_date"):
        values["signing_date"] = overrides["signing_date"]
    else:
        today = datetime.date.today()
        day = today.day
        suffix = "th" if 11 <= day <= 13 else {1: "st", 2: "nd", 3: "rd"}.get(day % 10, "th")
        values["signing_date"] = f"{day}{suffix} day of {today.strftime('%B, %Y')}"

    values["signing_place"] = overrides.get("signing_place") or "Gurugram, Haryana"

    return values
