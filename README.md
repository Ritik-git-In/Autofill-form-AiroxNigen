# Airox Nigen — Tender Document Autofill

Upload an incoming tender/RFP PDF → the app scans it once for the facts your
certificates need → you review/correct those facts in one screen → it
generates your standard Word documents (Times New Roman, 11pt, your
letterhead, your logo) with anything it couldn't find highlighted yellow
as `[FILL: ...]` for you to fill in manually.

## How it's organized

```
app/
  extraction.py   Step 1+2: read the PDF, ask Claude for one canonical
                   "Tender Profile" JSON (or falls back to a basic regex
                   extractor if no API key is set)
  schema.py        The ~14 canonical fields every template is built from
  filler.py        Fills {{placeholders}} into a .docx, preserving formatting,
                   highlights anything left unresolved
  registry.py       Lists templates, merges Tender Profile + your fixed
                   Company Profile into the final values dict
  main.py          The web app (Flask): upload -> review -> generate

templates_store/    Your actual template .docx files (starter set: 7,
                    built from your real documents — add more the same way)
config/
  company_profile.json   Airox's own fixed facts (address, GSTIN, signatory,
                          shareholders...) — edit this by hand when it changes,
                          it is NOT re-extracted from incoming PDFs
static/airox_logo.jpg
build_templates.py  The script that generated the starter templates_store/
                    (reference only — you won't normally run this again)
```

## Setup (first time)

1. Open this folder in VS Code (`E:\Form fill project\Autoform`).
2. Open a terminal in VS Code (`` Ctrl+` ``) and run:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   pip install -r requirements.txt
   ```
3. (Optional but recommended) Copy `.env.example` to `.env` and paste in an
   Anthropic API key, so extraction uses Claude instead of the basic regex
   fallback:
   ```
   copy .env.example .env
   ```
   Then add this line near the top of `app/main.py`, after the imports:
   ```python
   from dotenv import load_dotenv
   load_dotenv()
   ```
   Without a key, the app still runs — it just extracts fewer fields
   automatically and leaves more for you to fill in on the review screen.

## Run it

```
.venv\Scripts\activate
python -m app.main
```

— upload a PDF, review the extracted facts,
tick which documents you want, click Generate, and you'll get a `.zip` of
filled `.docx` files. Open them in Word for your final bold/highlight/edit
pass.

## Adding your other templates

Right now `templates_store/` has 7 of your documents (Anti-Collusion,
Anti-Blacklisting, Statement of Legal Capacity, Restrictions on Sourcing,
Certification for Not Availing Subsidy, Declaration of Shareholding, Form
of Tender) — built from the real ones you shared. Still to add: Power of
Attorney, Site Details, Financial Capacity, No Deviation Certificate (the
big table one), Tender Capacity.

To add one:
1. Open one of your real `.docx` files (or a copy of it) in Word.
2. Wherever the text changes tender-to-tender, replace it with a
   `{{field_name}}` tag typed directly into the sentence — reuse the field
   names in `app/schema.py` (`tender_ref_no`, `project_title`,
   `authority_name`, etc.) wherever the same fact applies. Add a new field
   to `schema.py`'s `TENDER_PROFILE_FIELDS` only if the document needs a
   fact none of the others do.
3. Save it into `templates_store/`.
4. Add a friendly label for it in `TEMPLATE_LABELS` in `app/registry.py`.

That's it — it'll show up on the review screen automatically, matched
against whatever fields it references.

## The one thing to watch

The review screen is the checkpoint. It exists because this fills legal
declarations — always glance over every field (especially anything not
highlighted "ok") before generating. The system never guesses a value it
isn't confident about; it leaves it for you instead.
