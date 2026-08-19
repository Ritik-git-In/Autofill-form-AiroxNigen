"""
The Tender Profile is the single source of truth extracted from an incoming
company PDF. Every template is filled FROM this profile -- no template ever
talks to the raw PDF directly. That's what keeps all generated documents
consistent with each other (same ref number, same dates, same project title
everywhere), and gives you exactly one place to review/correct before any
Word file is generated.

Each field is a FieldValue: a value, where it came from (page number in the
source PDF), and a confidence flag. If the extractor can't find something,
value stays None -- it must NEVER guess. None flows through to the filler,
which drops a visible, highlighted "[FILL: ...]" placeholder in the output
instead of silently leaving a blank that looks intentional.
"""
from dataclasses import dataclass, field, asdict
from typing import Optional


@dataclass
class FieldValue:
    value: Optional[str] = None
    source_page: Optional[int] = None
    confidence: str = "unset"  # "high" | "low" | "unset"

    def or_placeholder(self, label: str) -> str:
        if self.value:
            return self.value
        return f"[FILL: {label}]"

    def is_missing(self) -> bool:
        return not self.value


# Canonical field list. This is deliberately flat and small (~20 fields)
# because every one of Airox's boilerplate certificates is built from the
# same handful of underlying facts -- who the authority is, what the
# project is, what the money/dates are. Add a field here only when a
# template genuinely needs a new fact; don't grow this speculatively.
TENDER_PROFILE_FIELDS = [
    ("tender_ref_no", "Tender / Bid Enquiry Number"),
    ("tender_ref_date", "Tender Enquiry Date"),
    ("project_title", "Full project/scope title as written in the RFP (all-caps scope line)"),
    ("project_capacity", "Plant capacity, e.g. '300 kW'"),
    ("project_site", "Site / premises name, e.g. 'KPCL Bellary Premises'"),
    ("project_district", "District"),
    ("project_state", "State"),
    ("authority_name", "Full name of the procuring authority, with abbreviation"),
    ("authority_short", "Short abbreviation of the authority, e.g. 'KREDL'"),
    ("authority_address", "Full mailing address of the authority / Managing Director"),
    ("emd_amount", "Earnest Money Deposit / Bid Security amount"),
    ("bid_submission_date", "Bid / Technical Bid submission deadline date"),
    ("contract_period", "O&M or contract period, e.g. '10 years'"),
    ("financial_year", "Financial year referenced for tax/financial certificates, e.g. 'FY 24-25'"),
    ("site_taluk", "Taluk of the project site"),
    ("site_coordinates", "Site GPS coordinates (Lat/Long)"),
    ("site_contact_name", "Site coordinator name"),
    ("site_contact_phone", "Site coordinator phone number"),
    ("net_worth_inr_crores", "Airox's own net worth in INR Crores, as stated in a financial capacity / CA certificate"),
    ("financial_statements_year", "The year of the audited financial statements being referenced, e.g. '2024' for 'as on March 31, 2024'"),
    ("principal_place_of_business", "Airox's principal place of business, if stated separately from the registered office"),
]


def new_tender_profile() -> dict:
    return {key: FieldValue() for key, _ in TENDER_PROFILE_FIELDS}


def profile_to_json(profile: dict) -> dict:
    return {k: asdict(v) for k, v in profile.items()}


def profile_from_json(data: dict) -> dict:
    return {k: FieldValue(**v) for k, v in data.items()}
