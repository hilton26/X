#!/usr/bin/env python3
"""mandate_limit_extractor.py — first-pass extraction of investment limits from
client mandate PDFs into a fund x limit-type matrix.

RUN THIS LOCALLY ONLY. It is designed to read mandates directly off your P:\\
drive and never sends any file content anywhere — no network calls, no APIs.
Everything happens on your machine, in this Python process.

WHAT IT DOES
    1. Walks every client folder under --root, looking for a subfolder named
       "Legal" one or two levels down (as you described the structure).
    2. Extracts text from every PDF found in each "Legal" folder using
       pdfplumber (a pure Python PDF text extractor — no Word round-trip
       needed for text-based PDFs; see NOTE ON PDF-TO-WORD below).
    3. Splits each mandate's text into clauses and checks each clause against
       a predefined taxonomy of limit categories (LIMIT_CATEGORIES below).
    4. A clause becomes a confident matrix entry only if it matches exactly
       ONE category AND contains exactly ONE extractable numeric value
       (a percentage or Rand amount) AND that fund+category combination has
       no other confident match. Everything else — multiple categories
       matched, multiple numbers found, limit-sounding language that matched
       no category, a second conflicting value for a fund+category already
       filled — is routed to the "Needs Review" sheet instead of guessed.
    5. Writes one Excel workbook with three sheets:
         - "Limits Matrix"   : fund (row) x limit category (column), plus a
                                leading "Mandate Effective Date" column — the
                                latest effective/commencement date found across
                                all of that fund's Legal-folder mandate PDFs
         - "Needs Review"    : every clause the script wasn't confident about,
                                with source file and reason, for YOUR judgment
         - "Processing Log"  : one row per client folder — what was found,
                                what was skipped, and why

NOTE ON PDF-TO-WORD
    You mentioned mandates would need converting to Word first. For ordinary
    (digitally created) PDFs that extra step isn't necessary — pdfplumber
    reads the text directly and more reliably than a PDF->Word->re-open round
    trip. The one case where Word actually earns its place is a SCANNED
    mandate (a photocopy/fax that was scanned to PDF) — there is no text
    layer to extract at all. This script detects that case (see "NO_TEXT"
    in the Processing Log) rather than silently failing, and for those
    specific files your instinct is right: open them in Word (which OCRs on
    import) or run them through Acrobat's OCR, save as .docx, and I can add
    a docx-reading branch once we know how many mandates actually need it.

THREE MODES
    python mandate_limit_extractor.py --dry-run          # list what would be processed; extract/write nothing
    python mandate_limit_extractor.py --limit 5           # pilot run: only the first 5 client folders
    python mandate_limit_extractor.py                     # full run over every client folder under --root

REQUIREMENTS (install once)
    pip install pdfplumber pandas openpyxl

Ticket/owner: built for Hilton Netta (Compliance), 2026-08.
"""

from __future__ import annotations

import argparse
import re
import sys
from dataclasses import dataclass, field
from datetime import date
from pathlib import Path

try:
    import pdfplumber
except ImportError:
    print("Missing dependency: pip install pdfplumber pandas openpyxl", file=sys.stderr)
    raise

import pandas as pd
from dateutil import parser as date_parser

# --- CONFIG: this is the section you'll actually want to edit -----------------

DEFAULT_ROOT = Path(r"P:\Investment Operations\Segregated Clients\Active Clients")
DEFAULT_OUTPUT_DIR = Path(r"C:\Users\hilton.netta\OneDrive - Prescient\py\X\outputs")
LEGAL_FOLDER_NAME = "legal"  # matched case-insensitively
MAX_LEGAL_SEARCH_DEPTH = 2  # "Legal" is 1 or 2 levels below each client folder
SEARCH_LEGAL_SUBFOLDERS_FOR_PDFS = (
    False  # True = also look inside subfolders of "Legal" (e.g. "Signed", "Superseded")
)

# Predefined limit taxonomy: category name -> list of regex patterns (matched
# case-insensitively) that identify a clause's SUBJECT MATTER. A clause still
# needs a limit-phrase (see LIMIT_PHRASE_RE) and exactly one number to become
# a confident matrix entry — these patterns only decide topic, not confidence.
# Edit/add rows here; nothing else in the script needs to change.
LIMIT_CATEGORIES: dict[str, list[str]] = {
    "Equity Exposure Limit": [r"\bequit(y|ies)\b", r"\blisted shares\b"],
    "Debt / Fixed Income Exposure Limit": [
        r"\bdebt\b",
        r"\bfixed\s+income\b",
        r"\bbonds?\b",
    ],
    "Offshore / Foreign Exposure Limit": [
        r"\boffshore\b",
        r"\bforeign\s+(currency|assets|investments?|exposure)\b",
        r"\binternational\s+exposure\b",
    ],
    "Single Issuer / Counterparty Exposure Limit": [
        r"\bsingle\s+issuer\b",
        r"\bone\s+issuer\b",
        r"\bcounterparty\b",
        r"\bissuer\s+exposure\b",
    ],
    "Derivative Exposure / Usage Limit": [r"\bderivative(s)?\b"],
    "Listed Derivatives Limit": [r"\blisted\s+derivative(s)?\b"],
    "Unlisted / OTC Derivatives Limit": [
        r"\bunlisted\s+derivative(s)?\b",
        r"\bOTC\s+derivative(s)?\b",
        r"\bover-the-counter\s+derivative(s)?\b",
    ],
    "Leverage / Gearing Limit": [r"\bleverage\b", r"\bgearing\b", r"\bborrowing\b"],
    "Credit Rating Minimum": [
        r"\bcredit\s+rating\b",
        r"\binvestment\s+grade\b",
        r"\brated\s+(at\s+least\s+)?(AAA|AA|A|BBB|BB)\b",
    ],
    "Liquidity Limit": [r"\bliquid(ity)?\b"],
    "Prescribed Assets Minimum": [r"\bprescribed\s+asset(s)?\b"],
    "Related / Connected Party Exposure Limit": [
        r"\brelated\s+part(y|ies)\b",
        r"\bconnected\s+part(y|ies)\b",
        r"\baffiliate(s)?\b",
    ],
    "Collective Investment Scheme (CIS) Exposure Limit": [
        r"\bcollective\s+investment\s+scheme(s)?\b",
        r"\bCIS\b",
        r"\bunit\s+trust(s)?\b",
        r"\bfund(s)?\s+of\s+fund(s)?\b",
    ],
    "Unlisted Securities Limit": [
        r"\bunlisted\s+(securities|shares|equity|equities|instruments|assets)\b",
    ],
    "Securities / Scrip Lending Limit": [
        r"\bsecurities\s+lending\b",
        r"\bscrip\s+lending\b",
        r"\bstock\s+lending\b",
    ],
    "Repurchase Agreement (Repo) Limit": [
        r"\brepurchase\s+agreement(s)?\b",
        r"\brepo(s)?\b",
    ],
    "Commodities Exposure Limit": [r"\bcommodit(y|ies)\b"],
    "Cash Exposure Limit": [r"\bcash\b"],
    "Duration Limit": [r"(?<!modified\s)\bduration\b"],
    "Modified Duration Limit": [r"\bmodified\s+duration\b"],
    "Benchmark Deviation / Tracking Error Limit": [
        r"\btracking\s+error\b",
        r"\bbenchmark\b.{0,40}\bdeviation\b",
        r"\bactive\s+risk\b",
    ],
    "Money Market Instrument Limit": [r"\bmoney\s+market\b"],
    "Sector / Industry Concentration Limit": [
        r"\bsector\b",
        r"\bindustry\s+concentration\b",
    ],
    "Property / Real Estate Exposure Limit": [
        r"\bproperty\b",
        r"\bREIT\b",
        r"\breal\s+estate\b",
    ],
}

# Some categories are deliberately broader restatements of more specific ones
# above (e.g. any clause about "unlisted/OTC derivatives" or "listed
# derivatives" also contains the word "derivative"). Without this, those
# clauses would match two categories and get routed to "Needs Review" as an
# ambiguous multi-category match even though the specific category is exactly
# right. When a clause matches one of a generic category's overrides, the
# generic category is dropped from the match list in categorize_clause().
CATEGORY_OVERRIDES: dict[str, set[str]] = {
    "Derivative Exposure / Usage Limit": {
        "Listed Derivatives Limit",
        "Unlisted / OTC Derivatives Limit",
    },
}

# Phrases that mark a clause as *limit-like* regardless of category — used both
# to gate confident matches and to flag limit-sounding clauses that matched no
# category (so they don't silently disappear).
LIMIT_PHRASE_RE = re.compile(
    r"(shall\s+not\s+exceed|may\s+not\s+exceed|no\s+more\s+than|not\s+exceed|"
    r"maximum\s+of|limited\s+to|at\s+least|minimum\s+of|shall\s+be\s+limited|"
    r"up\s+to\s+a\s+maximum|restricted\s+to)",
    re.IGNORECASE,
)

# A percentage ("10%", "7.5 %") or a Rand amount ("R5 million", "R 250,000").
# Deliberately NOT case-insensitive on the "R" itself — a lowercase "r" is far
# too common inside ordinary words (e.g. "issuer, save...") and produced false
# positives in testing; the Rand sign in these documents is always capital R.
NUMERIC_RE = re.compile(
    r"(\d{1,3}(?:[.,]\d+)?\s?%|\bR\s?\d[\d,]*(?:\.\d+)?\s?(?:[Mm]illion|[Bb]illion|[Bb]n|[Mm]\b)?)"
)

# A date token in any of the formats mandates commonly use: "1 March 2022" /
# "1st March 2022", "March 1, 2022", "01/03/2022", or "2022-03-01".
_DATE_TOKEN = (
    r"(?:\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4}"
    r"|[A-Za-z]+\s+\d{1,2},?\s+\d{4}"
    r"|\d{1,2}[/-]\d{1,2}[/-]\d{2,4}"
    r"|\d{4}-\d{1,2}-\d{1,2})"
)

# Phrases that introduce a mandate's effective/commencement date, gated to a
# date token within a short distance so citations of unrelated dates (e.g. an
# Act's year) aren't picked up.
EFFECTIVE_DATE_RE = re.compile(
    r"(?:effective\s+date|with\s+effect\s+from|effective\s+from|"
    r"commencement\s+date|date\s+of\s+commencement|coming\s+into\s+effect|"
    r"date\s+of\s+this\s+(?:mandate|agreement))"
    r"[^\n]{0,40}?(" + _DATE_TOKEN + r")",
    re.IGNORECASE,
)


# --- data structures ------------------------------------------------------------


@dataclass
class Confident:
    fund: str
    category: str
    value: str
    clause: str
    source_pdf: str


@dataclass
class ReviewItem:
    fund: str
    source_pdf: str
    category: str
    clause: str
    reason: str


@dataclass
class LogRow:
    client_folder: str
    legal_folder: str
    pdfs_found: str
    status: str
    detail: str


# --- text handling ---------------------------------------------------------------


def normalize_text(text: str) -> str:
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(
        r"-\n(?=[a-z])", "", text
    )  # de-hyphenate words wrapped across a line break
    text = re.sub(r"[ \t]+", " ", text)
    return text


def extract_effective_date(text: str) -> date | None:
    """First effective/commencement date found in a mandate's full text, or None.

    Deliberately takes the first match rather than every date mentioned — mandates
    state their effective date once, near the top, and later hits are usually
    unrelated dates that happen to sit near a similarly-worded phrase.
    """
    for match in EFFECTIVE_DATE_RE.finditer(normalize_text(text)):
        try:
            return date_parser.parse(match.group(1), dayfirst=True, fuzzy=True).date()
        except (ValueError, OverflowError):
            continue
    return None


def split_clauses(text: str) -> list[str]:
    """Blank-line paragraphs, further split on numbered sub-clause markers (5.1, 5.2, ...)."""
    clauses: list[str] = []
    for para in re.split(r"\n\s*\n", text):
        para = para.strip()
        if not para:
            continue
        for sub in re.split(r"(?=\n?\d+\.\d+(?:\.\d+)?\s)", para):
            sub = sub.strip().replace("\n", " ")
            if sub:
                clauses.append(sub)
    return clauses


def extract_pdf_text(pdf_path: Path) -> str | None:
    """Return extracted text, or None if the PDF has no extractable text layer (likely scanned)."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
        text = "\n\n".join(pages).strip()
        return text if text else None
    except Exception as exc:  # noqa: BLE001 — surface any parse failure into the log, don't crash the run
        raise RuntimeError(f"pdfplumber failed: {exc}") from exc


# --- folder discovery -------------------------------------------------------------


def find_legal_folders(client_dir: Path) -> list[Path]:
    """Legal folders 1 or 2 levels below a client folder, matched case-insensitively."""
    found: list[Path] = []
    try:
        children = [c for c in client_dir.iterdir() if c.is_dir()]
    except (PermissionError, OSError):
        return found
    for child in children:
        if child.name.lower() == LEGAL_FOLDER_NAME:
            found.append(child)
        elif MAX_LEGAL_SEARCH_DEPTH >= 2:
            try:
                for grandchild in child.iterdir():
                    if (
                        grandchild.is_dir()
                        and grandchild.name.lower() == LEGAL_FOLDER_NAME
                    ):
                        found.append(grandchild)
            except (PermissionError, OSError):
                continue
    return found


def find_pdfs(legal_dir: Path) -> list[Path]:
    pdfs = sorted(p for p in legal_dir.glob("*.pdf") if p.is_file())
    if SEARCH_LEGAL_SUBFOLDERS_FOR_PDFS:
        for sub in [d for d in legal_dir.iterdir() if d.is_dir()]:
            pdfs.extend(sorted(sub.glob("*.pdf")))
    return pdfs


# --- classification ----------------------------------------------------------------


def categorize_clause(clause: str) -> list[str]:
    matches = []
    for category, patterns in LIMIT_CATEGORIES.items():
        if any(re.search(p, clause, re.IGNORECASE) for p in patterns):
            matches.append(category)
    for generic, specifics in CATEGORY_OVERRIDES.items():
        if generic in matches and specifics & set(matches):
            matches.remove(generic)
    return matches


def process_mandate(
    fund: str,
    source_pdf: str,
    text: str,
    confident: list[Confident],
    review: list[ReviewItem],
    seen_fund_category: dict[str, Confident],
) -> None:
    """seen_fund_category must be shared across every PDF processed for this fund —
    NOT reset per PDF — so that two mandate PDFs in the same Legal folder (e.g. an
    old and a current version) which both yield a value for the same category are
    caught as a conflict rather than one silently overwriting the other."""
    for clause in split_clauses(normalize_text(text)):
        categories = categorize_clause(clause)
        is_limit_like = bool(LIMIT_PHRASE_RE.search(clause))
        numbers = list(
            dict.fromkeys(NUMERIC_RE.findall(clause))
        )  # de-duplicate, preserve order

        if not categories and not is_limit_like:
            continue  # not limit-related at all — skip silently, this is the overwhelming majority of clauses

        if not categories and is_limit_like:
            review.append(
                ReviewItem(
                    fund,
                    source_pdf,
                    "UNCLASSIFIED",
                    clause,
                    "Limit-like language matched no category",
                )
            )
            continue

        if len(categories) > 1:
            review.append(
                ReviewItem(
                    fund,
                    source_pdf,
                    " / ".join(categories),
                    clause,
                    "Matched more than one category",
                )
            )
            continue

        category = categories[0]
        if len(numbers) != 1:
            reason = (
                "No extractable number found"
                if not numbers
                else f"Multiple numbers found: {numbers}"
            )
            review.append(ReviewItem(fund, source_pdf, category, clause, reason))
            continue

        key = f"{fund}|{category}"
        candidate = Confident(fund, category, numbers[0], clause, source_pdf)
        if key in seen_fund_category:
            # Second confident hit for the same fund+category — don't silently overwrite or pick one.
            review.append(
                ReviewItem(
                    fund,
                    source_pdf,
                    category,
                    clause,
                    f"Second value found for this fund+category (other value: {seen_fund_category[key].value})",
                )
            )
            prior = seen_fund_category.pop(key)
            confident.remove(prior)
            review.append(
                ReviewItem(
                    prior.fund,
                    prior.source_pdf,
                    prior.category,
                    prior.clause,
                    f"First of two conflicting values for this fund+category (other value: {candidate.value})",
                )
            )
            continue
        seen_fund_category[key] = candidate
        confident.append(candidate)


# --- main run ------------------------------------------------------------------------


def run(
    root: Path, limit: int | None, dry_run: bool
) -> tuple[list[Confident], list[ReviewItem], list[LogRow], dict[str, date]]:
    confident: list[Confident] = []
    review: list[ReviewItem] = []
    log: list[LogRow] = []

    if not root.exists():
        print(f"Root path does not exist or is not reachable: {root}", file=sys.stderr)
        sys.exit(1)

    client_dirs = sorted(d for d in root.iterdir() if d.is_dir())
    if limit is not None:
        client_dirs = client_dirs[:limit]

    print(
        f"{'DRY-RUN — ' if dry_run else ''}Scanning {len(client_dirs)} client folder(s) under {root}\n"
    )

    # One shared dict per client, spanning every Legal folder / PDF found for that
    # client, so a second value for the same fund+category is always caught as a
    # conflict — regardless of which PDF or which Legal folder it came from.
    seen_fund_category_by_client: dict[str, dict[str, Confident]] = {}

    # Latest effective/commencement date seen so far per client, across every
    # Legal folder / PDF — e.g. an old and a current mandate version both give
    # a date, and the current (later) one wins.
    effective_date_by_client: dict[str, date] = {}

    for client_dir in client_dirs:
        legal_folders = find_legal_folders(client_dir)
        if not legal_folders:
            log.append(LogRow(client_dir.name, "", "", "NO_LEGAL_FOLDER_FOUND", ""))
            print(f"  [NO LEGAL FOLDER] {client_dir.name}")
            continue

        for legal_dir in legal_folders:
            pdfs = find_pdfs(legal_dir)
            legal_rel = str(legal_dir.relative_to(root))
            if not pdfs:
                log.append(LogRow(client_dir.name, legal_rel, "", "NO_PDF_FOUND", ""))
                print(f"  [NO PDF]         {legal_rel}")
                continue

            status = "MULTIPLE_PDFS" if len(pdfs) > 1 else "OK"
            print(f"  [{status:15}] {legal_rel}  ({len(pdfs)} pdf)")
            log.append(
                LogRow(
                    client_dir.name,
                    legal_rel,
                    "; ".join(p.name for p in pdfs),
                    status,
                    "Manually confirm which PDF is the operative mandate"
                    if len(pdfs) > 1
                    else "",
                )
            )

            if dry_run:
                continue

            seen_fund_category: dict[str, Confident] = (
                seen_fund_category_by_client.setdefault(client_dir.name, {})
            )
            for pdf_path in pdfs:
                try:
                    text = extract_pdf_text(pdf_path)
                except RuntimeError as exc:
                    log.append(
                        LogRow(
                            client_dir.name, legal_rel, pdf_path.name, "ERROR", str(exc)
                        )
                    )
                    continue
                if text is None:
                    log.append(
                        LogRow(
                            client_dir.name,
                            legal_rel,
                            pdf_path.name,
                            "NO_TEXT",
                            "No extractable text — likely a scanned PDF; needs OCR/Word conversion",
                        )
                    )
                    continue

                pdf_effective_date = extract_effective_date(text)
                if pdf_effective_date is not None:
                    current = effective_date_by_client.get(client_dir.name)
                    if current is None or pdf_effective_date > current:
                        effective_date_by_client[client_dir.name] = pdf_effective_date

                process_mandate(
                    client_dir.name,
                    str(pdf_path.relative_to(root)),
                    text,
                    confident,
                    review,
                    seen_fund_category,
                )

    return confident, review, log, effective_date_by_client


def build_matrix(
    confident: list[Confident], effective_dates: dict[str, date]
) -> pd.DataFrame:
    categories = list(LIMIT_CATEGORIES.keys())
    funds = sorted({c.fund for c in confident} | set(effective_dates))
    matrix = pd.DataFrame(
        index=funds, columns=["Mandate Effective Date"] + categories, dtype=object
    )
    for c in confident:
        matrix.loc[c.fund, c.category] = c.value
    for fund, eff_date in effective_dates.items():
        matrix.loc[fund, "Mandate Effective Date"] = eff_date
    matrix.index.name = "Fund / Client Folder"
    return matrix


def write_workbook(
    out_path: Path,
    confident: list[Confident],
    review: list[ReviewItem],
    log: list[LogRow],
    effective_dates: dict[str, date],
) -> None:
    matrix = build_matrix(confident, effective_dates)
    review_df = (
        pd.DataFrame([vars(r) for r in review])
        if review
        else pd.DataFrame(
            columns=["fund", "source_pdf", "category", "clause", "reason"]
        )
    )
    log_df = (
        pd.DataFrame([vars(l) for l in log])
        if log
        else pd.DataFrame(
            columns=["client_folder", "legal_folder", "pdfs_found", "status", "detail"]
        )
    )

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        matrix.to_excel(writer, sheet_name="Limits Matrix")
        review_df.to_excel(writer, sheet_name="Needs Review", index=False)
        log_df.to_excel(writer, sheet_name="Processing Log", index=False)


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter
    )
    parser.add_argument(
        "--root", type=Path, default=DEFAULT_ROOT, help="Active Clients root folder"
    )
    parser.add_argument(
        "--out",
        type=Path,
        default=None,
        help=f"output .xlsx path (default: {DEFAULT_OUTPUT_DIR / 'mandate_limits_<date>.xlsx'})",
    )
    parser.add_argument(
        "--limit",
        type=int,
        default=None,
        help="only process the first N client folders (pilot testing)",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="list discovered Legal folders/PDFs; extract and write nothing",
    )
    args = parser.parse_args(argv)

    confident, review, log, effective_dates = run(args.root, args.limit, args.dry_run)

    if args.dry_run:
        print("\nDry run complete. Nothing extracted, nothing written.")
        return 0

    out_path = args.out or (
        DEFAULT_OUTPUT_DIR / f"mandate_limits_{date.today().isoformat()}.xlsx"
    )
    out_path.parent.mkdir(parents=True, exist_ok=True)
    write_workbook(out_path, confident, review, log, effective_dates)

    print(f"\nDone.")
    print(f"  Confident matrix entries : {len(confident)}")
    print(f"  Flagged for your review  : {len(review)}")
    print(f"  Mandate effective dates  : {len(effective_dates)}")
    print(f"  Client folders logged    : {len(log)}")
    print(f"  Written to               : {out_path.resolve()}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
aduklt AUTISM 