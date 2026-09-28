"""
LOCAL PHI SCANNER FOR WINDOWS  (improved)
==========================================

Scans:
    .txt  .csv  .tsv
    .docx
    .xlsx  .xlsm  .xls
    .pdf
    .json

Does NOT:
    - use an LLM
    - access the internet
    - call an API
    - upload files
    - modify scanned files
    - put detected PHI values into the report

IMPORTANT:
    This is a screening tool, not a guarantee that a file contains or
    does not contain PHI.

PDF NOTE:
    Scans PDFs that contain extractable text.
    Image-only / scanned PDFs are flagged separately (OCR needed).

USAGE:
    py phi_scanner.py C:\\Users\\YourName --report C:\\phi_report.csv
    py phi_scanner.py C:\\Users\\YourName --report C:\\phi_report.csv --min-risk MEDIUM
    py phi_scanner.py C:\\Users\\YourName --report C:\\phi_report.csv --workers 8

INSTALL DEPENDENCIES:
    py -m pip install openpyxl python-docx pymupdf xlrd
"""

import os
import re
import csv
import sys
import json
import argparse
from pathlib import Path
from collections import Counter
from concurrent.futures import ThreadPoolExecutor, as_completed


# ---------------------------------------------------------------------------
# Dependency check
# ---------------------------------------------------------------------------

try:
    import openpyxl
    from docx import Document
    import fitz           # PyMuPDF
    import xlrd           # legacy .xls
except ImportError as e:
    print("\nMissing required Python package.")
    print("Run:")
    print("    py -m pip install openpyxl python-docx pymupdf xlrd")
    print(f"\nDetails: {e}")
    sys.exit(1)


# ============================================================
# CONFIGURATION
# ============================================================

SUPPORTED_EXTENSIONS = {
    ".txt", ".csv", ".tsv",
    ".docx",
    ".xlsx", ".xlsm", ".xls",
    ".pdf",
    ".json",
    ".parquet",          # always flagged HIGH (binary columnar format)
}

SKIP_DIRECTORY_NAMES = {
    "$recycle.bin",
    "system volume information",
    "windows",
    "program files",
    "program files (x86)",
    "programdata",
    "$windows.~bt",
    "$windows.~ws",
    "node_modules",
    ".git",
    "__pycache__",
}

MAX_TEXT_CHARS = 5_000_000   # per file


# ============================================================
# PHI DETECTION PATTERNS
# ============================================================
#
# All patterns require contextual evidence (a field label, format
# specificity, or structural indicator) to reduce false positives.
# Matched values are NEVER stored or reported.

PATTERNS = {

    # ---- Strong identifiers (contextual label required) ----

    "MRN label": re.compile(
        r"\b(?:MRN|medical\s+record\s+(?:number|no\.?))"
        r"\s*[:#=\-]?\s*[A-Z0-9\-]{4,25}\b",
        re.IGNORECASE,
    ),

    "Patient ID label": re.compile(
        r"\b(?:patient\s*(?:id|identifier|number|no\.?))"
        r"\s*[:#=\-]?\s*[A-Z0-9\-]{3,30}\b",
        re.IGNORECASE,
    ),

    "DOB label": re.compile(
        r"\b(?:DOB|D\.O\.B\.|date\s+of\s+birth|birth\s+date)"
        r"\s*[:#=\-]?\s*"
        r"(?:"
        r"\d{1,2}[\/\-\.]\d{1,2}[\/\-\.](?:19|20)?\d{2}"
        r"|"
        r"(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|"
        r"May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|"
        r"Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
        r"\s+\d{1,2},?\s+(?:19|20)\d{2}"
        r")",
        re.IGNORECASE,
    ),

    "SSN label": re.compile(
        r"\b(?:SSN|social\s+security(?:\s+number)?)"
        r"\s*[:#=\-]?\s*\d{3}[- ]?\d{2}[- ]?\d{4}\b",
        re.IGNORECASE,
    ),

    # Bare SSN: full hyphenated form AND no surrounding digits or slashes
    # (reduces false positives from phone extensions / part numbers)
    "SSN pattern": re.compile(
        r"(?<![\/\d\-])\d{3}-\d{2}-\d{4}(?![\/\d\-])"
    ),

    "Patient name label": re.compile(
        r"\b(?:patient\s+name|pt\s+name)"
        r"\s*[:#=\-]\s*"
        r"[A-Z][A-Za-z'\-]+"
        r"(?:\s*,?\s+[A-Z][A-Za-z'\-]+){1,3}",
        re.IGNORECASE,
    ),

    # ---- Contact information ----

    # Require area code + separator to reduce false positives
    "Phone number": re.compile(
        r"(?<!\d)"
        r"(?:\+?1[\s\-.]?)?"
        r"(?:\(\d{3}\)|\d{3})"
        r"[\s\-.]\d{3}[\s\-.]\d{4}"
        r"(?!\d)"
    ),

    "Email address": re.compile(
        r"\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b",
        re.IGNORECASE,
    ),

    # ZIP+4 (HIPAA Safe Harbor: full ZIP+4 is a direct locator)
    "ZIP+4 code": re.compile(
        r"(?<!\d)\d{5}-\d{4}(?!\d)"
    ),

    # ---- Insurance / payer identifiers ----

    "Member ID label": re.compile(
        r"\b(?:member\s*(?:id|number|no\.?)|"
        r"subscriber\s*(?:id|number|no\.?)|"
        r"beneficiary\s*(?:id|number|no\.?))"
        r"\s*[:#=\-]?\s*[A-Z0-9\-]{4,30}\b",
        re.IGNORECASE,
    ),

    "Account number label": re.compile(
        r"\b(?:account|acct)\s*(?:number|no\.?|#)?"
        r"\s*[:#=\-]\s*[A-Z0-9\-]{4,30}\b",
        re.IGNORECASE,
    ),

    # ---- Provider / clinical identifiers ----

    # NPI: exactly 10 digits, label required
    "NPI number": re.compile(
        r"\b(?:NPI|national\s+provider\s+(?:id|identifier))"
        r"\s*[:#=\-]?\s*\d{10}\b",
        re.IGNORECASE,
    ),

    # DEA: two letters then 7 digits
    "DEA number": re.compile(
        r"\b(?:DEA(?:\s+number|\s+no\.?)?)"
        r"\s*[:#=\-]?\s*[A-Z]{2}\d{7}\b",
        re.IGNORECASE,
    ),

    # ICD-10 in context (diagnosis label required)
    "ICD-10 code label": re.compile(
        r"\b(?:diagnosis|dx|icd\s*(?:10|code|[-]?10[-]?cm)?)"
        r"\s*[:#=\-]?\s*[A-Z]\d{2}(?:\.\d{1,4})?\b",
        re.IGNORECASE,
    ),

    # CPT in context
    "CPT code label": re.compile(
        r"\b(?:CPT|procedure\s+code)"
        r"\s*[:#=\-]?\s*\d{5}(?:[A-Z])?\b",
        re.IGNORECASE,
    ),

    # ---- Genomic identifiers ----

    # dbSNP rs numbers
    "Genomic rsID": re.compile(
        r"\brs\d{5,}\b",
        re.IGNORECASE,
    ),

    # FHIR/HL7 resource references with patient context
    "FHIR Patient reference": re.compile(
        r'"(?:subject|patient|resourceType)"\s*:\s*"Patient(?:/[^"]{1,30})?"',
        re.IGNORECASE,
    ),
}


# ---- Header / column name dictionaries ----

HIGH_RISK_HEADERS = {
    "mrn", "medical record number", "medical record #",
    "patient name", "patient_name", "patientname",
    "date of birth", "birth date", "dob", "birthdate",
    "social security number", "ssn", "social security",
    "npi", "national provider id",
    "dea", "dea number",
}

# Substrings that, if found anywhere in a column header, trigger HIGH risk.
# Applied case-insensitively after stripping whitespace.
HEADER_PHI_SUBSTRINGS = {
    "name",      # first_name, last_name, patient_name, provider_name, etc.
    "mrn",       # any field containing MRN
    "date",      # encounter_date, visit_date, dob, dod, order_date, etc.
    "dob",       # date_of_birth, dob_year, etc.
    "dod",       # date_of_death, dod, etc.
    "death",     # death_date, date_of_death, cause_of_death
    "admit",     # admit_date, admission_date, ADMIT, admitdx, etc.
    "discharge", # discharge_date, DISCHARGE, dischargedx, etc.
    "order",     # order_date, ORDERDT, orderid, etc.
    "year",      # birth_year, admit_year, order_year, etc.
    "patient",   # patient_id, patient_name, patientid, etc.
    "person",    # personid, person_id, etc.
    "dt",        # orderdt, admit_dt, encounter_dt, etc.
    "id",        # any identifier field: studyid, encounterid, surgeryid, proc_id, etc.
}

MEDIUM_RISK_HEADERS = {
    "patient id", "patient_id", "patientid", "patient number",
    "member id", "member_id", "memberid",
    "subscriber id", "beneficiary id",
    "account number", "acct number",
    "phone", "telephone", "mobile", "cell",
    "email", "e-mail",
    "address", "street address", "street", "addr",
    "zip", "zip code", "zipcode", "postal code",
    "encounter id", "visit id", "admission id",
    "icd", "icd10", "icd-10", "diagnosis code", "dx code",
    "cpt", "procedure code",
}


# ============================================================
# HELPERS
# ============================================================

def normalize_header(value):
    if value is None:
        return ""
    value = str(value).strip().lower()
    value = re.sub(r"\s+", " ", value)
    return value


def analyze_text(text):
    """Return finding categories/counts. Matched values are discarded."""
    findings = Counter()
    if not text:
        return findings
    for name, pattern in PATTERNS.items():
        count = len(pattern.findall(text))
        if count:
            findings[name] += count
    return findings


def analyze_headers(headers):
    findings = Counter()
    for header in headers:
        normalized = normalize_header(header)
        if not normalized:
            continue
        if normalized in HIGH_RISK_HEADERS:
            findings[f"PHI column/header: {normalized}"] += 1
        elif normalized in MEDIUM_RISK_HEADERS:
            findings[f"Identifier column/header: {normalized}"] += 1
        else:
            # Substring check: flag if any PHI keyword appears anywhere in the header
            for kw in HEADER_PHI_SUBSTRINGS:
                if kw in normalized:
                    findings[f"PHI keyword in header: {normalized}"] += 1
                    break
    return findings


def merge_findings(destination, source):
    for key, value in source.items():
        destination[key] += value


def classify_risk(findings):
    if not findings:
        return "NONE"

    names = set(findings.keys())

    high_indicators = {
        "MRN label",
        "DOB label",
        "SSN label",
        "SSN pattern",
        "Patient name label",
        "NPI number",
        "DEA number",
        "FHIR Patient reference",
    }

    if names.intersection(high_indicators):
        return "HIGH"
    # Parquet files are always HIGH — binary columnar format may contain PHI
    # with no way to inspect content without pyarrow
    if any(n.startswith("Parquet file") for n in names):
        return "HIGH"
    if any(n.startswith("PHI column/header:") for n in names):
        return "HIGH"
    if any(n.startswith("PHI keyword in header:") for n in names):
        return "HIGH"
    if len(names) >= 3:
        return "HIGH"
    if len(names) >= 2:
        return "MEDIUM"

    medium_indicators = {
        "Patient ID label",
        "Member ID label",
        "Account number label",
        "ICD-10 code label",
        "CPT code label",
        "ZIP+4 code",
        "Genomic rsID",
    }

    if names.intersection(medium_indicators):
        return "MEDIUM"
    if any(n.startswith("Identifier column/header:") for n in names):
        return "MEDIUM"

    return "LOW"


def findings_to_string(findings):
    """Serialize categories to a readable string. No PHI values included."""
    return "; ".join(
        f"{k} ({v})"
        for k, v in sorted(findings.items())
    )


def safe_file_size(path):
    try:
        return path.stat().st_size
    except Exception:
        return ""


# ============================================================
# FILE READERS
# ============================================================

MAX_SCAN_ROWS = 100  # regex-scan only this many data rows per CSV/Excel file

def scan_plain_text(path):
    findings = Counter()
    ext = path.suffix.lower()
    is_tabular = ext in {".csv", ".tsv"}
    delimiter = "\t" if ext == ".tsv" else ","
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]

    lines = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                if is_tabular:
                    lines = [f.readline() for _ in range(MAX_SCAN_ROWS + 1)]
                else:
                    lines = [f.read(MAX_TEXT_CHARS)]
            break
        except UnicodeDecodeError:
            continue
    if lines is None:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            if is_tabular:
                lines = [f.readline() for _ in range(MAX_SCAN_ROWS + 1)]
            else:
                lines = [f.read(MAX_TEXT_CHARS)]

    if is_tabular:
        try:
            header_row = next(csv.reader([lines[0]], delimiter=delimiter), [])
            merge_findings(findings, analyze_headers(header_row))
        except Exception:
            pass
        text = "".join(lines[1:])
    else:
        text = lines[0] if lines else ""

    merge_findings(findings, analyze_text(text))
    return findings, None


def scan_docx(path):
    findings = Counter()
    doc = Document(path)
    total_chars = 0

    for para in doc.paragraphs:
        if para.text:
            merge_findings(findings, analyze_text(para.text))
            total_chars += len(para.text)
        if total_chars >= MAX_TEXT_CHARS:
            break

    if total_chars < MAX_TEXT_CHARS:
        for table in doc.tables:
            if table.rows:
                headers = [c.text for c in table.rows[0].cells]
                merge_findings(findings, analyze_headers(headers))
            for row in table.rows:
                for cell in row.cells:
                    if cell.text:
                        merge_findings(findings, analyze_text(cell.text))
                        total_chars += len(cell.text)
                    if total_chars >= MAX_TEXT_CHARS:
                        break
                if total_chars >= MAX_TEXT_CHARS:
                    break
            if total_chars >= MAX_TEXT_CHARS:
                break

    return findings, None


def scan_excel_xlsx(path):
    findings = Counter()
    wb = openpyxl.load_workbook(filename=path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            header_done = False
            data_rows = 0
            for row in ws.iter_rows(values_only=True):
                values = [str(v) for v in row if v is not None]
                if not values:
                    continue
                if not header_done:
                    merge_findings(findings, analyze_headers(values))
                    header_done = True
                    continue
                merge_findings(findings, analyze_text(" | ".join(values)))
                data_rows += 1
                if data_rows >= MAX_SCAN_ROWS:
                    break
    finally:
        wb.close()
    return findings, None


def scan_excel_xls(path):
    """Legacy .xls support via xlrd."""
    findings = Counter()
    try:
        wb = xlrd.open_workbook(str(path), on_demand=True)
    except Exception as e:
        return findings, f"XLS READ ERROR: {type(e).__name__}"

    for sheet_name in wb.sheet_names():
        ws = wb.sheet_by_name(sheet_name)
        if ws.nrows == 0:
            continue
        headers = [str(ws.cell_value(0, c)) for c in range(ws.ncols)
                   if ws.cell_value(0, c) not in (None, "")]
        merge_findings(findings, analyze_headers(headers))
        for row_idx in range(1, min(ws.nrows, MAX_SCAN_ROWS + 1)):
            values = [str(ws.cell_value(row_idx, c)) for c in range(ws.ncols)
                      if ws.cell_value(row_idx, c) not in (None, "")]
            if values:
                merge_findings(findings, analyze_text(" | ".join(values)))
    return findings, None


def scan_pdf(path):
    findings = Counter()
    pdf = fitz.open(str(path))
    total_chars = 0
    pages_with_text = 0
    try:
        for page in pdf:
            text = page.get_text("text")
            if text and text.strip():
                pages_with_text += 1
                merge_findings(findings, analyze_text(text))
                total_chars += len(text)
            if total_chars >= MAX_TEXT_CHARS:
                break
        if len(pdf) > 0 and pages_with_text == 0:
            return findings, "PDF HAS NO EXTRACTABLE TEXT - OCR NEEDED"
    finally:
        pdf.close()
    return findings, None


def scan_json(path):
    """
    Scan JSON including FHIR bundles and HL7 exports.
    Serializes to text for pattern matching; checks top-level keys as headers.
    """
    findings = Counter()
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin-1"]
    text = None
    for enc in encodings:
        try:
            with open(path, "r", encoding=enc, errors="strict") as f:
                text = f.read(MAX_TEXT_CHARS)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        with open(path, "r", encoding="utf-8", errors="ignore") as f:
            text = f.read(MAX_TEXT_CHARS)

    merge_findings(findings, analyze_text(text))

    # Extract JSON keys as potential headers
    try:
        obj = json.loads(text)
        keys = []
        if isinstance(obj, dict):
            keys = list(obj.keys())
            if obj.get("resourceType") in ("Bundle", "Patient"):
                findings["FHIR resource detected"] += 1
        elif isinstance(obj, list) and obj and isinstance(obj[0], dict):
            keys = list(obj[0].keys())
        merge_findings(findings, analyze_headers(keys))
    except Exception:
        pass  # Not valid JSON or truncated — text scan above still ran

    return findings, None


def scan_parquet(path):
    """Parquet files are binary columnar format — auto-flag HIGH.
    Attempt to read column names via pyarrow if available; otherwise flag blind."""
    findings = Counter()
    findings["Parquet file (binary columnar — header check only)"] += 1
    try:
        import pyarrow.parquet as pq  # optional — not required
        schema = pq.read_schema(str(path))
        headers = [str(f.name) for f in schema]
        merge_findings(findings, analyze_headers(headers))
        findings["Parquet columns read"] = len(headers)
    except ImportError:
        findings["Parquet (pyarrow not installed — column names unread)"] += 1
    except Exception as e:
        findings[f"Parquet read error: {type(e).__name__}"] += 1
    return findings, None


# ============================================================
# FILE DISPATCH
# ============================================================

def scan_file(path):
    ext = path.suffix.lower()
    if ext in {".txt", ".csv", ".tsv"}:
        return scan_plain_text(path)
    elif ext == ".docx":
        return scan_docx(path)
    elif ext in {".xlsx", ".xlsm"}:
        return scan_excel_xlsx(path)
    elif ext == ".xls":
        return scan_excel_xls(path)
    elif ext == ".pdf":
        return scan_pdf(path)
    elif ext == ".json":
        return scan_json(path)
    elif ext == ".parquet":
        return scan_parquet(path)
    return Counter(), None


# ============================================================
# FILENAME ANALYSIS
# ============================================================

_FILENAME_KEYWORDS = {
    "patient", "patients", "mrn", "medical_record",
    "medical record", "dob", "phi", "ssn", "subject_id",
    "participant", "encounter", "cohort",
}


def analyze_filename(path):
    findings = Counter()
    lower = path.stem.lower()
    merge_findings(findings, analyze_text(path.stem))
    for kw in _FILENAME_KEYWORDS:
        if kw in lower:
            findings["PHI-related filename"] += 1
            break
    return findings


# ============================================================
# DIRECTORY TRAVERSAL
# ============================================================

def should_skip_directory(path):
    try:
        name = path.name.lower()
        if name in SKIP_DIRECTORY_NAMES:
            return True
        if name.startswith("$"):
            return True
    except Exception:
        pass
    return False


def get_files(root):
    for current_root, dirs, files in os.walk(
        root, topdown=True, onerror=lambda e: None
    ):
        dirs[:] = [
            d for d in dirs
            if not should_skip_directory(Path(current_root) / d)
        ]
        for filename in files:
            path = Path(current_root) / filename
            if path.suffix.lower() in SUPPORTED_EXTENSIONS:
                yield path


# ============================================================
# WORKER (for thread pool)
# ============================================================

def process_file(path):
    """
    Scan a single file. Returns a result dict.
    Never returns matched PHI values.
    """
    findings = Counter()
    status = "OK"

    merge_findings(findings, analyze_filename(path))

    try:
        file_findings, warning = scan_file(path)
        merge_findings(findings, file_findings)
        if warning:
            status = warning
    except PermissionError:
        status = "ACCESS DENIED"
    except Exception as e:
        status = f"READ ERROR: {type(e).__name__}"

    return {
        "path": path,
        "risk": classify_risk(findings),
        "findings": findings,
        "status": status,
        "size_bytes": safe_file_size(path),
        "extension": path.suffix.lower(),
    }


# ============================================================
# MAIN SCANNER
# ============================================================

RISK_ORDER = {"NONE": 0, "LOW": 1, "MEDIUM": 2, "HIGH": 3}


def run_scan(root, report_path, min_risk="LOW", workers=4):

    min_risk_level = RISK_ORDER.get(min_risk.upper(), 1)
    root = Path(root)
    report_path = Path(report_path)

    if not root.exists():
        print(f"\nERROR: Path does not exist: {root}")
        sys.exit(1)

    print()
    print("LOCAL PHI SCANNER")
    print("=" * 60)
    print(f"Scanning : {root}")
    print(f"Report   : {report_path}")
    print(f"Min risk : {min_risk.upper()}")
    print(f"Workers  : {workers}")
    print()
    print("No file contents are transmitted anywhere.")
    print("The scanner does not modify scanned files.")
    print("=" * 60)
    print()

    print("Collecting file list...", end="", flush=True)
    all_files = list(get_files(root))
    total = len(all_files)
    print(f" {total:,} files found.\n")

    scanned = 0
    flagged = 0
    errors = 0
    ocr_needed = 0
    risk_counts = Counter()
    extension_counts = Counter()

    with open(report_path, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([
            "risk",
            "file",
            "extension",
            "size_bytes",
            "findings",
            "status",
        ])

        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(process_file, p): p for p in all_files}

            for future in as_completed(futures):
                scanned += 1

                current_path = futures[future]
                # Emit structured line parseable by phi_app.py
                print(
                    f"FILE:{current_path}\t"
                    f"PROGRESS:{scanned}/{total}\t"
                    f"FLAGGED:{flagged}\t"
                    f"ERRORS:{errors}",
                    flush=True,
                )

                try:
                    result = future.result()
                except Exception:
                    errors += 1
                    continue

                risk = result["risk"]
                status = result["status"]
                extension_counts[result["extension"]] += 1
                risk_counts[risk] += 1

                if "ACCESS DENIED" in status or "READ ERROR" in status:
                    errors += 1
                if "OCR NEEDED" in status:
                    ocr_needed += 1

                risk_level = RISK_ORDER.get(risk, 0)

                # Write if risk meets threshold, or if there was an error/warning
                if risk_level >= min_risk_level or (status != "OK" and risk == "NONE"):
                    if risk != "NONE":
                        flagged += 1
                    writer.writerow([
                        risk,
                        str(result["path"]),
                        result["extension"],
                        result["size_bytes"],
                        findings_to_string(result["findings"]),
                        status,
                    ])

    print(
        f"\r  {scanned:,}/{total:,} (100%)  "
        f"Flagged: {flagged:,}  Errors: {errors:,}     "
    )
    print()
    print("=" * 60)
    print("  SUMMARY")
    print("=" * 60)
    print(f"  Root               : {root}")
    print(f"  Total scanned      : {scanned:,}")
    print(f"  Flagged (non-NONE) : {flagged:,}")
    print(f"  Access errors      : {errors:,}")
    print(f"  PDFs needing OCR   : {ocr_needed:,}")
    print()
    print("  By risk level:")
    for level in ["HIGH", "MEDIUM", "LOW", "NONE"]:
        print(f"    {level:<8}: {risk_counts[level]:,}")
    print()
    print("  By extension:")
    for ext, count in sorted(extension_counts.items(), key=lambda x: -x[1]):
        print(f"    {ext:<10}: {count:,}")
    print()
    print(f"  Report saved to: {report_path}")
    print("=" * 60)
    print()

    # Write machine-readable summary for phi_app.py to display
    from datetime import datetime as _dt
    summary = {
        "scan_root":        str(root),
        "scan_date":        _dt.now().isoformat(),
        "total_scanned":    scanned,
        "total_flagged":    flagged,
        "access_errors":    errors,
        "ocr_needed":       ocr_needed,
        "by_risk":          dict(risk_counts),
        "by_extension":     dict(sorted(extension_counts.items(), key=lambda x: -x[1])[:20]),
        "report_path":      str(report_path),
    }
    summary_path = Path(report_path).with_name("phi_scan_summary.json")
    try:
        with open(summary_path, "w", encoding="utf-8") as _sf:
            import json as _json
            _json.dump(summary, _sf, indent=2)
        print(f"SUMMARY_JSON:{summary_path}")
    except Exception as _e:
        print(f"Warning: could not write summary JSON: {_e}")


# ============================================================
# ENTRY POINT
# ============================================================

def main():
    parser = argparse.ArgumentParser(
        description="Local PHI scanner -- no network, no API, no cloud.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  py phi_scanner.py C:\\Users\\YourName
  py phi_scanner.py C:\\Users\\YourName --report C:\\phi_report.csv
  py phi_scanner.py C:\\Users\\YourName --min-risk MEDIUM --workers 8
        """,
    )
    parser.add_argument("root", help="Root directory to scan")
    parser.add_argument(
        "--report",
        default="phi_report.csv",
        help="Output CSV report path (default: phi_report.csv in current dir)",
    )
    parser.add_argument(
        "--min-risk",
        default="LOW",
        choices=["LOW", "MEDIUM", "HIGH"],
        help="Minimum risk level to include in report (default: LOW)",
    )
    parser.add_argument(
        "--workers",
        type=int,
        default=4,
        help="Number of parallel worker threads (default: 4)",
    )

    args = parser.parse_args()
    run_scan(
        root=args.root,
        report_path=args.report,
        min_risk=args.min_risk,
        workers=args.workers,
    )


if __name__ == "__main__":
    main()
