import csv, os, sys

SUPPORTED = {".txt",".csv",".tsv",".docx",".xlsx",".xlsm",".xls",".pdf",".json"}
SKIP_DIRS = {"$recycle.bin","system volume information","windows","program files",
             "program files (x86)","programdata","appdata",".gemini"}

root = r"F:\01_Raw_Data"

all_files = set()
for dirpath, dirs, files in os.walk(root):
    dirs[:] = [d for d in dirs if d.lower() not in SKIP_DIRS and not d.startswith(".")]
    for f in files:
        if os.path.splitext(f)[1].lower() in SUPPORTED:
            all_files.add(os.path.normpath(os.path.join(dirpath, f)))

flagged = set()
report = r"D:\Active research\OpenSearch\phi_scan_report.csv"
with open(report, encoding="utf-8-sig") as f:
    for row in csv.DictReader(f):
        flagged.add(os.path.normpath(row.get("file","")))

no_hit = sorted(all_files - flagged)

out = open(r"D:\Active research\OpenSearch\no_phi_fields.txt", "w", encoding="utf-8")

def pr(s=""):
    out.write(s + "\n")

def get_fields(fpath):
    ext = os.path.splitext(fpath)[1].lower()
    try:
        if ext in (".csv",".tsv",".txt"):
            delim = "\t" if ext == ".tsv" else ","
            with open(fpath, encoding="utf-8-sig", errors="replace") as f:
                reader = csv.reader(f, delimiter=delim)
                headers = next(reader, [])
                rows = sum(1 for _ in reader)
            return headers, rows
        elif ext in (".xlsx",".xlsm"):
            import openpyxl
            wb = openpyxl.load_workbook(fpath, read_only=True, data_only=True)
            ws = wb.active
            headers = [str(c.value) if c.value is not None else "" for c in next(ws.iter_rows(max_row=1))]
            rows = ws.max_row - 1 if ws.max_row else 0
            wb.close()
            return headers, rows
        elif ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(fpath)
            ws = wb.sheet_by_index(0)
            headers = [str(ws.cell_value(0, c)) for c in range(ws.ncols)]
            return headers, ws.nrows - 1
        elif ext == ".docx":
            from docx import Document
            doc = Document(fpath)
            if doc.tables:
                headers = [cell.text for cell in doc.tables[0].rows[0].cells]
                return headers, len(doc.tables[0].rows) - 1
            return ["(text document — no table)"], len(doc.paragraphs)
        elif ext == ".json":
            import json
            with open(fpath, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and data and isinstance(data[0], dict):
                return list(data[0].keys()), len(data)
            elif isinstance(data, dict):
                return list(data.keys()), 1
            return ["(unknown structure)"], 0
    except Exception as e:
        return [f"ERROR: {e}"], 0
    return ["(unsupported)"], 0

for fpath in no_hit:
    size = os.path.getsize(fpath)
    fields, rows = get_fields(fpath)
    pr()
    pr(fpath)
    pr(f"  {rows:,} rows  |  {size:,} bytes  |  {len(fields)} fields")
    for i, h in enumerate(fields, 1):
        pr(f"    {i:>3}. {h}")

out.close()
print(f"Done. {len(no_hit)} files written to no_phi_fields.txt")
