import csv, os, zipfile, xml.etree.ElementTree as ET

SUPPORTED = {".txt",".csv",".tsv",".docx",".xlsx",".xlsm",".xls",".json"}
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
with open(r"D:\Active research\OpenSearch\phi_scan_report.csv", encoding="utf-8-sig") as f:
    for row in csv.DictReader(f):
        flagged.add(os.path.normpath(row.get("file","")))

no_hit = sorted(all_files - flagged)

def get_xlsx_headers_fast(fpath):
    """Read only the first row from xlsx without loading entire file."""
    try:
        ns = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
        with zipfile.ZipFile(fpath, "r") as z:
            # Find first sheet
            with z.open("xl/workbook.xml") as wb_xml:
                tree = ET.parse(wb_xml)
                sheets = tree.findall(f".//{{{ns}}}sheet")
                rId = sheets[0].get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id") if sheets else "rId1"
            with z.open("xl/_rels/workbook.xml.rels") as rels_xml:
                rtree = ET.parse(rels_xml)
                sheet_path = "xl/worksheets/sheet1.xml"
                for r in rtree.findall(".//{http://schemas.openxmlformats.org/package/2006/relationships}Relationship"):
                    if r.get("Id") == rId:
                        sheet_path = "xl/" + r.get("Target").lstrip("/")
                        break
            # Read shared strings
            shared = []
            if "xl/sharedStrings.xml" in z.namelist():
                with z.open("xl/sharedStrings.xml") as ss:
                    ss_tree = ET.parse(ss)
                    for si in ss_tree.findall(f".//{{{ns}}}si"):
                        t = "".join(n.text or "" for n in si.iter(f"{{{ns}}}t"))
                        shared.append(t)
            # Read first row only
            headers = []
            with z.open(sheet_path) as sh:
                for event, elem in ET.iterparse(sh, events=("end",)):
                    if elem.tag == f"{{{ns}}}row":
                        for c in elem:
                            t = c.get("t","")
                            v = c.find(f"{{{ns}}}v")
                            val = ""
                            if t == "s" and v is not None:
                                try: val = shared[int(v.text)]
                                except: val = v.text or ""
                            elif v is not None:
                                val = v.text or ""
                            headers.append(val)
                        break  # only first row
            return headers
    except Exception as e:
        return [f"ERROR: {e}"]

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
            return get_xlsx_headers_fast(fpath), "?"
        elif ext == ".xls":
            import xlrd
            wb = xlrd.open_workbook(fpath, on_demand=True)
            ws = wb.sheet_by_index(0)
            headers = [str(ws.cell_value(0, c)) for c in range(ws.ncols)]
            return headers, ws.nrows - 1
        elif ext == ".docx":
            from docx import Document
            doc = Document(fpath)
            if doc.tables:
                return [cell.text for cell in doc.tables[0].rows[0].cells], len(doc.tables[0].rows)-1
            return ["(text document)"], len(doc.paragraphs)
        elif ext == ".json":
            import json
            with open(fpath, encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list) and data and isinstance(data[0], dict):
                return list(data[0].keys()), len(data)
            elif isinstance(data, dict):
                return list(data.keys()), 1
    except Exception as e:
        return [f"ERROR: {e}"], "?"
    return ["(unsupported)"], "?"

out_path = r"D:\Active research\OpenSearch\no_phi_fields.txt"
with open(out_path, "w", encoding="utf-8") as out:
    out.write(f"NO-PHI-HIT FILES — F:\\01_Raw_Data\n")
    out.write(f"Total: {len(no_hit)} files\n")
    out.write("=" * 70 + "\n")
    for fpath in no_hit:
        size = os.path.getsize(fpath)
        fields, rows = get_fields(fpath)
        out.write(f"\n{fpath}\n")
        out.write(f"  {rows} rows  |  {size:,} bytes  |  {len(fields)} fields\n")
        for i, h in enumerate(fields, 1):
            out.write(f"    {i:>3}. {h}\n")

print(f"Done — {len(no_hit)} files. Output: {out_path}")
