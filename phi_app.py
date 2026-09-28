"""
PHI Scanner Web App  —  http://localhost:8081
=============================================

Standalone companion to OpenSearch (search_app.py on port 8080).

Provides:
  - Browser UI to launch phi_scanner.py across configured directories
  - Real-time scan progress via /api/phi_status
  - Results indexed into a separate OpenSearch index ("phi-scan")
  - Risk-badged PIL thumbnail cards (HIGH / MEDIUM / LOW)
  - Search/filter of PHI scan results
  - Open file / open folder actions (same as search_app.py)

Does NOT touch search_app.py, the "documents" index, or .cache_thumbnails/.

To remove this feature entirely:
  1. Delete phi_app.py, phi_scanner.py (if desired)
  2. Remove the one <a> link added to search_app.py
  Done.

Dependencies (same as search_app.py + phi_scanner.py):
  py -m pip install opensearch-py pillow openpyxl python-docx pymupdf xlrd psutil
"""

import os
import io
import re
import sys
import csv
import json
import html
import math
import time
import string
import hashlib
import threading
import subprocess
import urllib.parse
from pathlib import Path
from datetime import datetime
from collections import Counter
from http.server import ThreadingHTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
from concurrent.futures import ThreadPoolExecutor

import psutil
from opensearchpy import OpenSearch
from PIL import Image, ImageDraw, ImageFont

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

PHI_PORT            = 8081
SEARCH_APP_URL      = "http://localhost:8080"
OPENSEARCH_HOST     = "localhost"
OPENSEARCH_PORT     = 9200
PHI_INDEX_NAME      = "phi-scan"
PAGE_SIZE           = 100

BASE_DIR            = os.path.dirname(os.path.abspath(__file__))
PHI_THUMB_DIR       = os.path.join(BASE_DIR, ".cache_phi_thumbnails")
PHI_CONFIG_FILE     = os.path.join(BASE_DIR, "phi_scanner_config.json")
PHI_PROGRESS_FILE   = os.path.join(BASE_DIR, "phi_scanner_progress.json")
PHI_REPORT_FILE     = os.path.join(BASE_DIR, "phi_scan_report.csv")
PHI_SCANNER_SCRIPT  = os.path.join(BASE_DIR, "phi_scanner.py")

DOPUS_RT  = r"C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DOPUS_EXE = r"C:\Program Files\GPSoftware\Directory Opus\dopus.exe"

HIDDEN_DIRS = {
    "$recycle.bin", "system volume information", "windows",
    "program files", "program files (x86)", "programdata",
    "appdata", ".gemini", "deidentifier", "identified",
}

os.makedirs(PHI_THUMB_DIR, exist_ok=True)

# Global scanner process handle
PHI_SCANNER_PROCESS = None
PHI_INGEST_THREAD   = None


# ---------------------------------------------------------------------------
# OpenSearch client
# ---------------------------------------------------------------------------

def get_client():
    return OpenSearch(
        hosts=[{"host": OPENSEARCH_HOST, "port": OPENSEARCH_PORT}],
        use_ssl=False,
        timeout=15,
        max_retries=3,
        retry_on_timeout=True,
    )


def ensure_phi_index():
    client = get_client()
    if client.indices.exists(index=PHI_INDEX_NAME):
        return
    client.indices.create(
        index=PHI_INDEX_NAME,
        body={
            "settings": {"number_of_shards": 1, "number_of_replicas": 0},
            "mappings": {
                "properties": {
                    "file_path":     {"type": "keyword"},
                    "file_name":     {"type": "text",
                                      "fields": {"keyword": {"type": "keyword"}}},
                    "file_type":     {"type": "keyword"},
                    "size_bytes":    {"type": "long"},
                    "last_modified": {"type": "date", "ignore_malformed": True},
                    "scan_date":     {"type": "date"},
                    "risk":          {"type": "keyword"},
                    "findings":      {"type": "text"},
                    "finding_keys":  {"type": "keyword"},
                    "status":        {"type": "keyword"},
                }
            },
        },
    )


def get_phi_count():
    try:
        client = get_client()
        if client.indices.exists(index=PHI_INDEX_NAME):
            return client.count(index=PHI_INDEX_NAME).get("count", 0)
    except Exception:
        pass
    return 0


# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

def load_phi_config():
    if os.path.exists(PHI_CONFIG_FILE):
        try:
            with open(PHI_CONFIG_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"selected_directories": [], "min_risk": "LOW", "workers": 4}


def save_phi_config(cfg):
    with open(PHI_CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2)


def load_phi_progress():
    if os.path.exists(PHI_PROGRESS_FILE):
        try:
            with open(PHI_PROGRESS_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {"is_running": False, "scanned": 0, "flagged": 0,
            "status_message": "", "ingested": 0}


def save_phi_progress(data):
    try:
        with open(PHI_PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Directory helpers
# ---------------------------------------------------------------------------

def get_available_drives():
    drives = []
    for letter in string.ascii_uppercase:
        drive = f"{letter}:\\"
        if os.path.exists(drive):
            drives.append(drive)
    return drives


def list_subdirectories(parent_path):
    subdirs = []
    if not parent_path or not os.path.exists(parent_path):
        return subdirs
    try:
        with os.scandir(parent_path) as entries:
            for entry in entries:
                if entry.is_dir(follow_symlinks=False):
                    if entry.name.lower() not in HIDDEN_DIRS and not entry.name.startswith("."):
                        subdirs.append({"name": entry.name, "path": entry.path})
    except Exception:
        pass
    subdirs.sort(key=lambda x: x["name"].lower())
    return subdirs


# ---------------------------------------------------------------------------
# PHI Thumbnail Generator
# ---------------------------------------------------------------------------

RISK_PALETTE = {
    "HIGH":   {"header": "#c0392b", "bg": "#fff5f5", "badge_bg": "#f8d7da", "badge_fg": "#721c24"},
    "MEDIUM": {"header": "#d35400", "bg": "#fff8f0", "badge_bg": "#ffeeba", "badge_fg": "#856404"},
    "LOW":    {"header": "#b7950b", "bg": "#fffdf0", "badge_bg": "#fff3cd", "badge_fg": "#664d03"},
    "NONE":   {"header": "#6c757d", "bg": "#f8f9fa", "badge_bg": "#e9ecef", "badge_fg": "#495057"},
}


def get_phi_thumb_hash(file_path):
    norm = os.path.normpath(os.path.abspath(file_path)).lower()
    return hashlib.md5(norm.encode("utf-8")).hexdigest()


def generate_phi_thumbnail(file_path, risk, findings_str, cache_path):
    palette      = RISK_PALETTE.get(risk, RISK_PALETTE["NONE"])
    header_color = palette["header"]
    bg_color     = palette["bg"]
    badge_bg     = palette["badge_bg"]
    badge_fg     = palette["badge_fg"]

    img  = Image.new("RGB", (600, 720), bg_color)
    draw = ImageDraw.Draw(img)

    try:
        font_hdr   = ImageFont.truetype("arialbd.ttf", 13)
        font_title = ImageFont.truetype("arialbd.ttf", 22)
        font_md    = ImageFont.truetype("arialbd.ttf", 12)
        font_sm    = ImageFont.truetype("arial.ttf",   11)
    except Exception:
        font_hdr = font_title = font_md = font_sm = ImageFont.load_default()

    draw.rectangle([0, 0, 600, 44], fill=header_color)
    draw.text((14, 13), f"PHI SCAN  \u2022  {risk} RISK", fill="#ffffff", font=font_hdr)

    draw.rectangle([14, 58, 586, 130], fill=badge_bg, outline=header_color, width=2)
    badge_text = {"HIGH": "\u26a0 HIGH RISK", "MEDIUM": "\u26a0 MEDIUM RISK",
                  "LOW":  "\u26a0 LOW RISK"}.get(risk, risk)
    draw.text((24, 72), badge_text, fill=badge_fg, font=font_title)

    file_name = os.path.basename(file_path)
    file_ext  = os.path.splitext(file_name)[1].upper().lstrip(".")
    draw.text((14, 142), "File:", fill="#6c757d", font=font_sm)
    draw.text((14, 158), file_name[:54], fill="#212529", font=font_md)
    draw.text((14, 180), os.path.dirname(file_path)[:72], fill="#6c757d", font=font_sm)

    ext_x = 586 - len(file_ext) * 8 - 16
    draw.rectangle([ext_x, 56, 586, 78], fill=header_color)
    draw.text((ext_x + 8, 62), file_ext, fill="#ffffff", font=font_sm)

    draw.rectangle([14, 208, 586, 210], fill="#dee2e6")
    draw.text((14, 218), "FINDING CATEGORIES (no matched values stored):", fill="#6c757d", font=font_sm)

    y = 240
    cats = [c.strip() for c in findings_str.split(";") if c.strip()]
    for i, cat in enumerate(cats):
        if y > 675:
            draw.text((28, y), f"... and {len(cats) - i} more", fill="#6c757d", font=font_sm)
            break
        draw.rectangle([14, y + 2, 18, y + 12], fill=header_color)
        draw.text((26, y), cat[:72], fill="#343a40", font=font_sm)
        y += 22

    if not cats:
        draw.text((26, y), "(filename match only)", fill="#6c757d", font=font_sm)

    draw.rectangle([0, 696, 600, 720], fill="#f1f3f5")
    draw.text((14, 704), "PHI Scanner \u2014 categories only, no matched values stored",
              fill="#6c757d", font=font_sm)
    draw.rectangle([0, 0, 599, 719], outline=header_color, width=2)

    img.save(cache_path, "JPEG", quality=92)


def get_phi_thumbnail_bytes(file_path, risk, findings_str):
    thumb_hash = get_phi_thumb_hash(file_path)
    cache_path = os.path.join(PHI_THUMB_DIR, f"{thumb_hash}.jpg")
    if not os.path.exists(cache_path):
        try:
            generate_phi_thumbnail(file_path, risk, findings_str, cache_path)
        except Exception:
            pass
    if os.path.exists(cache_path):
        try:
            with open(cache_path, "rb") as f:
                return f.read(), "image/jpeg"
        except Exception:
            pass
    return None, None


# ---------------------------------------------------------------------------
# Scanner process management
# ---------------------------------------------------------------------------

def is_scanner_running():
    global PHI_SCANNER_PROCESS
    if PHI_SCANNER_PROCESS is not None and PHI_SCANNER_PROCESS.poll() is None:
        return True
    for p in psutil.process_iter(["name", "cmdline"]):
        try:
            if p.info["name"] and "python" in p.info["name"].lower():
                if any("phi_scanner.py" in str(a) for a in (p.info.get("cmdline") or [])):
                    return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False


def stop_scanner():
    global PHI_SCANNER_PROCESS
    if PHI_SCANNER_PROCESS and PHI_SCANNER_PROCESS.poll() is None:
        try:
            PHI_SCANNER_PROCESS.terminate()
            PHI_SCANNER_PROCESS.kill()
        except Exception:
            pass
    PHI_SCANNER_PROCESS = None
    prog = load_phi_progress()
    prog["is_running"] = False
    prog["status_message"] = "Scan stopped."
    save_phi_progress(prog)


def ingest_phi_report(report_path):
    if not os.path.exists(report_path):
        prog = load_phi_progress()
        prog["is_running"]     = False
        prog["status_message"] = "Scan complete \u2014 no report file found (0 files matched)."
        save_phi_progress(prog)
        return
    try:
        ensure_phi_index()
        client = get_client()
    except Exception as e:
        prog = load_phi_progress()
        prog["status_message"] = f"OpenSearch connection failed: {e}"
        save_phi_progress(prog)
        return

    from opensearchpy import helpers
    actions   = []
    row_count = 0

    try:
        with open(report_path, "r", encoding="utf-8-sig", newline="") as f:
            reader = csv.DictReader(f)
            for row in reader:
                file_path  = row.get("file", "")
                risk       = row.get("risk", "NONE")
                findings   = row.get("findings", "")
                status     = row.get("status", "OK")
                ext        = row.get("extension", "")
                size_bytes = row.get("size_bytes", "0")

                if not file_path:
                    continue

                finding_keys = [
                    re.sub(r"\s*\(\d+\)\s*$", "", p.strip())
                    for p in findings.split(";") if p.strip()
                ]

                last_modified = None
                try:
                    mtime = os.path.getmtime(file_path)
                    last_modified = datetime.fromtimestamp(mtime).isoformat()
                except Exception:
                    pass

                doc_id = hashlib.md5(
                    os.path.normpath(file_path).lower().encode("utf-8")
                ).hexdigest()

                actions.append({
                    "_index": PHI_INDEX_NAME,
                    "_id":    doc_id,
                    "_source": {
                        "file_path":     file_path,
                        "file_name":     os.path.basename(file_path),
                        "file_type":     ext.lstrip(".").lower(),
                        "size_bytes":    int(size_bytes) if str(size_bytes).isdigit() else 0,
                        "last_modified": last_modified,
                        "scan_date":     datetime.now().isoformat(),
                        "risk":          risk,
                        "findings":      findings,
                        "finding_keys":  finding_keys,
                        "status":        status,
                    },
                })

                if risk != "NONE" and os.path.exists(file_path):
                    cache_path = os.path.join(PHI_THUMB_DIR, f"{get_phi_thumb_hash(file_path)}.jpg")
                    if not os.path.exists(cache_path):
                        try:
                            generate_phi_thumbnail(file_path, risk, findings, cache_path)
                        except Exception:
                            pass

                row_count += 1

                if len(actions) >= 200:
                    try:
                        helpers.bulk(client, actions, raise_on_error=False)
                    except Exception:
                        pass
                    actions = []
                    prog = load_phi_progress()
                    prog["ingested"] = row_count
                    prog["status_message"] = f"Ingesting... {row_count:,} rows processed"
                    save_phi_progress(prog)

        if actions:
            try:
                helpers.bulk(client, actions, raise_on_error=False)
            except Exception:
                pass
        client.indices.refresh(index=PHI_INDEX_NAME)

    except Exception as e:
        prog = load_phi_progress()
        prog["status_message"] = f"Ingest error: {e}"
        save_phi_progress(prog)
        return

    prog = load_phi_progress()
    prog["ingested"]       = row_count
    prog["is_running"]     = False
    prog["status_message"] = f"\u2705 Scan complete \u2014 {row_count:,} results ingested"
    save_phi_progress(prog)


def run_phi_scan(directories, min_risk="LOW", workers=4):
    global PHI_SCANNER_PROCESS

    save_phi_progress({
        "is_running": True, "scanned": 0, "flagged": 0, "ingested": 0,
        "status_message": "\u26a1 PHI scan starting...", "timestamp": time.time(),
    })

    try:
        for directory in directories:
            if not os.path.exists(directory):
                continue

            prog = load_phi_progress()
            prog["status_message"] = f"\U0001f50d Scanning: {directory}"
            save_phi_progress(prog)

            cmd = [
                sys.executable, PHI_SCANNER_SCRIPT,
                directory,
                "--report",    PHI_REPORT_FILE,
                "--min-risk",  min_risk,
                "--workers",   str(workers),
            ]

            PHI_SCANNER_PROCESS = subprocess.Popen(
                cmd, cwd=BASE_DIR,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace",
            )

            for line in PHI_SCANNER_PROCESS.stdout:
                line = line.strip()
                if not line:
                    continue
                # Structured FILE: line from phi_scanner.py
                if line.startswith("FILE:"):
                    parts = {}
                    for chunk in line.split("\t"):
                        if ":" in chunk:
                            k, _, v = chunk.partition(":")
                            parts[k] = v
                    current_file = parts.get("FILE", "")
                    current_dir  = os.path.dirname(current_file)
                    prog_str     = parts.get("PROGRESS", "0/0")
                    scanned_n, _, total_n = prog_str.partition("/")
                    prog = load_phi_progress()
                    prog["scanned"]      = int(scanned_n) if scanned_n.isdigit() else prog.get("scanned", 0)
                    prog["total"]        = int(total_n)   if total_n.isdigit()   else prog.get("total", 0)
                    prog["flagged"]      = int(parts.get("FLAGGED", "0"))
                    prog["current_file"] = current_file
                    prog["current_dir"]  = current_dir
                    prog["status_message"] = f"\U0001f50d Scanning: {current_dir}"
                    save_phi_progress(prog)
                elif line and not line.startswith("="):
                    # Collect file totals line: "X,XXX files found"
                    m_total = re.search(r"([\d,]+)\s+files found", line)
                    if m_total:
                        prog = load_phi_progress()
                        prog["total"] = int(m_total.group(1).replace(",", ""))
                        prog["status_message"] = f"Found {prog['total']:,} files to scan..."
                        save_phi_progress(prog)

            PHI_SCANNER_PROCESS.wait()

        prog = load_phi_progress()
        prog["status_message"] = "\u23f3 Ingesting results into OpenSearch..."
        save_phi_progress(prog)

        ingest_phi_report(PHI_REPORT_FILE)

    except Exception as e:
        prog = load_phi_progress()
        prog["is_running"]     = False
        prog["status_message"] = f"Scan error: {e}"
        save_phi_progress(prog)
    finally:
        PHI_SCANNER_PROCESS = None


# ---------------------------------------------------------------------------
# Query builder
# ---------------------------------------------------------------------------

def build_phi_query(query_str, risk_filter, sort_by, page):
    must = []
    if risk_filter in ("HIGH", "MEDIUM", "LOW"):
        must.append({"term": {"risk": risk_filter}})
    if query_str:
        must.append({"multi_match": {
            "query":  query_str,
            "fields": ["file_name^3", "findings^2", "file_path", "finding_keys^2"],
            "type":   "best_fields",
        }})

    body = {
        "query": {"bool": {"must": must}} if must else {"match_all": {}},
        "from":  (page - 1) * PAGE_SIZE,
        "size":  PAGE_SIZE,
    }

    if sort_by == "risk_desc":
        body["sort"] = [{"_script": {
            "type":   "number",
            "script": {"source":
                "if(doc['risk'].value=='HIGH')return 3;"
                "if(doc['risk'].value=='MEDIUM')return 2;"
                "if(doc['risk'].value=='LOW')return 1;return 0;"},
            "order": "desc",
        }}, {"scan_date": {"order": "desc"}}]
    elif sort_by == "date_desc":
        body["sort"] = [{"last_modified": {"order": "desc"}}]
    elif sort_by == "date_asc":
        body["sort"] = [{"last_modified": {"order": "asc"}}]
    elif sort_by == "name_asc":
        body["sort"] = [{"file_name.keyword": {"order": "asc"}}]
    elif sort_by == "size_desc":
        body["sort"] = [{"size_bytes": {"order": "desc"}}]
    else:
        body["sort"] = [{"scan_date": {"order": "desc"}}]

    return body


# ---------------------------------------------------------------------------
# HTML Template
# ---------------------------------------------------------------------------

HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>PHI Scanner</title>
    <style>
        body{font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;background:#1a1a2e;margin:0;padding:20px;color:#e0e0e0}
        .container{max-width:1350px;margin:0 auto}
        .header{display:flex;justify-content:space-between;align-items:center;margin-bottom:20px;border-bottom:2px solid #2d2d4e;padding-bottom:15px}
        .header-title h2{margin:0;font-size:24px;color:#e74c3c}
        .header-title p{margin:4px 0 0 0;color:#9e9e9e;font-size:14px}
        .header-buttons{display:flex;gap:10px;align-items:center}
        .count-badge{background:#2d1f1f;color:#e74c3c;border:1px solid #5c2626;padding:9px 15px;border-radius:6px;font-size:14px;font-weight:bold}
        .btn-back{background:#2d2d4e;color:#a0a0c0;padding:10px 16px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:600;text-decoration:none;display:inline-flex;align-items:center;gap:6px}
        .btn-back:hover{background:#3d3d6e;color:white}
        .btn-settings{background:#6c757d;color:white;padding:10px 18px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:600}
        .btn-scan{background:#c0392b;color:white;padding:10px 18px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:600;transition:background 0.3s}
        .btn-scan.running{background:#2ecc71}
        .search-box{display:flex;gap:10px;margin-bottom:20px;align-items:center}
        input[type=text]{flex:1;padding:14px;font-size:16px;border:2px solid #3d3d5e;border-radius:6px;background:#16213e;color:#e0e0e0}
        input[type=text]:focus{outline:none;border-color:#e74c3c}
        select{padding:14px;font-size:15px;border:2px solid #3d3d5e;border-radius:6px;background:#16213e;color:#e0e0e0;cursor:pointer}
        .btn-search{padding:14px 28px;font-size:16px;background:#c0392b;color:white;border:none;border-radius:6px;cursor:pointer;font-weight:bold}
        .stats-bar{display:flex;justify-content:space-between;align-items:center;margin-bottom:15px}
        .stats{color:#9e9e9e;font-weight:500}
        .scan-badge{background:#2d1f0e;color:#e67e22;border:1px solid #6e3d1a;padding:6px 14px;border-radius:12px;font-size:13px;font-weight:bold;display:none;animation:pulse 1.5s infinite}
        @keyframes pulse{0%{opacity:1}50%{opacity:.5}100%{opacity:1}}
        .result-card{background:#16213e;border-radius:8px;padding:18px;margin-bottom:15px;box-shadow:0 2px 8px rgba(0,0,0,.3);display:flex;gap:24px;align-items:flex-start;border-left:4px solid #555}
        .result-card.risk-high{border-left-color:#c0392b}
        .result-card.risk-medium{border-left-color:#d35400}
        .result-card.risk-low{border-left-color:#b7950b}
        .card-left{flex:1;min-width:0}
        .card-right{flex-shrink:0;width:380px}
        .thumb-preview{width:100%;border-radius:6px;border:1px solid #3d3d5e;box-shadow:0 2px 8px rgba(0,0,0,.4)}
        .thumb-placeholder{width:100%;height:300px;background:#1a2744;border-radius:6px;border:1px dashed #3d3d5e;display:flex;align-items:center;justify-content:center;color:#6c757d;font-size:14px}
        .result-header{display:flex;justify-content:space-between;align-items:center;margin-bottom:6px}
        .file-title{font-weight:bold;font-size:17px;color:#e0e0e0}
        .card-actions{display:flex;align-items:center;gap:8px;flex-wrap:wrap}
        .risk-badge{padding:4px 12px;border-radius:12px;font-size:13px;font-weight:bold;display:inline-block;margin-bottom:6px}
        .risk-high{background:#f8d7da;color:#721c24}
        .risk-medium{background:#ffeeba;color:#856404}
        .risk-low{background:#fff3cd;color:#664d03}
        .btn-action{text-decoration:none;padding:6px 12px;border-radius:5px;font-size:12px;font-weight:bold;cursor:pointer;display:inline-flex;align-items:center;gap:4px;border:none}
        .btn-open-file{background:#2980b9;color:white}
        .btn-open-explorer{background:#17a2b8;color:white}
        .btn-open-folder{background:#6f42c1;color:white}
        .file-path{color:#7f8c8d;font-size:12px;margin:4px 0 8px 0;word-break:break-all}
        .file-meta{display:flex;gap:12px;font-size:12px;color:#7f8c8d;margin-bottom:8px;flex-wrap:wrap}
        .findings-list{background:#0d1b2a;border-left:3px solid #c0392b;padding:10px 14px;border-radius:4px;font-size:13px;color:#bdc3c7;line-height:1.8}
        .finding-tag{display:inline-block;background:#2d1f1f;color:#e74c3c;border:1px solid #5c2626;padding:2px 8px;border-radius:10px;font-size:11px;margin:2px 3px 2px 0}
        .pagination-bar{display:flex;justify-content:space-between;align-items:center;background:#16213e;border:1px solid #2d2d4e;border-radius:8px;padding:12px 20px;margin:15px 0}
        .page-info{font-weight:bold;color:#bdc3c7;font-size:15px}
        .btn-page{text-decoration:none;padding:9px 18px;border-radius:6px;background:#c0392b;color:white;font-weight:bold;font-size:14px}
        .btn-page.disabled{background:#2d2d4e;color:#6c757d;pointer-events:none}
        .toast{position:fixed;bottom:20px;right:20px;background:#27ae60;color:white;padding:14px 28px;border-radius:6px;font-size:15px;font-weight:bold;box-shadow:0 4px 12px rgba(0,0,0,.4);display:none;z-index:2000}
        .modal{display:none;position:fixed;z-index:1000;left:0;top:0;width:100%;height:100%;background:rgba(0,0,0,.7)}
        .modal-content{background:#1e1e3e;margin:40px auto;padding:25px;border-radius:8px;width:650px;max-width:90%;max-height:80vh;display:flex;flex-direction:column;border:1px solid #3d3d5e}
        .modal-header{display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid #3d3d5e;padding-bottom:12px;margin-bottom:15px}
        .modal-header h3{margin:0;color:#e0e0e0}
        .close{cursor:pointer;font-size:24px;color:#7f8c8d}
        .tree-container{flex:1;overflow-y:auto;border:1px solid #3d3d5e;border-radius:6px;padding:12px;font-size:14px;background:#16213e;color:#bdc3c7}
        .tree-item{margin:6px 0}
        .tree-toggle{cursor:pointer;display:inline-block;width:22px;height:22px;line-height:22px;text-align:center;font-weight:bold;color:#e74c3c;user-select:none}
        .tree-children{margin-left:24px;display:none}
        .tree-children.open{display:block}
        .modal-footer{display:flex;justify-content:space-between;align-items:center;margin-top:15px;border-top:1px solid #3d3d5e;padding-top:15px}
        .btn-save{background:#c0392b;color:white;padding:10px 20px;border:none;border-radius:6px;cursor:pointer;font-weight:bold}
        .config-row{margin:10px 0;display:flex;align-items:center;gap:10px;flex-wrap:wrap;color:#bdc3c7}
        .config-row select,.config-row input[type=number]{padding:6px 10px;background:#16213e;color:#e0e0e0;border:1px solid #3d3d5e;border-radius:4px}
    </style>
</head>
<body>
<div class="container">
    <div class="header">
        <div class="header-title">
            <h2>&#128274; PHI Scanner</h2>
            <p>Scans documents locally for PHI indicators &mdash; no network, no cloud, no PHI values stored</p>
        </div>
        <div class="header-buttons">
            <div class="count-badge" id="phiCountBadge">&#128274; PHI Results: ...</div>
            <a href="SEARCH_APP_URL_PLACEHOLDER" class="btn-back">&#8592; Document Search</a>
            <button type="button" class="btn-settings" onclick="openConfigModal()">&#128194; Scan Directories</button>
            <button type="button" class="btn-scan" id="scanBtn" onclick="toggleScan()">&#128274; Start PHI Scan</button>
        </div>
    </div>
    <form class="search-box" method="GET" action="/">
        <input type="text" name="q" value="{QUERY}" placeholder="Search file names, finding categories (e.g. MRN, HIGH, patient)..." autofocus>
        <select name="risk" onchange="this.form.submit()">
            <option value=""       {RISK_ALL}>All Risk Levels</option>
            <option value="HIGH"   {RISK_HIGH}>&#9888; HIGH Only</option>
            <option value="MEDIUM" {RISK_MEDIUM}>&#9888; MEDIUM Only</option>
            <option value="LOW"    {RISK_LOW}>&#9888; LOW Only</option>
        </select>
        <select name="sort" onchange="this.form.submit()">
            <option value="risk_desc" {SORT_RISK}>Risk: Highest First</option>
            <option value="date_desc" {SORT_DATE_DESC}>Date: Newest First</option>
            <option value="date_asc"  {SORT_DATE_ASC}>Date: Oldest First</option>
            <option value="name_asc"  {SORT_NAME_ASC}>File Name (A-Z)</option>
            <option value="size_desc" {SORT_SIZE_DESC}>Size: Largest First</option>
        </select>
        <button type="submit" class="btn-search">Search</button>
    </form>
    <div class="stats-bar">
        <div class="stats">{STATS}</div>
        <div class="scan-badge" id="scanBadge">&#9889; PHI scan running...</div>
    </div>

    <!-- Live scan progress panel — hidden when not scanning -->
    <div id="progressPanel" style="display:none; background:#0d1b2a; border:1px solid #c0392b; border-radius:8px; padding:16px 20px; margin-bottom:16px; font-family:monospace; font-size:13px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
            <span style="color:#e74c3c; font-weight:bold; font-size:14px;">&#9889; PHI Scan In Progress</span>
            <span id="progressFraction" style="color:#bdc3c7;">0 / 0 files</span>
        </div>
        <div style="background:#1a2744; border-radius:4px; height:8px; margin-bottom:12px; overflow:hidden;">
            <div id="progressBar" style="background:linear-gradient(90deg,#c0392b,#e74c3c); height:8px; width:0%; transition:width 0.4s ease; border-radius:4px;"></div>
        </div>
        <div style="display:flex; gap:24px; margin-bottom:10px; flex-wrap:wrap;">
            <span>&#128269; <strong id="progScanned" style="color:#e0e0e0;">0</strong> scanned</span>
            <span>&#9888; <strong id="progFlagged" style="color:#e74c3c;">0</strong> flagged</span>
            <span>&#128274; <strong id="progPct" style="color:#bdc3c7;">0%</strong> complete</span>
        </div>
        <div style="border-top:1px solid #2d2d4e; padding-top:10px;">
            <div style="color:#7f8c8d; font-size:11px; margin-bottom:3px;">Current folder:</div>
            <div id="progCurrentDir"  style="color:#3498db; word-break:break-all; margin-bottom:6px;">&mdash;</div>
            <div style="color:#7f8c8d; font-size:11px; margin-bottom:3px;">Current file:</div>
            <div id="progCurrentFile" style="color:#bdc3c7; word-break:break-all;">&mdash;</div>
        </div>
    </div>
    {PAGINATION_TOP}
    {RESULTS}
    {PAGINATION_BOTTOM}
</div>
<div class="toast" id="toastMsg"></div>
<div id="configModal" class="modal">
    <div class="modal-content">
        <div class="modal-header">
            <h3>&#128194; PHI Scan Configuration</h3>
            <span class="close" onclick="closeConfigModal()">&times;</span>
        </div>
        <p style="font-size:13px;color:#7f8c8d;margin-top:0">Select directories to scan for PHI:</p>
        <div class="tree-container" id="treeContainer">Loading...</div>
        <div class="config-row" style="margin-top:12px;border-top:1px solid #3d3d5e;padding-top:12px">
            <label>Min Risk:</label>
            <select id="minRiskSelect">
                <option value="LOW">LOW (all)</option>
                <option value="MEDIUM">MEDIUM+</option>
                <option value="HIGH">HIGH only</option>
            </select>
            <label>Workers:</label>
            <input type="number" id="workersInput" value="4" min="1" max="16" style="width:60px">
        </div>
        <div class="modal-footer">
            <span id="configStatus" style="font-size:13px;font-weight:bold;color:#e74c3c"></span>
            <button type="button" class="btn-save" onclick="saveConfig()">Save Configuration</button>
        </div>
    </div>
</div>
<script>
let selectedPaths=new Set(),isScanning=false;
document.addEventListener('DOMContentLoaded',()=>{checkStatus();setInterval(checkStatus,2000)});
function showToast(m,e=false){const t=document.getElementById('toastMsg');t.innerText=m;t.style.backgroundColor=e?'#c0392b':'#27ae60';t.style.display='block';setTimeout(()=>t.style.display='none',4000)}
async function checkStatus(){
  try{
    const r=await fetch('/api/phi_status'),d=await r.json();
    const btn=document.getElementById('scanBtn');
    const badge=document.getElementById('scanBadge');
    const cb=document.getElementById('phiCountBadge');
    const panel=document.getElementById('progressPanel');
    isScanning=d.is_running;
    if(d.phi_count!==undefined) cb.innerText='\uD83D\uDD12 PHI Results: '+d.phi_count.toLocaleString();
    if(btn){
      if(isScanning){
        btn.className='btn-scan running';
        btn.innerText='\u23F9 Stop Scan';
        if(badge) badge.style.display='none';
        if(panel){
          panel.style.display='block';
          const sc=d.scanned||0,tot=d.total||0,fl=d.flagged||0;
          const pct=tot>0?Math.round(sc/tot*100):0;
          document.getElementById('progressBar').style.width=pct+'%';
          document.getElementById('progressFraction').innerText=sc.toLocaleString()+' / '+(tot>0?tot.toLocaleString():'?')+' files';
          document.getElementById('progScanned').innerText=sc.toLocaleString();
          document.getElementById('progFlagged').innerText=fl.toLocaleString();
          document.getElementById('progPct').innerText=pct+'%';
          const curFile=d.current_file||'';
          const curDir=d.current_dir||'';
          const fname=curFile?curFile.replace(/.*[\\/]/,''):'';
          document.getElementById('progCurrentDir').innerText=curDir||'\u2014';
          document.getElementById('progCurrentFile').innerText=fname||'\u2014';
        }
      } else {
        btn.className='btn-scan';
        btn.innerText='\uD83D\uDD12 Start PHI Scan';
        if(badge) badge.style.display='none';
        if(panel) panel.style.display='none';
        if(d.status_message&&d.status_message.includes('\u2705')&&!window._reloaded){
          window._reloaded=true;
          setTimeout(()=>window.location.reload(),1500);
        }
      }
    }
  } catch(e){}
}
async function toggleScan(){if(isScanning){await fetch('/api/stop_phi_scan',{method:'POST'});showToast('PHI scan stopped.');checkStatus()}else{const r=await fetch('/api/phi_config'),c=await r.json();if(!c.selected_directories||c.selected_directories.length===0){showToast('Configure scan directories first.',true);openConfigModal();return}await fetch('/api/start_phi_scan',{method:'POST'});window._reloaded=false;showToast('PHI scan started...');checkStatus()}}
async function openConfigModal(){document.getElementById('configModal').style.display='block';const r=await fetch('/api/phi_config'),c=await r.json();selectedPaths=new Set(c.selected_directories||[]);document.getElementById('minRiskSelect').value=c.min_risk||'LOW';document.getElementById('workersInput').value=c.workers||4;loadDriveTree()}
function closeConfigModal(){document.getElementById('configModal').style.display='none'}
async function saveConfig(){const dirs=Array.from(selectedPaths),mr=document.getElementById('minRiskSelect').value,w=parseInt(document.getElementById('workersInput').value)||4;await fetch('/api/phi_config',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({selected_directories:dirs,min_risk:mr,workers:w})});document.getElementById('configStatus').innerText='\\u2713 Saved!';setTimeout(()=>{closeConfigModal();document.getElementById('configStatus').innerText=''},1000)}
async function loadDriveTree(){const c=document.getElementById('treeContainer');c.innerHTML='';const r=await fetch('/api/drives'),drives=await r.json();for(const d of drives)c.appendChild(buildItem(d,'\\uD83D\\uDCBB '+d,true))}
function buildItem(path,label,bold=false){const item=document.createElement('div');item.className='tree-item';const tg=document.createElement('span');tg.className='tree-toggle';tg.innerText='\\u25B6';tg.setAttribute('data-path',path);tg.onclick=function(){toggleF(this)};const cb=document.createElement('input');cb.type='checkbox';cb.value=path;cb.checked=selectedPaths.has(path);cb.onchange=function(){this.checked?selectedPaths.add(this.value):selectedPaths.delete(this.value)};const lb=document.createElement(bold?'strong':'span');lb.innerHTML=' '+label;const ch=document.createElement('div');ch.className='tree-children';item.appendChild(tg);item.appendChild(cb);item.appendChild(lb);item.appendChild(ch);return item}
async function toggleF(el){const path=el.getAttribute('data-path'),ch=el.parentElement.querySelector('.tree-children');if(el.innerText==='\\u25B6'){el.innerText='\\u25BC';ch.classList.add('open');if(!ch.children.length){const r=await fetch('/api/ls?path='+encodeURIComponent(path)),subs=await r.json();if(!subs.length)ch.innerHTML='<div style="margin-left:20px;color:#555;font-style:italic">(no subfolders)</div>';else for(const s of subs)ch.appendChild(buildItem(s.path,'\\uD83D\\uDCC1 '+s.name))}}else{el.innerText='\\u25B6';ch.classList.remove('open')}}
async function handleOpenFile(p){try{const r=await fetch('/api/open_file?path='+encodeURIComponent(p)),d=await r.json();showToast(d.status==='ok'?'Opening file...':d.message,d.status==='error')}catch(e){}}
async function handleOpenExplorer(p){try{const r=await fetch('/api/open_folder?explorer=1&path='+encodeURIComponent(p)),d=await r.json();showToast(d.status==='ok'?'Opening folder...':d.message,d.status==='error')}catch(e){}}
async function handleOpenFolder(p){try{const r=await fetch('/api/open_folder?path='+encodeURIComponent(p)),d=await r.json();showToast(d.status==='ok'?'Opening folder...':d.message,d.status==='error')}catch(e){}}
</script>
</body>
</html>
"""


# ---------------------------------------------------------------------------
# HTTP Handler
# ---------------------------------------------------------------------------

def format_size(size_bytes):
    try:
        b = int(size_bytes)
        if b < 1024:     return f"{b} B"
        if b < 1048576:  return f"{b/1024:.1f} KB"
        return f"{b/1048576:.1f} MB"
    except Exception:
        return ""


class PhiHandler(SimpleHTTPRequestHandler):

    def log_message(self, fmt, *args):
        pass

    def send_json(self, data, status=200):
        try:
            self.send_response(status)
            self.send_header("Content-Type", "application/json; charset=utf-8")
            self.end_headers()
            self.wfile.write(json.dumps(data).encode("utf-8"))
        except (ConnectionAbortedError, ConnectionResetError, BrokenPipeError, OSError):
            pass

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        if parsed.path == "/api/phi_status":
            prog = load_phi_progress()
            self.send_json({
                "is_running":     prog.get("is_running", False),
                "scanned":        prog.get("scanned", 0),
                "total":          prog.get("total", 0),
                "flagged":        prog.get("flagged", 0),
                "ingested":       prog.get("ingested", 0),
                "current_file":   prog.get("current_file", ""),
                "current_dir":    prog.get("current_dir", ""),
                "status_message": prog.get("status_message", ""),
                "phi_count":      get_phi_count(),
            })
            return

        if parsed.path == "/api/phi_config":
            self.send_json(load_phi_config())
            return

        if parsed.path == "/api/phi_thumbnail":
            file_path    = params.get("path",     [""])[0]
            risk         = params.get("risk",     ["NONE"])[0]
            findings_str = params.get("findings", [""])[0]
            img_data, ctype = get_phi_thumbnail_bytes(file_path, risk, findings_str)
            if img_data:
                try:
                    self.send_response(200)
                    self.send_header("Content-Type", ctype)
                    self.send_header("Cache-Control", "no-cache")
                    self.end_headers()
                    self.wfile.write(img_data)
                except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
                    pass
                return
            self.send_response(404)
            self.end_headers()
            return

        if parsed.path == "/api/drives":
            self.send_json(get_available_drives())
            return

        if parsed.path == "/api/ls":
            self.send_json(list_subdirectories(params.get("path", [""])[0]))
            return

        if parsed.path == "/api/open_file":
            fp = params.get("path", [""])[0]
            if fp:
                norm = os.path.normpath(fp)
                if os.path.exists(norm):
                    try:
                        os.startfile(norm)
                        self.send_json({"status": "ok", "message": "File opened."})
                    except Exception as e:
                        self.send_json({"status": "error", "message": str(e)}, 500)
                else:
                    parent = os.path.dirname(norm)
                    subprocess.Popen(["explorer", parent])
                    self.send_json({"status": "warning", "message": f"File not found. Opened: {parent}"})
            else:
                self.send_json({"status": "error", "message": "No path."}, 400)
            return

        if parsed.path == "/api/open_folder":
            fp = params.get("path", [""])[0]
            force_explorer = params.get("explorer", ["0"])[0] == "1"
            if fp:
                norm = os.path.normpath(fp)
                target = norm if os.path.exists(norm) else os.path.dirname(norm)
                try:
                    if force_explorer:
                        subprocess.Popen(["explorer", "/select,", target] if os.path.isfile(target) else ["explorer", target])
                    elif os.path.exists(DOPUS_RT):
                        parent = os.path.dirname(target) if os.path.isfile(target) else target
                        fname  = os.path.basename(target) if os.path.isfile(target) else ""
                        args   = [DOPUS_RT, "/cmd", "Go", parent, "NEW"]
                        if fname:
                            args.append(f"SELECT={fname}")
                        subprocess.Popen(args)
                    else:
                        subprocess.Popen(["explorer", target])
                    self.send_json({"status": "ok", "message": "Folder opened."})
                except Exception as e:
                    self.send_json({"status": "error", "message": str(e)}, 500)
            else:
                self.send_json({"status": "error", "message": "No path."}, 400)
            return

        # ---- Main page ----
        query_str   = params.get("q",    [""])[0].strip()
        risk_filter = params.get("risk", [""])[0].strip().upper()
        sort_by     = params.get("sort", ["risk_desc"])[0].strip()
        try:
            page = max(1, int(params.get("page", ["1"])[0]))
        except Exception:
            page = 1

        stats_html = results_html = pagination_html = ""

        sort_state = {
            "SORT_RISK":      "selected" if sort_by == "risk_desc"  else "",
            "SORT_DATE_DESC": "selected" if sort_by == "date_desc"  else "",
            "SORT_DATE_ASC":  "selected" if sort_by == "date_asc"   else "",
            "SORT_NAME_ASC":  "selected" if sort_by == "name_asc"   else "",
            "SORT_SIZE_DESC": "selected" if sort_by == "size_desc"  else "",
            "RISK_ALL":       "selected" if not risk_filter          else "",
            "RISK_HIGH":      "selected" if risk_filter == "HIGH"    else "",
            "RISK_MEDIUM":    "selected" if risk_filter == "MEDIUM"  else "",
            "RISK_LOW":       "selected" if risk_filter == "LOW"     else "",
        }

        try:
            client   = get_client()
            es_query = build_phi_query(query_str, risk_filter, sort_by, page)
            res      = client.search(index=PHI_INDEX_NAME, body=es_query)
            hits     = res["hits"]["hits"]
            total    = res["hits"]["total"]["value"]
            took     = res["took"]

            total_pages = math.ceil(total / PAGE_SIZE) if total > 0 else 1
            start_doc   = (page - 1) * PAGE_SIZE + 1 if total > 0 else 0
            end_doc     = min(page * PAGE_SIZE, total)
            label       = f"Filtered: {risk_filter} | " if risk_filter else ""
            stats_html  = f"{label}Found {total:,} result(s) in {took} ms"

            if total > PAGE_SIZE:
                qe = urllib.parse.quote(query_str)
                re_ = urllib.parse.quote(risk_filter)
                pd = "disabled" if page <= 1 else ""
                nd = "disabled" if end_doc >= total else ""
                nc = min(PAGE_SIZE, total - end_doc)
                pagination_html = (
                    f'<div class="pagination-bar">'
                    f'<a href="/?q={qe}&risk={re_}&sort={sort_by}&page={page-1}" class="btn-page {pd}">&larr; Previous</a>'
                    f'<span class="page-info">Showing {start_doc:,}&ndash;{end_doc:,} of {total:,} (Page {page} of {total_pages})</span>'
                    f'<a href="/?q={qe}&risk={re_}&sort={sort_by}&page={page+1}" class="btn-page {nd}">Next {nc:,} &rarr;</a>'
                    f'</div>'
                )

            if not hits:
                results_html = "<div class='result-card'><div class='card-left'>No results. Run a PHI scan first or adjust filters.</div></div>"
            else:
                cards = []
                for hit in hits:
                    src      = hit["_source"]
                    fpath    = src.get("file_path", "")
                    fname    = html.escape(src.get("file_name", os.path.basename(fpath)))
                    risk     = src.get("risk", "NONE")
                    findings = src.get("findings", "")
                    status   = src.get("status", "OK")
                    ftype    = src.get("file_type", "")
                    sz       = format_size(src.get("size_bytes", 0))
                    sd       = src.get("scan_date", "")[:10]
                    mt       = src.get("last_modified", "")[:10]
                    rl       = risk.lower()

                    jp  = html.escape(fpath.replace("\\", "\\\\").replace("'", "\\'"))
                    ep  = urllib.parse.quote(fpath)
                    ef  = urllib.parse.quote(findings[:500])
                    tu  = f"/api/phi_thumbnail?path={ep}&risk={risk}&findings={ef}"

                    cats = [re.sub(r"\s*\(\d+\)\s*$", "", c.strip()) for c in findings.split(";") if c.strip()]
                    tags = " ".join(f'<span class="finding-tag">{html.escape(c)}</span>' for c in cats[:12])
                    if len(cats) > 12:
                        tags += f' <span class="finding-tag">+{len(cats)-12} more</span>'

                    sn = f'<span style="color:#e74c3c;font-size:12px"> &#9888; {html.escape(status)}</span>' if status not in ("OK","") else ""

                    cards.append(f"""
                    <div class="result-card risk-{rl}">
                        <div class="card-left">
                            <span class="risk-badge risk-{rl}">&nbsp;&#9888; {risk} RISK&nbsp;</span>
                            <div class="result-header">
                                <span class="file-title">{fname}</span>
                                <div class="card-actions">
                                    <button class="btn-action btn-open-file"     onclick="handleOpenFile('{jp}')">&#8599; Open</button>
                                    <button class="btn-action btn-open-explorer" onclick="handleOpenExplorer('{jp}')">&#128193; Explorer</button>
                                    <button class="btn-action btn-open-folder"   onclick="handleOpenFolder('{jp}')">&#128193; Opus</button>
                                </div>
                            </div>
                            <div class="file-path">&#128193; {html.escape(fpath)}</div>
                            <div class="file-meta">
                                <span><b>Type:</b> {html.escape(ftype.upper())}</span>
                                <span><b>Size:</b> {sz}</span>
                                <span><b>Modified:</b> {html.escape(mt)}</span>
                                <span><b>Scanned:</b> {html.escape(sd)}</span>
                                {sn}
                            </div>
                            <div class="findings-list">{tags or "(no categories)"}</div>
                        </div>
                        <div class="card-right">
                            <img src="{tu}" class="thumb-preview" alt="PHI Risk Card"
                                 onerror="this.parentElement.innerHTML='<div class=\\'thumb-placeholder\\'>No preview</div>'">
                        </div>
                    </div>""")
                results_html = "\n".join(cards)

        except Exception as e:
            err = str(e)
            if "index_not_found" in err or "no such index" in err.lower():
                results_html = "<div class='result-card'><div class='card-left'>No PHI scan results yet. Click <b>Start PHI Scan</b> to begin.</div></div>"
                stats_html   = "No scan results indexed yet."
            else:
                results_html = f"<div class='result-card'><div class='card-left' style='color:#e74c3c'>Search error: {html.escape(err)}</div></div>"

        out = HTML_TEMPLATE.replace("SEARCH_APP_URL_PLACEHOLDER", SEARCH_APP_URL)
        out = out.replace("{QUERY}",           html.escape(query_str))
        out = out.replace("{STATS}",           stats_html)
        out = out.replace("{RESULTS}",         results_html)
        out = out.replace("{PAGINATION_TOP}",  pagination_html)
        out = out.replace("{PAGINATION_BOTTOM}", pagination_html)
        for k, v in sort_state.items():
            out = out.replace("{" + k + "}", v)

        try:
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            self.wfile.write(out.encode("utf-8"))
        except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError, OSError):
            pass

    def do_POST(self):
        global PHI_SCANNER_PROCESS, PHI_INGEST_THREAD
        parsed = urlparse(self.path)

        if parsed.path == "/api/start_phi_scan":
            cfg  = load_phi_config()
            dirs = cfg.get("selected_directories", [])
            if not dirs:
                self.send_json({"status": "error", "message": "No directories configured."}, 400)
                return
            if is_scanner_running():
                self.send_json({"status": "error", "message": "Scan already running."}, 409)
                return
            t = threading.Thread(
                target=run_phi_scan,
                args=(dirs, cfg.get("min_risk","LOW"), int(cfg.get("workers",4))),
                daemon=True,
            )
            t.start()
            PHI_INGEST_THREAD = t
            self.send_json({"status": "ok", "message": "PHI scan started."})
            return

        if parsed.path == "/api/stop_phi_scan":
            stop_scanner()
            self.send_json({"status": "ok", "message": "Scan stopped."})
            return

        if parsed.path == "/api/phi_config":
            length = int(self.headers.get("Content-Length", 0))
            try:
                save_phi_config(json.loads(self.rfile.read(length)))
                self.send_json({"status": "ok"})
            except Exception as e:
                self.send_json({"status": "error", "message": str(e)}, 500)
            return

        self.send_json({"status": "error", "message": "Unknown endpoint."}, 404)


# ---------------------------------------------------------------------------
# Entrypoint
# ---------------------------------------------------------------------------

def main():
    try:
        ensure_phi_index()
    except Exception as e:
        print(f"[!] OpenSearch not reachable at startup: {e}")
        print("    App will start anyway; connect OpenSearch before scanning.")

    server = ThreadingHTTPServer(("localhost", PHI_PORT), PhiHandler)
    print(f"[+] PHI Scanner  ->  http://localhost:{PHI_PORT}")
    print(f"    Document Search ->  http://localhost:8080")
    print(f"    PHI index       ->  {PHI_INDEX_NAME}")
    print(f"    Thumbnails      ->  {PHI_THUMB_DIR}")
    server.serve_forever()


if __name__ == "__main__":
    main()
