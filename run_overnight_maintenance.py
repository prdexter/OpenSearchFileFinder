"""
Overnight Automated Maintenance & Optimization Script
------------------------------------------------------
Performs an end-to-end unattended maintenance run:
1. Prevents Windows PC from sleeping while tasks are running.
2. Prunes orphaned thumbnails from .cache_thumbnails (reclaims ~60 GB).
3. Executes ingest_documents.py with full .csv indexing and backup.
4. Runs pregenerate_thumbnails.py multi-threaded for 100% thumbnail coverage.
5. Writes a detailed summary report to overnight_maintenance_report.txt.
"""

import os
import sys
import time
import json
import ctypes
import hashlib
import subprocess
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
REPORT_FILE = os.path.join(BASE_DIR, "overnight_maintenance_report.txt")
THUMB_CACHE_DIR = os.path.join(BASE_DIR, ".cache_thumbnails")
CONFIG_FILE = os.path.join(BASE_DIR, "indexer_config.json")

def prevent_sleep():
    if sys.platform == "win32":
        try:
            # ES_CONTINUOUS | ES_SYSTEM_REQUIRED (0x80000001)
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000 | 0x00000001)
            print("[+] Windows power management configured: PC sleep prevented during maintenance.")
        except Exception as e:
            print(f"[-] Could not set sleep prevention: {e}")

def allow_sleep():
    if sys.platform == "win32":
        try:
            # ES_CONTINUOUS (0x80000000)
            ctypes.windll.kernel32.SetThreadExecutionState(0x80000000)
            print("[+] Windows power management restored to normal.")
        except Exception:
            pass

def log(msg, report_lines):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    formatted = f"[{ts}] {msg}"
    print(formatted)
    report_lines.append(formatted)

def prune_orphans(report_lines):
    from opensearchpy import OpenSearch
    from opensearchpy.helpers import scan

    log("=" * 70, report_lines)
    log("PHASE 1: ORPHAN THUMBNAIL CACHE PRUNING", report_lines)
    log("=" * 70, report_lines)

    client = OpenSearch([{"host": "localhost", "port": 9200}], timeout=60)
    if not client.indices.exists(index="documents"):
        log("[-] Error: 'documents' index does not exist in OpenSearch.", report_lines)
        return 0, 0

    log("[*] Querying OpenSearch for all active document paths...", report_lines)
    active_hashes = set()
    for hit in scan(client, index="documents", query={"_source": ["file_path"]}):
        fp = hit["_source"].get("file_path")
        if fp:
            norm = os.path.normpath(os.path.abspath(fp)).lower()
            active_hashes.add(hashlib.md5(norm.encode("utf-8")).hexdigest() + ".jpg")

    log(f"[+] Loaded {len(active_hashes):,} active document hashes.", report_lines)

    if not os.path.exists(THUMB_CACHE_DIR):
        log("[-] Cache directory does not exist.", report_lines)
        return 0, 0

    all_files = [f for f in os.listdir(THUMB_CACHE_DIR) if f.endswith(".jpg")]
    log(f"[+] Current files in .cache_thumbnails: {len(all_files):,}", report_lines)

    deleted_count = 0
    reclaimed_bytes = 0

    for f in all_files:
        if f not in active_hashes:
            f_path = os.path.join(THUMB_CACHE_DIR, f)
            try:
                sz = os.path.getsize(f_path)
                os.remove(f_path)
                deleted_count += 1
                reclaimed_bytes += sz
            except Exception:
                pass

    reclaimed_gb = reclaimed_bytes / (1024 * 1024 * 1024)
    log(f"[+] Successfully pruned {deleted_count:,} orphaned thumbnail files.", report_lines)
    log(f"[+] Reclaimed {reclaimed_gb:.2f} GB of disk space.", report_lines)
    log(f"[+] Remaining cache thumbnails: {len(all_files) - deleted_count:,}", report_lines)
    return deleted_count, reclaimed_gb

def run_document_ingestion(report_lines):
    log("=" * 70, report_lines)
    log("PHASE 2: DOCUMENT & CSV INGESTION AND BACKUP", report_lines)
    log("=" * 70, report_lines)

    log("[*] Launching ingest_documents.py with full CSV and document delta scan...", report_lines)
    ingest_py = os.path.join(BASE_DIR, "ingest_documents.py")
    cmd = [sys.executable, ingest_py]

    start_t = time.time()
    proc = subprocess.Popen(
        cmd,
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )

    for line in iter(proc.stdout.readline, ""):
        line_clean = line.strip()
        if line_clean:
            if any(k in line_clean for k in ("Starting", "Completed", "Indexed", "Scanned", "Backup", "Total", "Pre-fetching", "Error", "Warning")):
                log(f"  [Ingest] {line_clean}", report_lines)

    proc.wait()
    elapsed = time.time() - start_t
    log(f"[+] Document ingestion completed in {elapsed:.1f} seconds (Exit Code: {proc.returncode}).", report_lines)

def run_pregeneration(report_lines):
    log("=" * 70, report_lines)
    log("PHASE 3: MULTI-THREADED THUMBNAIL PRE-GENERATION", report_lines)
    log("=" * 70, report_lines)

    pregen_py = os.path.join(BASE_DIR, "pregenerate_thumbnails.py")
    cmd = [sys.executable, pregen_py]

    start_t = time.time()
    proc = subprocess.Popen(
        cmd,
        cwd=BASE_DIR,
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        text=True,
        bufsize=1
    )

    for line in iter(proc.stdout.readline, ""):
        line_clean = line.strip()
        if line_clean:
            if any(k in line_clean for k in ("Total", "Progress", "Summary", "Generated", "Skipped", "Error")):
                log(f"  [PreGen] {line_clean}", report_lines)

    proc.wait()
    elapsed = time.time() - start_t
    log(f"[+] Thumbnail pre-generation completed in {elapsed:.1f} seconds (Exit Code: {proc.returncode}).", report_lines)

def final_audit(report_lines):
    from opensearchpy import OpenSearch
    from opensearchpy.helpers import scan

    log("=" * 70, report_lines)
    log("FINAL AUDIT & COVERAGE VERIFICATION", report_lines)
    log("=" * 70, report_lines)

    client = OpenSearch([{"host": "localhost", "port": 9200}], timeout=60)
    total_docs = client.count(index="documents")["count"]

    csv_count = 0
    for hit in scan(client, index="documents", query={"query": {"term": {"file_type": "csv"}}}):
        csv_count += 1

    cached_files = [f for f in os.listdir(THUMB_CACHE_DIR) if f.endswith(".jpg")] if os.path.exists(THUMB_CACHE_DIR) else []

    log(f"[+] Total Documents in OpenSearch Index : {total_docs:,}", report_lines)
    log(f"[+] CSV Documents Indexed               : {csv_count:,}", report_lines)
    log(f"[+] Total Thumbnails in Cache           : {len(cached_files):,}", report_lines)
    coverage = (len(cached_files) / total_docs * 100) if total_docs > 0 else 0
    log(f"[+] Overall Thumbnail Coverage Ratio    : {coverage:.1f}%", report_lines)

def main():
    report_lines = []
    log("=" * 70, report_lines)
    log("OPENSEARCH OVERNIGHT MAINTENANCE RUN STARTED", report_lines)
    log("=" * 70, report_lines)

    prevent_sleep()
    overall_start = time.time()

    try:
        # Phase 1: Prune ~60 GB of stale thumbnails
        prune_orphans(report_lines)

        # Phase 2: Ingest all documents including CSVs
        run_document_ingestion(report_lines)

        # Phase 3: Pre-generate all remaining thumbnails (multi-threaded)
        run_pregeneration(report_lines)

        # Final stats
        final_audit(report_lines)

    except Exception as e:
        log(f"[-] Critical exception during overnight maintenance: {e}", report_lines)
    finally:
        allow_sleep()
        overall_elapsed = time.time() - overall_start
        log(f"[+] Total Maintenance Run Duration: {overall_elapsed / 60:.1f} minutes", report_lines)
        log("=" * 70, report_lines)
        log("MAINTENANCE RUN COMPLETE - READY FOR MORNING USE", report_lines)
        log("=" * 70, report_lines)

        try:
            with open(REPORT_FILE, "w", encoding="utf-8") as f:
                f.write("\n".join(report_lines) + "\n")
            print(f"[+] Written detailed morning report to: {REPORT_FILE}")
        except Exception:
            pass

if __name__ == "__main__":
    main()
