"""
High-Performance Document Ingestion Pipeline for OpenSearch.
Includes:
- Integrated Page 1 Thumbnail Generation directly during Document Indexing
- Real-time chunking for live UI counter updates
"""

import os
import sys
import json
import time
import hashlib
import argparse
import threading
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from opensearchpy import OpenSearch, helpers
import fitz  # PyMuPDF for PDF thumbnail rendering

try:
    import pypdf
except ImportError:
    pypdf = None

try:
    import docx
except ImportError:
    docx = None

try:
    import openpyxl
except ImportError:
    openpyxl = None

try:
    import win32com.client
    HAS_WORD = True
except ImportError:
    HAS_WORD = False

# Default folders to skip
DEFAULT_EXCLUDE_DIRS = {
    r'c:\windows', r'c:\program files', r'c:\program files (x86)',
    r'c:\programdata', r'$recycle.bin', r'system volume information',
    '__pycache__', '.git', '.svn', 'node_modules', '.venv', 'venv',
    '.dropbox.cache', 'dropboxbackup', '.cache', 'appdata'
}

# File extensions to strictly SKIP
SKIP_EXTENSIONS = {
    '.dll', '.exe', '.sys', '.bin', '.so', '.dylib', '.pyc', '.class',
    '.zip', '.tar', '.gz', '.7z', '.rar', '.iso', '.obj', '.o', '.lib',
    '.pdb', '.dat', '.tmp', '.pyd', '.nupkg', '.cab', '.msi', '.lnk',
    '.cache', '.bak', '.ico', '.png', '.jpg', '.jpeg', '.gif', '.mp3', '.mp4',
    '.ini', '.log', '.csv', '.tsv'
}

# Pure Human Document extensions (PDF, Word, Excel, Text, Markdown, PPT)
STRICT_DOCUMENT_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.txt', '.md', '.rtf', '.pptx', '.ppt'
}

THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")
if not os.path.exists(THUMB_CACHE_DIR):
    os.makedirs(THUMB_CACHE_DIR, exist_ok=True)

WORD_LOCK = threading.Lock()


def generate_thumbnail_if_missing(file_path: str, ext: str):
    """
    Generates Page 1 thumbnail during indexing if not already cached on disk.
    """
    try:
        file_hash = hashlib.md5(file_path.encode('utf-8')).hexdigest()
        cache_path = os.path.join(THUMB_CACHE_DIR, f"{file_hash}.jpg")

        if os.path.exists(cache_path):
            return

        if ext == '.pdf':
            doc = fitz.open(file_path)
            if len(doc) > 0:
                page = doc[0]
                pix = page.get_pixmap(dpi=200)
                pix.save(cache_path)
                doc.close()

        elif ext in ('.docx', '.doc') and HAS_WORD:
            with WORD_LOCK:
                temp_pdf = os.path.join(THUMB_CACHE_DIR, f"temp_{file_hash}.pdf")
                try:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    doc = word.Documents.Open(os.path.abspath(file_path))
                    doc.SaveAs(temp_pdf, FileFormat=17)
                    doc.Close()
                    word.Quit()

                    fitz_doc = fitz.open(temp_pdf)
                    if len(fitz_doc) > 0:
                        pix = fitz_doc[0].get_pixmap(dpi=200)
                        pix.save(cache_path)
                    fitz_doc.close()

                    if os.path.exists(temp_pdf):
                        os.remove(temp_pdf)
                except Exception:
                    if os.path.exists(temp_pdf):
                        try:
                            os.remove(temp_pdf)
                        except Exception:
                            pass
    except Exception:
        pass


def get_opensearch_client(host="localhost", port=9200):
    return OpenSearch(
        hosts=[{"host": host, "port": port}],
        http_compress=True,
        use_ssl=False,
        verify_certs=False,
        ssl_assert_hostname=False,
        ssl_show_warn=False,
        timeout=30,
        max_retries=3,
        retry_on_timeout=True
    )


def ensure_index_exists(client, index_name="documents", settings_path="index_settings.json"):
    if client.indices.exists(index=index_name):
        return

    print(f"[*] Creating index '{index_name}' with custom settings...")
    if os.path.exists(settings_path):
        with open(settings_path, "r", encoding="utf-8") as f:
            index_body = json.load(f)
    else:
        index_body = {
            "mappings": {
                "properties": {
                    "file_name": {"type": "text", "fields": {"keyword": {"type": "keyword"}}},
                    "file_type": {"type": "text", "fields": {"keyword": {"type": "keyword"}}},
                    "file_path": {"type": "keyword"},
                    "file_directory": {"type": "keyword"},
                    "file_extension": {"type": "keyword"},
                    "file_size": {"type": "long"},
                    "created_date": {"type": "date"},
                    "modified_date": {"type": "date"},
                    "content": {"type": "text", "term_vector": "with_positions_offsets"}
                }
            }
        }
    
    client.indices.create(index=index_name, body=index_body)


def extract_file_content(file_path: str, max_bytes: int = 1_000_000) -> str:
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in SKIP_EXTENSIONS or not ext:
        return ""
        
    if ext in ['.txt', '.md', '.rtf']:
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read(max_bytes)
        except Exception:
            return ""

    elif ext == '.pdf' and pypdf is not None:
        try:
            reader = pypdf.PdfReader(file_path)
            text_parts = []
            for i, page in enumerate(reader.pages):
                if i >= 50:
                    break
                text = page.extract_text()
                if text:
                    text_parts.append(text)
            return "\n".join(text_parts)[:max_bytes]
        except Exception:
            return ""

    elif ext in ['.docx', '.doc'] and docx is not None:
        try:
            doc = docx.Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs if p.text])[:max_bytes]
        except Exception:
            return ""

    elif ext in ['.xlsx', '.xls'] and openpyxl is not None:
        try:
            wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
            text_parts = []
            for sheet in wb.worksheets:
                for row in sheet.iter_rows(values_only=True):
                    row_str = " ".join([str(val) for val in row if val is not None])
                    if row_str.strip():
                        text_parts.append(row_str)
                    if len(text_parts) > 2000:
                        break
            return "\n".join(text_parts)[:max_bytes]
        except Exception:
            return ""

    return ""


def process_single_file(file_path: str, index_name: str) -> dict:
    file_name = os.path.basename(file_path)
    ext = os.path.splitext(file_name)[1].lower()
    
    if not ext or ext not in STRICT_DOCUMENT_EXTENSIONS or file_name.startswith('~') or file_name.startswith('.'):
        return None
        
    try:
        stat = os.stat(file_path)
        doc_id = hashlib.sha256(file_path.encode('utf-8')).hexdigest()
        
        file_dir = os.path.dirname(file_path)
        file_type = ext.lstrip('.').lower()
        content = extract_file_content(file_path)

        # Generate Page 1 Thumbnail persistently during ingestion
        generate_thumbnail_if_missing(file_path, ext)
        
        doc = {
            "_op_type": "index",
            "_index": index_name,
            "_id": doc_id,
            "file_name": file_name,
            "file_type": file_type,
            "file_path": file_path,
            "file_directory": file_dir,
            "file_extension": ext,
            "file_size": stat.st_size,
            "created_date": datetime.fromtimestamp(stat.st_ctime).isoformat(),
            "modified_date": datetime.fromtimestamp(stat.st_mtime).isoformat(),
            "content": content
        }
        return doc
    except Exception:
        return None


def is_excluded_dir(dir_path: str) -> bool:
    norm_path = dir_path.lower()
    for excl in DEFAULT_EXCLUDE_DIRS:
        if excl in norm_path or norm_path.startswith(excl):
            return True
    return False


def scan_directories(target_dirs: list):
    for target_dir in target_dirs:
        if not os.path.exists(target_dir):
            continue
            
        print(f"[*] Scanning target: {target_dir}")
        for root, dirs, files in os.walk(target_dir, topdown=True):
            dirs[:] = [d for d in dirs if not is_excluded_dir(os.path.join(root, d))]
            for file in files:
                file_path = os.path.join(root, file)
                yield file_path


def main():
    parser = argparse.ArgumentParser(description="Ingest documents into OpenSearch")
    parser.add_argument("--dir", nargs="+", help="Directories to scan and index")
    parser.add_argument("--index", default="documents", help="Target OpenSearch index name")
    args = parser.parse_args()

    if not args.dir:
        config_path = os.path.join(os.path.dirname(__file__), "indexer_config.json")
        if os.path.exists(config_path):
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                args.dir = cfg.get("selected_directories", [r"D:\Active research"])
        else:
            args.dir = [r"D:\Active research"]

    print(f"[*] Ingestion starting for directories: {args.dir}")
    client = get_opensearch_client()
    ensure_index_exists(client, index_name=args.index)

    start_time = time.time()
    batch = []
    total_indexed = 0

    with ThreadPoolExecutor(max_workers=4) as executor:
        futures = set()
        for file_path in scan_directories(args.dir):
            futures.add(executor.submit(process_single_file, file_path, args.index))

            if len(futures) >= 100:
                completed = set()
                for fut in as_completed(futures):
                    doc = fut.result()
                    if doc:
                        batch.append(doc)
                    completed.add(fut)
                    if len(completed) >= 50:
                        break
                futures -= completed

                if len(batch) >= 100:
                    helpers.bulk(client, batch)
                    total_indexed += len(batch)
                    print(f"[+] Bulk indexed {total_indexed} documents...")
                    batch = []

        # Flush remaining futures
        for fut in as_completed(futures):
            doc = fut.result()
            if doc:
                batch.append(doc)

        if batch:
            helpers.bulk(client, batch)
            total_indexed += len(batch)

    elapsed = time.time() - start_time
    print(f"[✓] Ingestion complete! Total documents indexed: {total_indexed:,} in {elapsed:.2f} seconds.")


if __name__ == "__main__":
    main()
