"""
High-Performance Document Ingestion Pipeline for OpenSearch.
Supports strict document mode (PDF, Word, Excel, Text, Markdown, PowerPoint)
with clean file_type field (pdf, docx, xlsx, txt) for 1-word search filtering.
"""

import os
import sys
import json
import time
import hashlib
import argparse
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from opensearchpy import OpenSearch, helpers
try:
    from tqdm import tqdm
except ImportError:
    tqdm = None

# Optional text extraction imports
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

# Default folders to skip (OS binaries, cache, system temporary files, cloud caches)
DEFAULT_EXCLUDE_DIRS = {
    r'c:\windows', r'c:\program files', r'c:\program files (x86)',
    r'c:\programdata', r'$recycle.bin', r'system volume information',
    '__pycache__', '.git', '.svn', 'node_modules', '.venv', 'venv',
    '.dropbox.cache', 'dropboxbackup', '.cache', 'appdata'
}

# File extensions to strictly SKIP (Binaries, DLLs, Shortcuts, System caches, CSVs, Data files)
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
        print(f"[+] Index '{index_name}' already exists.")
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
    print(f"[+] Index '{index_name}' successfully created.")


def extract_file_content(file_path: str, max_bytes: int = 1_000_000) -> str:
    """Extracts text content from document formats safely (capped at 1MB per file)."""
    ext = os.path.splitext(file_path)[1].lower()
    
    if ext in SKIP_EXTENSIONS or not ext:
        return ""
        
    # Text & Markdown files
    if ext in ['.txt', '.md', '.rtf']:
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                return f.read(max_bytes)
        except Exception:
            return ""

    # PDF files
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

    # Word files
    elif ext in ['.docx', '.doc'] and docx is not None:
        try:
            doc = docx.Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs if p.text])[:max_bytes]
        except Exception:
            return ""

    # Excel files
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
    """Parses metadata and content for a single document file."""
    file_name = os.path.basename(file_path)
    ext = os.path.splitext(file_name)[1].lower()
    
    if not ext or ext not in STRICT_DOCUMENT_EXTENSIONS or file_name.startswith('~') or file_name.startswith('.'):
        return None
        
    try:
        stat = os.stat(file_path)
        doc_id = hashlib.sha256(file_path.encode('utf-8')).hexdigest()
        
        file_dir = os.path.dirname(file_path)
        file_type = ext.lstrip('.').lower() # 'pdf', 'docx', 'xlsx', 'txt', 'md'
        content = extract_file_content(file_path)
        
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
    """Checks if a directory path should be skipped."""
    norm_path = dir_path.lower()
    for excl in DEFAULT_EXCLUDE_DIRS:
        if excl in norm_path or norm_path.startswith(excl):
            return True
    return False


def scan_directories(target_dirs: list):
    """Recursively yields file paths from multiple target directories."""
    for target_dir in target_dirs:
        if not os.path.exists(target_dir):
            print(f"[!] Warning: Directory '{target_dir}' does not exist. Skipping.")
            continue
            
        print(f"[*] Scanning target: {target_dir}")
        for root, dirs, files in os.walk(target_dir, topdown=True):
            dirs[:] = [d for d in dirs if not is_excluded_dir(os.path.join(root, d))]
            
            if is_excluded_dir(root):
                continue
                
            for f in files:
                ext = os.path.splitext(f)[1].lower()
                if ext in STRICT_DOCUMENT_EXTENSIONS and not f.startswith('~') and not f.startswith('.'):
                    yield os.path.join(root, f)


def bulk_ingest(client, target_dirs: list, index_name="documents", batch_size=2000, num_threads=8):
    file_generator = scan_directories(target_dirs)
    
    actions_buffer = []
    total_indexed = 0
    total_failed = 0
    start_time = time.time()

    executor = ThreadPoolExecutor(max_workers=num_threads)
    
    print(f"[*] Starting ingestion pipeline using {num_threads} worker threads...")
    
    pending_futures = []
    
    try:
        for file_path in file_generator:
            future = executor.submit(process_single_file, file_path, index_name)
            pending_futures.append(future)
            
            if len(pending_futures) >= batch_size:
                for fut in as_completed(pending_futures):
                    res = fut.result()
                    if res:
                        actions_buffer.append(res)
                    else:
                        total_failed += 1
                        
                pending_futures = []
                
                if actions_buffer:
                    success_count, errors = helpers.bulk(client, actions_buffer, stats_only=False, raise_on_error=False)
                    total_indexed += success_count
                    if errors:
                        total_failed += len(errors)
                    actions_buffer = []
                    
                    elapsed = time.time() - start_time
                    rate = total_indexed / elapsed if elapsed > 0 else 0
                    print(f"[Progress] Indexed: {total_indexed:,} docs | Failed/Skipped: {total_failed:,} | Speed: {rate:.1f} docs/sec")
    except Exception as e:
        print(f"[!] Error during directory scan: {e}")

    # Flush remaining futures
    if pending_futures:
        for fut in as_completed(pending_futures):
            res = fut.result()
            if res:
                actions_buffer.append(res)
            else:
                total_failed += 1

    if actions_buffer:
        success_count, errors = helpers.bulk(client, actions_buffer, stats_only=False, raise_on_error=False)
        total_indexed += success_count
        if errors:
            total_failed += len(errors)

    executor.shutdown(wait=True)
    total_time = time.time() - start_time
    avg_rate = total_indexed / total_time if total_time > 0 else 0
    
    print("\n" + "="*50)
    print("INGESTION SUMMARY")
    print("="*50)
    print(f"Total Successfully Indexed: {total_indexed:,} documents")
    print(f"Total Failed/Skipped:       {total_failed:,} documents")
    print(f"Total Execution Time:        {total_time:.2f} seconds ({total_time/60:.2f} minutes)")
    print(f"Average Throughput:          {avg_rate:.1f} docs/sec")
    print("="*50)


def main():
    parser = argparse.ArgumentParser(description="Ingest documents into OpenSearch.")
    parser.add_argument("--dir", nargs="+", help="One or more directory paths to scan and index")
    parser.add_argument("--index", type=str, default="documents", help="OpenSearch index name (default: documents)")
    parser.add_argument("--threads", type=int, default=8, help="Number of worker threads (default: 8)")
    parser.add_argument("--batch", type=int, default=2000, help="Batch size for bulk indexing (default: 2000)")
    
    args = parser.parse_args()
    
    client = get_opensearch_client()
    
    try:
        info = client.info()
        print(f"[+] Connected to OpenSearch cluster: {info.get('cluster_name')} (v{info.get('version', {}).get('number')})")
    except Exception as e:
        print(f"[!] Failed to connect to OpenSearch at http://localhost:9200. Is Docker container running?")
        print(f"    Error details: {e}")
        sys.exit(1)
        
    ensure_index_exists(client, index_name=args.index)
    target_dirs = args.dir
    if not target_dirs:
        user_input = input("Enter directory path(s) to index (space-separated, e.g. C:\\Users D:\\): ").strip()
        target_dirs = user_input.split()
        
    bulk_ingest(client, target_dirs=target_dirs, index_name=args.index, batch_size=args.batch, num_threads=args.threads)


if __name__ == "__main__":
    main()
