"""
Background Asynchronous Pre-Generator for Document Thumbnails.
Scans all indexed documents in OpenSearch and pre-renders high-resolution
Page 1 JPEG thumbnails into .cache_thumbnails folder.
"""

import os
import sys
import time
import hashlib
import threading
from concurrent.futures import ThreadPoolExecutor
from opensearchpy import OpenSearch
import fitz  # PyMuPDF for PDF thumbnail rendering

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")

if not os.path.exists(THUMB_CACHE_DIR):
    os.makedirs(THUMB_CACHE_DIR, exist_ok=True)

try:
    import win32com.client
    HAS_WORD = True
except ImportError:
    HAS_WORD = False

WORD_LOCK = threading.Lock()


def get_client():
    return OpenSearch(hosts=[{"host": OPENSEARCH_HOST, "port": OPENSEARCH_PORT}], use_ssl=False)


def process_pdf_thumbnail(file_path, cache_path):
    try:
        doc = fitz.open(file_path)
        if len(doc) > 0:
            page = doc[0]
            pix = page.get_pixmap(dpi=200)
            pix.save(cache_path)
            doc.close()
            return True
    except Exception:
        pass
    return False


def process_docx_thumbnail(file_path, cache_path, file_hash):
    if not HAS_WORD:
        return False
    with WORD_LOCK:
        temp_pdf = os.path.join(THUMB_CACHE_DIR, f"temp_{file_hash}.pdf")
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(file_path))
            doc.SaveAs(temp_pdf, FileFormat=17)  # 17 = wdFormatPDF
            doc.Close()
            word.Quit()

            fitz_doc = fitz.open(temp_pdf)
            if len(fitz_doc) > 0:
                pix = fitz_doc[0].get_pixmap(dpi=200)
                pix.save(cache_path)
            fitz_doc.close()

            if os.path.exists(temp_pdf):
                os.remove(temp_pdf)
            return True
        except Exception:
            if os.path.exists(temp_pdf):
                try:
                    os.remove(temp_pdf)
                except Exception:
                    pass
    return False


def main():
    client = get_client()
    if not client.indices.exists(index="documents"):
        print("[-] Index 'documents' does not exist.")
        return

    print("[*] Fetching document list from OpenSearch...")
    res = client.search(
        index="documents",
        body={"query": {"match_all": {}}, "_source": ["file_path", "file_type"]},
        size=10000
    )

    hits = res['hits']['hits']
    print(f"[+] Total documents found in index: {len(hits)}")

    count = 0
    skipped = 0
    errors = 0

    for i, hit in enumerate(hits):
        src = hit['_source']
        fpath = src.get('file_path', '')
        ftype = src.get('file_type', '').lower()

        if not fpath or not os.path.exists(fpath):
            continue

        file_hash = hashlib.md5(fpath.encode('utf-8')).hexdigest()
        cache_path = os.path.join(THUMB_CACHE_DIR, f"{file_hash}.jpg")

        if os.path.exists(cache_path):
            skipped += 1
            continue

        if ftype == 'pdf':
            if process_pdf_thumbnail(fpath, cache_path):
                count += 1
            else:
                errors += 1
        elif ftype in ('docx', 'doc'):
            if process_docx_thumbnail(fpath, cache_path, file_hash):
                count += 1
            else:
                errors += 1

        if (i + 1) % 50 == 0:
            print(f"[{i + 1}/{len(hits)}] Pre-generated: {count} | Cached: {skipped} | Errors: {errors}")

    print(f"[✓] Pre-generation complete! Total new thumbnails generated: {count}")


if __name__ == "__main__":
    main()
