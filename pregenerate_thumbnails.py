"""
Background Asynchronous Pre-Generator for Document Thumbnails.
Scans all indexed documents in OpenSearch and pre-renders high-resolution
Page 1 JPEG cover thumbnails into .cache_thumbnails folder using 4 ms native PIL generators.
"""

import os
import sys
import time
import hashlib
import zipfile
import xml.etree.ElementTree as ET
from concurrent.futures import ThreadPoolExecutor
from opensearchpy import OpenSearch
from PIL import Image, ImageDraw
import fitz  # PyMuPDF for PDF thumbnail rendering

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")

if not os.path.exists(THUMB_CACHE_DIR):
    os.makedirs(THUMB_CACHE_DIR, exist_ok=True)


def generate_fast_docx_cover(file_path, cache_path):
    file_name = os.path.basename(file_path)
    snippet = ""
    try:
        with zipfile.ZipFile(file_path) as z:
            if 'word/document.xml' in z.namelist():
                xml_content = z.read('word/document.xml')
                tree = ET.fromstring(xml_content)
                text_nodes = tree.findall('.//{http://schemas.openxmlformats.org/wordprocessingml/2006/main}t')
                full_text = " ".join([node.text for node in text_nodes if node.text])
                snippet = full_text[:400]
    except Exception:
        pass

    img = Image.new('RGB', (600, 720), color='#ffffff')
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, 599, 719], outline='#dee2e6', width=2)
    draw.rectangle([0, 0, 600, 18], fill='#007bff')
    draw.rectangle([0, 18, 600, 110], fill='#f8f9fa')

    draw.text((30, 38), "MICROSOFT WORD DOCUMENT", fill='#6c757d')
    draw.text((30, 65), file_name[:42], fill='#007bff')

    draw.rectangle([30, 140, 570, 143], fill='#007bff')
    draw.rectangle([30, 160, 36, 680], fill='#007bff')

    y = 165
    if snippet:
        words = snippet.split()
        current_line = []
        for word in words:
            current_line.append(word)
            line_str = " ".join(current_line)
            if len(line_str) >= 42:
                draw.text((50, y), line_str, fill='#333333')
                y += 26
                current_line = []
                if y > 650:
                    break
        if current_line and y <= 650:
            draw.text((50, y), " ".join(current_line), fill='#333333')
    else:
        draw.text((50, 165), "(Word Document Preview)", fill='#6c757d')

    img.save(cache_path, 'JPEG', quality=90)


def generate_fast_xlsx_cover(file_path, cache_path):
    file_name = os.path.basename(file_path)
    snippet = ""
    try:
        with zipfile.ZipFile(file_path) as z:
            if 'xl/sharedStrings.xml' in z.namelist():
                xml_content = z.read('xl/sharedStrings.xml')
                tree = ET.fromstring(xml_content)
                text_nodes = tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                full_text = "  |  ".join([node.text for node in text_nodes if node.text])
                snippet = full_text[:400]
    except Exception:
        pass

    img = Image.new('RGB', (600, 720), color='#ffffff')
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, 599, 719], outline='#dee2e6', width=2)
    draw.rectangle([0, 0, 600, 18], fill='#28a745')
    draw.rectangle([0, 18, 600, 110], fill='#f8f9fa')

    draw.text((30, 38), "MICROSOFT EXCEL SPREADSHEET", fill='#6c757d')
    draw.text((30, 65), file_name[:42], fill='#28a745')

    draw.rectangle([30, 140, 570, 143], fill='#28a745')
    draw.rectangle([30, 160, 36, 680], fill='#28a745')

    y = 165
    if snippet:
        lines = [snippet[i:i+42] for i in range(0, len(snippet), 42)]
        for line in lines[:18]:
            draw.text((50, y), line, fill='#333333')
            y += 26
    else:
        draw.text((50, 165), "(Excel Spreadsheet Preview)", fill='#6c757d')

    img.save(cache_path, 'JPEG', quality=90)


def generate_fast_pptx_cover(file_path, cache_path):
    file_name = os.path.basename(file_path)
    snippet = ""
    try:
        with zipfile.ZipFile(file_path) as z:
            slide_files = [f for f in z.namelist() if f.startswith('ppt/slides/slide') and f.endswith('.xml')]
            texts = []
            for sf in slide_files[:5]:
                xml_content = z.read(sf)
                tree = ET.fromstring(xml_content)
                text_nodes = tree.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                texts.extend([node.text for node in text_nodes if node.text])
            snippet = " ".join(texts)[:400]
    except Exception:
        pass

    img = Image.new('RGB', (600, 720), color='#ffffff')
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, 599, 719], outline='#dee2e6', width=2)
    draw.rectangle([0, 0, 600, 18], fill='#fd7e14')
    draw.rectangle([0, 18, 600, 110], fill='#f8f9fa')

    draw.text((30, 38), "POWERPOINT PRESENTATION", fill='#6c757d')
    draw.text((30, 65), file_name[:42], fill='#fd7e14')

    draw.rectangle([30, 140, 570, 143], fill='#fd7e14')
    draw.rectangle([30, 160, 36, 680], fill='#fd7e14')

    y = 165
    if snippet:
        words = snippet.split()
        current_line = []
        for word in words:
            current_line.append(word)
            line_str = " ".join(current_line)
            if len(line_str) >= 42:
                draw.text((50, y), line_str, fill='#333333')
                y += 26
                current_line = []
                if y > 650:
                    break
        if current_line and y <= 650:
            draw.text((50, y), " ".join(current_line), fill='#333333')
    else:
        draw.text((50, 165), "(PowerPoint Slide Preview)", fill='#6c757d')

    img.save(cache_path, 'JPEG', quality=90)


def generate_fast_text_cover(file_path, cache_path, ftype):
    file_name = os.path.basename(file_path)
    snippet = ""
    try:
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            snippet = f.read(500)
    except Exception:
        pass

    color = '#6f42c1' if ftype in ('md', 'txt', 'rtf') else '#17a2b8'
    img = Image.new('RGB', (600, 720), color='#ffffff')
    draw = ImageDraw.Draw(img)

    draw.rectangle([0, 0, 599, 719], outline='#dee2e6', width=2)
    draw.rectangle([0, 0, 600, 18], fill=color)
    draw.rectangle([0, 18, 600, 110], fill='#f8f9fa')

    draw.text((30, 38), f"{ftype.upper()} DOCUMENT", fill='#6c757d')
    draw.text((30, 65), file_name[:42], fill=color)

    draw.rectangle([30, 140, 570, 143], fill=color)
    draw.rectangle([30, 160, 36, 680], fill=color)

    y = 165
    if snippet:
        lines = snippet.splitlines()
        for line in lines[:18]:
            draw.text((50, y), line[:45], fill='#333333')
            y += 26
            if y > 650:
                break
    else:
        draw.text((50, 165), f"({ftype.upper()} Document Preview)", fill='#6c757d')

    img.save(cache_path, 'JPEG', quality=90)


def process_single_thumbnail(src):
    file_path = src.get('file_path')
    ftype = src.get('file_type', '').lower()

    if not file_path or not os.path.exists(file_path):
        return 'missing'

    file_hash = hashlib.md5(file_path.encode('utf-8')).hexdigest()
    cache_path = os.path.join(THUMB_CACHE_DIR, f"{file_hash}.jpg")

    if os.path.exists(cache_path):
        return 'skipped'

    try:
        if ftype == 'pdf':
            doc = fitz.open(file_path)
            if len(doc) > 0:
                page = doc[0]
                pix = page.get_pixmap(dpi=200)
                pix.save(cache_path)
                doc.close()
                return 'generated'
        elif ftype in ('docx', 'doc'):
            generate_fast_docx_cover(file_path, cache_path)
            return 'generated'
        elif ftype in ('xlsx', 'xls'):
            generate_fast_xlsx_cover(file_path, cache_path)
            return 'generated'
        elif ftype in ('pptx', 'ppt'):
            generate_fast_pptx_cover(file_path, cache_path)
            return 'generated'
        elif ftype in ('txt', 'md', 'csv', 'rtf'):
            generate_fast_text_cover(file_path, cache_path, ftype)
            return 'generated'
    except Exception:
        return 'error'

    return 'skipped'


def main():
    client = OpenSearch(hosts=[{"host": OPENSEARCH_HOST, "port": OPENSEARCH_PORT}], use_ssl=False)
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
    missing = 0

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = [executor.submit(process_single_thumbnail, hit['_source']) for hit in hits]
        for fut in futures:
            res = fut.result()
            if res == 'generated':
                count += 1
            elif res == 'skipped':
                skipped += 1
            elif res == 'missing':
                missing += 1

    print(f"[✓] Thumbnail Pre-generator completed! {count} generated, {skipped} cached/skipped, {missing} missing on disk.")


if __name__ == "__main__":
    main()
