"""
Dedicated Local Web Search Interface & Index Manager for OpenSearch.
Includes:
- Full Search Results Pagination (Top & Bottom Navigation Bar, e.g. "Showing 101 - 200 of 350", "Next 100 of 350")
- 4 ms Native PIL DOCX & PyMuPDF Page 1 Document Thumbnail Generator & Cache (/api/thumbnail?path=...)
- Mammoth.js & HTML5 In-Browser Live Document Previewer (/api/raw_file?path=...)
- 300px 2-Column Result Cards with Page 1 Previews on the Right Side
- Direct Windows Custom URI Protocols (openfile://, openopus://, openexplorer://)
- Live Document Counter badge in main header ('📊 14,096 Documents Indexed')
"""

import os
import re
import json
import string
import sys
import html
import time
import urllib.parse
import threading
import subprocess
import hashlib
import zipfile
import math
import xml.etree.ElementTree as ET
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
from pathlib import Path
from opensearchpy import OpenSearch
from PIL import Image, ImageDraw
import fitz  # PyMuPDF for PDF thumbnail rendering

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "indexer_config.json")
THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")
KNOWN_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt', 'csv', 'md', 'rtf', 'pptx', 'ppt'}
PAGE_SIZE = 100

DOPUS_RT = r"C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DOPUS_EXE = r"C:\Program Files\GPSoftware\Directory Opus\dopus.exe"

# System folders to hide from folder tree picker
HIDDEN_DIRS = {'$recycle.bin', 'system volume information', 'windows', 'program files', 'program files (x86)', 'programdata', 'appdata', 'deidentifier', 'identified'}

# Global variables to track active background indexer process
INDEXER_PROCESS = None

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


def get_thumbnail_bytes(file_path):
    if not os.path.exists(file_path):
        return None, None

    file_hash = hashlib.md5(file_path.encode('utf-8')).hexdigest()
    cache_path = os.path.join(THUMB_CACHE_DIR, f"{file_hash}.jpg")

    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'rb') as f:
                return f.read(), "image/jpeg"
        except Exception:
            pass

    ext = os.path.splitext(file_path)[1].lower().lstrip('.')

    if ext == 'pdf':
        try:
            doc = fitz.open(file_path)
            if len(doc) > 0:
                page = doc[0]
                pix = page.get_pixmap(dpi=200)
                pix.save(cache_path)
                doc.close()
                with open(cache_path, 'rb') as f:
                    return f.read(), "image/jpeg"
        except Exception:
            pass

    elif ext in ('docx', 'doc'):
        try:
            generate_fast_docx_cover(file_path, cache_path)
            with open(cache_path, 'rb') as f:
                return f.read(), "image/jpeg"
        except Exception:
            pass

    elif ext in ('jpg', 'jpeg', 'png', 'bmp', 'webp'):
        try:
            with open(file_path, 'rb') as f:
                content_type = f"image/{ext if ext != 'jpg' else 'jpeg'}"
                return f.read(), content_type
        except Exception:
            pass

    return None, None


def get_client():
    return OpenSearch(hosts=[{"host": OPENSEARCH_HOST, "port": OPENSEARCH_PORT}], use_ssl=False)


def get_document_count():
    try:
        client = get_client()
        if client.indices.exists(index="documents"):
            res = client.count(index="documents")
            return res.get('count', 0)
    except Exception:
        pass
    return 0


def load_config():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {"selected_directories": [r"D:\Active research", r"C:\Users\Paul Dexter\Documents"]}


def save_config(cfg):
    with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, indent=2)


def get_available_drives():
    drives = []
    for letter in string.ascii_uppercase:
        drive_path = f"{letter}:\\"
        if os.path.exists(drive_path):
            drives.append(drive_path)
    return drives


def list_subdirectories(parent_path: str):
    subdirs = []
    if not parent_path or not os.path.exists(parent_path):
        return subdirs

    try:
        with os.scandir(parent_path) as entries:
            for entry in entries:
                if entry.is_dir(follow_symlinks=False):
                    name_lower = entry.name.lower()
                    if name_lower not in HIDDEN_DIRS and not name_lower.startswith('.'):
                        subdirs.append({
                            "name": entry.name,
                            "path": entry.path
                        })
    except PermissionError:
        pass
    except Exception:
        pass
        
    subdirs.sort(key=lambda x: x["name"].lower())
    return subdirs


def parse_smart_query(user_query: str, sort_by: str = "relevance", page: int = 1, page_size: int = PAGE_SIZE):
    raw_query = user_query.strip()
    raw_query = re.sub(r'NOT\s*\(\s*([^)]+)\s*\)', r' NOT \1 ', raw_query, flags=re.IGNORECASE)
    
    tokens = raw_query.split()
    file_types = []
    must_terms = []
    must_not_terms = []

    skip_next = False
    for i, token in enumerate(tokens):
        if skip_next:
            skip_next = False
            continue

        token_upper = token.upper()

        if token_upper in ('AND', 'OR'):
            continue

        if token_upper == 'NOT' or token_upper == 'AND NOT':
            if i + 1 < len(tokens):
                not_val = tokens[i + 1].strip('()')
                if not_val:
                    must_not_terms.append(not_val)
                skip_next = True
            continue

        if token.startswith('-') and len(token) > 1:
            must_not_terms.append(token[1:].strip('()'))
            continue

        clean_token = token.lower().lstrip('.')
        if clean_token in KNOWN_EXTENSIONS:
            file_types.append(clean_token)
        else:
            must_terms.append(token)

    must_conditions = []
    must_not_conditions = []
    
    if file_types:
        must_conditions.append({"terms": {"file_type": file_types}})
        
    if must_terms:
        for term in must_terms:
            must_conditions.append({
                "multi_match": {
                    "query": term,
                    "fields": ["content^3", "file_name^5"],
                    "type": "best_fields"
                }
            })

    if must_not_terms:
        for not_term in must_not_terms:
            must_not_conditions.append({
                "multi_match": {
                    "query": not_term,
                    "fields": ["content", "file_name", "file_path"]
                }
            })

    bool_query = {}
    if must_conditions:
        bool_query["must"] = must_conditions
    if must_not_conditions:
        bool_query["must_not"] = must_not_conditions

    if not bool_query:
        query_body = {"query": {"match_all": {}}}
    else:
        query_body = {"query": {"bool": bool_query}}

    if sort_by == "date_desc":
        query_body["sort"] = [{"modified_date": {"order": "desc"}}]
    elif sort_by == "date_asc":
        query_body["sort"] = [{"modified_date": {"order": "asc"}}]
    elif sort_by == "name_asc":
        query_body["sort"] = [{"file_name.keyword": {"order": "asc"}}]
    elif sort_by == "size_desc":
        query_body["sort"] = [{"file_size": {"order": "desc"}}]
    else:
        if must_conditions and must_terms:
            query_body["sort"] = [{"_score": {"order": "desc"}}]
        else:
            query_body["sort"] = [{"modified_date": {"order": "desc"}}]

    query_body["highlight"] = {
        "pre_tags": ["<mark>"],
        "post_tags": ["</mark>"],
        "fields": {
            "content": {"fragment_size": 200, "number_of_fragments": 2}
        }
    }
    
    # Pagination
    query_body["from"] = (page - 1) * page_size
    query_body["size"] = page_size
    return query_body


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>OpenSearch File Finder & Indexer</title>
    <!-- Mammoth.js for client-side Word DOCX rendering -->
    <script src="https://cdnjs.cloudflare.com/ajax/libs/mammoth/1.4.21/mammoth.browser.min.js"></script>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1150px; margin: 0 auto; }
        .header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 20px; border-bottom: 2px solid #e9ecef; padding-bottom: 15px; }
        .header-title h2 { margin: 0; font-size: 24px; color: #007bff; }
        .header-title p { margin: 4px 0 0 0; color: #6c757d; font-size: 14px; }
        .header-buttons { display: flex; gap: 10px; align-items: center; }
        .count-badge { background-color: #e3f2fd; color: #0d47a1; border: 1px solid #bbdefb; padding: 9px 15px; border-radius: 6px; font-size: 14px; font-weight: bold; display: flex; align-items: center; gap: 6px; }
        .btn-settings { background-color: #6c757d; color: white; padding: 10px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 6px; }
        .btn-settings:hover { background-color: #5a6268; }
        
        .btn-index { background-color: #28a745; color: white; padding: 10px 18px; border: none; border-radius: 6px; cursor: pointer; font-size: 14px; font-weight: 600; display: flex; align-items: center; gap: 6px; transition: background-color 0.3s; }
        .btn-index:hover { background-color: #218838; }
        .btn-index.running { background-color: #dc3545; }
        .btn-index.running:hover { background-color: #c82333; }
        
        .search-box { display: flex; gap: 10px; margin-bottom: 20px; align-items: center; }
        input[type="text"] { flex: 1; padding: 14px; font-size: 16px; border: 2px solid #ced4da; border-radius: 6px; }
        select { padding: 14px; font-size: 15px; border: 2px solid #ced4da; border-radius: 6px; background: white; cursor: pointer; }
        .btn-search { padding: 14px 28px; font-size: 16px; background-color: #007bff; color: white; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; }
        .btn-search:hover { background-color: #0056b3; }

        .stats-bar { display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; }
        .stats { color: #6c757d; font-weight: 500; }
        .toast-notification { position: fixed; bottom: 20px; right: 20px; background: #28a745; color: white; padding: 14px 28px; border-radius: 6px; font-size: 15px; font-weight: bold; box-shadow: 0 4px 12px rgba(0,0,0,0.2); display: none; z-index: 2000; }
        .index-badge { background: #fff3e0; color: #e65100; border: 1px solid #ffe0b2; padding: 4px 12px; border-radius: 12px; font-size: 13px; font-weight: bold; display: none; }

        /* Pagination Bar Styles */
        .pagination-bar { display: flex; justify-content: space-between; align-items: center; background: white; border: 1px solid #dee2e6; border-radius: 8px; padding: 12px 20px; margin: 15px 0; box-shadow: 0 2px 4px rgba(0,0,0,0.04); }
        .pagination-bar .page-info { font-weight: bold; color: #495057; font-size: 15px; }
        .btn-page { text-decoration: none; padding: 9px 18px; border-radius: 6px; background-color: #007bff; color: white; font-weight: bold; font-size: 14px; transition: background-color 0.2s; display: inline-flex; align-items: center; gap: 6px; }
        .btn-page:hover { background-color: #0056b3; color: white; }
        .btn-page.disabled { background-color: #e9ecef; color: #adb5bd; pointer-events: none; cursor: default; }

        /* 2-Column Result Card Layout with 300px Right Column Preview */
        .result-card { background: white; border-radius: 8px; padding: 18px; margin-bottom: 15px; box-shadow: 0 2px 6px rgba(0,0,0,0.06); display: flex; gap: 20px; align-items: flex-start; }
        .card-left { flex: 1; min-width: 0; }
        .card-right { flex-shrink: 0; width: 300px; }
        
        .thumb-preview { width: 100%; border-radius: 6px; border: 1px solid #dee2e6; box-shadow: 0 2px 5px rgba(0,0,0,0.1); cursor: pointer; transition: transform 0.2s, box-shadow 0.2s; background: #fff; }
        .thumb-preview:hover { transform: scale(1.03); box-shadow: 0 4px 12px rgba(0,0,0,0.2); }
        .thumb-placeholder { width: 100%; height: 360px; background: #f1f3f5; border-radius: 6px; border: 1px dashed #ced4da; display: flex; align-items: center; justify-content: center; color: #adb5bd; font-size: 14px; font-weight: 500; text-align: center; padding: 10px; }

        .result-header { display: flex; justify-content: space-between; align-items: center; margin-bottom: 6px; }
        .file-title { font-weight: bold; font-size: 17px; color: #007bff; text-decoration: none; cursor: pointer; }
        .file-title:hover { text-decoration: underline; color: #0056b3; }
        
        .card-actions { display: flex; align-items: center; gap: 8px; }
        .btn-action { text-decoration: none; padding: 6px 12px; border-radius: 5px; font-size: 13px; font-weight: bold; cursor: pointer; display: inline-flex; align-items: center; gap: 4px; border: none; }
        .btn-preview { background-color: #fd7e14; color: white; }
        .btn-preview:hover { background-color: #e8590c; color: white; }
        .btn-open-file { background-color: #007bff; color: white; }
        .btn-open-file:hover { background-color: #0056b3; color: white; }
        .btn-open-explorer { background-color: #17a2b8; color: white; }
        .btn-open-explorer:hover { background-color: #138496; color: white; }
        .btn-open-folder { background-color: #6f42c1; color: white; }
        .btn-open-folder:hover { background-color: #593196; color: white; }

        .badge { background: #e9ecef; padding: 4px 10px; border-radius: 12px; font-size: 13px; text-transform: uppercase; font-weight: 600; color: #495057; }
        .file-path { color: #6c757d; font-size: 13px; margin: 4px 0 10px 0; word-break: break-all; text-decoration: none; display: inline-block; cursor: pointer; }
        .file-path:hover { text-decoration: underline; color: #007bff; }
        .snippet { background: #f8f9fa; padding: 12px; border-left: 4px solid #007bff; border-radius: 4px; font-size: 14px; color: #495057; line-height: 1.5; margin-top: 6px; }
        mark { background-color: #ffe066; padding: 2px 4px; border-radius: 3px; font-weight: bold; }

        /* Full In-Browser Document Viewer Modal */
        .viewer-modal { display: none; position: fixed; z-index: 3000; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); align-items: center; justify-content: center; }
        .viewer-box { background: white; width: 92%; height: 92%; border-radius: 10px; display: flex; flex-direction: column; overflow: hidden; box-shadow: 0 10px 30px rgba(0,0,0,0.5); }
        .viewer-header { background: #343a40; color: white; padding: 14px 20px; display: flex; justify-content: space-between; align-items: center; font-weight: bold; }
        .viewer-header span { font-size: 16px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 80%; }
        .viewer-body { flex: 1; border: none; width: 100%; height: 100%; background: #f8f9fa; overflow: auto; padding: 25px; box-sizing: border-box; }
        .viewer-close { cursor: pointer; font-size: 24px; color: #adb5bd; }
        .viewer-close:hover { color: white; }

        /* Modal Styles */
        .modal { display: none; position: fixed; z-index: 1000; left: 0; top: 0; width: 100%; height: 100%; background-color: rgba(0,0,0,0.5); }
        .modal-content { background-color: white; margin: 40px auto; padding: 25px; border-radius: 8px; width: 650px; max-width: 90%; max-height: 80vh; display: flex; flex-direction: column; }
        .modal-header { display: flex; justify-content: space-between; align-items: center; border-bottom: 1px solid #dee2e6; padding-bottom: 12px; margin-bottom: 15px; }
        .modal-header h3 { margin: 0; }
        .close { cursor: pointer; font-size: 24px; color: #aaa; }
        .close:hover { color: #000; }

        .tree-container { flex: 1; overflow-y: auto; border: 1px solid #ced4da; border-radius: 6px; padding: 12px; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-size: 14px; background: #fafafa; }
        .tree-item { margin: 6px 0; }
        .tree-toggle { cursor: pointer; display: inline-block; width: 22px; height: 22px; line-height: 22px; text-align: center; font-weight: bold; color: #007bff; user-select: none; border-radius: 3px; }
        .tree-toggle:hover { background-color: #e9ecef; }
        .tree-children { margin-left: 24px; display: none; }
        .tree-children.open { display: block; }
        
        .modal-footer { display: flex; justify-content: space-between; align-items: center; margin-top: 15px; border-top: 1px solid #dee2e6; padding-top: 15px; }
        .btn-save { background-color: #28a745; color: white; padding: 10px 20px; border: none; border-radius: 6px; cursor: pointer; font-weight: bold; }
        .btn-save:hover { background-color: #218838; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <div class="header-title">
                <h2>🔍 OpenSearch File Finder</h2>
                <p>Type extension names (e.g. <code>pdf</code>, <code>docx</code>, <code>xlsx</code>, <code>md</code>) and exclusion terms (e.g. <code>quality NOT IUH</code>) directly into the search bar!</p>
            </div>
            <div class="header-buttons">
                <div class="count-badge" id="docCountBadge">📊 Total Indexed: ...</div>
                <button type="button" class="btn-settings" onclick="openTreeModal()">📁 Index Directories</button>
                <button type="button" class="btn-index" id="indexBtn" onclick="toggleIndexing()">⚡ Start Indexing</button>
            </div>
        </div>

        <form class="search-box" method="GET" action="/">
            <input type="text" name="q" value="{QUERY}" placeholder="Try: 'pdf quality disparities', 'quality NOT IUH', 'xlsx pathology'..." autofocus>
            <select name="sort" onchange="this.form.submit()">
                <option value="relevance" {SORT_RELEVANCE}>Best Match (Relevance)</option>
                <option value="date_desc" {SORT_DATE_DESC}>Date: Newest First</option>
                <option value="date_asc" {SORT_DATE_ASC}>Date: Oldest First</option>
                <option value="name_asc" {SORT_NAME_ASC}>File Name (A - Z)</option>
                <option value="size_desc" {SORT_SIZE_DESC}>File Size: Largest First</option>
            </select>
            <button type="submit" class="btn-search">Search</button>
        </form>

        <div class="stats-bar">
            <div class="stats">{STATS}</div>
            <div class="index-badge" id="indexStatusBadge">⏳ Indexing active in background...</div>
        </div>

        {PAGINATION_TOP}

        {RESULTS}

        {PAGINATION_BOTTOM}
    </div>

    <div class="toast-notification" id="toastMsg"></div>

    <!-- Full Document Viewer Modal -->
    <div class="viewer-modal" id="docViewerModal">
        <div class="viewer-box">
            <div class="viewer-header">
                <span id="viewerTitle">📄 Document Viewer</span>
                <span class="viewer-close" onclick="closeDocViewer()">&times;</span>
            </div>
            <div class="viewer-body" id="viewerContainer"></div>
        </div>
    </div>

    <!-- Directory Tree Picker Modal -->
    <div id="treeModal" class="modal">
        <div class="modal-content">
            <div class="modal-header">
                <h3>📁 Select Directories to Index</h3>
                <span class="close" onclick="closeTreeModal()">&times;</span>
            </div>
            <p style="font-size:13px; color:#6c757d; margin-top:0;">Check the folders you want OpenSearch to scan and index across your drives:</p>
            <div class="tree-container" id="treeContainer">
                Loading drive and directory tree...
            </div>
            <div class="modal-footer">
                <span id="statusMsg" style="font-size:13px; font-weight:bold; color:#007bff;"></span>
                <button type="button" class="btn-save" onclick="saveSelectedDirectories()">Save & Start Indexer</button>
            </div>
        </div>
    </div>

    <script>
        let selectedPaths = new Set();
        let isIndexing = false;

        document.addEventListener('DOMContentLoaded', function() {
            checkStatus();
            setInterval(checkStatus, 3000);
        });

        function showToast(msg, isError=false) {
            const toast = document.getElementById('toastMsg');
            toast.innerText = msg;
            toast.style.backgroundColor = isError ? '#dc3545' : '#28a745';
            toast.style.display = 'block';
            setTimeout(() => { toast.style.display = 'none'; }, 4000);
        }

        async function openDocViewer(filePath, title, ftype) {
            const modal = document.getElementById('docViewerModal');
            const container = document.getElementById('viewerContainer');
            const titleElem = document.getElementById('viewerTitle');

            titleElem.innerText = '📄 Live Preview: ' + title;
            container.innerHTML = '<div style="text-align:center; padding:40px; font-size:18px; color:#6c757d;">⏳ Loading document preview...</div>';
            modal.style.display = 'flex';

            const rawUrl = '/api/raw_file?path=' + encodeURIComponent(filePath);

            if (ftype === 'pdf') {
                container.innerHTML = '<iframe src="' + rawUrl + '" style="width:100%; height:100%; border:none;"></iframe>';
            } else if (ftype === 'docx' || ftype === 'doc') {
                try {
                    const response = await fetch(rawUrl);
                    const arrayBuffer = await response.arrayBuffer();
                    if (window.mammoth) {
                        const result = await mammoth.convertToHtml({arrayBuffer: arrayBuffer});
                        container.innerHTML = '<div style="max-width:850px; margin:0 auto; background:white; padding:40px; border-radius:6px; box-shadow:0 2px 10px rgba(0,0,0,0.1); line-height:1.6; font-size:16px;">' + result.value + '</div>';
                    } else {
                        container.innerHTML = '<div style="color:red;">Mammoth.js not loaded</div>';
                    }
                } catch (e) {
                    container.innerHTML = '<div style="color:red; padding:20px;">Error rendering DOCX document: ' + e + '</div>';
                }
            } else {
                container.innerHTML = '<iframe src="' + rawUrl + '" style="width:100%; height:100%; border:none;"></iframe>';
            }
        }

        function closeDocViewer() {
            document.getElementById('docViewerModal').style.display = 'none';
            document.getElementById('viewerContainer').innerHTML = '';
        }

        async function handleOpenFile(filePath, customUrl) {
            showToast('Opening file in default application...');
            window.location.href = customUrl;
            try {
                await fetch('/api/open_file?path=' + encodeURIComponent(filePath));
            } catch (e) {}
        }

        async function handleOpenExplorer(filePath, customUrl) {
            showToast('Opening folder in Windows File Explorer...');
            window.location.href = customUrl;
            try {
                await fetch('/api/open_folder?explorer=1&path=' + encodeURIComponent(filePath));
            } catch (e) {}
        }

        async function handleOpenFolder(filePath, customUrl) {
            showToast('Opening folder in Directory Opus...');
            window.location.href = customUrl;
            try {
                await fetch('/api/open_folder?path=' + encodeURIComponent(filePath));
            } catch (e) {}
        }

        async function checkStatus() {
            try {
                const res = await fetch('/api/status');
                const data = await res.json();
                const btn = document.getElementById('indexBtn');
                const badge = document.getElementById('indexStatusBadge');
                const countBadge = document.getElementById('docCountBadge');

                isIndexing = data.indexing_running;
                if (data.total_docs !== undefined) {
                    countBadge.innerText = '📊 Total Indexed: ' + data.total_docs.toLocaleString() + ' docs';
                }

                if (btn && badge) {
                    if (isIndexing) {
                        btn.className = 'btn-index running';
                        btn.innerText = '⏹️ Stop Indexing';
                        badge.style.display = 'block';
                    } else {
                        btn.className = 'btn-index';
                        btn.innerText = '⚡ Start Indexing';
                        badge.style.display = 'none';
                    }
                }
            } catch (e) {}
        }

        async function toggleIndexing() {
            const btn = document.getElementById('indexBtn');
            const badge = document.getElementById('indexStatusBadge');

            if (isIndexing) {
                btn.innerText = '⏳ Stopping...';
                await fetch('/api/stop_indexing', { method: 'POST' });
                checkStatus();
            } else {
                if (btn) btn.className = 'btn-index running';
                if (btn) btn.innerText = '⏹️ Stop Indexing';
                if (badge) badge.style.display = 'block';

                const cfgRes = await fetch('/api/config');
                const cfg = await cfgRes.json();
                const arr = cfg.selected_directories || [];

                await fetch('/api/config', {
                    method: 'POST',
                    headers: {'Content-Type': 'application/json'},
                    body: JSON.stringify({selected_directories: arr})
                });
                checkStatus();
            }
        }

        async function openTreeModal() {
            document.getElementById('treeModal').style.display = 'block';
            const cfgRes = await fetch('/api/config');
            const cfg = await cfgRes.json();
            selectedPaths = new Set(cfg.selected_directories || []);
            loadDriveTree();
        }

        function closeTreeModal() {
            document.getElementById('treeModal').style.display = 'none';
        }

        function norm(p) {
            return p ? p.toLowerCase().replace(/\\\\/g, '/').replace(/\/$/, '') : '';
        }

        function getCheckboxState(checkPath) {
            const cPath = norm(checkPath);
            let exactMatch = false;
            let childMatch = false;

            for (const sp of selectedPaths) {
                const normSp = norm(sp);
                if (normSp === cPath) {
                    exactMatch = true;
                } else if (normSp.startsWith(cPath + '/')) {
                    childMatch = true;
                }
            }

            if (exactMatch) return 'checked';
            if (childMatch) return 'indeterminate';
            return 'unchecked';
        }

        function updateCheckboxVisual(cb, state) {
            if (state === 'checked') {
                cb.checked = true;
                cb.indeterminate = false;
            } else if (state === 'indeterminate') {
                cb.checked = false;
                cb.indeterminate = true;
            } else {
                cb.checked = false;
                cb.indeterminate = false;
            }
        }

        async function loadDriveTree() {
            const container = document.getElementById('treeContainer');
            container.innerHTML = '';

            const res = await fetch('/api/drives');
            const drives = await res.json();

            for (const drive of drives) {
                const driveItem = document.createElement('div');
                driveItem.className = 'tree-item';

                const toggleSpan = document.createElement('span');
                toggleSpan.className = 'tree-toggle';
                toggleSpan.innerText = '▶';
                toggleSpan.setAttribute('data-path', drive);
                toggleSpan.onclick = function() { toggleFolder(this); };

                const cb = document.createElement('input');
                cb.type = 'checkbox';
                cb.value = drive;
                updateCheckboxVisual(cb, getCheckboxState(drive));
                cb.onchange = function() { onCheckboxChange(this); };

                const label = document.createElement('strong');
                label.innerHTML = ' 💻 ' + drive;

                const childrenDiv = document.createElement('div');
                childrenDiv.className = 'tree-children';

                driveItem.appendChild(toggleSpan);
                driveItem.appendChild(cb);
                driveItem.appendChild(label);
                driveItem.appendChild(childrenDiv);
                container.appendChild(driveItem);
            }
        }

        async function toggleFolder(element) {
            const path = element.getAttribute('data-path');
            const treeItem = element.parentElement;
            const childrenDiv = treeItem.querySelector('.tree-children');

            if (element.innerText === '▶') {
                element.innerText = '▼';
                childrenDiv.classList.add('open');

                if (childrenDiv.children.length === 0) {
                    const res = await fetch('/api/ls?path=' + encodeURIComponent(path));
                    const subdirs = await res.json();

                    if (subdirs.length === 0) {
                        childrenDiv.innerHTML = '<div style="margin-left:20px; color:#aaa; font-style:italic;">(Empty or No subfolders)</div>';
                    } else {
                        for (const sub of subdirs) {
                            const subItem = document.createElement('div');
                            subItem.className = 'tree-item';

                            const subToggle = document.createElement('span');
                            subToggle.className = 'tree-toggle';
                            subToggle.innerText = '▶';
                            subToggle.setAttribute('data-path', sub.path);
                            subToggle.onclick = function() { toggleFolder(this); };

                            const subCb = document.createElement('input');
                            subCb.type = 'checkbox';
                            subCb.value = sub.path;
                            updateCheckboxVisual(subCb, getCheckboxState(sub.path));
                            subCb.onchange = function() { onCheckboxChange(this); };

                            const subLabel = document.createElement('span');
                            subLabel.innerHTML = ' 📁 ' + sub.name;

                            const subChildren = document.createElement('div');
                            subChildren.className = 'tree-children';

                            subItem.appendChild(subToggle);
                            subItem.appendChild(subCb);
                            subItem.appendChild(subLabel);
                            subItem.appendChild(subChildren);
                            childrenDiv.appendChild(subItem);
                        }
                    }
                }
            } else {
                element.innerText = '▶';
                childrenDiv.classList.remove('open');
            }
        }

        function onCheckboxChange(cb) {
            cb.indeterminate = false;
            if (cb.checked) {
                selectedPaths.add(cb.value);
            } else {
                selectedPaths.delete(cb.value);
            }
        }

        async function saveSelectedDirectories() {
            const arr = Array.from(selectedPaths);
            const statusMsg = document.getElementById('statusMsg');
            statusMsg.innerText = 'Saving configuration & starting indexer...';

            await fetch('/api/config', {
                method: 'POST',
                headers: {'Content-Type': 'application/json'},
                body: JSON.stringify({selected_directories: arr})
            });

            statusMsg.innerText = '✓ Indexer triggered in background!';
            setTimeout(() => {
                closeTreeModal();
                statusMsg.innerText = '';
                checkStatus();
            }, 1500);
        }
    </script>
</body>
</html>
"""


class SearchHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        # API: Raw File Content Endpoint for In-Browser Document Viewer
        if parsed.path == '/api/raw_file':
            file_path = params.get('path', [''])[0]
            if file_path and os.path.exists(file_path):
                ext = os.path.splitext(file_path)[1].lower().lstrip('.')
                content_types = {
                    'pdf': 'application/pdf',
                    'docx': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                    'doc': 'application/msword',
                    'txt': 'text/plain; charset=utf-8',
                    'html': 'text/html; charset=utf-8',
                    'png': 'image/png',
                    'jpg': 'image/jpeg',
                    'jpeg': 'image/jpeg',
                    'webp': 'image/webp'
                }
                ctype = content_types.get(ext, 'application/octet-stream')
                try:
                    with open(file_path, 'rb') as f:
                        data = f.read()
                    self.send_response(200)
                    self.send_header("Content-Type", ctype)
                    self.send_header("Content-Length", str(len(data)))
                    self.end_headers()
                    self.wfile.write(data)
                    return
                except Exception as e:
                    self.send_response(500)
                    self.end_headers()
                    return
            self.send_response(404)
            self.end_headers()
            return

        # API: Render & Serve Page 1 Thumbnail Image
        if parsed.path == '/api/thumbnail':
            file_path = params.get('path', [''])[0]
            if file_path:
                img_data, content_type = get_thumbnail_bytes(file_path)
                if img_data:
                    try:
                        self.send_response(200)
                        self.send_header("Content-Type", content_type)
                        self.send_header("Cache-Control", "public, max-age=86400")
                        self.end_headers()
                        self.wfile.write(img_data)
                    except (BrokenPipeError, ConnectionAbortedError, ConnectionResetError):
                        pass
                    return
            self.send_response(404)
            self.end_headers()
            return

        # API: Open File in Windows Default App
        if parsed.path == '/api/open_file':
            file_path = params.get('path', [''])[0]
            if file_path and os.path.exists(file_path):
                try:
                    norm_path = os.path.normpath(file_path)
                    os.startfile(norm_path)
                    self.send_json({"status": "ok", "message": "File opened successfully"})
                except Exception as e:
                    self.send_json({"status": "error", "message": str(e)}, status=500)
            else:
                self.send_json({"status": "error", "message": f"File not found: {file_path}"}, status=404)
            return

        # API: Open Containing Folder directly in Windows File Explorer or Directory Opus
        if parsed.path == '/api/open_folder':
            file_path = params.get('path', [''])[0]
            force_explorer = params.get('explorer', ['0'])[0] == '1'

            if file_path and os.path.exists(file_path):
                try:
                    norm_path = os.path.normpath(file_path)
                    
                    if force_explorer:
                        subprocess.Popen(['explorer', '/select,', norm_path])
                        self.send_json({"status": "ok", "message": "Folder opened in Windows Explorer"})
                    else:
                        if os.path.exists(DOPUS_RT):
                            subprocess.Popen([DOPUS_RT, "/cmd", "Go", norm_path, "NEW", "SELECT"])
                        elif os.path.exists(DOPUS_EXE):
                            subprocess.Popen([DOPUS_EXE, "/select", norm_path])
                        else:
                            subprocess.Popen(['explorer', '/select,', norm_path])
                            
                        self.send_json({"status": "ok", "message": "Folder opened in Directory Opus"})
                except Exception as e:
                    self.send_json({"status": "error", "message": str(e)}, status=500)
            else:
                self.send_json({"status": "error", "message": f"Path not found: {file_path}"}, status=404)
            return

        # API: Indexing Status & Total Count
        if parsed.path == '/api/status':
            is_running = INDEXER_PROCESS is not None and INDEXER_PROCESS.poll() is None
            total_count = get_document_count()
            self.send_json({
                "indexing_running": is_running,
                "total_docs": total_count
            })
            return

        # API: Available Drives
        if parsed.path == '/api/drives':
            drives = get_available_drives()
            self.send_json(drives)
            return

        # API: Subdirectories
        if parsed.path == '/api/ls':
            parent_path = params.get('path', [''])[0]
            subdirs = list_subdirectories(parent_path)
            self.send_json(subdirs)
            return

        # API: Current Config
        if parsed.path == '/api/config':
            cfg = load_config()
            self.send_json(cfg)
            return

        # Render Main Search HTML
        query_str = params.get('q', [''])[0].strip()
        sort_by = params.get('sort', ['relevance'])[0].strip()
        
        try:
            page = int(params.get('page', ['1'])[0])
            if page < 1:
                page = 1
        except Exception:
            page = 1

        cfg = load_config()
        selected_json = json.dumps(cfg.get('selected_directories', []))

        stats_html = ""
        results_html = ""
        pagination_html = ""

        sort_state = {
            "SORT_RELEVANCE": "selected" if sort_by == "relevance" else "",
            "SORT_DATE_DESC": "selected" if sort_by == "date_desc" else "",
            "SORT_DATE_ASC": "selected" if sort_by == "date_asc" else "",
            "SORT_NAME_ASC": "selected" if sort_by == "name_asc" else "",
            "SORT_SIZE_DESC": "selected" if sort_by == "size_desc" else ""
        }

        if query_str:
            try:
                client = get_client()
                es_query = parse_smart_query(query_str, sort_by=sort_by, page=page, page_size=PAGE_SIZE)
                res = client.search(index="documents", body=es_query)

                hits = res['hits']['hits']
                total = res['hits']['total']['value']
                took = res['took']

                total_pages = math.ceil(total / PAGE_SIZE) if total > 0 else 1
                start_doc = (page - 1) * PAGE_SIZE + 1 if total > 0 else 0
                end_doc = min(page * PAGE_SIZE, total)

                stats_html = f"Found {total:,} matching document(s) in {took} ms"

                # Build Pagination Controls
                if total > 0:
                    prev_disabled = "disabled" if page <= 1 else ""
                    prev_page = page - 1 if page > 1 else 1
                    prev_url = f"/?q={urllib.parse.quote(query_str)}&sort={sort_by}&page={prev_page}"

                    next_disabled = "disabled" if end_doc >= total else ""
                    next_page = page + 1 if end_doc < total else page
                    next_count = min(PAGE_SIZE, total - end_doc)
                    next_label = f"Next {next_count} of {total:,} ➡️" if next_count > 0 else "Next ➡️"
                    next_url = f"/?q={urllib.parse.quote(query_str)}&sort={sort_by}&page={next_page}"

                    pagination_html = f"""
                    <div class="pagination-bar">
                        <a href="{prev_url}" class="btn-page {prev_disabled}">⬅️ Previous 100</a>
                        <span class="page-info">Showing {start_doc:,} - {end_doc:,} of {total:,} documents (Page {page} of {total_pages})</span>
                        <a href="{next_url}" class="btn-page {next_disabled}">{next_label}</a>
                    </div>
                    """

                if not hits:
                    results_html = "<div class='result-card'>No matching documents found.</div>"
                else:
                    cards = []
                    for hit in hits:
                        src = hit['_source']
                        fname = html.escape(src.get('file_name', 'Unnamed File'))
                        fpath = src.get('file_path', '')
                        
                        # Generate direct Windows Custom Protocol URLs & escaped JS paths
                        encoded_path = urllib.parse.quote(fpath.replace('\\', '/'))
                        escaped_js_path = html.escape(fpath.replace('\\', '\\\\').replace("'", "\\'"))
                        escaped_title = html.escape(fname.replace("'", "\\'"))
                        
                        open_file_url = f"openfile://{encoded_path}"
                        open_opus_url = f"openopus://{encoded_path}"
                        open_explorer_url = f"openexplorer://{encoded_path}"
                        thumb_url = f"/api/thumbnail?path={encoded_path}"
                        
                        escaped_display_path = html.escape(fpath)
                        ftype = src.get('file_type', 'doc').lower()

                        highlights = hit.get('highlight', {}).get('content', [])
                        if highlights:
                            snippet_text = " ... ".join(highlights)
                        else:
                            snippet_text = (src.get('content', '')[:300] + "...") if src.get('content') else "No preview text available."

                        # Render Right Column Thumbnail (Page 1 preview for PDF, DOCX, and Image)
                        if ftype in ('pdf', 'docx', 'doc', 'jpg', 'jpeg', 'png', 'bmp', 'webp'):
                            thumb_html = f"""
                            <div class="card-right">
                                <img src="{thumb_url}" class="thumb-preview" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}', '{ftype}')" title="Click to view live full document preview" alt="Page 1 Preview">
                            </div>
                            """
                        else:
                            thumb_html = f"""
                            <div class="card-right">
                                <div class="thumb-placeholder">📄 {ftype.upper()}<br>Document Preview</div>
                            </div>
                            """

                        card = f"""
                        <div class="result-card">
                            <div class="card-left">
                                <div class="result-header">
                                    <a class="file-title" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}', '{ftype}')" title="Click to view live full document preview in browser">{fname}</a>
                                    <div class="card-actions">
                                        <button type="button" class="btn-action btn-preview" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}', '{ftype}')">👁️ Preview</button>
                                        <button type="button" class="btn-action btn-open-file" onclick="handleOpenFile('{escaped_js_path}', '{open_file_url}')">↗️ Open File</button>
                                        <button type="button" class="btn-action btn-open-explorer" onclick="handleOpenExplorer('{escaped_js_path}', '{open_explorer_url}')">📁 Explorer</button>
                                        <button type="button" class="btn-action btn-open-folder" onclick="handleOpenFolder('{escaped_js_path}', '{open_opus_url}')">📁 Opus</button>
                                        <span class="badge">{ftype}</span>
                                    </div>
                                </div>
                                <div class="file-path" onclick="handleOpenExplorer('{escaped_js_path}', '{open_explorer_url}')" title="Click to open folder in Windows File Explorer">📁 {escaped_display_path}</div>
                                <div class="snippet">{snippet_text}</div>
                            </div>
                            {thumb_html}
                        </div>
                        """
                        cards.append(card)
                    results_html = "\n".join(cards)
            except Exception as e:
                results_html = f"<div class='result-card' style='color:red;'>Error executing search: {e}</div>"

        html_out = HTML_TEMPLATE.replace("{QUERY}", query_str).replace("{STATS}", stats_html).replace("{RESULTS}", results_html).replace("{SELECTED_JSON}", selected_json)
        html_out = html_out.replace("{PAGINATION_TOP}", pagination_html).replace("{PAGINATION_BOTTOM}", pagination_html)
        
        for key, val in sort_state.items():
            html_out = html_out.replace(f"{{{key}}}", val)

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html_out.encode('utf-8'))

    def do_POST(self):
        global INDEXER_PROCESS
        parsed = urlparse(self.path)

        # API: Stop Indexing
        if parsed.path == '/api/stop_indexing':
            if INDEXER_PROCESS and INDEXER_PROCESS.poll() is None:
                try:
                    INDEXER_PROCESS.terminate()
                    INDEXER_PROCESS.wait(timeout=2)
                except Exception:
                    try:
                        INDEXER_PROCESS.kill()
                    except Exception:
                        pass
            INDEXER_PROCESS = None
            self.send_json({"status": "ok", "message": "Indexing stopped."})
            return

        # API: Config and Start Indexing
        if parsed.path == '/api/config':
            content_length = int(self.headers.get('Content-Length', 0))
            body = self.rfile.read(content_length)
            try:
                data = json.loads(body)
                save_config(data)

                # Terminate any running instance first
                if INDEXER_PROCESS and INDEXER_PROCESS.poll() is None:
                    try:
                        INDEXER_PROCESS.terminate()
                    except Exception:
                        pass

                # Trigger ingest_documents.py as manageable subprocess
                dirs = data.get("selected_directories", [])
                if dirs:
                    def run_ingest():
                        global INDEXER_PROCESS
                        try:
                            cmd = [sys.executable, "ingest_documents.py", "--dir"] + dirs
                            INDEXER_PROCESS = subprocess.Popen(cmd, cwd=os.path.dirname(__file__))
                            INDEXER_PROCESS.wait()
                        finally:
                            INDEXER_PROCESS = None

                    t = threading.Thread(target=run_ingest)
                    t.daemon = True
                    t.start()

                self.send_json({"status": "ok", "message": "Saved and indexing started!"})
            except Exception as e:
                self.send_json({"status": "error", "message": str(e)}, status=500)

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header("Content-Type", "json/application; charset=utf-8")
        self.end_headers()
        self.wfile.write(json.dumps(data).encode('utf-8'))


def main():
    server = HTTPServer(('localhost', 8080), SearchHandler)
    print("[+] OpenSearch Smart Search Server running at http://localhost:8080")
    server.serve_forever()


if __name__ == "__main__":
    main()
