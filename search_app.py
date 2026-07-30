"""
Dedicated Local Web Search Interface & Index Manager for OpenSearch.
Includes:
- Integrated In-Browser PDF.js & HTML5 Document Viewer Modal (No external app required!)
- High-Speed PyMuPDF Page 1 PDF Document Thumbnail Generator & Cache (/api/thumbnail?path=...)
- 2-Column Result Cards with Page 1 Thumbnail Previews & Live Preview button ('👁️ Preview')
- Direct Windows Custom URI Protocols (openfile://, openopus://, openexplorer://)
- Live Document Counter badge in main header ('📊 14,219 Documents Indexed')
- Dynamic '⚡ Start Indexing' / '⏹️ Stop Indexing' control button
- Strict AND matching for multi-word search terms
- Full Boolean NOT support (e.g. 'quality NOT IUH', 'quality -IUH')
"""

import os
import re
import json
import string
import sys
import html
import urllib.parse
import threading
import subprocess
import hashlib
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
from pathlib import Path
from opensearchpy import OpenSearch
import fitz  # PyMuPDF for lightning-fast Page 1 thumbnail generation

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "indexer_config.json")
THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")
KNOWN_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt', 'csv', 'md', 'rtf', 'pptx', 'ppt'}

DOPUS_RT = r"C:\Program Files\GPSoftware\Directory Opus\dopusrt.exe"
DOPUS_EXE = r"C:\Program Files\GPSoftware\Directory Opus\dopus.exe"

# System folders to hide from folder tree picker
HIDDEN_DIRS = {'$recycle.bin', 'system volume information', 'windows', 'program files', 'program files (x86)', 'programdata', 'appdata'}

# Global variables to track active background indexer process
INDEXER_PROCESS = None

if not os.path.exists(THUMB_CACHE_DIR):
    os.makedirs(THUMB_CACHE_DIR, exist_ok=True)


def get_thumbnail_bytes(file_path):
    """
    Generates or retrieves Page 1 thumbnail JPEG bytes for PDF/Image files.
    """
    if not os.path.exists(file_path):
        return None, None

    file_hash = hashlib.md5(file_path.encode('utf-8')).hexdigest()
    cache_path = os.path.join(THUMB_CACHE_DIR, f"{file_hash}.jpg")

    # Serve cached thumbnail if it exists
    if os.path.exists(cache_path):
        try:
            with open(cache_path, 'rb') as f:
                return f.read(), "image/jpeg"
        except Exception:
            pass

    ext = os.path.splitext(file_path)[1].lower().lstrip('.')

    # 1. Render PDF Page 1 Thumbnail
    if ext == 'pdf':
        try:
            doc = fitz.open(file_path)
            if len(doc) > 0:
                page = doc[0]
                pix = page.get_pixmap(dpi=200)  # Ultra-high resolution 300px width Page 1 render
                pix.save(cache_path)
                doc.close()
                with open(cache_path, 'rb') as f:
                    return f.read(), "image/jpeg"
        except Exception:
            pass

    # 2. Render Image Files Directly
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


def parse_smart_query(user_query: str, sort_by: str = "relevance"):
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
    query_body["size"] = 100
    return query_body


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>OpenSearch File Finder & Indexer</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1100px; margin: 0 auto; }
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

        /* 2-Column Result Card Layout */
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
        .viewer-body { flex: 1; border: none; width: 100%; height: 100%; background: #f8f9fa; }
        .viewer-close { cursor: pointer; font-size: 24px; color: #adb5bd; }
        .viewer-close:hover { color: white; }

        /* Lightbox Image Preview */
        .lightbox-modal { display: none; position: fixed; z-index: 3500; left: 0; top: 0; width: 100%; height: 100%; background: rgba(0,0,0,0.85); align-items: center; justify-content: center; }
        .lightbox-content { max-width: 90%; max-height: 90%; border-radius: 8px; box-shadow: 0 5px 25px rgba(0,0,0,0.5); }
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

        {RESULTS}
    </div>

    <div class="toast-notification" id="toastMsg"></div>

    <!-- Full Document Viewer Modal -->
    <div class="viewer-modal" id="docViewerModal">
        <div class="viewer-box">
            <div class="viewer-header">
                <span id="viewerTitle">📄 Document Viewer</span>
                <span class="viewer-close" onclick="closeDocViewer()">&times;</span>
            </div>
            <iframe class="viewer-body" id="viewerFrame" src="about:blank"></iframe>
        </div>
    </div>

    <div class="lightbox-modal" id="lightboxModal" onclick="closeLightbox()">
        <img class="lightbox-content" id="lightboxImg">
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

        function openDocViewer(filePath, title) {
            const modal = document.getElementById('docViewerModal');
            const frame = document.getElementById('viewerFrame');
            const titleElem = document.getElementById('viewerTitle');

            titleElem.innerText = '📄 Previewing: ' + title;
            frame.src = '/api/raw_file?path=' + encodeURIComponent(filePath);
            modal.style.display = 'flex';
        }

        function closeDocViewer() {
            document.getElementById('docViewerModal').style.display = 'none';
            document.getElementById('viewerFrame').src = 'about:blank';
        }

        function openLightbox(src) {
            const modal = document.getElementById('lightboxModal');
            const img = document.getElementById('lightboxImg');
            img.src = src;
            modal.style.display = 'flex';
        }

        function closeLightbox() {
            document.getElementById('lightboxModal').style.display = 'none';
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

                if (isIndexing) {
                    btn.className = 'btn-index running';
                    btn.innerText = '⏹️ Stop Indexing';
                    badge.style.display = 'block';
                } else {
                    btn.className = 'btn-index';
                    btn.innerText = '⚡ Start Indexing';
                    badge.style.display = 'none';
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
                btn.className = 'btn-index running';
                btn.innerText = '⏹️ Stop Indexing';
                badge.style.display = 'block';

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
                    self.send_response(200)
                    self.send_header("Content-Type", content_type)
                    self.send_header("Cache-Control", "public, max-age=86400")
                    self.end_headers()
                    self.wfile.write(img_data)
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
                        # Direct Windows Explorer launch with file selected
                        subprocess.Popen(['explorer', '/select,', norm_path])
                        self.send_json({"status": "ok", "message": "Folder opened in Windows Explorer"})
                    else:
                        # Try Directory Opus first
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

        # Render Main Search HTML
        query_str = params.get('q', [''])[0].strip()
        sort_by = params.get('sort', ['relevance'])[0].strip()

        cfg = load_config()
        selected_json = json.dumps(cfg.get('selected_directories', []))

        stats_html = ""
        results_html = ""

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
                es_query = parse_smart_query(query_str, sort_by=sort_by)
                res = client.search(index="documents", body=es_query)

                hits = res['hits']['hits']
                total = res['hits']['total']['value']
                took = res['took']

                stats_html = f"Found {total:,} matching document(s) in {took} ms"

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

                        # Render Right Column Thumbnail (Page 1 preview for PDF / Image)
                        if ftype in ('pdf', 'jpg', 'jpeg', 'png', 'bmp', 'webp'):
                            thumb_html = f"""
                            <div class="card-right">
                                <img src="{thumb_url}" class="thumb-preview" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}')" title="Click to open live full document previewer" alt="Page 1 Preview">
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
                                    <a class="file-title" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}')" title="Click to view live full document preview in browser">{fname}</a>
                                    <div class="card-actions">
                                        <button type="button" class="btn-action btn-preview" onclick="openDocViewer('{escaped_js_path}', '{escaped_title}')">👁️ Preview</button>
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
