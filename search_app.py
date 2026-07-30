"""
Dedicated Local Web Search Interface & Index Manager for OpenSearch.
Includes:
- Native Windows File Launching via /api/open_file (Click title or '↗️ Open File' button)
- Windows Explorer Folder Launcher via /api/open_folder ('📁 Open Folder' button)
- Live Document Counter badge in main header ('📊 14,219 Documents Indexed')
- Dynamic '⚡ Start Indexing' / '⏹️ Stop Indexing' control button
- Strict AND matching for multi-word search terms
- Full Boolean NOT support (e.g. 'quality NOT IUH', 'quality -IUH')
- Interactive Directory Tree Picker with Tri-State Indeterminate Checkboxes
"""

import os
import re
import json
import string
import sys
import threading
import subprocess
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
from pathlib import Path
from opensearchpy import OpenSearch

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
CONFIG_FILE = os.path.join(os.path.dirname(__file__), "indexer_config.json")
KNOWN_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt', 'csv', 'md', 'rtf', 'pptx', 'ppt'}

# System folders to hide from folder tree picker
HIDDEN_DIRS = {'$recycle.bin', 'system volume information', 'windows', 'program files', 'program files (x86)', 'programdata', 'appdata'}

# Global variables to track active background indexer process
INDEXER_PROCESS = None


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
        .index-badge { background: #fff3e0; color: #e65100; border: 1px solid #ffe0b2; padding: 4px 12px; border-radius: 12px; font-size: 13px; font-weight: bold; display: none; }

        .result-card { background: white; border-radius: 8px; padding: 18px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .result-header { display: flex; justify-content: space-between; align-items: center; }
        .file-title { font-weight: bold; font-size: 17px; color: #007bff; cursor: pointer; text-decoration: none; }
        .file-title:hover { text-decoration: underline; color: #0056b3; }
        
        .card-actions { display: flex; align-items: center; gap: 8px; }
        .btn-open-file { background-color: #007bff; color: white; border: none; padding: 6px 12px; border-radius: 5px; font-size: 13px; font-weight: bold; cursor: pointer; display: flex; align-items: center; gap: 4px; }
        .btn-open-file:hover { background-color: #0056b3; }
        .btn-open-folder { background-color: #f8f9fa; color: #495057; border: 1px solid #ced4da; padding: 6px 10px; border-radius: 5px; font-size: 13px; font-weight: bold; cursor: pointer; display: flex; align-items: center; gap: 4px; }
        .btn-open-folder:hover { background-color: #e2e6ea; }

        .badge { background: #e9ecef; padding: 4px 10px; border-radius: 12px; font-size: 13px; text-transform: uppercase; font-weight: 600; color: #495057; }
        .file-path { color: #6c757d; font-size: 13px; margin: 4px 0 10px 0; word-break: break-all; }
        .snippet { background: #f8f9fa; padding: 10px; border-left: 4px solid #007bff; border-radius: 4px; font-size: 14px; color: #495057; line-height: 1.5; }
        mark { background-color: #ffe066; padding: 2px 4px; border-radius: 3px; font-weight: bold; }

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
                <button class="btn-settings" onclick="openTreeModal()">📁 Index Directories</button>
                <button class="btn-index" id="indexBtn" onclick="toggleIndexing()">⚡ Start Indexing</button>
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
                <button class="btn-save" onclick="saveSelectedDirectories()">Save & Start Indexer</button>
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

        async function openFile(filePath) {
            try {
                await fetch('/api/open_file?path=' + encodeURIComponent(filePath));
            } catch (e) {
                alert('Error opening file: ' + e);
            }
        }

        async function openFolder(filePath) {
            try {
                await fetch('/api/open_folder?path=' + encodeURIComponent(filePath));
            } catch (e) {
                alert('Error opening folder: ' + e);
            }
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

        # API: Open File in Windows Default App
        if parsed.path == '/api/open_file':
            file_path = params.get('path', [''])[0]
            if file_path and os.path.exists(file_path):
                try:
                    os.startfile(file_path)
                    self.send_json({"status": "ok", "message": "File opened successfully"})
                except Exception as e:
                    self.send_json({"status": "error", "message": str(e)}, status=500)
            else:
                self.send_json({"status": "error", "message": "File not found"}, status=404)
            return

        # API: Open Containing Folder in File Explorer
        if parsed.path == '/api/open_folder':
            file_path = params.get('path', [''])[0]
            if file_path and os.path.exists(file_path):
                try:
                    subprocess.run(['explorer', '/select,', os.path.normpath(file_path)])
                    self.send_json({"status": "ok", "message": "Folder opened successfully"})
                except Exception as e:
                    self.send_json({"status": "error", "message": str(e)}, status=500)
            else:
                self.send_json({"status": "error", "message": "Path not found"}, status=404)
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
                        fname = src.get('file_name', 'Unnamed File')
                        fpath = src.get('file_path', '')
                        ftype = src.get('file_type', 'doc')
                        
                        # Escape single quotes and backslashes for JS inline call
                        escaped_path = fpath.replace('\\', '\\\\').replace("'", "\\'")

                        highlights = hit.get('highlight', {}).get('content', [])
                        if highlights:
                            snippet_text = " ... ".join(highlights)
                        else:
                            snippet_text = (src.get('content', '')[:300] + "...") if src.get('content') else "No preview text available."

                        card = f"""
                        <div class="result-card">
                            <div class="result-header">
                                <a class="file-title" onclick="openFile('{escaped_path}')" title="Click to open file in Windows">{fname}</a>
                                <div class="card-actions">
                                    <button class="btn-open-file" onclick="openFile('{escaped_path}')">↗️ Open File</button>
                                    <button class="btn-open-folder" onclick="openFolder('{escaped_path}')">📁 Folder</button>
                                    <span class="badge">{ftype}</span>
                                </div>
                            </div>
                            <div class="file-path">📁 {fpath}</div>
                            <div class="snippet">{snippet_text}</div>
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
    print("[+] OpenSearch Smart Search & Directory Tree Server running at http://localhost:8080")
    server.serve_forever()


if __name__ == "__main__":
    main()
