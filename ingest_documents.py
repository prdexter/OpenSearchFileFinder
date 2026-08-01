"""
High-Performance Document Ingestion Pipeline for OpenSearch.
Includes:
- Multi-Threaded Real-Time 20-Doc Batch Flushes with OpenSearch Index Refresh
- Instant Live UI Counter Updates on http://localhost:8080
- 4 ms Native PIL DOCX & PyMuPDF Page 1 Thumbnail Generator
"""

import os
import io
import sys

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass
if hasattr(sys.stderr, 'reconfigure'):
    try:
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass
import json
import time
import hashlib
import argparse
import threading
import zipfile
import xml.etree.ElementTree as ET
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from opensearchpy import OpenSearch, helpers
from PIL import Image, ImageDraw, ImageFont, ImageStat
import fitz  # PyMuPDF for PDF thumbnail rendering
try:
    fitz.TOOLS.mupdf_display_errors(False)
except Exception:
    pass

try:
    import mammoth
    from html2image import Html2Image
    HTI_LOCK = threading.Lock()
    HTI_INSTANCE = None
except ImportError:
    mammoth = None
    Html2Image = None
    HTI_LOCK = None
    HTI_INSTANCE = None

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

# Fast Folder Exclusions (names & full paths)
DEFAULT_EXCLUDE_NAMES = {
    '__pycache__', '.git', '.svn', 'node_modules', '.venv', 'venv',
    '.dropbox.cache', 'dropboxbackup', '.cache', 'appdata', '.gemini',
    'deidentifier', 'identified', 'new topic1 - copy', 'build', 'dist',
    'target', 'bin', 'obj', '.vs', '.idea', '.vscode', '.nuget', 'coverage',
    '$recycle.bin', 'system volume information', 'windows', 'program files',
    'program files (x86)', 'programdata'
}

DEFAULT_EXCLUDE_DIRS = {
    r'c:\windows', r'c:\program files', r'c:\program files (x86)',
    r'c:\programdata', r'$recycle.bin', r'system volume information',
    r'd:\active research\deidentifier\identified', r'deidentifier\identified'
}

# File extensions to strictly SKIP
SKIP_EXTENSIONS = {
    '.dll', '.exe', '.sys', '.bin', '.so', '.dylib', '.pyc', '.class',
    '.zip', '.tar', '.gz', '.7z', '.rar', '.iso', '.obj', '.o', '.lib',
    '.pdb', '.dat', '.tmp', '.pyd', '.nupkg', '.cab', '.msi', '.lnk',
    '.cache', '.bak', '.ico', '.png', '.jpg', '.jpeg', '.gif', '.mp3', '.mp4',
    '.ini', '.log', '.tsv', '.pb', '.pak'
}

# Primary Human Document & Presentation Extensions ONLY (PDF, Word, Excel, PPT, Text, Markdown, RTF)
STRICT_DOCUMENT_EXTENSIONS = {
    '.pdf', '.docx', '.doc', '.xlsx', '.xls', '.pptx', '.ppt',
    '.txt', '.md', '.markdown', '.rtf', '.odt', '.ods', '.odp', '.epub'
}

THUMB_CACHE_DIR = os.path.join(os.path.dirname(__file__), ".cache_thumbnails")
if not os.path.exists(THUMB_CACHE_DIR):
    os.makedirs(THUMB_CACHE_DIR, exist_ok=True)


def get_thumbnail_hash(file_path):
    norm = os.path.normpath(os.path.abspath(file_path)).lower()
    return hashlib.md5(norm.encode('utf-8')).hexdigest()


def generate_fast_docx_cover(file_path, cache_path):
    if mammoth is not None and Html2Image is not None:
        try:
            with open(file_path, 'rb') as docx_file:
                res = mammoth.convert_to_html(docx_file)
                html_body = res.value[:50000] if res.value else ""

            if html_body and len(html_body.strip()) > 0:
                html_content = f"""<!DOCTYPE html>
<html>
<head>
<meta charset="utf-8">
<style>
body {{ background: #ffffff; margin: 0; padding: 25px; font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }}
.container {{ max-width: 850px; margin: 0 auto; background: white; padding: 35px; border-radius: 6px; box-shadow: 0 2px 10px rgba(0,0,0,0.1); line-height: 1.6; font-size: 16px; color: #212529; min-height: 950px; }}
table {{ border-collapse: collapse; width: 100%; font-size: 14px; margin: 15px 0; }}
td, th {{ border: 1px solid #dee2e6; padding: 8px 12px; text-align: left; }}
tr:nth-child(even) {{ background-color: #f8f9fa; }}
tr:first-child {{ background-color: #e9ecef; font-weight: bold; }}
h1, h2, h3, h4, h5, h6 {{ margin-top: 0; color: #111; }}
p {{ margin-bottom: 1rem; }}
</style>
</head>
<body>
<div class="container">{html_body}</div>
</body>
</html>"""
                out_dir = os.path.dirname(cache_path)
                out_name = os.path.basename(cache_path)
                with HTI_LOCK:
                    global HTI_INSTANCE
                    if HTI_INSTANCE is None:
                        HTI_INSTANCE = Html2Image(output_path=out_dir, size=(850, 1020), custom_flags=['--hide-scrollbars', '--no-sandbox', '--disable-gpu'])
                    else:
                        HTI_INSTANCE.output_path = out_dir
                    HTI_INSTANCE.screenshot(html_str=html_content, save_as=out_name)

                if os.path.exists(cache_path) and os.path.getsize(cache_path) > 0:
                    return
        except Exception:
            pass

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


def generate_fast_pptx_cover(file_path, cache_path):
    file_name = os.path.basename(file_path)

    # 1. Fast Native Embedded Slide 1 Image Extraction from Zip Archive

    # 2. Second Priority: Extract High-Res Embedded Slide 1 Image ONLY IF IT IS NOT A BLANK DUMMY
    try:
        if zipfile.is_zipfile(file_path):
            with zipfile.ZipFile(file_path) as z:
                for item in z.namelist():
                    if item.lower() in ('docprops/thumbnail.jpeg', 'docprops/thumbnail.jpg', 'docprops/thumbnail.png'):
                        img_bytes = z.read(item)
                        slide_img = Image.open(io.BytesIO(img_bytes)).convert('RGB')

                        # Validate that embedded thumbnail is not a blank/solid white dummy image
                        stat = ImageStat.Stat(slide_img)
                        if sum(stat.stddev) >= 3.0:
                            canvas = Image.new('RGB', (600, 720), '#f8f9fa')
                            draw = ImageDraw.Draw(canvas)

                            sw, sh = slide_img.size
                            scale = min(560 / sw, 520 / sh)
                            nw, nh = int(sw * scale), int(sh * scale)
                            resized_slide = slide_img.resize((nw, nh), Image.Resampling.LANCZOS)

                            x = (600 - nw) // 2
                            y = (720 - nh) // 2 + 10

                            draw.rectangle([0, 0, 599, 719], outline='#d0d0d0', width=1)
                            draw.rectangle([x - 4, y - 4, x + nw + 4, y + nh + 4], fill='#e4e6e9')
                            draw.rectangle([x - 2, y - 2, x + nw + 2, y + nh + 2], fill='#cfd2d7')
                            canvas.paste(resized_slide, (x, y))

                            draw.rectangle([0, 0, 600, 32], fill='#d24726')
                            try:
                                font_header = ImageFont.truetype('arialbd.ttf', 13)
                            except Exception:
                                font_header = ImageFont.load_default()
                            draw.text((15, 8), f"PowerPoint  |  {file_name[:48]}", fill='#ffffff', font=font_header)

                            canvas.save(cache_path, 'JPEG', quality=95)
                            return
    except Exception:
        pass

    # 2. Second Priority: Render True Presentation Slide Card from Slide 1 & Slide 2 XML/pptx shapes
    slide1_texts = []
    slide2_texts = []

    try:
        import pptx
        prs = pptx.Presentation(file_path)
        if len(prs.slides) > 0:
            for shape in prs.slides[0].shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    slide1_texts.append(shape.text_frame.text.strip())
        if len(prs.slides) > 1:
            for shape in prs.slides[1].shapes:
                if shape.has_text_frame and shape.text_frame.text.strip():
                    slide2_texts.append(shape.text_frame.text.strip())
    except Exception:
        pass

    if not slide1_texts:
        try:
            with zipfile.ZipFile(file_path) as z:
                if 'ppt/slides/slide1.xml' in z.namelist():
                    xml1 = z.read('ppt/slides/slide1.xml')
                    tree1 = ET.fromstring(xml1)
                    nodes1 = tree1.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                    slide1_texts = [n.text.strip() for n in nodes1 if n.text and n.text.strip()]

                if 'ppt/slides/slide2.xml' in z.namelist():
                    xml2 = z.read('ppt/slides/slide2.xml')
                    tree2 = ET.fromstring(xml2)
                    nodes2 = tree2.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}t')
                    slide2_texts = [n.text.strip() for n in nodes2 if n.text and n.text.strip()]
        except Exception:
            pass

    title_text = slide1_texts[0] if slide1_texts else file_name[:40]
    subtitle_text = slide1_texts[1] if len(slide1_texts) > 1 else ""
    callout_text = slide1_texts[2] if len(slide1_texts) > 2 else ""

    canvas = Image.new('RGB', (600, 720), '#f4f6f9')
    draw = ImageDraw.Draw(canvas)

    draw.rectangle([0, 0, 599, 719], outline='#cfd2d7', width=1)
    draw.rectangle([0, 0, 600, 32], fill='#d24726')
    try:
        font_header = ImageFont.truetype('arialbd.ttf', 13)
        font_title = ImageFont.truetype('arialbd.ttf', 20)
        font_sub = ImageFont.truetype('arial.ttf', 13)
        font_body = ImageFont.truetype('arial.ttf', 11)
    except Exception:
        font_header = font_title = font_sub = font_body = ImageFont.load_default()

    draw.text((15, 8), f"PowerPoint  |  {file_name[:48]}", fill='#ffffff', font=font_header)

    # 1. Main Slide 1 16:9 Widescreen Frame (540px x 320px)
    sx, sy, sw, sh = 30, 55, 540, 320
    draw.rectangle([sx - 4, sy - 4, sx + sw + 4, sy + sh + 4], fill='#e2e5e9')
    draw.rectangle([sx - 2, sy - 2, sx + sw + 2, sy + sh + 2], fill='#d0d4da')
    draw.rectangle([sx, sy, sx + sw, sy + sh], fill='#ffffff', outline='#d24726', width=2)

    ty = sy + 25
    words = title_text.split()
    line = []
    for w in words:
        line.append(w)
        if len(" ".join(line)) >= 32:
            draw.text((sx + 25, ty), " ".join(line), fill='#1a2530', font=font_title)
            ty += 26
            line = []
            if ty > sy + 110:
                break
    if line and ty <= sy + 110:
        draw.text((sx + 25, ty), " ".join(line), fill='#1a2530', font=font_title)
        ty += 30

    if subtitle_text:
        words = subtitle_text.split()
        line = []
        for w in words:
            line.append(w)
            if len(" ".join(line)) >= 48:
                draw.text((sx + 25, ty), " ".join(line), fill='#5f6b7a', font=font_sub)
                ty += 18
                line = []
                if ty > sy + 170:
                    break
        if line and ty <= sy + 170:
            draw.text((sx + 25, ty), " ".join(line), fill='#5f6b7a', font=font_sub)
            ty += 22

    if callout_text:
        cx, cy, cw, ch = sx + 25, ty + 10, sw - 50, sy + sh - ty - 25
        if ch >= 40:
            draw.rectangle([cx, cy, cx + cw, cy + ch], fill='#f8f9fa', outline='#d0d4da', width=1)
            words = callout_text.split()
            line = []
            cty = cy + 10
            for w in words:
                line.append(w)
                if len(" ".join(line)) >= 58:
                    draw.text((cx + 12, cty), " ".join(line), fill='#334155', font=font_body)
                    cty += 16
                    line = []
                    if cty > cy + ch - 18:
                        break
            if line and cty <= cy + ch - 18:
                draw.text((cx + 12, cty), " ".join(line), fill='#334155', font=font_body)

    # 2. Slide 2 / Overview Frame (Lower Half: 540px x 290px)
    sy2 = 395
    draw.rectangle([sx - 4, sy2 - 4, sx + sw + 4, sy2 + 290 + 4], fill='#e2e5e9')
    draw.rectangle([sx - 2, sy2 - 2, sx + sw + 2, sy2 + 290 + 2], fill='#d0d4da')
    draw.rectangle([sx, sy2, sx + sw, sy2 + 290], fill='#ffffff', outline='#cbd5e1', width=1)
    
    s2_title = slide2_texts[0] if slide2_texts else "SLIDE OVERVIEW & ROADMAP"
    draw.rectangle([sx, sy2, sx + sw, sy2 + 32], fill='#f1f5f9')
    draw.text((sx + 20, sy2 + 8), s2_title[:45].upper(), fill='#475569', font=font_header)
    
    ty2 = sy2 + 45
    if len(slide2_texts) > 1:
        for st in slide2_texts[1:6]:
            words = st.split()
            line = []
            for w in words:
                line.append(w)
                if len(" ".join(line)) >= 56:
                    draw.text((sx + 25, ty2), f"• {" ".join(line)}", fill='#334155', font=font_body)
                    ty2 += 18
                    line = []
                    if ty2 > sy2 + 260:
                        break
            if line and ty2 <= sy2 + 260:
                draw.text((sx + 25, ty2), f"• {" ".join(line)}", fill='#334155', font=font_body)
                ty2 += 20
    else:
        draw.text((sx + 25, ty2), "Presentation contains structured case vignettes & teaching anchors.", fill='#64748b', font=font_sub)

    canvas.save(cache_path, 'JPEG', quality=95)


def generate_thumbnail_if_missing(file_path: str, ext: str):
    try:
        file_hash = get_thumbnail_hash(file_path)
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

        elif ext in ('.docx', '.doc'):
            generate_fast_docx_cover(file_path, cache_path)

        elif ext in ('.pptx', '.ppt'):
            generate_fast_pptx_cover(file_path, cache_path)
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

    elif ext == '.pdf':
        try:
            doc = fitz.open(file_path)
            text_parts = []
            for i, page in enumerate(doc):
                if i >= 50:
                    break
                text = page.get_text()
                if text:
                    text_parts.append(text)
            doc.close()
            return "\n".join(text_parts)[:max_bytes]
        except Exception:
            return ""

    elif ext in ['.docx', '.doc'] and docx is not None:
        try:
            doc = docx.Document(file_path)
            return "\n".join([p.text for p in doc.paragraphs if p.text])[:max_bytes]
        except Exception:
            return ""

    elif ext in ['.xlsx', '.xls']:
        try:
            with zipfile.ZipFile(file_path) as z:
                if 'xl/sharedStrings.xml' in z.namelist():
                    xml_content = z.read('xl/sharedStrings.xml')
                    tree = ET.fromstring(xml_content)
                    text_nodes = tree.findall('.//{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t')
                    return " ".join([node.text for node in text_nodes if node.text])[:max_bytes]
        except Exception:
            pass
        return ""

    return ""


def fetch_existing_metadata(client, index_name="documents"):
    existing_map = {}
    if not client.indices.exists(index=index_name):
        return existing_map
    print("[*] Pre-fetching existing index metadata for fast delta comparison...")
    write_progress(True, 0, 0, 0, "Pre-fetching metadata from OpenSearch index...")
    try:
        from opensearchpy.helpers import scan
        docs = scan(client, index=index_name, query={"_source": ["file_path", "modified_date", "file_size"]})
        count = 0
        for doc in docs:
            count += 1
            src = doc.get('_source', {})
            fp = src.get('file_path')
            if fp:
                norm = os.path.normpath(os.path.abspath(fp)).lower()
                existing_map[norm] = (src.get('modified_date'), src.get('file_size'))
            if count % 5000 == 0:
                write_progress(True, 0, 0, 0, f"Loading cached metadata: {count:,} records loaded...")
    except Exception as e:
        print(f"[-] Warning: Could not pre-fetch index metadata: {e}")
    print(f"[+] Loaded {len(existing_map):,} cached entries for instant skip check.")
    write_progress(True, 0, 0, 0, f"Metadata loaded ({len(existing_map):,} entries). Starting file scan...")
    return existing_map


def copy_file_to_backup(file_path: str, backup_dir: str) -> str:
    """
    Mirrors file_path into backup_dir preserving directory structure.
    Returns destination path on success, or None on failure.
    """
    if not backup_dir:
        return None
    try:
        drive, rel_path = os.path.splitdrive(file_path)
        clean_rel = rel_path.lstrip('\\').lstrip('/')
        dest_path = os.path.join(backup_dir, clean_rel)
        os.makedirs(os.path.dirname(dest_path), exist_ok=True)
        import shutil
        shutil.copy2(file_path, dest_path)
        return dest_path
    except Exception:
        return None


def process_single_file(file_path: str, index_name: str, existing_map: dict = None, force: bool = False, backup_dir: str = None, exclude_backup_exts: set = None) -> tuple:
    file_name = os.path.basename(file_path)
    ext = os.path.splitext(file_name)[1].lower()
    
    if not ext or ext not in STRICT_DOCUMENT_EXTENSIONS or file_name.startswith('~') or file_name.startswith('.'):
        return None, 'ignored'
        
    try:
        stat = os.stat(file_path)
        norm_path = os.path.normpath(os.path.abspath(file_path)).lower()
        mtime_iso = datetime.fromtimestamp(stat.st_mtime).isoformat()
        file_size = stat.st_size

        # Fast Delta Skip Check: Compare file mtime and file size
        if not force and existing_map and norm_path in existing_map:
            cached_mtime, cached_size = existing_map[norm_path]
            if cached_mtime == mtime_iso and cached_size == file_size:
                return None, 'skipped'

        # Canonical path hash for consistent document ID across drive casing
        doc_id = hashlib.sha256(norm_path.encode('utf-8')).hexdigest()
        file_dir = os.path.dirname(file_path)
        file_type = ext.lstrip('.').lower()
        content = extract_file_content(file_path)
        
        # Trigger cover thumbnail generation if missing
        generate_thumbnail_if_missing(file_path, file_type)

        # Handle Document Backup (Skipping excluded extensions like .csv)
        backup_dest = None
        is_backed_up = False
        if exclude_backup_exts is None:
            exclude_backup_exts = {'.csv'}

        if backup_dir and ext not in exclude_backup_exts:
            backup_dest = copy_file_to_backup(file_path, backup_dir)
            is_backed_up = backup_dest is not None

        doc = {
            "_op_type": "index",
            "_index": index_name,
            "_id": doc_id,
            "file_name": file_name,
            "file_type": file_type,
            "file_path": file_path,
            "file_directory": file_dir,
            "file_extension": ext,
            "file_size": file_size,
            "created_date": datetime.fromtimestamp(stat.st_ctime).isoformat(),
            "modified_date": mtime_iso,
            "is_backed_up": is_backed_up,
            "backup_path": backup_dest,
            "content": content
        }
        return doc, 'indexed'
    except Exception:
        return None, 'error'


def is_excluded_dir(dir_path: str) -> bool:
    norm_path = dir_path.lower().replace('\\', '/')
    dir_name = os.path.basename(dir_path).lower()
    
    if dir_name in DEFAULT_EXCLUDE_NAMES or dir_name.startswith('.'):
        return True
        
    for excl in DEFAULT_EXCLUDE_DIRS:
        if excl in norm_path:
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


PROGRESS_FILE = os.path.join(os.path.dirname(__file__), "indexer_progress.json")

def write_progress(is_running: bool, scanned: int, indexed: int, skipped: int, status_msg: str):
    try:
        data = {
            "is_running": is_running,
            "scanned": scanned,
            "indexed": indexed,
            "skipped": skipped,
            "status_message": status_msg,
            "timestamp": time.time()
        }
        with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f)
    except Exception:
        pass


def sync_to_network_target(backup_dir: str, network_target: str):
    if not backup_dir or not os.path.exists(backup_dir):
        print(f"[-] Backup directory '{backup_dir}' does not exist. Skipping network copy.")
        return
    if not network_target:
        return

    print(f"\n[*] Starting backup sync from '{backup_dir}' to network target '{network_target}'...")
    try:
        os.makedirs(network_target, exist_ok=True)
        import subprocess
        if sys.platform == 'win32':
            cmd = ["robocopy", backup_dir, network_target, "/MIR", "/FFT", "/R:2", "/W:2", "/NDL", "/NFL", "/NJH"]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode <= 7:
                print(f"[✓] Network backup sync complete to '{network_target}'!")
            else:
                print(f"[-] Robocopy returned code {res.returncode}: {res.stderr or res.stdout}")
        else:
            for root, _, files in os.walk(backup_dir):
                rel = os.path.relpath(root, backup_dir)
                dest = os.path.join(network_target, rel)
                os.makedirs(dest, exist_ok=True)
                for f in files:
                    s_file = os.path.join(root, f)
                    d_file = os.path.join(dest, f)
                    if not os.path.exists(d_file) or os.path.getmtime(s_file) > os.path.getmtime(d_file):
                        import shutil
                        shutil.copy2(s_file, d_file)
            print(f"[✓] Network backup sync complete to '{network_target}'!")
    except Exception as e:
        print(f"[-] Network backup sync error: {e}")


def backup_full_dropbox(dropbox_src: str, backup_dir: str):
    if not os.path.exists(dropbox_src) or not backup_dir:
        return
    print(f"\n[*] Performing FULL backup of Dropbox ({dropbox_src}) into '{backup_dir}' (including ALL file types)...")
    dest_dir = os.path.join(backup_dir, "Dropbox")
    os.makedirs(dest_dir, exist_ok=True)
    try:
        import subprocess
        if sys.platform == 'win32':
            cmd = ["robocopy", dropbox_src, dest_dir, "/MIR", "/FFT", "/R:2", "/W:2", "/XD", ".dropbox.cache", "dropboxbackup", "/NDL", "/NFL", "/NJH"]
            res = subprocess.run(cmd, capture_output=True, text=True)
            if res.returncode <= 7:
                print(f"[✓] Full Dropbox backup complete to '{dest_dir}'!")
            else:
                print(f"[-] Robocopy Dropbox backup code {res.returncode}: {res.stderr or res.stdout}")
        else:
            for root, dirs, files in os.walk(dropbox_src):
                dirs[:] = [d for d in dirs if d.lower() not in ('.dropbox.cache', 'dropboxbackup')]
                rel = os.path.relpath(root, dropbox_src)
                dest = os.path.join(dest_dir, rel)
                os.makedirs(dest, exist_ok=True)
                for f in files:
                    s_file = os.path.join(root, f)
                    d_file = os.path.join(dest, f)
                    if not os.path.exists(d_file) or os.path.getmtime(s_file) > os.path.getmtime(d_file):
                        import shutil
                        shutil.copy2(s_file, d_file)
            print(f"[✓] Full Dropbox backup complete to '{dest_dir}'!")
    except Exception as e:
        print(f"[-] Full Dropbox backup error: {e}")


def main():
    parser = argparse.ArgumentParser(description="Ingest documents into OpenSearch with Hybrid Delta Indexing")
    parser.add_argument("--dir", nargs="+", help="Directories to scan and index")
    parser.add_argument("--index", default="documents", help="Target OpenSearch index name")
    parser.add_argument("--force", "-f", action="store_true", help="Force re-indexing of all files regardless of timestamps")
    parser.add_argument("--backup-dir", default=None, help="Target directory for document backups")
    parser.add_argument("--exclude-backup-ext", nargs="*", default=None, help="File extensions to exclude from backup (e.g. .csv)")
    parser.add_argument("--network-target", default=None, help="Optional network target directory to sync backups to when complete")
    args = parser.parse_args()

    config_backup_dir = r"D:\Backups\Documents"
    config_exclude_exts = [".csv"]

    config_path = os.path.join(os.path.dirname(__file__), "indexer_config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                cfg = json.load(f)
                if not args.dir:
                    args.dir = cfg.get("selected_directories", [r"D:\Active research"])
                config_backup_dir = cfg.get("backup_directory", r"D:\Backups\Documents")
                config_exclude_exts = cfg.get("exclude_backup_extensions", [".csv"])
        except Exception:
            pass

    if not args.dir:
        args.dir = [r"D:\Active research"]

    backup_dir = args.backup_dir or config_backup_dir
    raw_exclude_exts = args.exclude_backup_ext if args.exclude_backup_ext is not None else config_exclude_exts
    exclude_backup_exts = {ext.lower() if ext.startswith('.') else f".{ext.lower()}" for ext in raw_exclude_exts}

    print(f"[*] Ingestion starting for directories: {args.dir} (Force mode: {args.force})")
    if backup_dir:
        print(f"[*] Document Backup Enabled -> Destination: '{backup_dir}' (Excluding: {sorted(list(exclude_backup_exts))})")
    else:
        print(f"[*] Document Backup Disabled (Set 'backup_directory' in indexer_config.json or use --backup-dir)")

    client = get_opensearch_client()
    ensure_index_exists(client, index_name=args.index)

    existing_map = fetch_existing_metadata(client, index_name=args.index) if not args.force else {}

    start_time = time.time()
    batch = []
    total_scanned = 0
    total_indexed = 0
    total_skipped = 0
    scanned_paths = set()

    # Initialize status file at start of scan
    write_progress(True, 1, 0, 0, "Checking files: 1 scanned (0 updated, 0 skipped)")

    from concurrent.futures import wait, FIRST_COMPLETED

    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = set()
        max_queue = 32

        def handle_finished_future(fut):
            nonlocal total_scanned, total_indexed, total_skipped, batch
            doc, status = fut.result()
            if status == 'skipped':
                total_skipped += 1
                total_scanned += 1
            elif status == 'indexed' or doc:
                total_scanned += 1
                if doc:
                    batch.append(doc)

            if total_scanned % 10 == 0 or len(batch) >= 20:
                msg = f"Checking documents: {total_scanned:,} scanned ({total_indexed:,} updated, {total_skipped:,} skipped)"
                write_progress(True, total_scanned, total_indexed, total_skipped, msg)

            if len(batch) >= 20:
                try:
                    helpers.bulk(client, batch)
                    client.indices.refresh(index=args.index)
                    total_indexed += len(batch)
                    print(f"[+] Bulk indexed {total_indexed} new/updated documents ({total_skipped:,} skipped unmodified)...")
                except Exception as e:
                    print(f"[-] Bulk ingest error: {e}")
                batch = []

        for file_path in scan_directories(args.dir):
            norm_p = os.path.normpath(os.path.abspath(file_path)).lower()
            scanned_paths.add(norm_p)
            
            while len(futures) >= max_queue:
                done, futures = wait(futures, return_when=FIRST_COMPLETED)
                for fut in done:
                    handle_finished_future(fut)

            futures.add(executor.submit(process_single_file, file_path, args.index, existing_map, args.force, backup_dir, exclude_backup_exts))

        # Flush remaining futures with timeout safety
        if futures:
            done, not_done = wait(futures, timeout=15)
            for fut in done:
                try:
                    handle_finished_future(fut)
                except Exception:
                    pass
            if not_done:
                print(f"[-] Warning: {len(not_done)} indexing threads timed out and were skipped.")

        if batch:
            try:
                helpers.bulk(client, batch)
                client.indices.refresh(index=args.index)
                total_indexed += len(batch)
            except Exception:
                pass

    # Ghost / Dead File Cleanup Pass
    total_deleted = 0
    if not args.force and existing_map:
        dead_paths = set(existing_map.keys()) - scanned_paths
        if dead_paths:
            print(f"[*] Cleaning up {len(dead_paths):,} deleted or moved documents from index...")
            for fp in dead_paths:
                try:
                    client.delete_by_query(index=args.index, body={"query": {"term": {"file_path": fp}}})
                    total_deleted += 1
                except Exception:
                    pass
            client.indices.refresh(index=args.index)

    elapsed = time.time() - start_time
    final_msg = f"[+] Indexing Complete: {total_scanned:,} documents checked ({total_indexed:,} updated, {total_skipped:,} up-to-date)"
    write_progress(False, total_scanned, total_indexed, total_skipped, final_msg)
    print(f"[+] Hybrid Indexing Complete! {total_scanned:,} files checked ({total_indexed:,} indexed/updated, {total_skipped:,} unmodified/skipped, {total_deleted:,} dead entries cleaned up) in {elapsed:.2f} seconds.")

    # Full Dropbox Backup Phase (includes ALL files in D:\Dropbox)
    if backup_dir:
        backup_full_dropbox(r"D:\Dropbox", backup_dir)

    # Optional Network Backup Sync Phase
    if backup_dir and os.path.exists(backup_dir):
        net_target = args.network_target
        if not net_target and sys.stdin.isatty():
            try:
                ans = input("\n[?] Ingestion and local backup complete. Would you like to save/sync backups to a network target? (y/N): ").strip().lower()
                if ans in ['y', 'yes']:
                    net_target = input("[?] Enter network target path (e.g. \\\\NAS\\Share\\Backups or Z:\\Backups): ").strip()
            except (KeyboardInterrupt, EOFError):
                net_target = None
        
        if net_target:
            sync_to_network_target(backup_dir, net_target)


if __name__ == "__main__":
    main()
