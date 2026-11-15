# OpenSearch Local File Finder & Ingestion Engine 🔍

A high-performance, 100% open-source local document search engine and web finder powered by **OpenSearch** and **Python**. Designed as a fast, private, zero-cost replacement for desktop search tools (such as X1 or dtSearch) on large document collections (1M+ files).

---

## ✨ Features

- **🚀 High-Throughput Multithreaded Ingestion:** Multi-worker directory scanner (`ingest_documents.py`) extracting text from PDFs, Word docs (`.docx`), Excel spreadsheets (`.xlsx`), Text, Markdown (`.md`), and PowerPoint files with SHA-256 deduplication.
- **⚡ Sub-Second Search Response:** Sub-50ms query response times even across millions of documents.
- **Search controls:** Search contents, filenames, or both; use exact phrases, `AND`, `OR`, `NOT`, and parentheses. Filter by folder, file type, modified dates, and indexed local/network paths.
- **↕️ Custom Sorting Options:** Sort search results instantly by Relevance (Best Match), Date (Newest/Oldest), File Name (A-Z), or File Size (Largest First).
- **Local indexing and search:** OpenSearch and the Python apps run locally. Review deployment security and data-handling requirements before using sensitive documents; the browser previews load some JavaScript libraries from public CDNs.
- **🛡️ Memory Optimized:** Includes pre-configured WSL2 ceiling settings (`memory=6GB`) to prevent host RAM exhaustion and system UI freezes.

---

## 📂 Repository Structure

```
OpenSearch/
├── docker-compose.yml       # Docker configuration for OpenSearch (port 9200) & Dashboards (port 5601)
├── index_settings.json      # Schema mappings, term vectors, and analyzer configurations
├── ingest_documents.py      # High-performance multithreaded document ingestion pipeline
├── search_app.py            # Lightweight Smart Search web finder (port 8080)
├── Launch_OpenSearch.bat    # 1-Click launcher (starts Docker, containers, server & opens Edge)
├── requirements.txt         # Python dependencies (opensearch-py, pypdf, python-docx, etc.)
├── LICENSE                  # Apache 2.0 Open Source License
└── README.md                # Project documentation
```

---

## ⚙️ Prerequisites

1. **[Docker Desktop](https://www.docker.com/products/docker-desktop/)** (version 24.0+ with WSL2 backend enabled).
2. **Python 3.10+** (with `pip` installed).

---

## 🚀 Quick Start

### 1. Clone & Install Dependencies
```bash
git clone https://github.com/prdexter/OpenSearchFileFinder.git
cd OpenSearchFileFinder
pip install -r requirements.txt
```

### 2. Start OpenSearch Engine
```bash
docker compose up -d
```

### 3. Ingest Documents
Run the ingestion pipeline on your target directories:
```bash
python ingest_documents.py --dir "C:\Users\YourName\Documents" "D:\Research"
```

### 4. Launch Smart Search Finder
Start the local search web finder on port `8080`:
```bash
python search_app.py
```
Open **`http://localhost:8080`** in your browser!

---

## 🖥️ 1-Click Windows Desktop Launcher

Run `Launch_OpenSearch.bat` or use the generated desktop shortcut to:
1. Verify and start Docker Desktop.
2. Spin up OpenSearch Docker containers.
3. Start `search_app.py` server in the background.
4. Launch Microsoft Edge directly to `http://localhost:8080`.

---

## Local configuration and backups

Copy `indexer_config.example.json` to `indexer_config.json` and adjust the directories and backup destinations for your machine. The app also saves these settings through its directory configuration dialog. Local configuration, reports, thumbnail caches, and progress files are excluded from Git.

Search uses the existing `documents` index; changing the search interface does not rebuild it. Network filters only match indexed network paths and do not scan NAS backups.

Backup errors propagate to the final app status. Windows permission repair can be installed once with `install_backup_permission_helper.ps1` in Administrator PowerShell. The helper is configured for the original workstation/account and restricts repairs to `D:\Active research`; review and adapt those settings before installing on another computer. It runs on demand while that account is signed in. `retry_active_research_backup.py` retries that configured mirror and then performs a read-only comparison.

`run_overnight_maintenance.py` and `Run_Overnight_Maintenance.bat` perform ingestion and thumbnail maintenance. The `allfields*.py` scripts are local PHI field-review utilities; review their input/output paths before running them.

## Tests

On Windows, from the project directory:

```powershell
python -m unittest discover -s tests -p '*_test.py' -v
```

The suite includes small temporary-file Robocopy and permission-helper tests. It does not run production ingestion or NAS backups. Keep the separately backed-up source documents and OpenSearch Docker data volume when restoring the app; this Git repository contains code, not the search database or document backups.

## 📄 License

Distributed under the **Apache 2.0 License**. See [`LICENSE`](LICENSE) for details.
