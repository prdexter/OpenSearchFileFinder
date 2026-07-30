# OpenSearch Local File Finder & Ingestion Engine 🔍

A high-performance, 100% open-source local document search engine and web finder powered by **OpenSearch** and **Python**. Designed as a fast, private, zero-cost replacement for desktop search tools (such as X1 or dtSearch) on large document collections (1M+ files).

---

## ✨ Features

- **🚀 High-Throughput Multithreaded Ingestion:** Multi-worker directory scanner (`ingest_documents.py`) extracting text from PDFs, Word docs (`.docx`), Excel spreadsheets (`.xlsx`), Text, Markdown (`.md`), and PowerPoint files with SHA-256 deduplication.
- **⚡ Sub-Second Search Response:** Sub-50ms query response times even across millions of documents.
- **🧠 Smart Query Parsing:** Automatically parses 1-word extension queries (`pdf`, `docx`, `xlsx`, `md`) as file type filters without requiring complex syntax or quotes (`pdf patient` ➔ `file_type: pdf AND content: patient`).
- **↕️ Custom Sorting Options:** Sort search results instantly by Relevance (Best Match), Date (Newest/Oldest), File Name (A-Z), or File Size (Largest First).
- **🔒 100% Local & PHI Compliant:** All data processing, indexing, and UI servers run strictly locally inside Docker containers. Zero data is sent to external clouds.
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
git clone https://github.com/your-username/opensearch-file-finder.git
cd opensearch-file-finder
pip install -r requirements.txt
```

### 2. Start OpenSearch Engine
```bash
docker compose up -d
```

### 3. Ingest Documents
Run the ingestion pipeline on your target directories:
```bash
python ingest_documents.py --dir "C:\Users\YourName\Documents" "D:\Research" --threads 8
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

## 📄 License

Distributed under the **Apache 2.0 License**. See [`LICENSE`](file:///D:/Active%20research/OpenSearch/LICENSE) for details.
