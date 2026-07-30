"""
Dedicated Local Web Search Interface for OpenSearch.
Includes smart extension parsing (pdf, docx, xlsx, txt, md, pptx) AND sorting options
(Relevance, Newest First, Oldest First, File Name A-Z, File Size Largest First).
"""

import os
import re
import json
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import parse_qs, urlparse
from opensearchpy import OpenSearch

OPENSEARCH_HOST = "localhost"
OPENSEARCH_PORT = 9200
KNOWN_EXTENSIONS = {'pdf', 'docx', 'doc', 'xlsx', 'xls', 'txt', 'csv', 'md', 'rtf', 'pptx', 'ppt'}


def get_client():
    return OpenSearch(hosts=[{"host": OPENSEARCH_HOST, "port": OPENSEARCH_PORT}], use_ssl=False)


def parse_smart_query(user_query: str, sort_by: str = "relevance"):
    """
    Parses plain search queries and applies selected sort order.
    """
    tokens = user_query.strip().split()
    file_types = []
    text_terms = []

    for token in tokens:
        clean_token = token.lower().lstrip('.')
        if clean_token in KNOWN_EXTENSIONS:
            file_types.append(clean_token)
        else:
            text_terms.append(token)

    must_conditions = []
    
    if file_types:
        must_conditions.append({"terms": {"file_type": file_types}})
        
    if text_terms:
        search_text = " ".join(text_terms)
        must_conditions.append({
            "multi_match": {
                "query": search_text,
                "fields": ["content^3", "file_name^5"],
                "type": "best_fields"
            }
        })
        
    if not must_conditions:
        query_body = {"query": {"match_all": {}}}
    else:
        query_body = {"query": {"bool": {"must": must_conditions}}}

    # Apply sorting options
    if sort_by == "date_desc":
        query_body["sort"] = [{"modified_date": {"order": "desc"}}]
    elif sort_by == "date_asc":
        query_body["sort"] = [{"modified_date": {"order": "asc"}}]
    elif sort_by == "name_asc":
        query_body["sort"] = [{"file_name.keyword": {"order": "asc"}}]
    elif sort_by == "size_desc":
        query_body["sort"] = [{"file_size": {"order": "desc"}}]
    else:
        # Default: Best Match / Relevance
        if must_conditions and text_terms:
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
    <title>OpenSearch File Finder</title>
    <style>
        body { font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; background-color: #f8f9fa; margin: 0; padding: 20px; color: #333; }
        .container { max-width: 1100px; margin: 0 auto; }
        .header { text-align: center; margin-bottom: 25px; }
        .search-box { display: flex; gap: 10px; margin-bottom: 20px; align-items: center; }
        input[type="text"] { flex: 1; padding: 14px; font-size: 16px; border: 2px solid #ced4da; border-radius: 6px; }
        select { padding: 14px; font-size: 15px; border: 2px solid #ced4da; border-radius: 6px; background: white; cursor: pointer; }
        button { padding: 14px 28px; font-size: 16px; background-color: #007bff; color: white; border: none; border-radius: 6px; cursor: pointer; }
        button:hover { background-color: #0056b3; }
        .stats { margin-bottom: 15px; color: #6c757d; font-weight: 500; }
        .result-card { background: white; border-radius: 8px; padding: 18px; margin-bottom: 12px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); }
        .result-header { display: flex; justify-content: space-between; align-items: center; }
        .file-title { font-weight: bold; font-size: 17px; color: #007bff; text-decoration: none; }
        .badge { background: #e9ecef; padding: 4px 10px; border-radius: 12px; font-size: 13px; text-transform: uppercase; font-weight: 600; color: #495057; }
        .file-path { color: #6c757d; font-size: 13px; margin: 4px 0 10px 0; word-break: break-all; }
        .snippet { background: #f8f9fa; padding: 10px; border-left: 4px solid #007bff; border-radius: 4px; font-size: 14px; color: #495057; line-height: 1.5; }
        mark { background-color: #ffe066; padding: 2px 4px; border-radius: 3px; font-weight: bold; }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>🔍 OpenSearch File Finder</h2>
            <p>Smart search across your documents with custom sorting and automatic file type parsing (<code>pdf</code>, <code>docx</code>, <code>xlsx</code>, <code>md</code>).</p>
        </div>
        <form class="search-box" method="GET" action="/">
            <input type="text" name="q" value="{QUERY}" placeholder="Try: 'pdf patient', 'docx quality', 'xlsx', 'pathology'..." autofocus>
            <select name="sort" onchange="this.form.submit()">
                <option value="relevance" {SORT_RELEVANCE}>Best Match (Relevance)</option>
                <option value="date_desc" {SORT_DATE_DESC}>Date: Newest First</option>
                <option value="date_asc" {SORT_DATE_ASC}>Date: Oldest First</option>
                <option value="name_asc" {SORT_NAME_ASC}>File Name (A - Z)</option>
                <option value="size_desc" {SORT_SIZE_DESC}>File Size: Largest First</option>
            </select>
            <button type="submit">Search</button>
        </form>
        <div class="stats">{STATS}</div>
        {RESULTS}
    </div>
</body>
</html>
"""


class SearchHandler(SimpleHTTPRequestHandler):
    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)
        query_str = params.get('q', [''])[0].strip()
        sort_by = params.get('sort', ['relevance'])[0].strip()

        stats_html = ""
        results_html = ""

        # Map selected sort state
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
                        
                        highlights = hit.get('highlight', {}).get('content', [])
                        if highlights:
                            snippet_text = " ... ".join(highlights)
                        else:
                            snippet_text = (src.get('content', '')[:300] + "...") if src.get('content') else "No preview text available."

                        card = f"""
                        <div class="result-card">
                            <div class="result-header">
                                <span class="file-title">{fname}</span>
                                <span class="badge">{ftype}</span>
                            </div>
                            <div class="file-path">📁 {fpath}</div>
                            <div class="snippet">{snippet_text}</div>
                        </div>
                        """
                        cards.append(card)
                    results_html = "\n".join(cards)
            except Exception as e:
                results_html = f"<div class='result-card' style='color:red;'>Error executing search: {e}</div>"

        html_out = HTML_TEMPLATE.replace("{QUERY}", query_str).replace("{STATS}", stats_html).replace("{RESULTS}", results_html)
        for key, val in sort_state.items():
            html_out = html_out.replace(f"{{{key}}}", val)

        self.send_response(200)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.end_headers()
        self.wfile.write(html_out.encode('utf-8'))


def main():
    server = HTTPServer(('localhost', 8080), SearchHandler)
    print("[+] OpenSearch Smart Search Server running at http://localhost:8080")
    server.serve_forever()


if __name__ == "__main__":
    main()
