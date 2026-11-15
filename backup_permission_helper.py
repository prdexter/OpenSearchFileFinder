"""Request narrowly scoped, elevated read-permission repair from Windows Task Scheduler."""
import json
import os
from pathlib import Path
import re
import subprocess
import time
import uuid

TASK_NAME = "OpenSearch-RepairActiveResearchPermissions"
SOURCE_ROOT = r"D:\Active research"


def denied_source_paths(log_path):
    text = Path(log_path).read_text(encoding="utf-16", errors="replace")
    matches = re.findall(
        r"ERROR 5 \([^\r\n]*?\) (?:Copying Directory|Copying File|Accessing Source Directory) (D:\\Active research\\[^\r\n]+)",
        text, flags=re.IGNORECASE,
    )
    root = os.path.normcase(os.path.abspath(SOURCE_ROOT)) + os.sep
    paths = sorted({os.path.abspath(path.strip().rstrip('\\')) for path in matches}, key=len)
    selected = []
    for path in paths:
        normalized = os.path.normcase(path)
        if not normalized.startswith(root):
            continue
        if not any(normalized.startswith(os.path.normcase(parent) + os.sep) for parent in selected):
            selected.append(path)
    return selected


def repair_denied_paths(log_path, timeout=600):
    paths = denied_source_paths(log_path)
    if not paths:
        return False
    state = Path(os.environ["LOCALAPPDATA"]) / "OpenSearch" / "permission_repair"
    if not (state / "installed.json").exists():
        print("[-] Automatic permission repair is not installed. Run install_backup_permission_helper.ps1 as Administrator once.")
        return False
    request_id = uuid.uuid4().hex
    request_path = state / f"{request_id}.request.json"
    result_path = state / f"{request_id}.result.json"
    temporary_path = request_path.with_suffix(".tmp")
    temporary_path.write_text(json.dumps({"request_id": request_id, "paths": paths}), encoding="utf-8")
    temporary_path.replace(request_path)
    launch = subprocess.run(["schtasks.exe", "/Run", "/TN", TASK_NAME], capture_output=True, text=True)
    if launch.returncode:
        print(f"[-] Could not start permission helper: {launch.stdout} {launch.stderr}")
        return False
    print(f"[*] Repairing read access on {len(paths):,} source paths automatically...")
    deadline = time.monotonic() + timeout
    last_notice = time.monotonic()
    while time.monotonic() < deadline:
        if result_path.exists():
            result = json.loads(result_path.read_text(encoding="utf-8-sig"))
            if result.get("request_id") != request_id or not result.get("success"):
                print(f"[-] Permission helper could not finish: {result.get('errors', [])}")
                return False
            return True
        if time.monotonic() - last_notice >= 30:
            print("[*] Still waiting for the permission helper...")
            last_notice = time.monotonic()
        time.sleep(1)
    print("[-] Permission helper timed out; backup remains failed.")
    return False
