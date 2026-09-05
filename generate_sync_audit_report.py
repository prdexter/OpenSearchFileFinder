import os
import sys
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)
        sys.stderr.reconfigure(encoding='utf-8', line_buffering=True)
    except Exception:
        pass

REPORT_FILE = os.path.join(os.path.dirname(__file__), "sync_audit_report.txt")

def to_long_path(p: str) -> str:
    if not p:
        return p
    p_str = str(p)
    if p_str.startswith("\\\\?\\") or p_str.startswith("\\\\?\\UNC\\"):
        return p_str
    if p_str.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p_str[2:]
    return "\\\\?\\" + os.path.abspath(p_str)

def check_file_pair(item):
    label, src_path, dst_path, rel_path = item
    try:
        long_src = to_long_path(src_path)
        long_dst = to_long_path(dst_path)

        s_stat = os.stat(long_src)
        src_size = s_stat.st_size
        src_mtime = s_stat.st_mtime
        s_dt = datetime.fromtimestamp(src_mtime).strftime("%Y-%m-%d %H:%M:%S")

        if not os.path.exists(long_dst):
            return {
                "label": label,
                "file_path": src_path,
                "rel_path": rel_path,
                "dst_path": dst_path,
                "reason": "NEW_FILE (Target file missing on backup)",
                "src_size": src_size,
                "dst_size": None,
                "src_mtime": s_dt,
                "dst_mtime": "N/A",
                "time_delta": "N/A"
            }

        d_stat = os.stat(long_dst)
        dst_size = d_stat.st_size
        dst_mtime = d_stat.st_mtime
        d_dt = datetime.fromtimestamp(dst_mtime).strftime("%Y-%m-%d %H:%M:%S")
        delta = abs(dst_mtime - src_mtime)

        if src_size != dst_size:
            return {
                "label": label,
                "file_path": src_path,
                "rel_path": rel_path,
                "dst_path": dst_path,
                "reason": f"SIZE_MISMATCH (Source: {src_size:,} bytes | Target: {dst_size:,} bytes)",
                "src_size": src_size,
                "dst_size": dst_size,
                "src_mtime": s_dt,
                "dst_mtime": d_dt,
                "time_delta": f"{delta:.2f}s"
            }
        elif delta > 3.0:
            return {
                "label": label,
                "file_path": src_path,
                "rel_path": rel_path,
                "dst_path": dst_path,
                "reason": f"TIMESTAMP_DRIFT (Delta: {delta:.2f}s > 3.0s threshold)",
                "src_size": src_size,
                "dst_size": dst_size,
                "src_mtime": s_dt,
                "dst_mtime": d_dt,
                "time_delta": f"{delta:.2f}s"
            }
    except Exception:
        pass
    return None


def collect_fast_pairs(label, src_dir, dst_dir, max_files=None, progress_callback=None, initial_count=0):
    if not os.path.exists(src_dir):
        return []
    exclude_names = {'__pycache__', '.git', 'node_modules', '.venv', 'venv', '.cache', '.cache_thumbnails', 'appdata', 'identified', 'scratch', 'temp'}
    exclude_exts = {'.csv', '.pyc', '.tmp', '.log', '.dat', '.cache'}
    exclude_filenames = {'indexer_config.json', 'indexer_progress.json', 'sync_progress.json', 'robocopy_monitor.log', 'sync_audit_report.txt'}

    pairs = []
    stack = [(src_dir, dst_dir, "")]
    last_cb = 0.0

    while stack:
        s_curr, d_curr, rel_curr = stack.pop()
        if max_files and len(pairs) >= max_files:
            break
        try:
            for entry in os.scandir(to_long_path(s_curr)):
                if max_files and len(pairs) >= max_files:
                    break
                if entry.name.lower() in exclude_names or entry.name.startswith('~') or entry.name.startswith('.'):
                    continue
                rel_path = os.path.join(rel_curr, entry.name) if rel_curr else entry.name
                if entry.is_file(follow_symlinks=False):
                    ext = os.path.splitext(entry.name)[1].lower()
                    if ext in exclude_exts or entry.name.lower() in exclude_filenames:
                        continue
                    pairs.append((label, entry.path, os.path.join(dst_dir, rel_path), rel_path))
                    now = time.time()
                    if progress_callback and now - last_cb >= 0.3:
                        progress_callback(initial_count + len(pairs), 0, 0, is_collecting=True)
                        last_cb = now
                elif entry.is_dir(follow_symlinks=False):
                    stack.append((entry.path, dst_dir, rel_path))
        except Exception:
            pass
    return pairs


def generate_full_audit_report(progress_callback=None):
    print("[*] Generating Fast Sync Audit Report...", flush=True)
    start_time = time.time()
    
    pairs = []
    nas_target_dir = r"\\Synology_NAS\Videos and pics\Backups"

    print(" -> Collecting file pairs for Local D:\\Backups vs NAS...", flush=True)
    b_pairs = collect_fast_pairs("SAN Sync: D:\\Backups -> NAS", r"D:\Backups", nas_target_dir, progress_callback=progress_callback, initial_count=len(pairs))
    pairs.extend(b_pairs)
    print(f"    Collected {len(b_pairs):,} file pairs for D:\\Backups.", flush=True)

    print(" -> Collecting file pairs for OpenSearch project vs NAS...", flush=True)
    o_pairs = collect_fast_pairs("SAN Sync: OpenSearch -> NAS", r"D:\Active research\OpenSearch", os.path.join(nas_target_dir, "Active research", "OpenSearch"), progress_callback=progress_callback, initial_count=len(pairs))
    pairs.extend(o_pairs)
    print(f"    Collected {len(o_pairs):,} file pairs for OpenSearch.", flush=True)

    endnote_src = r"D:\Endnote" if os.path.exists(r"D:\Endnote") else r"D:\Backups\Documents\Endnote"
    print(" -> Collecting file pairs for EndNote vs NAS...", flush=True)
    e_pairs = collect_fast_pairs("SAN Sync: EndNote -> NAS", endnote_src, r"\\Synology_NAS\Videos and pics\Endnote", progress_callback=progress_callback, initial_count=len(pairs))
    pairs.extend(e_pairs)
    print(f"    Collected {len(e_pairs):,} file pairs for EndNote.", flush=True)

    quicken_src = r"C:\Users\Paul Dexter\OneDrive\Finances and family\Quicken"
    print(" -> Collecting file pairs for Quicken vs NAS...", flush=True)
    q_pairs = collect_fast_pairs("SAN Sync: Quicken -> NAS", quicken_src, r"\\Synology_NAS\Videos and pics\Quicken", progress_callback=progress_callback, initial_count=len(pairs))
    pairs.extend(q_pairs)
    print(f"    Collected {len(q_pairs):,} file pairs for Quicken.", flush=True)

    print(f"[*] Checking {len(pairs):,} total candidate files against NAS using 128 parallel threads...", flush=True)
    all_events = []
    checked_count = 0
    mismatch_count = 0
    total_pairs = len(pairs)
    last_cb = 0.0

    with ThreadPoolExecutor(max_workers=128) as executor:
        futures = [executor.submit(check_file_pair, p) for p in pairs]
        for fut in as_completed(futures):
            res = fut.result()
            checked_count += 1
            if res:
                all_events.append(res)
                mismatch_count += 1
            now = time.time()
            if progress_callback and (now - last_cb >= 0.3 or checked_count == total_pairs):
                progress_callback(checked_count, total_pairs, mismatch_count, is_collecting=False)
                last_cb = now
            if now - last_cb >= 1.0 or checked_count == total_pairs:
                print(f"\r -> Checked {checked_count:,}/{total_pairs:,} candidate files... ({mismatch_count:,} drift/missing)", end="", flush=True)
                if not progress_callback:
                    last_cb = now
    print("", flush=True)

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    elapsed = time.time() - start_time

    header = (
        "========================================================================================\n"
        "                  OPENSEARCH BACKUP & SAN SYNC AUDIT REPORT REPORT                      \n"
        "========================================================================================\n"
        f" Report Timestamp          : {now_str}\n"
        f" Audit Execution Time      : {elapsed:.2f} seconds\n"
        f" Candidate Files Inspected : {len(pairs):,}\n"
        f" Total Files Requiring Sync: {len(all_events):,}\n"
        "========================================================================================\n\n"
    )

    with open(REPORT_FILE, "a", encoding="utf-8") as f:
        f.write(header)
        if not all_events:
            f.write("✅ NO FILES REQUIRE SYNC (All local backup files and Synology NAS copies are 100% up to date and in parity).\n")
        else:
            for idx, ev in enumerate(all_events, 1):
                dst_sz = f"{ev['dst_size']:,} bytes" if ev['dst_size'] is not None else "N/A"
                entry = (
                    f"[{idx}] SYNC EVENT: {ev['label']}\n"
                    f"     Prompt Reason : {ev['reason']}\n"
                    f"     File Name     : {os.path.basename(ev['file_path'])}\n"
                    f"     Relative Path : {ev['rel_path']}\n"
                    f"     Source Path   : {ev['file_path']}\n"
                    f"     Target Path   : {ev['dst_path']}\n"
                    f"     Source Size   : {ev['src_size']:,} bytes\n"
                    f"     Target Size   : {dst_sz}\n"
                    f"     Source mtime  : {ev['src_mtime']}\n"
                    f"     Target mtime  : {ev['dst_mtime']}\n"
                    f"     Time Delta    : {ev['time_delta']}\n"
                    "----------------------------------------------------------------------------------------\n"
                )
                f.write(entry)

    print(f"[✓] Sync Audit Report successfully written to: {REPORT_FILE}", flush=True)
    print(f"[+] Total files requiring sync across all targets: {len(all_events):,} (Completed in {elapsed:.2f}s)", flush=True)
    return all_events


if __name__ == "__main__":
    generate_full_audit_report()
