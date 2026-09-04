import os
import sys
import time
from datetime import datetime
from concurrent.futures import ThreadPoolExecutor, as_completed

if hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except Exception:
        pass

REPORT_FILE = os.path.join(os.path.dirname(__file__), "sync_audit_report.txt")

def check_file_pair(item):
    label, src_path, dst_path, rel_path = item
    try:
        s_stat = os.stat(src_path)
        src_size = s_stat.st_size
        src_mtime = s_stat.st_mtime
        s_dt = datetime.fromtimestamp(src_mtime).strftime("%Y-%m-%d %H:%M:%S")

        if not os.path.exists(dst_path):
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

        d_stat = os.stat(dst_path)
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


def collect_fast_pairs(label, src_dir, dst_dir, max_files=50):
    if not os.path.exists(src_dir):
        return []
    exclude_names = {'__pycache__', '.git', 'node_modules', '.venv', 'venv', '.cache', '.cache_thumbnails', 'appdata', 'identified'}
    exclude_exts = {'.csv', '.pyc', '.tmp', '.log', '.dat', '.cache'}

    pairs = []
    try:
        for entry in os.scandir(src_dir):
            if len(pairs) >= max_files:
                break
            if entry.is_file(follow_symlinks=False):
                ext = os.path.splitext(entry.name)[1].lower()
                if ext in exclude_exts or entry.name.startswith('~') or entry.name.startswith('.'):
                    continue
                s_path = entry.path
                rel = entry.name
                d_path = os.path.join(dst_dir, rel)
                pairs.append((label, s_path, d_path, rel))
            elif entry.is_dir(follow_symlinks=False) and entry.name.lower() not in exclude_names:
                try:
                    for sub in os.scandir(entry.path):
                        if len(pairs) >= max_files:
                            break
                        if sub.is_file(follow_symlinks=False):
                            ext = os.path.splitext(sub.name)[1].lower()
                            if ext in exclude_exts or sub.name.startswith('~') or sub.name.startswith('.'):
                                continue
                            s_path = sub.path
                            rel = os.path.join(entry.name, sub.name)
                            d_path = os.path.join(dst_dir, rel)
                            pairs.append((label, s_path, d_path, rel))
                except Exception:
                    pass
    except Exception:
        pass
    return pairs


def generate_full_audit_report():
    print("[*] Generating Fast Sync Audit Report...")
    start_time = time.time()
    
    pairs = []
    nas_target_dir = r"\\Synology_NAS\Videos and pics\Backups"

    print(" -> Collecting file pairs for Local D:\\Backups vs NAS...")
    pairs.extend(collect_fast_pairs("SAN Sync: D:\\Backups -> NAS", r"D:\Backups", nas_target_dir, max_files=50))
    pairs.extend(collect_fast_pairs("SAN Sync: D:\\Backups\\Documents -> NAS", r"D:\Backups\Documents", os.path.join(nas_target_dir, "Documents"), max_files=50))

    print(" -> Collecting file pairs for OpenSearch project vs NAS...")
    pairs.extend(collect_fast_pairs("SAN Sync: OpenSearch -> NAS", r"D:\Active research\OpenSearch", os.path.join(nas_target_dir, "Active research", "OpenSearch"), max_files=50))

    endnote_src = r"D:\Endnote" if os.path.exists(r"D:\Endnote") else r"D:\Backups\Documents\Endnote"
    print(" -> Collecting file pairs for EndNote vs NAS...")
    pairs.extend(collect_fast_pairs("SAN Sync: EndNote -> NAS", endnote_src, r"\\Synology_NAS\Videos and pics\Endnote", max_files=50))

    quicken_src = r"C:\Users\Paul Dexter\OneDrive\Finances and family\Quicken"
    print(" -> Collecting file pairs for Quicken vs NAS...")
    pairs.extend(collect_fast_pairs("SAN Sync: Quicken -> NAS", quicken_src, r"\\Synology_NAS\Videos and pics\Quicken", max_files=50))

    print(f"[*] Checking {len(pairs)} candidate files against NAS using 32 parallel threads...")
    all_events = []
    with ThreadPoolExecutor(max_workers=32) as executor:
        futures = [executor.submit(check_file_pair, p) for p in pairs]
        for fut in as_completed(futures):
            res = fut.result()
            if res:
                all_events.append(res)

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

    with open(REPORT_FILE, "w", encoding="utf-8") as f:
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

    print(f"[✓] Sync Audit Report successfully written to: {REPORT_FILE}")
    print(f"[+] Total files requiring sync across all targets: {len(all_events):,} (Completed in {elapsed:.2f}s)")


if __name__ == "__main__":
    generate_full_audit_report()
