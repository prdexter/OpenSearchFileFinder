"""
Automated Watchdog & Auto-Restarter for Multi-Job Robocopy Network Backup.
Monitors Robocopy jobs every 30 seconds for Quicken, EndNote, and Documents backup.
"""
import os
import sys
import time
import subprocess
import psutil

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

# Quicken Backup Paths
QUICKEN_SRC = r"C:\Users\Paul Dexter\OneDrive\Finances and family\Quicken"
QUICKEN_LOCAL_DST = r"D:\Quicken"
QUICKEN_NAS_DST = r"\\Synology_NAS\Videos and pics\Quicken"

# EndNote Backup Paths
ENDNOTE_SRC = r"D:\Endnote"
ENDNOTE_FALLBACK_SRC = r"D:\Backups\Documents\Endnote"
ENDNOTE_LOCAL_DST = r"D:\Endnote"
ENDNOTE_NAS_DST = r"\\Synology_NAS\Videos and pics\Endnote"

# Documents Backup Paths
DOCS_SRC = r"D:\Backups"
DOCS_NAS_DST = r"\\Synology_NAS\Videos and pics\Backups"

LOG_FILE = os.path.join(os.path.dirname(__file__), "robocopy_monitor.log")


def log(msg):
    timestamp = time.strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{timestamp}] {msg}"
    try:
        print(line, flush=True)
    except Exception:
        try:
            print(line.encode('ascii', errors='replace').decode('ascii'), flush=True)
        except Exception:
            pass
    try:
        with open(LOG_FILE, "a", encoding="utf-8", errors="replace") as f:
            f.write(line + "\n")
    except Exception:
        pass


def run_quicken_sync():
    log(f"[*] Syncing Quicken: '{QUICKEN_SRC}' -> '{QUICKEN_LOCAL_DST}'...")
    try:
        cmd1 = ["robocopy", QUICKEN_SRC, QUICKEN_LOCAL_DST, "/E", "/FFT", "/DCOPY:DAT", "/TIMFIX", "/J", "/R:1", "/W:1", "/MT:64", "/NFL", "/NDL"]
        subprocess.run(cmd1, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as e:
        log(f"[-] Quicken Local Sync error: {e}")

    log(f"[*] Syncing Quicken: '{QUICKEN_LOCAL_DST}' -> '{QUICKEN_NAS_DST}'...")
    try:
        cmd2 = ["robocopy", QUICKEN_LOCAL_DST, QUICKEN_NAS_DST, "/E", "/FFT", "/DCOPY:DAT", "/TIMFIX", "/J", "/R:1", "/W:1", "/MT:64", "/NFL", "/NDL"]
        subprocess.run(cmd2, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        log(f"[OK] Quicken sync completed successfully to '{QUICKEN_NAS_DST}'")
    except Exception as e:
        log(f"[-] Quicken NAS Sync error: {e}")


def run_endnote_sync():
    src = ENDNOTE_SRC if os.path.exists(ENDNOTE_SRC) else (ENDNOTE_FALLBACK_SRC if os.path.exists(ENDNOTE_FALLBACK_SRC) else ENDNOTE_LOCAL_DST)
    if not os.path.exists(src):
        log(f"[-] EndNote source path not found. Skipping sync.")
        return
    if os.path.normpath(src) != os.path.normpath(ENDNOTE_LOCAL_DST):
        log(f"[*] Syncing EndNote: '{src}' -> '{ENDNOTE_LOCAL_DST}'...")
        try:
            cmd1 = ["robocopy", src, ENDNOTE_LOCAL_DST, "/E", "/FFT", "/DCOPY:DAT", "/TIMFIX", "/J", "/R:1", "/W:1", "/MT:64", "/NFL", "/NDL"]
            subprocess.run(cmd1, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception as e:
            log(f"[-] EndNote Local Sync error: {e}")

    log(f"[*] Syncing EndNote: '{ENDNOTE_LOCAL_DST}' -> '{ENDNOTE_NAS_DST}'...")
    try:
        cmd2 = ["robocopy", ENDNOTE_LOCAL_DST, ENDNOTE_NAS_DST, "/E", "/FFT", "/DCOPY:DAT", "/TIMFIX", "/J", "/R:1", "/W:1", "/MT:64", "/NFL", "/NDL"]
        subprocess.run(cmd2, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        log(f"[OK] EndNote sync completed successfully to '{ENDNOTE_NAS_DST}'")
    except Exception as e:
        log(f"[-] EndNote NAS Sync error: {e}")


def is_indexer_running():
    for p in psutil.process_iter(['name', 'cmdline']):
        try:
            if p.info['name'] and 'python' in p.info['name'].lower():
                cmd = p.info.get('cmdline') or []
                if any('ingest_documents.py' in str(arg) for arg in cmd):
                    return True
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False


def is_docs_robocopy_running():
    for p in psutil.process_iter(['pid', 'name', 'cmdline']):
        try:
            if p.info['name'] and 'robocopy' in p.info['name'].lower():
                cmd = p.info.get('cmdline') or []
                if any('Synology_NAS' in str(arg) or 'Backup' in str(arg) for arg in cmd) or len(cmd) > 1:
                    return True, p.pid
        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass
    return False, None


def start_docs_robocopy():
    log(f"[*] Launching Documents Robocopy transfer from '{DOCS_SRC}' to '{DOCS_NAS_DST}'...")
    cmd = ["robocopy", DOCS_SRC, DOCS_NAS_DST, "/E", "/FFT", "/DCOPY:DAT", "/TIMFIX", "/J", "/R:1", "/W:1", "/MT:64", "/XD", "__pycache__", ".git", "node_modules", ".venv", "venv", ".cache", ".cache_thumbnails", "appdata", "identified"]
    try:
        proc = subprocess.Popen(cmd, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        log(f"[OK] Documents Robocopy started successfully with PID: {proc.pid}")
        return proc.pid
    except Exception as e:
        log(f"[-] Failed to start Documents Robocopy: {e}")
        return None


def main():
    log("==================================================")
    log("Multi-Job Robocopy Watchdog & Auto-Sync Started")
    log("==================================================")

    # Initial Sync on startup
    if not is_indexer_running():
        run_quicken_sync()
        run_endnote_sync()

    last_io_bytes = 0
    stuck_counter = 0
    periodic_sync_counter = 0

    while True:
        try:
            if is_indexer_running():
                # Yield control to primary ingest_documents.py process
                time.sleep(15)
                continue

            # 1. Documents Robocopy Check
            running, pid = is_docs_robocopy_running()
            if not running:
                log("⚠️ Documents Robocopy process is NOT running. Restarting now...")
                start_docs_robocopy()
            else:
                try:
                    p = psutil.Process(pid)
                    io = p.io_counters()
                    current_io = io.read_bytes + io.write_bytes
                    if current_io == last_io_bytes:
                        stuck_counter += 1
                        if stuck_counter >= 10:  # ~5 minutes of 0 IO
                            log(f"⚠️ Documents Robocopy PID {pid} idle for 5 minutes. Restarting...")
                            p.kill()
                            time.sleep(2)
                            start_docs_robocopy()
                            stuck_counter = 0
                    else:
                        stuck_counter = 0
                        last_io_bytes = current_io
                except Exception:
                    pass

            # 2. Periodic Quicken & EndNote Sync (every 5 minutes = 10 loops of 30s)
            periodic_sync_counter += 1
            if periodic_sync_counter >= 10:
                run_quicken_sync()
                run_endnote_sync()
                periodic_sync_counter = 0

        except Exception as e:
            log(f"[-] Monitor loop error: {e}")

        time.sleep(30)


if __name__ == "__main__":
    main()
