"""Retry and verify Active research without reindexing or rerunning other backups."""
import json
import os
import subprocess
from pathlib import Path

import ingest_documents as indexer


def main():
    config = json.loads(Path(__file__).with_name("indexer_config.json").read_text(encoding="utf-8"))
    target = config["network_backup_target"]
    command = indexer.active_research_backup_command(target)
    source = os.path.normcase(os.path.abspath(command[1]))
    destination = os.path.normcase(os.path.abspath(command[2]))
    # This retry is restricted to the existing configured Active research mirror.
    if source != os.path.normcase(r"D:\Active research") or destination != os.path.normcase(
        r"\\Synology_NAS\Videos and pics\Backups\Active research"
    ):
        raise RuntimeError("Unexpected source or destination; refusing to mirror")
    indexer.prevent_sleep()
    copied = skipped = 0
    try:
        code, copied, skipped = indexer.run_robocopy_with_live_progress(command, label="Active research NAS backup")
        if not 0 <= code <= 7:
            raise RuntimeError(f"Active research backup failed: Robocopy code {code}; see backup_logs")
        # A separate read-only pass must find no remaining copies, mismatches, or failures.
        verify_command = command + ["/L", "/NFL", "/NDL"]
        verify_code, _, _ = indexer.run_robocopy_with_live_progress(verify_command, label="Verifying Active research NAS backup")
        if verify_code != 0:
            raise RuntimeError(f"Active research verification found remaining differences or errors (code {verify_code})")
        message = f"✅ Active research NAS backup verified ({copied:,} copied, {skipped:,} skipped)"
        indexer.write_progress(False, copied + skipped, copied, skipped, message, status="completed")
        print(message)
    except Exception as error:
        indexer.write_progress(False, copied + skipped, copied, skipped, "❌ Active research NAS backup failed",
                               status="failed", error_details=str(error))
        raise
    finally:
        indexer.allow_sleep()


if __name__ == "__main__":
    main()
