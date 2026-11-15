"""Regression checks for backup failures and the app's final status."""
import ast
import contextlib
import io
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import ingest_documents as indexer


class BackupStatusTests(unittest.TestCase):
    def run_network_jobs(self, codes):
        with contextlib.ExitStack() as stack:
            stack.enter_context(contextlib.redirect_stdout(io.StringIO()))
            stack.enter_context(patch.object(indexer.os.path, "exists", side_effect=lambda p: p != r"D:\Endnote"))
            stack.enter_context(patch.object(indexer.os, "makedirs"))
            stack.enter_context(patch.object(indexer, "build_san_summary_dict", return_value={}))
            stack.enter_context(patch.object(indexer, "get_latest_file_info", return_value=("file", "time", "path")))
            stack.enter_context(patch("threading.Thread"))
            runner = stack.enter_context(patch.object(indexer, "run_robocopy_with_live_progress",
                side_effect=[(code, 0, 0) for code in codes]))
            progress = stack.enter_context(patch.object(indexer, "write_progress"))
            try:
                indexer.sync_to_network_target(r"D:\Backups\Documents", r"\\NAS\Backups")
            except RuntimeError:
                self.assertEqual(runner.call_count, 6, "A failed job must not prevent checking other jobs")
                self.assertEqual(progress.call_args.kwargs["status"], "failed")
                raise
            return runner, progress

    def test_each_network_and_local_staging_failure_propagates(self):
        for failed_job in range(6):
            for code in (8, 11, 16, -1):
                with self.subTest(job=failed_job, code=code):
                    codes = [0] * 6
                    codes[failed_job] = code
                    with self.assertRaisesRegex(RuntimeError, "NAS backup failed"):
                        self.run_network_jobs(codes)

    def test_robocopy_success_codes_are_accepted_without_premature_completion(self):
        for code in (0, 1, 7):
            runner, progress = self.run_network_jobs([code] * 6)
            self.assertEqual(runner.call_count, 6)
            self.assertTrue(progress.call_args.args[0], "Only main may publish final completion")

    def test_full_directory_failure_propagates(self):
        with patch.object(indexer.os.path, "exists", return_value=True), \
             patch.object(indexer.os, "makedirs"), \
             patch.object(indexer, "run_robocopy_with_live_progress", return_value=(11, 0, 0)), \
             contextlib.redirect_stdout(io.StringIO()):
            with self.assertRaisesRegex(RuntimeError, "code 11"):
                indexer.backup_full_directory(r"D:\source", r"D:\backup")

    def test_individual_copy_failure_is_retained(self):
        indexer.BACKUP_FAILURES.clear()
        with patch.object(indexer.os, "stat", side_effect=PermissionError("denied")), \
             contextlib.redirect_stdout(io.StringIO()):
            self.assertIsNone(indexer.copy_file_to_backup(r"C:\docs\file.pdf", r"D:\backup"))
        self.assertIn("denied", indexer.BACKUP_FAILURES[0])
        indexer.BACKUP_FAILURES.clear()

    def test_app_does_not_replace_failure_with_success(self):
        # Execute the real subprocess supervisor without starting the web server.
        source = Path(indexer.__file__).with_name("search_app.py").read_text(encoding="utf-8")
        node = next(n for n in ast.walk(ast.parse(source)) if isinstance(n, ast.FunctionDef) and n.name == "run_ingest")
        for returncode, saved_status, expected in (
            (0, "failed", "failed"), (0, "running", "failed"),
            (0, "completed", "completed"), (1, "completed", "failed"),
        ):
            with self.subTest(returncode=returncode, saved_status=saved_status), tempfile.TemporaryDirectory() as directory:
                progress_path = Path(directory) / "indexer_progress.json"
                progress_path.write_text(json.dumps({"status": saved_status, "status_message": "message", "is_running": False}))
                namespace = dict(os=os, sys=sys, json=json, subprocess=Mock(), dirs=[directory],
                    __file__=str(Path(directory) / "search_app.py"), INDEXER_RUN_LOG=str(Path(directory) / "run.log"))
                namespace["subprocess"].Popen.return_value.wait.return_value = returncode
                exec(compile(ast.Module(body=[node], type_ignores=[]), "search_app.py", "exec"), namespace)
                namespace["run_ingest"]()
                result = json.loads(progress_path.read_text())
                self.assertEqual(result["status"], expected)
                self.assertFalse(result["is_running"])

    def test_real_robocopy_runner_copies_and_logs(self):
        with tempfile.TemporaryDirectory() as directory, patch.object(indexer, "write_progress"), \
             contextlib.redirect_stdout(io.StringIO()):
            root = Path(directory)
            source, target = root / "source", root / "target"
            source.mkdir()
            (source / "sample.txt").write_text("backup verification")
            with patch.dict(os.environ, {"LOCALAPPDATA": directory}):
                code, _, _ = indexer.run_robocopy_with_live_progress(["robocopy", str(source), str(target), "/E", "/R:0", "/W:0"])
            self.assertIn(code, range(8))
            self.assertEqual((target / "sample.txt").read_text(), "backup verification")
            logs = list((root / "OpenSearch" / "backup_logs").glob("*.log"))
            self.assertEqual(len(logs), 1)
            self.assertIn("sample.txt", logs[0].read_text(encoding="utf-16"))


if __name__ == "__main__":
    unittest.main()
