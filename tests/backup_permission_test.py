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
import uuid

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import backup_permission_helper as helper
import ingest_documents as indexer


class AutomaticPermissionTests(unittest.TestCase):
    def test_only_source_access_denials_are_selected(self):
        with tempfile.TemporaryDirectory() as directory:
            log = Path(directory) / "copy.log"
            log.write_text(
                'ERROR 5 (0x00000005) Copying Directory D:\\Active research\\project\\locked\\\n'
                'ERROR 5 (0x00000005) Copying File D:\\Active research\\project\\locked\\child.txt\n'
                'ERROR 5 (0x00000005) Copying Directory D:\\Other\\private\\\n'
                'ERROR 5 (0x00000005) Copying File \\\\NAS\\share\\private.txt\n'
                'ERROR 32 (0x00000020) Copying File D:\\Active research\\busy.txt\n', encoding="utf-16")
            self.assertEqual(helper.denied_source_paths(log), [r"D:\Active research\project\locked"])

    def test_helper_installation_is_required(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"LOCALAPPDATA": directory}), \
             patch.object(helper, "denied_source_paths", return_value=[r"D:\Active research\locked"]), \
             patch.object(helper.subprocess, "run") as launch, contextlib.redirect_stdout(io.StringIO()):
            self.assertFalse(helper.repair_denied_paths("unused"))
            launch.assert_not_called()

    def test_helper_waits_for_matching_success_receipt(self):
        with tempfile.TemporaryDirectory() as directory, patch.dict(os.environ, {"LOCALAPPDATA": directory}), \
             patch.object(helper, "denied_source_paths", return_value=[r"D:\Active research\locked"]):
            state = Path(directory) / "OpenSearch" / "permission_repair"
            state.mkdir(parents=True)
            (state / "installed.json").write_text("{}")
            def launch(command, **kwargs):
                self.assertEqual(command[-1], helper.TASK_NAME)
                request_path = next(state.glob("*.request.json"))
                request = json.loads(request_path.read_text())
                result = state / f"{request['request_id']}.result.json"
                result.write_text(json.dumps({"request_id": request["request_id"], "success": True}))
                return Mock(returncode=0)
            with patch.object(helper.subprocess, "run", side_effect=launch), contextlib.redirect_stdout(io.StringIO()):
                self.assertTrue(helper.repair_denied_paths("unused"))

    def test_retry_is_bounded_and_verification_never_repairs(self):
        for codes, extra_args, expected_calls, expected_repairs in (
            ([9, 0], [], 2, 1), ([9, 9], [], 2, 1), ([9], ["/L"], 1, 0),
        ):
            with self.subTest(codes=codes, extra_args=extra_args), tempfile.TemporaryDirectory() as directory, \
                 patch.dict(os.environ, {"LOCALAPPDATA": directory}), \
                 patch.object(indexer, "write_progress"), \
                 patch.object(helper, "repair_denied_paths", return_value=True) as repair, \
                 patch.object(indexer.subprocess, "Popen") as launch, \
                 contextlib.redirect_stdout(io.StringIO()):
                processes = []
                for code in codes:
                    process = Mock(returncode=code)
                    process.stdout = io.StringIO("")
                    processes.append(process)
                launch.side_effect = processes
                result = indexer.run_robocopy_with_live_progress(
                    ["robocopy", r"D:\Active research", r"\\NAS\backup"] + extra_args)
                self.assertEqual(result[0], codes[-1])
                self.assertEqual(launch.call_count, expected_calls)
                self.assertEqual(repair.call_count, expected_repairs)

    def test_real_worker_repairs_local_fixture_and_rejects_outside_source(self):
        workspace = Path(indexer.__file__).parent
        (workspace / "scratch").mkdir(exist_ok=True)
        with tempfile.TemporaryDirectory(dir=workspace / "scratch") as directory:
            root = Path(directory)
            fixture = root / "fixture"
            fixture.mkdir()
            (fixture / "sample.txt").write_text("permission helper fixture")
            state = root / "state"
            state.mkdir()
            sid = subprocess.check_output(["powershell.exe", "-NoProfile", "-Command",
                "[Security.Principal.WindowsIdentity]::GetCurrent().User.Value"], text=True).strip()
            for requested_path, expected in ((str(fixture), True), (r"C:\Windows", False),
                                             (r"D:\Active research\..\private", False)):
                request_id = uuid.uuid4().hex
                (state / f"{request_id}.request.json").write_text(json.dumps(
                    {"request_id": request_id, "paths": [requested_path]}))
                subprocess.run(["powershell.exe", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File",
                    str(workspace / "backup_permission_worker.ps1"), "-StateDirectory", str(state),
                    "-BackupAccountSid", sid], check=True, capture_output=True)
                result = json.loads((state / f"{request_id}.result.json").read_text(encoding="utf-8-sig"))
                self.assertEqual(result["success"], expected, result)


if __name__ == "__main__":
    unittest.main()
