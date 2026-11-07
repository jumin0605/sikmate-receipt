"""Restart classification must not delete receipt files."""
import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from job_resume import classify_on_restart


def write_job(root: Path, job_id: str, state: str, with_source: bool) -> None:
    job = root / job_id
    (job / "원본").mkdir(parents=True)
    if with_source:
        (job / "원본" / "receipt.jpg").write_bytes(b"jpg")
    (job / "status.json").write_text(json.dumps({"state": state, "file_count": 1}), encoding="utf-8")


class RestartClassificationTests(unittest.TestCase):
    def test_running_and_queued_jobs_with_files_can_resume(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_job(root, "a" * 16, "running", True)
            write_job(root, "b" * 16, "queued", True)
            write_job(root, "c" * 16, "done", True)
            write_job(root, "d" * 16, "error", True)
            (root / "notes.txt").write_text("skip", encoding="utf-8")
            (root / ("e" * 16)).mkdir()
            resumable, stalled = classify_on_restart(root)
            self.assertEqual(resumable, ["a" * 16, "b" * 16])
            self.assertEqual(stalled, [])
            self.assertTrue((root / ("a" * 16) / "원본" / "receipt.jpg").is_file())

    def test_open_job_without_originals_is_stalled_and_kept(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            write_job(root, "f" * 16, "running", False)
            resumable, stalled = classify_on_restart(root)
            self.assertEqual(resumable, [])
            self.assertEqual(stalled, ["f" * 16])
            self.assertTrue((root / ("f" * 16) / "status.json").is_file())

    def test_bad_status_and_missing_directory_are_ignored(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            job = root / ("1" * 16)
            job.mkdir()
            (job / "status.json").write_text("{", encoding="utf-8")
            self.assertEqual(classify_on_restart(root), ([], []))
            self.assertEqual(classify_on_restart(root / "missing"), ([], []))


if __name__ == "__main__":
    unittest.main()
