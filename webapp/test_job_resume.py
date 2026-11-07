"""Restart classification must not delete receipt files."""
import json
import unittest
from pathlib import Path
from tempfile import TemporaryDirectory

from job_resume import ALLOWED_EXTENSIONS, classify_on_restart, clear_derived_outputs


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
            self.assertEqual(stalled, [("f" * 16, "incomplete")])
            self.assertTrue((root / ("f" * 16) / "status.json").is_file())
            self.assertTrue((root / ("f" * 16) / "원본").is_dir())

    def test_bad_status_and_missing_directory_are_ignored(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            job = root / ("1" * 16)
            job.mkdir()
            (job / "status.json").write_text("{", encoding="utf-8")
            self.assertEqual(classify_on_restart(root), ([], []))
            self.assertEqual(classify_on_restart(root / "missing"), ([], []))

    def test_extensions_match_the_upload_allow_list(self):
        source = Path(__file__).with_name("app.py").read_text(encoding="utf-8")
        block = source.split("ALLOWED_EXT = {", 1)[1].split("}", 1)[0]
        backend = set(__import__("re").findall(r'"(\.[a-z0-9]+)"', block))
        self.assertEqual(ALLOWED_EXTENSIONS, backend)

    def test_partial_or_non_file_originals_are_not_resumed_or_deleted(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            outside = root / "outside.jpg"
            outside.write_bytes(b"keep-me")
            job = "a" * 16
            source = root / job / "원본"
            source.mkdir(parents=True)
            (source / "one.jpg").write_bytes(b"jpg")
            (source / "nested").mkdir()
            (source / "linked.jpg").symlink_to(outside)
            (source / "empty.jpg").write_bytes(b"")
            (source / "notes.txt").write_bytes(b"txt")
            (root / job / "status.json").write_text(
                json.dumps({"state": "running", "file_count": 1}), encoding="utf-8")
            resumable, stalled = classify_on_restart(root)
            self.assertEqual(resumable, [])
            self.assertEqual(stalled, [(job, "incomplete")])
            self.assertEqual((source / "one.jpg").read_bytes(), b"jpg")
            self.assertTrue((source / "nested").is_dir())
            self.assertTrue((source / "linked.jpg").is_symlink())
            self.assertEqual(outside.read_bytes(), b"keep-me")

    def test_exact_readable_set_can_resume(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            job = "b" * 16
            source = root / job / "원본"
            source.mkdir(parents=True)
            (source / "one.jpg").write_bytes(b"one")
            (source / "two.pdf").write_bytes(b"two")
            (root / job / "status.json").write_text(
                json.dumps({"state": "queued", "file_count": 2}), encoding="utf-8")
            resumable, stalled = classify_on_restart(root)
            self.assertEqual(resumable, [job])
            self.assertEqual(stalled, [])

    def test_replay_drops_stale_exports_and_keeps_original_bytes(self):
        with TemporaryDirectory() as tmp:
            root = Path(tmp)
            job = root / ("c" * 16)
            (job / "원본").mkdir(parents=True)
            (job / "원본" / "a.jpg").write_bytes(b"original")
            (job / "status.json").write_text("{}", encoding="utf-8")
            stale = job / "정리결과"
            stale.mkdir()
            (stale / "old.jpg").write_bytes(b"stale")
            (job / "json").mkdir()
            (job / "json" / "old.json").write_text("{}", encoding="utf-8")
            (job / "지출현황.xlsx").write_bytes(b"xlsx")
            (job / "정리된_영수증.zip").write_bytes(b"zip")
            target = root / "safe-target"
            target.mkdir()
            (target / "secret.jpg").write_bytes(b"secret")
            link = job / "정리결과"
            # Replace the real directory with a link after moving it aside.
            moved = root / "moved-output"
            stale.rename(moved)
            link.symlink_to(target, target_is_directory=True)
            removed = clear_derived_outputs(job)
            self.assertIn("정리결과", removed)
            self.assertFalse(link.exists())
            self.assertFalse(link.is_symlink())
            self.assertEqual((target / "secret.jpg").read_bytes(), b"secret")
            self.assertEqual((job / "원본" / "a.jpg").read_bytes(), b"original")
            self.assertTrue((job / "status.json").is_file())
            self.assertEqual(clear_derived_outputs(job), [])
            fresh = job / "정리결과"
            fresh.mkdir()
            (fresh / "a_정리.jpg").write_bytes(b"new")
            self.assertEqual([path.name for path in fresh.iterdir()], ["a_정리.jpg"])
            self.assertTrue((moved / "old.jpg").is_file())


if __name__ == "__main__":
    unittest.main()
