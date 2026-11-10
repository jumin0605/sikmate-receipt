"""Keep the browser allow-list and page flow aligned with the server."""
import re
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parent
APP = (ROOT / "app.py").read_text(encoding="utf-8")
HTML = (ROOT / "static" / "index.html").read_text(encoding="utf-8")
JS = (ROOT / "static" / "receipt-upload.js").read_text(encoding="utf-8")


def extensions(source, pattern):
    block = re.search(pattern, source, re.S)
    if block is None:
        raise AssertionError(pattern)
    return set(re.findall(r"""['"](\.[a-z0-9]+)['"]""", block.group(1)))


class UploadContractTests(unittest.TestCase):
    def test_browser_extensions_match_backend(self):
        backend = extensions(APP, r"ALLOWED_EXT = \{([^}]+)\}")
        browser = extensions(JS, r"ALLOWED_EXT = new Set\(\[([^\]]+)\]\)")
        self.assertEqual(browser, backend)
        self.assertIn(".jpg", backend)
        self.assertIn(".jpeg", backend)
        self.assertIn(".png", backend)
        self.assertIn(".webp", backend)
        self.assertIn(".bmp", backend)
        self.assertIn(".tif", backend)
        self.assertIn(".tiff", backend)
        self.assertIn(".pdf", backend)
        self.assertNotIn(".gif", backend)
        self.assertNotIn(".heic", backend)

    def test_page_keeps_existing_job_flows(self):
        self.assertIn('src="/static/receipt-upload.js"', HTML)
        self.assertNotIn("dataTransfer.files", HTML)
        self.assertNotIn("image/*", HTML)
        pick = re.search(r"#pick\{([^}]+)\}", HTML)
        self.assertIsNotNone(pick)
        self.assertNotIn("display:none", re.sub(r"\s+", "", pick.group(1)))
        for marker in (
            "localStorage.getItem('lastJob')",
            "localStorage.setItem('lastJob'",
            "setInterval(poll, 2000)",
            "download/ledger",
            "download/images",
            "download/json",
            "async function okRow",
            "async function saveRow",
            "function startEdit",
            "snapshotDataTransfer",
            "collectSnapshot",
            "submitJob",
            "한 번에 40개까지",
            "10MB",
            'id="retryPoll"',
            'id="newUpload"',
            'id="relogin"',
        ):
            self.assertIn(marker, HTML)
        self.assertNotIn("보관 기간이 지났을 수 있습니다", HTML)


if __name__ == "__main__":
    unittest.main()
