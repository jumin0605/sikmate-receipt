"""사진 정리 옵션(자르기·기울기·밝기·배경)을 서버가 안전하게 받아들이는지 확인한다."""
import ast
import unittest
from pathlib import Path

TREE = ast.parse((Path(__file__).resolve().parent / "app.py").read_text(encoding="utf-8"))
NAMES = ("DEFAULT_OPTIONS", "clean_options", "_flag")
BODY = [n for n in TREE.body if getattr(n, "name", None) in NAMES
        or any(getattr(t, "id", None) in NAMES for t in getattr(n, "targets", []))]
NS = {}
exec(compile(ast.Module(body=BODY, type_ignores=[]), "app.py", "exec"), NS)
clean_options, flag = NS["clean_options"], NS["_flag"]


class CleanOptions(unittest.TestCase):
    def test_defaults_when_missing(self):
        self.assertEqual(clean_options(None), {"crop": True, "deskew": True, "whiten": True, "bg": "white"})

    def test_values_are_kept(self):
        o = clean_options({"crop": False, "deskew": False, "whiten": False, "bg": "transparent"})
        self.assertEqual(o, {"crop": False, "deskew": False, "whiten": False, "bg": "transparent"})

    def test_bad_values_fall_back(self):
        o = clean_options({"crop": "no", "bg": "red; rm -rf"})
        self.assertEqual(o, {"crop": True, "deskew": True, "whiten": True, "bg": "white"})

    def test_form_flags(self):
        self.assertTrue(flag(None))
        self.assertTrue(flag("1"))
        self.assertFalse(flag("0"))
        self.assertFalse(flag("false"))

    def test_app_passes_options_to_core(self):
        src = (Path(__file__).resolve().parent / "app.py").read_text(encoding="utf-8")
        for needle in ('crop=opts["crop"]', 'deskew=opts["deskew"]', 'clean=opts["whiten"]', 'bg=opts["bg"]'):
            self.assertIn(needle, src)


if __name__ == "__main__":
    unittest.main()
