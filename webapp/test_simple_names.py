"""업로드 파일 이름을 단순하게 정리하는 규칙을 확인한다. (cv2 없이도 돌도록 필요한 정의만 꺼내 쓴다)"""
import ast
import os
import re
import unittest
from pathlib import Path

HERE = Path(__file__).resolve().parent


def extract(path, names):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    body = [n for n in tree.body
            if getattr(n, "name", None) in names
            or any(getattr(t, "id", None) in names for t in getattr(n, "targets", []))]
    return body


CORE_NAMES = ("_KAKAO_PREFIX", "_KAKAO_RE", "_kakao_stem", "_PREFIXED_DATE_RE", "_JUNK_REST_RE",
              "_prefixed_date_stem", "normalize_stem", "guess_year")
ns = {"re": re, "os": os}
exec(compile(ast.Module(body=extract(HERE.parent / "영수증_정리_도구.py", CORE_NAMES), type_ignores=[]),
             "core", "exec"), ns)


class Core:
    normalize_stem = staticmethod(ns["normalize_stem"])
    guess_year = staticmethod(ns["guess_year"])


import json, shutil, tempfile
ans = {"re": re, "core": Core, "Path": Path, "json": json, "shutil": shutil}
exec(compile(ast.Module(body=extract(HERE / "app.py", ("_decode_filename", "_guess_year_hint", "simple_upload_stem", "simplify_cleaned_names",
                                                "_unique_stem", "merge_jobs")),
                        type_ignores=[]), "app", "exec"), ans)
simple = ans["simple_upload_stem"]


class SimpleNames(unittest.TestCase):
    def test_kakao_variants(self):
        cases = {
            "KakaoTalk_20260813_142355123": "26.08.13",
            "KakaoTalk_20260813_142355123_01": "26.08.13",
            "KakaoTalk_Photo_2026-08-13-14-23-55": "26.08.13",
            "KakaoTalk_Photo_2026-08-13-14-23-55 001": "26.08.13",
            "KakaoTalk_Image_2026-08-13-14-23-55": "26.08.13",
            "카카오톡_20260813_142355": "26.08.13",
            "kakaotalk_20260813_142355 (1)": "26.08.13",
        }
        for src, want in cases.items():
            self.assertEqual(simple(src), want, src)

    def test_same_day_numbering_in_upload(self):
        # create_job 의 번호 붙이는 규칙: 첫 장은 그대로, 다음부터 _2, _3
        used, out = set(), []
        for s in ("KakaoTalk_20260813_142355123", "KakaoTalk_20260813_142401456", "KakaoTalk_Photo_2026-08-13-15-00-00"):
            stem = simple(s)
            final, n = stem, 2
            while final.lower() in used:
                final = f"{stem}_{n}"
                n += 1
            used.add(final.lower())
            out.append(final)
        self.assertEqual(out, ["26.08.13", "26.08.13_2", "26.08.13_3"])

    def test_document_scan_names(self):
        # 스크린샷의 이름들: Document_YYMMDD_HHMMSS
        for s, want in {"Document_261103_171938": "26.11.03", "Document_261106_173827": "26.11.06",
                        "Document_261110_145501_정리": "26.11.10", "IMG_20260813_142355": "26.08.13",
                        "Scan 2026-08-13 14.23.55": "26.08.13"}.items():
            self.assertEqual(simple(s), want, s)

    def test_text_after_date_is_kept(self):
        self.assertEqual(simple("Document_261106 합계"), "Document_261106 합계")
        self.assertEqual(simple("IMG_1234"), "IMG_1234")

    def test_cleaned_names_lose_suffix(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            made = []
            for n in ("26.11.06_정리.jpg", "26.11.06_2_정리.jpg", "충북농산_정리_p1.jpg"):
                (td / n).write_bytes(b"x"); made.append(td / n)
            out = ans["simplify_cleaned_names"](made)
            self.assertEqual([p.name for p in out], ["26.11.06.jpg", "26.11.06_2.jpg", "충북농산_p1.jpg"])
            self.assertTrue(all(p.exists() for p in out))

    def test_unique_stem_for_merge(self):
        u = ans["_unique_stem"]
        used = {"26.11.06", "26.11.06_2"}
        self.assertEqual(u("26.11.06", used), "26.11.06_3")
        self.assertEqual(u("26.11.07", used), "26.11.07")
        self.assertEqual(u("26.11.06_2", used), "26.11.06_3")

    def test_merge_jobs_numbers_across_runs(self):
        with tempfile.TemporaryDirectory() as td:
            td = Path(td)
            ans["JOBS_DIR"] = td / "jobs"
            ids = ["a" * 16, "b" * 16]
            for jid in ids:
                (td / "jobs" / jid / "json").mkdir(parents=True)
                (td / "jobs" / jid / "정리결과").mkdir(parents=True)
                (td / "jobs" / jid / "json" / "26.11.06.json").write_text(
                    json.dumps({"_원본": "x/26.11.06.jpg", "상호": jid[:1]}), encoding="utf-8")
                (td / "jobs" / jid / "정리결과" / "26.11.06.jpg").write_bytes(jid[:1].encode())
            dest = td / "m"
            n = ans["merge_jobs"](ids, dest)
            self.assertEqual(n, 2)
            self.assertEqual(sorted(p.name for p in (dest / "json").glob("*.json")), ["26.11.06.json", "26.11.06_2.json"])
            self.assertEqual((dest / "정리결과" / "26.11.06_2.jpg").read_bytes(), b"b")
            d2 = json.loads((dest / "json" / "26.11.06_2.json").read_text(encoding="utf-8"))
            self.assertTrue(d2["_원본"].endswith("26.11.06_2.jpg"))
            self.assertEqual(d2["상호"], "b")

    def test_not_kakao_or_invalid_left_alone(self):
        for s in ("영수증 사진", "IMG_1234", "KakaoTalk_20261399_142355123"):
            self.assertEqual(simple(s), s)

    def test_date_names_still_work(self):
        self.assertEqual(simple("20260810_3"), "26.08.10_3")
        self.assertEqual(simple("8.21 충북농산", "26"), "26.08.21 충북농산")

    def test_year_hint_from_names(self):
        self.assertEqual(ans["_guess_year_hint"](["26.8.13 킹마트.jpg", "26.8.14.jpg"]), "26")


if __name__ == "__main__":
    unittest.main()
