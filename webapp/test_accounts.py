# -*- coding: utf-8 -*-
"""개별 로그인: 계정 저장소(accounts.py)와 서버(app.py) 연결 시험"""
import json
import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path

from accounts import AccountError, AccountStore, LoginThrottle, check_password, hash_password


class AccountStoreTests(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.store = AccountStore(self.tmp)

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_password_is_hashed_not_stored(self):
        rec, first = self.store.create("kim", "secret12", "김영양")
        self.assertTrue(first)
        raw = (self.tmp / "users.json").read_text(encoding="utf-8")
        self.assertNotIn("secret12", raw)
        self.assertTrue(check_password("secret12", rec["hash"]))
        self.assertFalse(check_password("secret13", rec["hash"]))

    def test_hash_is_salted(self):
        self.assertNotEqual(hash_password("same-pass"), hash_password("same-pass"))

    def test_only_first_account_is_first(self):
        self.assertTrue(self.store.create("a1", "password1")[1])
        self.assertFalse(self.store.create("b2", "password2")[1])

    def test_username_rules_and_duplicates(self):
        for bad in ("", "a", "a b", "x" * 61, "a/b", "../x", "a@b c"):
            with self.assertRaises(AccountError):
                self.store.create(bad, "password1")
        self.store.create("Kim", "password1")
        with self.assertRaises(AccountError):
            self.store.create("kim", "password1")       # 대소문자만 다른 같은 아이디
        with self.assertRaises(AccountError):
            self.store.create("park", "12345")          # 비밀번호가 너무 짧음

    def test_email_address_as_username(self):
        rec, _ = self.store.create("Jumin0605.adsens2025@gmail.com", "password1")
        self.assertIsNotNone(self.store.verify("jumin0605.adsens2025@GMAIL.com", "password1"))
        self.assertEqual(self.store.read_token(self.store.make_token(rec)), "jumin0605.adsens2025@gmail.com")

    def test_korean_username_and_display_name_default(self):
        rec, _ = self.store.create("손성임", "password1")
        self.assertEqual(rec["name"], "손성임")
        self.assertIsNotNone(self.store.verify("손성임", "password1"))

    def test_verify(self):
        self.store.create("kim", "password1")
        self.assertIsNotNone(self.store.verify("KIM", "password1"))
        self.assertIsNone(self.store.verify("kim", "wrong-pass"))
        self.assertIsNone(self.store.verify("nobody", "password1"))

    def test_token_roundtrip_and_tamper(self):
        rec, _ = self.store.create("kim", "password1")
        tok = self.store.make_token(rec)
        self.assertEqual(self.store.read_token(tok), "kim")
        self.assertIsNone(self.store.read_token(tok[:-2] + "xx"))
        self.assertIsNone(self.store.read_token("a.b.c.d"))
        self.assertIsNone(self.store.read_token(""))
        self.assertIsNone(self.store.read_token(None))

    def test_token_cannot_be_moved_to_another_user(self):
        a, _ = self.store.create("aaa", "password1")
        self.store.create("bbb", "password2")
        tok = self.store.make_token(a)
        head, exp, pwv, sig = tok.split(".")
        from accounts import _b64e
        forged = ".".join([_b64e(b"bbb"), exp, pwv, sig])
        self.assertIsNone(self.store.read_token(forged))

    def test_token_expires(self):
        rec, _ = self.store.create("kim", "password1")
        self.assertIsNone(self.store.read_token(self.store.make_token(rec, days=-1)))

    def test_token_survives_restart_with_same_secret_file(self):
        rec, _ = self.store.create("kim", "password1")
        tok = self.store.make_token(rec)
        again = AccountStore(self.tmp)                  # 서버 재시작
        self.assertEqual(again.read_token(tok), "kim")

    def test_password_change_logs_out_old_sessions(self):
        rec, _ = self.store.create("kim", "password1")
        old = self.store.make_token(rec)
        with self.assertRaises(AccountError):
            self.store.change_password("kim", "wrong", "newpass12")
        new_rec = self.store.change_password("kim", "password1", "newpass12")
        self.assertIsNone(self.store.read_token(old))
        self.assertEqual(self.store.read_token(self.store.make_token(new_rec)), "kim")
        self.assertIsNone(self.store.verify("kim", "password1"))
        self.assertIsNotNone(self.store.verify("kim", "newpass12"))

    def test_each_user_gets_own_folder(self):
        a, _ = self.store.create("aaa", "password1")
        b, _ = self.store.create("bbb", "password2")
        self.assertNotEqual(self.store.user_dir(a), self.store.user_dir(b))
        self.assertTrue(self.store.user_dir(a).is_dir())
        self.assertEqual(self.store.user_dir(a).parent, self.tmp / "users")


class ThrottleTests(unittest.TestCase):
    def test_blocks_after_max_fails_then_recovers(self):
        t = LoginThrottle(max_fails=3, window=60, lock_for=1)
        for _ in range(3):
            self.assertEqual(t.blocked("k"), 0)
            t.fail("k")
        self.assertGreater(t.blocked("k"), 0)
        self.assertEqual(t.blocked("other"), 0)
        time.sleep(1.1)
        self.assertEqual(t.blocked("k"), 0)

    def test_success_clears_count(self):
        t = LoginThrottle(max_fails=3)
        t.fail("k"); t.fail("k"); t.ok("k"); t.fail("k")
        self.assertEqual(t.blocked("k"), 0)


class AppAccountTests(unittest.TestCase):
    """서버 전체를 임시 폴더로 띄워 두 사람이 서로의 기록을 못 보는지 확인한다."""

    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp())
        cls.store_dir = cls.tmp / "store"
        cls.store_dir.mkdir()
        # 계정 기능이 생기기 전 기록(예전 파일)
        legacy = {"rev": 7, "saved_at": 1.0, "data": {"settings": {"facility": "신태양", "dietitian": "손성임 영양사",
                                                                      "meatVendor": "총각축산", "martVendor": "남한강마트"},
                                                         "settingsT": 5}}
        (cls.store_dir / "menu_state.json").write_text(json.dumps(legacy, ensure_ascii=False), encoding="utf-8")
        (cls.store_dir / "weekly_book.json").write_text(json.dumps({"menus": {"테스트메뉴": [""] * 11}, "settings": {}}),
                                                         encoding="utf-8")
        os.environ["DATA_DIR"] = str(cls.tmp / "data")
        os.environ["MENU_STORE_DIR"] = str(cls.store_dir)
        os.environ["APP_ACCESS_CODE"] = "가입코드123"
        os.environ.pop("GEMINI_API_KEY", None)
        import app as app_module
        from fastapi.testclient import TestClient
        cls.app_module = app_module
        cls.TestClient = TestClient

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def client(self):
        return self.TestClient(self.app_module.app, follow_redirects=False)

    def signup(self, c, username, password="password1", name="", code="가입코드123", next_="/app"):
        return c.post("/signup", data={"username": username, "password": password, "password2": password,
                                       "name": name, "code": code, "next": next_})

    # 맨 처음 가입하는 계정이 예전 기록을 이어받으므로 이 시험이 가장 먼저 돌아야 한다(이름 앞의 1)
    def test_1_flow_isolation_and_legacy_adoption(self):
        m = self.app_module
        anon = self.client()
        # 로그인하지 않으면 기록·작업·화면 모두 막힌다
        for url in ("/api/menu/state", "/api/weekly/menudb", "/api/me"):
            self.assertEqual(anon.get(url).status_code, 401, url)
        self.assertEqual(anon.put("/api/menu/state", json={"base_rev": 0, "data": {}}).status_code, 401)
        page = anon.get("/app")
        self.assertIn("로그인", page.text)
        self.assertNotIn("shellTab", page.text)          # 도구 화면이 아니라 로그인 화면이다

        # 가입 코드가 틀리면 가입되지 않는다
        bad = self.signup(self.client(), "hacker", code="틀린코드")
        self.assertEqual(bad.status_code, 401)
        self.assertIsNone(m.ACCOUNTS.get("hacker"))

        # 맨 처음 가입한 계정이 예전 기록을 이어받는다
        c1 = self.client()
        r = self.signup(c1, "sung", name="손성임 영양사")
        self.assertEqual(r.status_code, 303)
        self.assertEqual(r.headers["location"], "/app")
        self.assertIn("rc_auth", r.cookies)
        st1 = c1.get("/api/menu/state").json()
        self.assertEqual(st1["rev"], 7)
        self.assertEqual(st1["data"]["settings"]["meatVendor"], "총각축산")   # 서버 사본은 그대로 (화면에서 청년축산으로 바뀜)
        self.assertIn("테스트메뉴", c1.get("/api/weekly/menudb").json()["extra"])
        self.assertTrue((self.store_dir / "menu_state.json").is_file())     # 예전 파일은 지우지 않는다

        # 두 번째 사람은 빈 설정으로 시작하고, 첫 사람의 기록은 보이지 않는다
        c2 = self.client()
        self.assertEqual(self.signup(c2, "kim", name="김영양").status_code, 303)
        st2 = c2.get("/api/menu/state").json()
        self.assertEqual(st2["data"]["settings"]["facility"], "")
        self.assertEqual(st2["data"]["settings"]["dietitian"], "김영양")
        self.assertNotIn("테스트메뉴", c2.get("/api/weekly/menudb").json()["extra"])

        # 서로 쓴 내용이 섞이지 않는다
        rev = st2["rev"]
        self.assertEqual(c2.put("/api/menu/state", json={"base_rev": rev, "data": {"settings": {"facility": "김시설"}}}).status_code, 200)
        self.assertEqual(c1.get("/api/menu/state").json()["data"]["settings"]["facility"], "신태양")
        self.assertEqual(c2.get("/api/menu/state").json()["data"]["settings"]["facility"], "김시설")
        c2.put("/api/weekly/menudb", json={"name": "김메뉴", "values": ["a"] * 11})
        self.assertNotIn("김메뉴", c1.get("/api/weekly/menudb").json()["extra"])
        self.assertIn("김메뉴", c2.get("/api/weekly/menudb").json()["extra"])

        # 엑셀 양식(템플릿)도 사람마다 따로: 한 사람이 바꿔도 다른 사람의 기본 양식은 그대로
        self.assertIsNone(c1.get("/api/weekly/menudb").json()["template"])
        self.assertIsNone(c2.get("/api/weekly/menudb").json()["template"])

        # 같은 아이디로 다시 가입은 안 된다
        self.assertEqual(self.signup(self.client(), "KIM").status_code, 400)

        # 로그아웃하면 기록을 다시 못 본다 / 다시 로그인하면 본다
        c2.post("/logout")
        c2.cookies.clear()
        self.assertEqual(c2.get("/api/menu/state").status_code, 401)
        wrong = c2.post("/login", data={"username": "kim", "password": "nope-nope"})
        self.assertEqual(wrong.status_code, 401)
        ok = c2.post("/login", data={"username": "kim", "password": "password1", "next": "/menu"})
        self.assertEqual((ok.status_code, ok.headers["location"]), (303, "/menu"))
        self.assertEqual(c2.get("/api/menu/state").json()["data"]["settings"]["facility"], "김시설")

        # 다른 곳으로 돌려보내는 next 값은 무시한다
        evil = c2.post("/login", data={"username": "kim", "password": "password1", "next": "https://evil.example"})
        self.assertEqual(evil.headers["location"], "/app")

    def test_job_owner_guard(self):
        m = self.app_module
        c1, c2 = self.client(), self.client()
        self.signup(c1, "own1")
        self.signup(c2, "own2")
        jid = "ab" * 8
        jd = m.JOBS_DIR / jid
        (jd / "json").mkdir(parents=True)
        (jd / "정리결과").mkdir()
        m.write_status(jid, state="done", owner="own1", file_count=0)
        self.assertEqual(c1.get(f"/api/jobs/{jid}").status_code, 200)
        self.assertEqual(c2.get(f"/api/jobs/{jid}").status_code, 404)
        self.assertEqual(c2.get(f"/api/jobs/{jid}/receipts").status_code, 404)
        self.assertEqual(c2.delete(f"/api/jobs/{jid}").status_code, 404)
        self.assertTrue(jd.is_dir())
        # 월 누적에도 남의 작업은 합쳐지지 않는다
        sync = c2.post("/api/accumulate/sync", json={"job_ids": [jid]}).json()
        self.assertFalse(sync["jobs"][0]["exists"])
        self.assertEqual(c2.post("/api/accumulate/ledger", json={"job_ids": [jid]}).status_code, 404)
        # 주인이 적히지 않은 예전 작업은 로그인한 누구나 열 수 있다
        jid2 = "cd" * 8
        (m.JOBS_DIR / jid2 / "json").mkdir(parents=True)
        m.write_status(jid2, state="done", file_count=0)
        self.assertEqual(c2.get(f"/api/jobs/{jid2}").status_code, 200)

    def test_password_change_and_throttle(self):
        c = self.client()
        self.signup(c, "pwuser")
        self.assertEqual(c.get("/api/me").json()["username"], "pwuser")
        old_cookie = c.cookies.get("rc_auth")
        r = c.post("/account/password", data={"old": "password1", "new": "newpass12", "new2": "newpass12"})
        self.assertEqual(r.status_code, 200)
        self.assertEqual(c.get("/api/me").status_code, 200)             # 지금 기기는 새로 로그인된 상태
        other = self.client()
        other.cookies.set("rc_auth", old_cookie)
        self.assertEqual(other.get("/api/me").status_code, 401)         # 예전 로그인은 풀린다
        # 틀린 비밀번호를 계속 넣으면 잠시 막힌다
        t = self.client()
        codes = [t.post("/login", data={"username": "pwuser", "password": f"bad-{i}-pass"}).status_code for i in range(8)]
        self.assertEqual(codes[0], 401)
        self.assertEqual(codes[-1], 429)
        good = t.post("/login", data={"username": "pwuser", "password": "newpass12"})
        self.assertEqual(good.status_code, 429)                         # 막힌 동안은 맞는 비밀번호도 안 받는다


if __name__ == "__main__":
    unittest.main()

