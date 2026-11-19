# -*- coding: utf-8 -*-
"""개별 로그인(계정) 저장소 - 표준 라이브러리만 쓴다.

- 계정 목록: <root>/users.json  (비밀번호는 PBKDF2-SHA256 으로 해시해 저장, 원문은 어디에도 남기지 않는다)
- 사용자별 폴더: <root>/users/<id>/   (식단표·발주서 기록, 주간 서류 설정, 엑셀 양식을 이 안에 따로 둔다)
- 로그인 표시: 서명한 쿠키 값(토큰). 서버가 재시작돼도 같은 비밀 키(<root>/session.key)를 쓰므로 유지된다.
"""

from __future__ import annotations

import base64
import hashlib
import hmac
import json
import os
import re
import secrets
import threading
import time
from pathlib import Path

USERNAME_RE = re.compile(r"^[0-9A-Za-z가-힣_.\-@+]{2,60}$")
MIN_PASSWORD = 6
MAX_PASSWORD = 100
PBKDF2_ROUNDS = 200_000
TOKEN_DAYS = 30


class AccountError(ValueError):
    """화면에 그대로 보여 줘도 되는 안내 문구를 담는다."""


def _b64e(b: bytes) -> str:
    return base64.urlsafe_b64encode(b).decode().rstrip("=")


def _b64d(s: str) -> bytes:
    return base64.urlsafe_b64decode(s + "=" * (-len(s) % 4))


def hash_password(password: str, rounds: int = PBKDF2_ROUNDS) -> str:
    salt = secrets.token_bytes(16)
    dk = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), salt, rounds)
    return f"pbkdf2_sha256${rounds}${_b64e(salt)}${_b64e(dk)}"


def check_password(password: str, stored: str) -> bool:
    try:
        algo, rounds, salt, dk = stored.split("$")
        if algo != "pbkdf2_sha256":
            return False
        calc = hashlib.pbkdf2_hmac("sha256", password.encode("utf-8"), _b64d(salt), int(rounds))
        return hmac.compare_digest(calc, _b64d(dk))
    except Exception:  # noqa: BLE001
        return False


# 없는 아이디로 시도해도 걸리는 시간이 비슷하도록 쓰는 가짜 해시
_DUMMY_HASH = hash_password("dummy-password-for-timing", 1000)


def clean_username(raw: str) -> str:
    name = (raw or "").strip()
    if not USERNAME_RE.fullmatch(name):
        raise AccountError("아이디는 2~60자의 한글·영문·숫자와 _ . - @ + 만 쓸 수 있습니다. 공백은 쓸 수 없습니다. (이메일 주소도 아이디로 쓸 수 있습니다)")
    return name


def clean_display(raw: str, fallback: str) -> str:
    name = re.sub(r"\s+", " ", (raw or "")).strip()[:20]
    return name or fallback


def check_new_password(password: str) -> str:
    if not isinstance(password, str) or len(password) < MIN_PASSWORD:
        raise AccountError(f"비밀번호는 {MIN_PASSWORD}자 이상으로 정해 주세요.")
    if len(password) > MAX_PASSWORD:
        raise AccountError("비밀번호가 너무 깁니다.")
    return password


class AccountStore:
    def __init__(self, root: Path, session_secret: str = ""):
        self.root = Path(root)
        self.users_path = self.root / "users.json"
        self._lock = threading.RLock()
        self._cache: dict | None = None
        self._cache_mtime = None
        self._secret_env = (session_secret or "").strip()
        self._secret: bytes | None = None

    # -- 저장소 --------------------------------------------------------
    def _read(self) -> dict:
        with self._lock:
            try:
                mt = self.users_path.stat().st_mtime_ns
            except FileNotFoundError:
                self._cache, self._cache_mtime = {"users": {}}, None
                return self._cache
            if self._cache is None or mt != self._cache_mtime:
                try:
                    d = json.loads(self.users_path.read_text(encoding="utf-8"))
                    if not isinstance(d, dict) or not isinstance(d.get("users"), dict):
                        d = {"users": {}}
                except Exception:  # noqa: BLE001
                    d = {"users": {}}
                self._cache, self._cache_mtime = d, mt
            return self._cache

    def _write(self, d: dict):
        with self._lock:
            self.root.mkdir(parents=True, exist_ok=True)
            tmp = self.users_path.with_suffix(".tmp")
            tmp.write_text(json.dumps(d, ensure_ascii=False, indent=1), encoding="utf-8")
            if self.users_path.is_file():
                self.users_path.with_suffix(".bak").write_bytes(self.users_path.read_bytes())
            os.replace(tmp, self.users_path)
            self._cache = d
            self._cache_mtime = self.users_path.stat().st_mtime_ns

    def _secret_key(self) -> bytes:
        if self._secret is None:
            if self._secret_env:
                self._secret = hashlib.sha256(("sikmate-session|" + self._secret_env).encode()).digest()
            else:
                p = self.root / "session.key"
                if p.is_file():
                    self._secret = p.read_bytes()
                else:
                    self.root.mkdir(parents=True, exist_ok=True)
                    self._secret = secrets.token_bytes(32)
                    p.write_bytes(self._secret)
        return self._secret

    # -- 계정 ------------------------------------------------------------
    def count(self) -> int:
        return len(self._read()["users"])

    def get(self, username: str) -> dict | None:
        return self._read()["users"].get((username or "").strip().lower())

    def create(self, username: str, password: str, display: str = "") -> tuple[dict, bool]:
        """새 계정을 만든다. (계정, 맨 처음 만든 계정인지)를 돌려준다."""
        name = clean_username(username)
        check_new_password(password)
        key = name.lower()
        with self._lock:
            d = dict(self._read())
            users = dict(d["users"])
            if key in users:
                raise AccountError("이미 쓰고 있는 아이디입니다. 다른 아이디를 써 주세요.")
            first = not users
            rec = {
                "id": secrets.token_hex(6),
                "username": name,
                "name": clean_display(display, name),
                "hash": hash_password(password),
                "pwv": 1,
                "created": int(time.time()),
            }
            users[key] = rec
            d["users"] = users
            self._write(d)
            self.user_dir(rec).mkdir(parents=True, exist_ok=True)
            return rec, first

    def verify(self, username: str, password: str) -> dict | None:
        rec = self.get(username)
        if rec is None:
            check_password(password or "", _DUMMY_HASH)
            return None
        return rec if check_password(password or "", rec["hash"]) else None

    def change_password(self, username: str, old: str, new: str) -> dict:
        check_new_password(new)
        with self._lock:
            rec = self.verify(username, old)
            if rec is None:
                raise AccountError("현재 비밀번호가 맞지 않습니다.")
            d = dict(self._read())
            users = dict(d["users"])
            rec = dict(rec)
            rec["hash"] = hash_password(new)
            rec["pwv"] = int(rec.get("pwv", 1)) + 1     # 이전에 로그인해 둔 곳은 모두 로그아웃된다
            users[rec["username"].lower()] = rec
            d["users"] = users
            self._write(d)
            return rec

    def user_dir(self, rec: dict) -> Path:
        return self.root / "users" / rec["id"]

    # -- 로그인 표시(토큰) -----------------------------------------------
    def _sign(self, payload: str) -> str:
        return _b64e(hmac.new(self._secret_key(), payload.encode("utf-8"), hashlib.sha256).digest())

    def make_token(self, rec: dict, days: int = TOKEN_DAYS) -> str:
        exp = int(time.time()) + days * 86400
        payload = f"{_b64e(rec['username'].encode('utf-8'))}.{exp}.{rec.get('pwv', 1)}"
        return f"{payload}.{self._sign(payload)}"

    def read_token(self, token: str | None) -> str | None:
        """올바른 토큰이면 계정의 아이디를 돌려준다."""
        if not token or token.count(".") != 3:
            return None
        try:
            u64, exp, pwv, sig = token.split(".")
            payload = f"{u64}.{exp}.{pwv}"
            if not hmac.compare_digest(sig, self._sign(payload)):
                return None
            if int(exp) < time.time():
                return None
            rec = self.get(_b64d(u64).decode("utf-8"))
            if rec is None or int(pwv) != int(rec.get("pwv", 1)):
                return None
            return rec["username"].lower()
        except Exception:  # noqa: BLE001
            return None


class LoginThrottle:
    """같은 곳에서 계속 틀리면 잠시 막는다 (비밀번호 무작위 대입 방지). 메모리에만 둔다."""

    def __init__(self, max_fails: int = 6, window: int = 600, lock_for: int = 600):
        self.max_fails, self.window, self.lock_for = max_fails, window, lock_for
        self._fails: dict[str, list[float]] = {}
        self._locked: dict[str, float] = {}
        self._lock = threading.Lock()

    def blocked(self, key: str) -> int:
        """막혀 있으면 남은 초, 아니면 0"""
        with self._lock:
            until = self._locked.get(key, 0)
            remaining = until - time.time()
            if remaining > 0:
                return int(remaining) + 1
            self._locked.pop(key, None)
            return 0

    def fail(self, key: str):
        now = time.time()
        with self._lock:
            hist = [t for t in self._fails.get(key, []) if now - t < self.window] + [now]
            self._fails[key] = hist
            if len(hist) >= self.max_fails:
                self._locked[key] = now + self.lock_for
                self._fails.pop(key, None)
            if len(self._fails) > 5000:       # 메모리가 끝없이 커지지 않게
                self._fails.clear()

    def ok(self, key: str):
        with self._lock:
            self._fails.pop(key, None)
