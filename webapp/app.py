# -*- coding: utf-8 -*-
"""영양사 영수증 정리 - 웹 서버 (FastAPI)

기존 '영수증_정리_도구.py' 의 함수를 그대로 불러 쓴다. 이 파일은 웹 껍데기만 맡는다.
  업로드 -> 사진 정리(run) -> 내용 판독(extract_json) -> 지출현황 장부(make_ledger)

환경변수
  GEMINI_API_KEY    (필수) 결제 계정이 연결된 프로젝트의 키를 쓸 것
  APP_ACCESS_CODE   (권장) 파일럿 참여자에게만 알려 주는 접속 코드. 비우면 누구나 접속
  DATA_DIR          작업 파일 저장 위치 (기본 ./data)
  RETENTION_DAYS    작업 파일 보관 일수 (기본 7)
  MAX_FILES         한 번에 올릴 수 있는 파일 수 (기본 40)
  MAX_FILE_MB       파일 하나 크기 상한 (기본 25)
  DAILY_FILE_LIMIT  하루 전체 처리 장수 상한. 요금 폭주 방지 (기본 400)
  RECEIPT_MAX_SIDE  사진 긴 변 상한(px). 메모리 절약용 (기본 2400, 0이면 끔)
  MAX_RESUMES       서버가 재시작될 때 같은 작업을 다시 시도하는 최대 횟수 (기본 1)
"""

from __future__ import annotations

import hashlib
import importlib.util
import json
import os
import queue
import re
import secrets
import shutil
import sys
import threading
import time
import csv
import io
import tempfile
import zipfile
from datetime import datetime
from pathlib import Path

from fastapi import Cookie, FastAPI, File, Form, HTTPException, UploadFile
from fastapi.responses import FileResponse, HTMLResponse, JSONResponse, RedirectResponse
from fastapi.staticfiles import StaticFiles
from starlette.background import BackgroundTask
from pydantic import BaseModel

from env_config import load_project_env
from job_resume import classify_on_restart, clear_derived_outputs

HERE = Path(__file__).resolve().parent
load_project_env(HERE.parent)
DATA_DIR = Path(os.environ.get("DATA_DIR", HERE / "data")).resolve()
RETENTION_DAYS = int(os.environ.get("RETENTION_DAYS", "7"))
# 월 누적에 담아 둔 작업은 한 달치를 모을 수 있게 더 오래 보관한다.
ACCUM_RETENTION_DAYS = int(os.environ.get("ACCUM_RETENTION_DAYS", "45"))
MAX_FILES = int(os.environ.get("MAX_FILES", "40"))
MAX_FILE_MB = int(os.environ.get("MAX_FILE_MB", "25"))
DAILY_FILE_LIMIT = int(os.environ.get("DAILY_FILE_LIMIT", "400"))
ACCESS_CODE = os.environ.get("APP_ACCESS_CODE", "").strip()
ALLOWED_EXT = {".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tif", ".tiff", ".pdf"}

JOBS_DIR = DATA_DIR / "jobs"
JOBS_DIR.mkdir(parents=True, exist_ok=True)

# 판독 모델. 기존 도구의 기본값(gemini-2.5-flash / gemini-3.1-pro-preview)은 2026-11 현재
# 신규 키에서 막혀 있어서 웹에서는 따로 지정한다. 환경변수로 언제든 바꿀 수 있다.
# (영수증_정리_도구.py 는 import 될 때 환경변수를 읽으므로 load_core() 보다 먼저 설정)
MODEL = os.environ.setdefault("GEMINI_MODEL", "gemini-3.8-flash")
PRO_MODEL = os.environ.setdefault("GEMINI_PRO_MODEL", "gemini-4-argon")
# 위 모델이 없어졌을 때 먼저 시도할 대체 모델(쉼표로 구분). 비워 두면 이 키로 쓸 수 있는 최신 정식 모델을 자동으로 찾는다.
os.environ.setdefault("GEMINI_FALLBACK_MODELS", "")
# 서버는 메모리가 작아서(512MB) 휴대폰 원본 사진(4000px 이상)을 그대로 처리하면 한 장에도 한도를 넘는다.
# 사진을 읽을 때 긴 변을 이 값으로 줄인다. 0이면 줄이지 않는다. (영수증_정리_도구.py 가 import 될 때 읽음)
# 3000px 로는 큰 사진(약 3500x4600)에서 512MB 를 넘어 서버가 죽었다(2026-11-10). 재판독 해상도와 같은 2400px 로 낮춘다.
os.environ.setdefault("RECEIPT_MAX_SIDE", "2400")
# 같은 작업 때문에 서버가 계속 죽으면 재시작마다 다시 시도하며 무한 반복된다. 이 횟수까지만 다시 시도한다.
MAX_RESUMES = int(os.environ.get("MAX_RESUMES", "1"))


def _decode_filename(name: str) -> str:
    """일부 업로드 프로그램은 한글 파일명을 =?utf-8?B?...?= 로 감싸 보낸다. 풀어서 쓴다."""
    if name.startswith("=?") and name.endswith("?="):
        from email.header import decode_header, make_header
        try:
            return str(make_header(decode_header(name)))
        except Exception:  # noqa: BLE001
            return name
    return name


# ---------------------------------------------------------------------------
# 기존 도구 불러오기 (프로젝트 폴더, 또는 webapp 옆에 복사해 둔 것)
# ---------------------------------------------------------------------------
def load_core():
    for cand in (HERE.parent / "영수증_정리_도구.py", HERE / "영수증_정리_도구.py"):
        if cand.exists():
            spec = importlib.util.spec_from_file_location("receipt_core", cand)
            mod = importlib.util.module_from_spec(spec)
            sys.modules["receipt_core"] = mod
            spec.loader.exec_module(mod)
            return mod
    raise RuntimeError("영수증_정리_도구.py 를 찾지 못했습니다.")


core = load_core()

app = FastAPI(title="영양사 영수증 정리")
_token_salt = os.environ.get("APP_SESSION_SALT", "receipt-web")


def _token() -> str:
    return hashlib.sha256((ACCESS_CODE + _token_salt).encode()).hexdigest()


def _check_auth(token: str | None):
    if ACCESS_CODE and token != _token():
        raise HTTPException(status_code=401, detail="접속 코드가 필요합니다.")


# ---------------------------------------------------------------------------
# 작업 상태 (파일 한 개: jobs/<id>/status.json)
# ---------------------------------------------------------------------------
_status_lock = threading.Lock()


_rebuild_locks: dict[str, threading.Lock] = {}


def rebuild_ledger(job_id: str, log=None) -> Path:
    """json 폴더의 현재 값(사용자가 고친 값 포함)으로 지출현황 엑셀을 다시 만든다."""
    jd = JOBS_DIR / job_id
    lock = _rebuild_locks.setdefault(job_id, threading.Lock())
    with lock:
        xlsx = core.make_ledger(str(jd / "json"), out_path=str(jd / "지출현황.xlsx"),
                                photos=str(jd / "정리결과"),
                                log=log or (lambda m: add_log(job_id, m)))
        write_status(job_id, ledger_dirty=False)
        return Path(xlsx)


def job_dir(job_id: str) -> Path:
    if not re.fullmatch(r"[0-9a-f]{16}", job_id):
        raise HTTPException(status_code=404, detail="없는 작업입니다.")
    p = JOBS_DIR / job_id
    if not p.is_dir():
        raise HTTPException(status_code=404, detail="없는 작업입니다.")
    return p


def read_status(job_id: str) -> dict:
    p = job_dir(job_id) / "status.json"
    with _status_lock:
        return json.loads(p.read_text(encoding="utf-8"))


def write_status(job_id: str, **changes) -> dict:
    p = JOBS_DIR / job_id / "status.json"
    with _status_lock:
        st = json.loads(p.read_text(encoding="utf-8")) if p.exists() else {}
        st.update(changes)
        tmp = p.with_suffix(".tmp")
        tmp.write_text(json.dumps(st, ensure_ascii=False), encoding="utf-8")
        tmp.replace(p)
        return st


def add_log(job_id: str, line: str):
    line = str(line).rstrip()
    if not line:
        return
    p = JOBS_DIR / job_id / "log.txt"
    with open(p, "a", encoding="utf-8") as f:
        f.write(line + "\n")


# ---------------------------------------------------------------------------
# 처리 (한 번에 한 작업만. 기존 도구가 전역 상태를 쓰므로 줄을 세운다)
# work_q 는 이 프로세스 안에서만 유효하다. 현재 Docker 명령은 uvicorn 프로세스
# 하나다. 프로세스를 여러 개 띄우면 같은 작업을 한 쪽이 차지했다고 보장할 수 없다.
# ---------------------------------------------------------------------------
work_q: "queue.Queue[str]" = queue.Queue()


SERVICE_TROUBLE_SHORT = "판독 서비스에 문제가 생겨 읽기를 중단했습니다. 잠시 후 다시 올려 주세요. 계속되면 관리자에게 알려 주세요."
SERVICE_TROUBLE = ("영수증 판독 서비스에 일시적인 문제가 있어 처리하지 못했습니다. "
                   "잠시 후 다시 시도해 주세요. 계속되면 관리자에게 알려 주세요.")


def admin_note(job_id: str, msg: str):
    """관리자가 봐야 할 내용은 사용자 화면이 아니라 서버 로그(Render Logs)에 남긴다."""
    print(f"[관리자 확인] job={job_id} {msg}", flush=True)


def clean_receipt_name(name: str) -> str:
    """정리본 이름(예: 26.08.05_정리.jpg)을 올린 파일 이름처럼 보이게 되돌린다."""
    return re.sub(r"_정리(?=\.\w+(?: p\d+)?$)", "", name.strip())


def _guess_year_hint(names: list[str]) -> str | None:
    """올린 파일 이름들에서 가장 많이 보이는 두 자리 연도(날짜에 연도가 빠진 이름용)."""
    try:
        return core.guess_year([_decode_filename(n) for n in names], "")
    except Exception:  # noqa: BLE001
        return None


def simple_upload_stem(stem: str, year_hint: str | None = None) -> str:
    """올린 파일 이름을 단순하게 정리한다. 카카오톡 사진이면 'KakaoTalk_20260813_142355123' 이
    '26.08.13' 으로, 'YYYYMMDD', '8.21 충북농산' 같은 날짜 이름은 'YY.MM.DD ...' 로 바뀐다.
    날짜로 읽히지 않는 이름은 그대로 둔다. 같은 이름이 겹치면 호출한 쪽에서 번호를 붙인다."""
    try:
        new = core.normalize_stem(stem, year_hint)
    except Exception:  # noqa: BLE001
        return stem
    new = re.sub(r'[\\/:*?"<>|]', "_", new or "").strip()
    return new or stem


class ProblemTracker:
    """처리 로그 줄에서 '읽지 못한 파일'을 골라낸다.

    기존 도구는 한 장이 실패해도 나머지를 계속 처리하고 실패는 로그 한 줄로만 남긴다.
    로그는 화면에 마지막 12줄만 보이므로, 실패한 파일 이름과 이유를 따로 모아 둔다.
    """

    def __init__(self):
        self.items: list[dict] = []
        self._label: str | None = None
        self._recheck = False       # 재판독 단계에서는 중단돼도 1차 결과가 남는다
        self._seen: set[tuple] = set()

    def _add(self, name: str, reason: str):
        item = {"name": clean_receipt_name(name), "reason": reason.strip()[:120]}
        key = (item["name"], item["reason"])
        if key in self._seen:
            return None
        self._seen.add(key)
        self.items.append(item)
        return item

    def feed(self, msg: str):
        """새로 모은 항목이 있으면 돌려준다. 없으면 None."""
        msg = str(msg).rstrip("\n")
        m = re.match(r"\[실패\] (.+?): (.*)$", msg)               # 사진 정리 단계
        if m:
            return self._add(m.group(1), m.group(2))
        m = re.match(r"\[\d+/\d+\] 실패 (.+?): (.*)$", msg)        # 판독 단계: 파일을 열지 못함
        if m:
            return self._add(m.group(1), m.group(2))
        m = re.match(r"\[\d+/\d+\] (?!실패|건너뜀|다시 읽기)(.+)$", msg)   # 판독 시작 줄
        if m:
            self._label = m.group(1).strip()
            return None
        if msg.startswith("[재판독"):
            self._label = None      # 재판독 실패는 1차 결과가 그대로 남으므로 문제로 보지 않는다
            self._recheck = True
            return None
        m = re.match(r"\s+실패: (.*)$", msg)                       # 판독 단계: AI 응답 실패
        if m and self._label:
            return self._add(self._label, m.group(1))
        m = re.match(r"\[중단\] (.*)$", msg)                       # 키/권한 문제로 남은 장을 못 읽음
        if m and not self._recheck:
            # 자세한 원인(모델 이름, 키 문제)은 사용자에게 필요 없고 서버 로그에만 남긴다
            return self._add("남은 영수증", SERVICE_TROUBLE_SHORT)
        return None


def problem_lines(st: dict) -> list[str]:
    """화면에 보여 줄 '제외되었거나 읽지 못한 파일' 목록."""
    out = [str(s) for s in (st.get("skipped") or [])]
    for f in st.get("failed") or []:
        out.append(f"{f.get('name')} ({f.get('reason')})")
    return out

def simplify_cleaned_names(paths) -> list:
    """정리된 사진 이름 끝의 '_정리' 를 뗀다. (26.11.06_정리.jpg -> 26.11.06.jpg)
    떼면 이미 있는 이름과 겹칠 때만 그대로 둔다."""
    out = []
    for p in paths:
        p = Path(p)
        new = p.with_name(re.sub(r"_정리(?=(?:_\d+)?(?:_p\d+)?$)", "", p.stem, count=1) + p.suffix)
        if new != p and not new.exists():
            try:
                p.rename(new)
                p = new
            except OSError:
                pass
        out.append(p)
    return out

def run_job(job_id: str):
    jd = JOBS_DIR / job_id
    src, cleaned, jdir = jd / "원본", jd / "정리결과", jd / "json"
    st = read_status(job_id)
    total = st["file_count"]
    state = {"phase": "정리", "done": 0}
    tracker = ProblemTracker()

    def log(msg=""):
        msg = str(msg)
        add_log(job_id, msg)
        if msg.startswith(("[중단]", "[모델 대체]", "[재판독 건너뜀]")):
            admin_note(job_id, msg)
        found = tracker.feed(msg)
        if found:
            write_status(job_id, failed=list(tracker.items))
        # 진행률: 정리 0~25%, 판독 25~90%, 장부 90~100%
        if state["phase"] == "정리" and (msg.startswith("[완료]") or msg.startswith("[실패]")):
            state["done"] += 1
            write_status(job_id, progress=int(25 * state["done"] / max(total, 1)))
        elif state["phase"] == "판독":
            m = re.match(r"\[(\d+)/(\d+)\]", msg)
            if m:
                i, n = int(m.group(1)), int(m.group(2))
                write_status(job_id, progress=25 + int(65 * (i - 1) / max(n, 1)))

    try:
        # 이전 시도의 정리본·json·엑셀·zip 이 남아 있으면 새 결과와 섞인다.
        # 원본 폴더는 지우지 않는다.
        clear_derived_outputs(jd)
        write_status(job_id, state="running", phase="사진 정리 중", progress=1, failed=[],
                     started=datetime.now().isoformat(timespec="seconds"))
        written = core.run(str(src), out_dir=str(cleaned), bg="white", fmt="jpg",
                           deskew=True, clean=True, rotate_upright=False,
                           perspective=False, log=log)
        if not written:
            raise RuntimeError("정리된 영수증이 없습니다. 사진이 올바른지 확인해 주세요.")
        written = simplify_cleaned_names(written)

        state["phase"] = "판독"
        write_status(job_id, phase="영수증 내용 읽는 중", progress=25)
        made = core.extract_json([str(p) for p in written], out_dir=str(jdir),
                                 model=MODEL, pro_model=PRO_MODEL, log=log)
        if not made:
            raise RuntimeError("영수증 내용을 읽지 못했습니다. 잠시 후 다시 시도해 주세요.")

        state["phase"] = "장부"
        write_status(job_id, phase="지출현황 엑셀 만드는 중", progress=90)
        xlsx = rebuild_ledger(job_id, log=log)

        # 정리된 사진 묶음(zip)도 함께 제공
        with zipfile.ZipFile(jd / "정리된_영수증.zip", "w", zipfile.ZIP_DEFLATED) as z:
            for p in sorted(cleaned.glob("*")):
                z.write(p, p.name)

        write_status(job_id, state="done", phase="완료", progress=100,
                     finished=datetime.now().isoformat(timespec="seconds"),
                     ledger=Path(xlsx).name)
    except core.JsonSetupError as e:
        # 키·모델 문제는 사용자가 고칠 수 없다. 쉬운 말로 알리고 원인은 서버 로그에 남긴다.
        admin_note(job_id, f"판독 서비스 설정 문제: {e}")
        add_log(job_id, "[오류] 판독 서비스에 문제가 있어 처리하지 못했습니다.")
        write_status(job_id, state="error", phase="오류", error=SERVICE_TROUBLE)
    except Exception as e:  # noqa: BLE001
        add_log(job_id, f"[오류] {e}")
        write_status(job_id, state="error", phase="오류", error=str(e)[:300])


def worker():
    while True:
        job_id = work_q.get()
        try:
            run_job(job_id)
        finally:
            work_q.task_done()


threading.Thread(target=worker, daemon=True).start()


def _resume_jobs_on_startup():
    """프로세스만 다시 떴고 디스크에 원본이 남아 있으면 그 작업만 다시 줄에 넣는다.

    파일을 지우지 않는다. 재배포로 데이터 디렉터리가 비면 복구할 항목이 없다.
    """
    resumable, stalled = classify_on_restart(JOBS_DIR, ALLOWED_EXT)
    for job_id in resumable:
        try:
            tries = int(read_status(job_id).get("resume_count") or 0)
            if tries >= MAX_RESUMES:
                add_log(job_id, "[중단] 이 작업을 처리하다 서버가 반복해서 멈춰 더 시도하지 않습니다.")
                write_status(job_id, state="error", phase="오류",
                             error="처리 중 서버 메모리가 부족해 멈췄습니다. 사진 수를 줄이거나 큰 사진을 나눠서 다시 올려 주세요.")
                continue
            write_status(job_id, state="queued", phase="서버가 다시 시작되어 대기 중", progress=0,
                         resume_count=tries + 1)
            add_log(job_id, "[재시작] 디스크에 원본이 모두 남아 있어 다시 처리합니다.")
            work_q.put(job_id)
        except OSError:
            continue
    incomplete = "서버가 다시 시작되었고 원본이 기대 개수와 다르거나 읽을 수 없습니다. 일부만 처리하지 않았습니다. 영수증을 다시 올려 주세요."
    missing = "서버가 다시 시작되었지만 원본 파일이 없어 이어서 처리하지 못했습니다."
    for job_id, reason in stalled:
        try:
            write_status(job_id, state="error", phase="오류", error=incomplete if reason == "incomplete" else missing)
        except OSError:
            continue


_resume_jobs_on_startup()


# ---------------------------------------------------------------------------
# 보관 기간이 지난 작업 삭제 / 하루 사용량 집계
# ---------------------------------------------------------------------------
def purge_old():
    now = time.time()
    for p in JOBS_DIR.iterdir():
        try:
            if not p.is_dir():
                continue
            days = RETENTION_DAYS
            sp = p / "status.json"
            if sp.exists():
                try:
                    if json.loads(sp.read_text(encoding="utf-8")).get("keep"):
                        days = max(days, ACCUM_RETENTION_DAYS)
                except Exception:  # noqa: BLE001
                    pass
            newest = max(p.stat().st_mtime, sp.stat().st_mtime if sp.exists() else 0)
            if newest < now - days * 86400:
                shutil.rmtree(p, ignore_errors=True)
        except OSError:
            pass


def files_today() -> int:
    today = datetime.now().strftime("%Y-%m-%d")
    n = 0
    for p in JOBS_DIR.iterdir():
        sp = p / "status.json"
        if sp.exists():
            try:
                st = json.loads(sp.read_text(encoding="utf-8"))
                if st.get("created", "").startswith(today):
                    n += int(st.get("file_count", 0))
            except Exception:  # noqa: BLE001
                pass
    return n


def _purge_loop():
    while True:
        purge_old()
        time.sleep(3600)


threading.Thread(target=_purge_loop, daemon=True).start()


# ---------------------------------------------------------------------------
# 화면 / API
# ---------------------------------------------------------------------------
LOGIN_HTML = """<!doctype html><meta charset=utf-8><meta name=viewport content="width=device-width,initial-scale=1">
<title>접속 코드 | 식담온 영수증 정리</title><style>body{font-family:'Malgun Gothic',sans-serif;display:grid;place-items:center;min-height:100vh;margin:0;background:#f0fdfa}
form{background:#fff;padding:32px;border-radius:16px;box-shadow:0 6px 24px #0f172a14;width:min(340px,92vw);box-sizing:border-box}
h3{margin:0 0 6px;color:#115e59}
input,button{width:100%;padding:13px;font-size:16px;margin-top:12px;box-sizing:border-box}input{border:1px solid #cbd5e1;border-radius:10px}button{background:#0f766e;color:#fff;border:0;border-radius:10px;font-weight:700;cursor:pointer}
button:hover{background:#115e59}p{color:#b91c1c;margin:10px 0 0}a{display:block;text-align:center;margin-top:16px;color:#64748b;font-size:14px;text-decoration:none}</style><form method=post action=/login><h3>식담온 영수증 정리</h3>
<div>안내받은 접속 코드를 입력해 주세요.</div><input type=password name=code autofocus>
<button>들어가기</button>{msg}<a href="/">← 소개 페이지로</a></form>"""


@app.get("/", response_class=HTMLResponse)
def landing():
    """소개 페이지 (누구나 볼 수 있음)"""
    return FileResponse(HERE / "static" / "landing.html")


@app.get("/privacy", response_class=HTMLResponse)
def privacy():
    return FileResponse(HERE / "static" / "privacy.html")


@app.get("/app", response_class=HTMLResponse)
def index(rc_auth: str | None = Cookie(default=None)):
    """실제 도구. 접속 코드가 설정돼 있으면 코드 확인 후에만 열린다."""
    if ACCESS_CODE and rc_auth != _token():
        return HTMLResponse(LOGIN_HTML.replace("{msg}", ""))
    return FileResponse(HERE / "static" / "index.html")


@app.post("/login")
def login(code: str = Form(...)):
    if ACCESS_CODE and not secrets.compare_digest(code.strip(), ACCESS_CODE):
        return HTMLResponse(LOGIN_HTML.replace("{msg}", "<p>코드가 맞지 않습니다.</p>"), status_code=401)
    r = RedirectResponse("/app", status_code=303)
    r.set_cookie("rc_auth", _token(), max_age=30 * 86400, httponly=True, samesite="lax")
    return r


@app.get("/healthz")
def healthz():
    return {"ok": True, "gemini_key": bool(os.environ.get("GEMINI_API_KEY"))}


@app.post("/api/jobs")
async def create_job(files: list[UploadFile] = File(...), rc_auth: str | None = Cookie(default=None)):
    _check_auth(rc_auth)
    if not os.environ.get("GEMINI_API_KEY"):
        raise HTTPException(500, "서버에 GEMINI_API_KEY 가 설정되어 있지 않습니다.")
    if not files:
        raise HTTPException(400, "파일을 선택해 주세요.")
    if len(files) > MAX_FILES:
        raise HTTPException(400, f"한 번에 {MAX_FILES}개까지 올릴 수 있습니다.")
    if files_today() + len(files) > DAILY_FILE_LIMIT:
        raise HTTPException(429, "오늘 처리 가능한 양을 넘었습니다. 내일 다시 시도해 주세요.")

    job_id = secrets.token_hex(8)
    jd = JOBS_DIR / job_id
    (jd / "원본").mkdir(parents=True)
    saved, skipped = 0, []
    used = set()
    year_hint = _guess_year_hint([uf.filename or "" for uf in files])
    for uf in files:
        name = Path(_decode_filename(uf.filename or "")).name
        ext = Path(name).suffix.lower()
        if ext not in ALLOWED_EXT:
            skipped.append(f"{name} (지원하지 않는 형식)")
            continue
        stem = re.sub(r'[\\/:*?"<>|]', "_", Path(name).stem).strip() or "영수증"
        stem = simple_upload_stem(stem, year_hint)
        final = f"{stem}{ext}"
        k = 2
        while final.lower() in used:
            final = f"{stem}_{k}{ext}"
            k += 1
        used.add(final.lower())
        # 거대한 업로드를 한꺼번에 메모리에 읽지 않는다.
        data = await uf.read(MAX_FILE_MB * 1024 * 1024 + 1)
        if len(data) > MAX_FILE_MB * 1024 * 1024:
            skipped.append(f"{name} ({MAX_FILE_MB}MB 초과)")
            continue
        (jd / "원본" / final).write_bytes(data)
        saved += 1
    if not saved:
        shutil.rmtree(jd, ignore_errors=True)
        raise HTTPException(400, "처리할 수 있는 파일이 없습니다. " + ", ".join(skipped))

    write_status(job_id, state="queued", phase="대기 중", progress=0, file_count=saved,
                 skipped=skipped, created=datetime.now().isoformat(timespec="seconds"))
    work_q.put(job_id)
    return {"job_id": job_id, "file_count": saved, "skipped": skipped}


@app.get("/api/jobs/{job_id}")
def job_info(job_id: str, rc_auth: str | None = Cookie(default=None)):
    _check_auth(rc_auth)
    st = read_status(job_id)
    st["queue_ahead"] = max(0, list(work_q.queue).index(job_id)) if job_id in list(work_q.queue) else 0
    logp = JOBS_DIR / job_id / "log.txt"
    st["log_tail"] = logp.read_text(encoding="utf-8").splitlines()[-12:] if logp.exists() else []
    return st


def _receipt_rows(job_id: str) -> list[dict]:
    """판독 결과를 화면용 행으로 만든다. 사람이 고친 값이 있으면 그 값을 먼저 보여 준다."""
    jdir = job_dir(job_id) / "json"
    rows = []
    for p in sorted(jdir.glob("*.json")):
        if p.name.startswith("_"):
            continue
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            continue
        base = Path(d.get("_원본", p.stem)).name
        kids = d.get("영수증") if isinstance(d.get("영수증"), list) else None
        # 사진 한 장에 영수증이 여러 장이면 장마다 한 줄로 보여 준다
        for n, r in enumerate(kids or [d], 1):
            if not isinstance(r, dict):
                continue
            v = r.get("_검산") or d.get("_검산") or {}
            fixed_total = r.get("장부금액")
            rows.append({
                "key": f"{p.stem}#{n if kids else 0}",
                "file": f"{base} ({n}/{len(kids)})" if kids else base,
                "image": base,
                "vendor": r.get("장부거래처") or r.get("상호"),
                # 장부와 똑같이 파일 이름의 날짜를 우선한다 (손글씨 날짜를 2020년 등으로 잘못 읽는 일이 있음)
                "date": r.get("장부날짜") or core.stem_date(Path(base).stem) or r.get("거래일자"),
                "total": fixed_total if isinstance(fixed_total, (int, float)) else r.get("총액"),
                "items": len(r.get("품목") or []),
                "needs_check": bool(v.get("확인필요")) and not r.get("_확인됨"),
                "confirmed": bool(r.get("_확인됨")),
                "edited": bool(r.get("_수정")),
                "excluded": r.get("장부반영") == "제외(중복)",
                "note": core.json_note(v) if v.get("확인필요") else "",
                "hint": "" if v.get("확인필요") else core.json_hint(v),
                "unsure": r.get("판독주의") or [],
            })
    rows.sort(key=lambda r: (not r["needs_check"], r["date"] or "9999", r["file"]))
    return rows


@app.get("/api/jobs/{job_id}/receipts")
def job_receipts(job_id: str, rc_auth: str | None = Cookie(default=None)):
    """판독 결과 요약 (화면 표용). 확인이 필요한 영수증을 맨 위로 올린다."""
    _check_auth(rc_auth)
    rows = _receipt_rows(job_id)
    return {"rows": rows, "need_count": sum(r["needs_check"] for r in rows),
            "edited_count": sum(r["edited"] for r in rows),
            "ledger_dirty": bool(read_status(job_id).get("ledger_dirty")),
            "problems": problem_lines(read_status(job_id))}


class ReceiptFix(BaseModel):
    vendor: str | None = None
    date: str | None = None
    total: int | None = None
    exclude: bool | None = None
    confirm: bool | None = None


@app.patch("/api/jobs/{job_id}/receipts/{key}")
def fix_receipt(job_id: str, key: str, fix: ReceiptFix, rc_auth: str | None = Cookie(default=None)):
    """영양사가 영수증 한 장의 값을 직접 고치거나 '맞음'으로 확인한다.

    판독 원본 값(상호, 거래일자, 총액)은 그대로 두고, 고친 값은 장부거래처/장부날짜/장부금액에
    따로 적는다. 나중에 '모델이 어디서 틀렸는지' 비교할 수 있게 하려는 것이다.
    """
    _check_auth(rc_auth)
    jd = job_dir(job_id)
    m = re.fullmatch(r"(.+)#(\d+)", key)
    if not m:
        raise HTTPException(400, "잘못된 영수증 번호입니다.")
    jpath = (jd / "json" / f"{m.group(1)}.json").resolve()
    if jpath.parent != (jd / "json").resolve() or not jpath.exists():
        raise HTTPException(404, "없는 영수증입니다.")
    idx = int(m.group(2))

    if fix.vendor is not None and not (1 <= len(fix.vendor.strip()) <= 60):
        raise HTTPException(400, "상호는 1~60자로 적어 주세요.")
    if fix.total is not None and not (0 <= fix.total <= 2_000_000_000):
        raise HTTPException(400, "금액이 올바르지 않습니다.")
    if fix.date is not None:
        try:
            datetime.strptime(fix.date, "%Y-%m-%d")
        except ValueError:
            raise HTTPException(400, "날짜는 2026-08-05 같은 형식으로 적어 주세요.")

    lock = _rebuild_locks.setdefault(job_id, threading.Lock())
    with lock:
        d = json.loads(jpath.read_text(encoding="utf-8"))
        kids = d.get("영수증") if isinstance(d.get("영수증"), list) else None
        target = (kids[idx - 1] if kids and 1 <= idx <= len(kids) else None) if kids else (d if idx == 0 else None)
        if not isinstance(target, dict):
            raise HTTPException(404, "없는 영수증입니다.")

        changes = {}
        def put(field, shown_key, new, orig_key):
            before = target.get(field) if target.get(field) not in (None, "") else target.get(orig_key)
            if before != new:
                changes[shown_key] = {"전": before, "후": new}
            target[field] = new
        if fix.vendor is not None:
            put("장부거래처", "상호", fix.vendor.strip(), "상호")
        if fix.date is not None:
            put("장부날짜", "날짜", fix.date, "거래일자")
        if fix.total is not None:
            put("장부금액", "총액", fix.total, "총액")
        if fix.exclude is not None:
            if fix.exclude:
                target["장부반영"] = "제외(중복)"
                target["중복사유"] = "사용자가 장부에서 제외"
            else:
                target.pop("장부반영", None); target.pop("중복사유", None)
            changes["장부제외"] = {"전": not fix.exclude, "후": fix.exclude}
        if fix.confirm or changes:
            target["_확인됨"] = True            # 사람이 보고 확인했으니 '확인 필요'는 끈다
        if changes:
            hist = target.setdefault("_수정", [])
            hist.append({"시각": datetime.now().isoformat(timespec="seconds"), "변경": changes})
        if kids:
            d["_검산"] = dict(d.get("_검산") or {}, 확인필요=any(
                (k.get("_검산") or {}).get("확인필요") and not k.get("_확인됨") for k in kids if isinstance(k, dict)))
        jpath.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        if changes:
            with open(jd / "corrections.jsonl", "a", encoding="utf-8") as f:
                f.write(json.dumps({"key": key, "시각": datetime.now().isoformat(timespec="seconds"),
                                    "변경": changes}, ensure_ascii=False) + "\n")
    write_status(job_id, ledger_dirty=True)
    row = next((r for r in _receipt_rows(job_id) if r["key"] == key), None)
    return {"row": row}


def _receipt_json_path(jd: Path, key: str):
    m = re.fullmatch(r"(.+)#(\d+)", key)
    if not m:
        raise HTTPException(400, "잘못된 영수증 번호입니다.")
    jpath = (jd / "json" / f"{m.group(1)}.json").resolve()
    if jpath.parent != (jd / "json").resolve() or not jpath.exists():
        raise HTTPException(404, "없는 영수증입니다.")
    return jpath, int(m.group(2))


def _write_images_zip(jd: Path):
    cleaned = jd / "정리결과"
    with zipfile.ZipFile(jd / "정리된_영수증.zip", "w", zipfile.ZIP_DEFLATED) as z:
        for p in sorted(cleaned.glob("*")):
            z.write(p, p.name)


@app.delete("/api/jobs/{job_id}/receipts/{key}")
def delete_receipt(job_id: str, key: str, rc_auth: str | None = Cookie(default=None)):
    """영수증이 아니거나 필요 없는 사진을 결과에서 완전히 지운다. (엑셀·사진 묶음에서도 빠진다)"""
    _check_auth(rc_auth)
    jd = job_dir(job_id)
    jpath, idx = _receipt_json_path(jd, key)
    lock = _rebuild_locks.setdefault(job_id, threading.Lock())
    with lock:
        d = json.loads(jpath.read_text(encoding="utf-8"))
        kids = d.get("영수증") if isinstance(d.get("영수증"), list) else None
        if kids and len(kids) > 1:
            # 사진 한 장에 영수증이 여러 장이면 그 영수증만 뺀다
            if not 1 <= idx <= len(kids):
                raise HTTPException(404, "없는 영수증입니다.")
            del kids[idx - 1]
            jpath.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
        else:
            stem = jpath.stem
            jpath.unlink()
            for p in (jd / "정리결과").glob("*"):
                if p.stem == stem:
                    p.unlink(missing_ok=True)
            # 원본 사진도 지운다 (같은 원본에서 나온 다른 쪽이 남아 있으면 둔다)
            base = re.sub(r"_p\d+$", "", stem)
            still = any(p.stem == base or re.fullmatch(re.escape(base) + r"_p\d+", p.stem)
                        for p in (jd / "json").glob("*.json"))
            if not still:
                for p in (jd / "원본").glob("*"):
                    if p.stem == base:
                        p.unlink(missing_ok=True)
        # 묶음 파일에는 지운 영수증이 남아 있으므로 같이 지운다
        for extra in ("_모두.json", "_요약.csv"):
            (jd / "json" / extra).unlink(missing_ok=True)
        _write_images_zip(jd)
    write_status(job_id, ledger_dirty=True)
    return {"deleted": True}


@app.get("/api/jobs/{job_id}/image/{name}")
def receipt_image(job_id: str, name: str, rc_auth: str | None = Cookie(default=None)):
    """수정할 때 원본과 비교할 수 있게 정리된 영수증 사진을 보여 준다."""
    _check_auth(rc_auth)
    folder = (job_dir(job_id) / "정리결과").resolve()
    path = (folder / Path(name).name).resolve()
    if path.parent != folder or not path.is_file():
        raise HTTPException(404, "사진이 없습니다.")
    return FileResponse(path)


# ---------------------------------------------------------------------------
# 월 누적: 여러 번 올린 작업(회차)을 한 장부로 합친다.
# 어떤 작업을 담았는지는 사용자 브라우저가 기억하고, 서버는 받은 작업 번호만 합친다.
# ---------------------------------------------------------------------------
_accum_lock = threading.Lock()
MAX_ACCUM_JOBS = 40


class AccumReq(BaseModel):
    job_ids: list[str]
    month: str | None = None


def _valid_job_ids(ids: list[str]) -> list[str]:
    out = []
    for jid in ids[:MAX_ACCUM_JOBS]:
        if isinstance(jid, str) and re.fullmatch(r"[0-9a-f]{16}", jid) and (JOBS_DIR / jid).is_dir() and jid not in out:
            out.append(jid)
    return out


def _unique_stem(stem: str, used: set) -> str:
    """여러 회차를 합칠 때 같은 이름이 겹치면 26.08.13 -> 26.08.13_2 처럼 번호를 붙인다."""
    if stem.lower() not in used:
        return stem
    m = re.match(r"^(?P<head>.+?)(?P<tail>_정리.*)?$", stem)
    head, tail = m.group("head"), m.group("tail") or ""
    head = re.sub(r"_\d+$", "", head)
    n = 2
    while f"{head}_{n}{tail}".lower() in used:
        n += 1
    return f"{head}_{n}{tail}"


def merge_jobs(job_ids: list[str], dest: Path) -> int:
    """여러 작업의 판독 결과(json)와 정리된 사진을 한 폴더로 모은다. 모은 영수증 수를 돌려준다."""
    jdir, pdir = dest / "json", dest / "정리결과"
    jdir.mkdir(parents=True)
    pdir.mkdir(parents=True)
    used, count = set(), 0
    for jid in job_ids:
        jd = JOBS_DIR / jid
        imgs: dict[str, list[Path]] = {}
        for p in (jd / "정리결과").glob("*"):
            imgs.setdefault(p.stem, []).append(p)
        handled = set()
        for jp in sorted((jd / "json").glob("*.json")):
            if jp.name.startswith("_"):
                continue
            try:
                d = json.loads(jp.read_text(encoding="utf-8"))
            except Exception:  # noqa: BLE001
                continue
            new = _unique_stem(jp.stem, used)
            used.add(new.lower())
            first_img = None
            for p in imgs.get(jp.stem, []):
                shutil.copy2(p, pdir / (new + p.suffix))
                first_img = first_img or (pdir / (new + p.suffix))
            handled.add(jp.stem)
            kids = d.get("영수증") if isinstance(d.get("영수증"), list) else []
            for obj in [d] + [k for k in kids if isinstance(k, dict)]:
                if "_원본" in obj and first_img:
                    obj["_원본"] = str(first_img)
                if "파일명" in obj:
                    obj["파일명"] = first_img.name if first_img else new
            (jdir / f"{new}.json").write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
            count += 1
        for stem, ps in imgs.items():          # 판독 결과가 없는 사진도 이미지 묶음에는 넣는다
            if stem in handled:
                continue
            new = _unique_stem(stem, used)
            used.add(new.lower())
            for p in ps:
                shutil.copy2(p, pdir / (new + p.suffix))
    return count


@app.post("/api/accumulate/sync")
def accumulate_sync(req: AccumReq, rc_auth: str | None = Cookie(default=None)):
    """누적에 담은 회차들의 현재 상태를 알려 주고, 한 달 동안 지워지지 않게 표시한다."""
    _check_auth(rc_auth)
    out = []
    for jid in req.job_ids[:MAX_ACCUM_JOBS]:
        if not (isinstance(jid, str) and re.fullmatch(r"[0-9a-f]{16}", jid)) or not (JOBS_DIR / jid).is_dir():
            out.append({"job_id": jid, "exists": False})
            continue
        st = read_status(jid)
        if st.get("state") == "done":
            write_status(jid, keep=True)
        rows = _receipt_rows(jid) if st.get("state") == "done" else []
        live = [r for r in rows if not r["excluded"]]
        out.append({"job_id": jid, "exists": True, "state": st.get("state"),
                    "file_count": st.get("file_count", 0), "created": st.get("created"),
                    "receipts": len(live), "need_count": sum(r["needs_check"] for r in live),
                    "total": sum(r["total"] for r in live if isinstance(r["total"], (int, float)))})
    return {"jobs": out}


@app.post("/api/accumulate/{kind}")
def accumulate_download(kind: str, req: AccumReq, rc_auth: str | None = Cookie(default=None)):
    """담아 둔 회차를 한꺼번에 합쳐 지출현황 엑셀(ledger) 또는 정리된 사진 zip(images)으로 내려 준다."""
    _check_auth(rc_auth)
    if kind not in ("ledger", "images"):
        raise HTTPException(404, "알 수 없는 요청입니다.")
    ids = [j for j in _valid_job_ids(req.job_ids) if read_status(j).get("state") == "done"]
    if not ids:
        raise HTTPException(404, "합칠 수 있는 작업이 없습니다. 서버가 다시 시작되면 이전 작업이 지워질 수 있습니다.")
    month = req.month if req.month and re.fullmatch(r"\d{4}-\d{2}", req.month) else None
    suffix = f"_{month}" if month else ""
    base = DATA_DIR / "acc"
    base.mkdir(parents=True, exist_ok=True)
    tmp = Path(tempfile.mkdtemp(dir=base))
    try:
        with _accum_lock:                      # 한 번에 하나만 (메모리 보호)
            n = merge_jobs(ids, tmp)
            if kind == "images":
                out = tmp / "누적_정리된_영수증.zip"
                with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
                    for p in sorted((tmp / "정리결과").glob("*")):
                        z.write(p, p.name)
                nice = f"정리된_영수증_누적{suffix}.zip"
            else:
                if not n:
                    raise HTTPException(400, "장부에 넣을 영수증이 없습니다.")
                try:
                    out = Path(core.make_ledger(str(tmp / "json"), out_path=str(tmp / "누적_지출현황.xlsx"),
                                                photos=str(tmp / "정리결과"), log=lambda m: None))
                except HTTPException:
                    raise
                except BaseException as e:  # noqa: BLE001  (SystemExit 포함)
                    raise HTTPException(400, f"누적 엑셀을 만들지 못했습니다. ({str(e)[:80]})")
                nice = f"식자재_지출현황_누적{suffix}.xlsx"
    except BaseException:
        shutil.rmtree(tmp, ignore_errors=True)
        raise
    return FileResponse(out, filename=nice, background=BackgroundTask(shutil.rmtree, tmp, True))


@app.get("/api/jobs/{job_id}/download/{kind}")
def download(job_id: str, kind: str, rc_auth: str | None = Cookie(default=None)):
    _check_auth(rc_auth)
    jd = job_dir(job_id)
    files = {
        "ledger": (jd / "지출현황.xlsx", "식자재_지출현황.xlsx"),
        "images": (jd / "정리된_영수증.zip", "정리된_영수증.zip"),
        "json": (jd / "json" / "_요약.csv", "영수증_요약.csv"),
    }
    if kind == "ledger" and read_status(job_id).get("ledger_dirty"):
        try:
            rebuild_ledger(job_id)         # 영양사가 고친 값을 반영해 엑셀을 새로 만든다
        except Exception as e:  # noqa: BLE001
            raise HTTPException(400, f"엑셀을 다시 만들지 못했습니다. 장부에 넣을 영수증이 남아 있는지 확인해 주세요. ({str(e)[:80]})")
    if kind == "json":
        buf = io.StringIO()
        w = csv.writer(buf)
        w.writerow(["파일", "상호", "날짜", "총액", "품목수", "상태", "직접 수정", "참고"])
        for r in _receipt_rows(job_id):
            status = "장부 제외" if r["excluded"] else ("확인 필요" if r["needs_check"] else "정상")
            w.writerow([r["file"], r["vendor"] or "", r["date"] or "", r["total"] if r["total"] is not None else "",
                        r["items"], status, "예" if r["edited"] else "", r["note"] or r["hint"]])
        from fastapi.responses import Response
        return Response(("\ufeff" + buf.getvalue()).encode("utf-8"), media_type="text/csv; charset=utf-8",
                        headers={"Content-Disposition": "attachment; filename*=UTF-8''" +
                                 "%EC%98%81%EC%88%98%EC%A6%9D_%EC%9A%94%EC%95%BD.csv"})
    if kind not in files or not files[kind][0].exists():
        raise HTTPException(404, "파일이 아직 없습니다.")
    path, nice = files[kind]
    return FileResponse(path, filename=nice)


@app.delete("/api/jobs/{job_id}")
def delete_job(job_id: str, rc_auth: str | None = Cookie(default=None)):
    """사용자가 직접 자기 작업(영수증 포함)을 즉시 삭제"""
    _check_auth(rc_auth)
    shutil.rmtree(job_dir(job_id), ignore_errors=True)
    return {"deleted": True}


app.mount("/static", StaticFiles(directory=HERE / "static"), name="static")
