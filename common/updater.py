# -*- coding: utf-8 -*-
"""
자동 업데이트
------------------------------------------------------------
프로그램이 켜질 때 GitHub에서 최신 버전을 확인하고,
있으면 사용자에게 물어본 뒤 **exe 파일 하나만 받아서** 폴더 안에서 교체합니다.

zip 을 따로 올릴 필요가 없습니다. (zip 은 처음 한 번 전달할 때만 씀)
저장소에 exe 만 올려두면, 사용자는 exe 를 켤 때마다 자동으로 최신이 됩니다.

동작 순서
  1. 저장소의 version.json 에서 최신 버전 확인
  2. 현재 버전보다 높으면 "업데이트할까요?" 창을 띄움
  3. 사용자가 예 → 새 exe 다운로드 (약 5MB)
  4. 별도 worker 가 폴더 안의 exe 를 새 것으로 교체
     (실행 중인 exe 는 덮어쓸 수 없으므로, worker 가 %TEMP% 에서 돌아감)
  5. 새 프로그램 자동 실행

  ※ _internal 폴더(파이썬·라이브러리)가 바뀐 버전은 exe 만으로는 안 된다.
     version.json 의 deps 지문이 다르면 full(zip) 쪽으로 내려받는다.

절대 건드리지 않는 것
  - JE_automation_settings.json (사용자 설정)
  - 예전 설정 파일들 (최초 1회 이관용)
  - _internal 폴더 (exe 만 교체할 때는 그대로 둔다)
"""

import hashlib
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from urllib.request import urlopen, Request
from urllib.error import URLError

# ── GitHub 저장소 ────────────────────────────────────────────
# 저장소명은 반드시 영어/숫자여야 합니다 (한글이면 API URL이 깨집니다).
# build.py 의 GITHUB_REPO 와 반드시 같아야 합니다.
GITHUB_OWNER = "doctorjay-code"
GITHUB_REPO = "JE_automation"

RAW_BASE = "https://raw.githubusercontent.com/{}/{}/main".format(
    GITHUB_OWNER, GITHUB_REPO
)

# raw.githubusercontent.com 은 같은 주소를 최대 5분간 캐시한다.
# 그래서 version.json 은 API(캐시 없음)로 먼저 읽고, 실패하면 raw 로 넘어간다.
VERSION_API = "https://api.github.com/repos/{}/{}/contents/version.json?ref=main".format(
    GITHUB_OWNER, GITHUB_REPO
)

# ── 설정 ────────────────────────────────────────────────────
CHECK_TIMEOUT = 8
DOWNLOAD_TIMEOUT = 300
# 업데이트해도 살아남아야 하는 파일 (사용자 설정 · 받아둔 변경 이력)
PRESERVE_FILES = [
    "JE_automation_settings.json",           # 지금 이름
    "JE_Automation_settings.json",           # 예전 이름 (대소문자만 다름)
    "JE_업무자동화_설정.json",                # 예전(한글) 이름
    "운용지시대장_자동반영도구_설정.json",
    "자금판_자동반영도구_설정.json",
    "은대_합계_검증_설정.json",
]


def app_dir():
    """exe로 실행하면 exe가 있는 폴더, 소스로 실행하면 프로젝트 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def deps_fingerprint(folder):
    """
    `_internal` 폴더의 지문. (파일 이름 + 크기만 본다 — 빠름)

    exe 만 바꿔도 되는지 판단하는 데 쓴다.
    라이브러리 구성이 그대로면 exe 하나만 받아도 되고,
    구성이 바뀌었으면 _internal 까지 새로 받아야 한다.
    """
    internal = os.path.join(folder, "_internal")
    if not os.path.isdir(internal):
        return ""
    items = []
    try:
        for root, _dirs, files in os.walk(internal):
            for name in files:
                path = os.path.join(root, name)
                try:
                    size = os.path.getsize(path)
                except OSError:
                    size = -1
                rel = os.path.relpath(path, internal).replace("\\", "/")
                items.append("{}|{}".format(rel, size))
    except OSError:
        return ""
    if not items:
        return ""
    h = hashlib.md5()
    for line in sorted(items):
        h.update(line.encode("utf-8", "ignore"))
        h.update(b"\n")
    return h.hexdigest()[:16]


# ── 버전 비교 ────────────────────────────────────────────────
def parse_version(text):
    """'v1.0.2' → (1, 0, 2)"""
    parts = []
    for chunk in str(text).strip().lstrip("vV").split("."):
        digits = ""
        for ch in chunk:
            if ch.isdigit():
                digits += ch
            else:
                break
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def _http_get(url, timeout, accept=None):
    req = Request(url, headers={
        "User-Agent": "JE-Updater",
        "Accept": accept or "application/vnd.github+json",
    })
    with urlopen(req, timeout=timeout) as resp:
        return resp.read()


def _normalize_notes(value):
    """notes 는 목록(list) 또는 한 줄 문자열. 항상 목록으로 맞춘다."""
    if isinstance(value, (list, tuple)):
        return [str(x).strip() for x in value if str(x).strip()]
    text = str(value or "").strip()
    return [text] if text else []


def _from_version_json():
    """
    저장소의 version.json 을 읽는다.

    raw.githubusercontent.com 은 같은 주소를 최대 5분간 캐시해서
    push 직후에는 옛 내용이 올 수 있다.
    그래서 API(캐시 없음)로 먼저 읽고, 실패하면 raw 로 넘어간다.
    """
    # API 응답도 60초 캐시된다(Cache-Control: max-age=60).
    # push 직후 바로 반영되도록 주소 뒤에 매번 다른 값을 붙인다.
    candidates = (
        ("{}&t={}".format(VERSION_API, int(time.time())), "application/vnd.github.raw"),
        ("{}/version.json?t={}".format(RAW_BASE, int(time.time())), None),
    )

    raw = None
    for url, accept in candidates:
        try:
            raw = _http_get(url, CHECK_TIMEOUT, accept=accept)
            break
        except (URLError, OSError, ValueError):
            continue
    if raw is None:
        return None

    try:
        data = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return None
    if not isinstance(data, dict):
        return None

    version = str(data.get("version", "")).strip()
    if not version:
        return None

    # exe 하나만 받는 주소 (기본 방식)
    exe_url = str(data.get("exe_url", "")).strip()
    exe_file = str(data.get("exe_file", "")).strip()
    if not exe_url and exe_file:
        exe_url = "{}/{}".format(RAW_BASE, exe_file.lstrip("/"))

    # 예전 형식 호환 (url / file 이 zip 인 경우)
    zip_url = str(data.get("url", "")).strip()
    file_name = str(data.get("file", "")).strip()
    if not zip_url and file_name and not exe_url:
        zip_url = "{}/{}".format(RAW_BASE, file_name.lstrip("/"))

    if not exe_url and not zip_url:
        return None

    return {
        "version": version.lstrip("vV"),
        "tag": "v" + version.lstrip("vV"),
        "exe_url": exe_url,
        "exe_file": exe_file or os.path.basename(exe_url),
        "full_url": str(data.get("full_url", "")).strip(),
        "deps": str(data.get("deps", "")).strip(),
        "url": zip_url,                      # 예전 형식 (zip)
        "notes": _normalize_notes(data.get("notes")),
    }


def check_latest(current_version):
    """
    최신 버전 정보를 찾는다.
    반환: {"version":..., "exe_url":..., ...} 또는 None (확인 실패 / 최신)
    """
    info = _from_version_json()
    if info is None:
        return None
    try:
        if parse_version(info["version"]) <= parse_version(current_version):
            return None
    except Exception:
        return None
    return info


# ── 다운로드 ────────────────────────────────────────────────
# ── 변경 이력 (changelog) ───────────────────────────────────
def fetch_changelog():
    """
    저장소의 changelog.json 을 읽는다. {"1.1.7": ["...", "..."], ...}
    인터넷이 안 되면 None. (따로 저장해 두지 않는다 — 파일을 만들지 않기 위해)
    """
    url = "{}/changelog.json?t={}".format(RAW_BASE, int(time.time()))
    try:
        raw = _http_get(url, CHECK_TIMEOUT)
    except (URLError, OSError, ValueError):
        return None
    try:
        data = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeDecodeError):
        return None
    if not isinstance(data, dict):
        return None

    out = {}
    for key, value in data.items():
        notes = _normalize_notes(value)
        if notes:
            out[str(key).strip().lstrip("vV")] = notes
    return out or None


def download(url, dest_path, progress_cb=None):
    """파일을 내려받는다. progress_cb(퍼센트)를 주기적으로 부른다."""
    req = Request(url, headers={"User-Agent": "JE-Updater"})
    with urlopen(req, timeout=DOWNLOAD_TIMEOUT) as resp:
        total = int(resp.headers.get("Content-Length", 0) or 0)
        got = 0
        with open(dest_path, "wb") as f:
            while True:
                chunk = resp.read(64 * 1024)
                if not chunk:
                    break
                f.write(chunk)
                got += len(chunk)
                if progress_cb and total > 0:
                    progress_cb(min(100, int(got * 100 / total)))


# ── 폴더 교체 ────────────────────────────────────────────────
def _safe_remove(path):
    import stat

    def _on_error(func, p, _exc):
        try:
            os.chmod(p, stat.S_IWRITE)
            func(p)
        except Exception:
            pass

    try:
        if os.path.isdir(path):
            shutil.rmtree(path, onerror=_on_error)
        elif os.path.exists(path):
            os.remove(path)
    except Exception:
        pass


def _wait_until_unlocked(target, timeout=120):
    """실행 중인 프로그램이 끝나 exe 를 놓아줄 때까지 기다린다."""
    deadline = time.time() + timeout
    while time.time() < deadline:
        try:
            with open(target, "r+b"):
                return True
        except OSError:
            time.sleep(0.5)
        except Exception:
            return True
    return False


def _collect_exe_names(folder):
    """폴더 안의 exe 파일이름을 모은다. (바로가기 대상을 살리기 위해)"""
    names = []
    try:
        for name in os.listdir(folder):
            if name.lower().endswith(".exe"):
                names.append(name)
    except OSError:
        pass
    return names


def apply_update(zip_path, target_dir):
    """
    zip 으로 target_dir 을 교체한다.
    반환: (성공여부, 메시지)
    """
    # 1) 사용자 설정은 미리 따로 빼둔다
    preserved = {}
    for name in PRESERVE_FILES:
        path = os.path.join(target_dir, name)
        if os.path.exists(path):
            try:
                with open(path, "r", encoding="utf-8") as f:
                    preserved[name] = f.read()
            except Exception:
                pass

    # 2) 바로가기 대상 exe 이름 (예전 이름이어도 새 버전으로 덮어쓴다)
    keep_exe_names = set(_collect_exe_names(target_dir))

    # 3) 백업
    backup_dir = os.path.join(tempfile.gettempdir(), "JE_automation_backup")
    _safe_remove(backup_dir)
    try:
        shutil.copytree(target_dir, backup_dir)
    except Exception as e:
        return False, "백업을 만들지 못했습니다: {}".format(e)

    # 4) 압축 풀기
    extract_dir = os.path.join(tempfile.gettempdir(), "JE_automation_extract")
    _safe_remove(extract_dir)
    try:
        os.makedirs(extract_dir, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(extract_dir)
    except Exception as e:
        _safe_remove(backup_dir)
        return False, "압축을 풀지 못했습니다: {}".format(e)

    # 5) 실제 프로그램 폴더 찾기 (zip 안에 폴더가 하나 더 있을 수 있다)
    source_root = extract_dir
    try:
        entries = [n for n in os.listdir(extract_dir) if not n.startswith("__MACOSX")]
        if len(entries) == 1 and os.path.isdir(os.path.join(extract_dir, entries[0])):
            source_root = os.path.join(extract_dir, entries[0])
    except OSError:
        pass

    # 6) 새 exe 찾기
    new_exe = None
    try:
        for name in os.listdir(source_root):
            if name.lower().endswith(".exe"):
                new_exe = os.path.join(source_root, name)
                break
    except OSError:
        pass
    if not new_exe:
        _safe_remove(backup_dir)
        _safe_remove(extract_dir)
        return False, "다운로드한 파일에 실행 파일이 없습니다."

    # 7) 교체
    try:
        # 7-1) 기존 파일 지우기 (설정 파일과 바로가기용 exe 이름은 남긴다)
        for name in os.listdir(target_dir):
            if name in PRESERVE_FILES or name in keep_exe_names:
                continue
            _safe_remove(os.path.join(target_dir, name))

        # 7-2) 새 파일 복사
        for name in os.listdir(source_root):
            src = os.path.join(source_root, name)
            dst = os.path.join(target_dir, name)
            if os.path.isdir(src):
                _safe_remove(dst)
                shutil.copytree(src, dst)
            else:
                shutil.copy2(src, dst)

        # 7-3) 예전 이름의 바로가기용 exe 도 새 버전으로 갱신
        for name in keep_exe_names:
            dst = os.path.join(target_dir, name)
            if os.path.abspath(dst) != os.path.abspath(new_exe):
                try:
                    shutil.copy2(new_exe, dst)
                except Exception:
                    pass
    except Exception as e:
        # 실패하면 백업으로 되돌린다
        try:
            for name in os.listdir(target_dir):
                _safe_remove(os.path.join(target_dir, name))
            for name in os.listdir(backup_dir):
                src = os.path.join(backup_dir, name)
                dst = os.path.join(target_dir, name)
                if os.path.isdir(src):
                    shutil.copytree(src, dst)
                else:
                    shutil.copy2(src, dst)
        except Exception:
            pass
        _safe_remove(backup_dir)
        _safe_remove(extract_dir)
        return False, "파일을 교체하지 못했습니다: {}".format(e)

    # 8) 사용자 설정 되돌리기
    for name, content in preserved.items():
        try:
            with open(os.path.join(target_dir, name), "w", encoding="utf-8") as f:
                f.write(content)
        except Exception:
            pass

    _safe_remove(extract_dir)
    _safe_remove(backup_dir)
    return True, "업데이트를 마쳤습니다."


# ── exe 하나만 교체 (기본 업데이트 방식) ────────────────────
def apply_exe_update(new_exe_path, target_dir):
    """
    새 exe 를 폴더에 넣고, 기존 exe 이름들도 새 내용으로 덮어쓴다.

    - `_internal` 폴더와 설정 파일은 **건드리지 않는다.** (그래서 빠르고 안전)
    - 바로가기가 옛 이름을 가리키고 있어도 계속 동작한다.

    반환: (성공여부, 메시지)
    """
    if not os.path.exists(new_exe_path):
        return False, "받은 파일을 찾지 못했습니다."
    if not os.path.isdir(target_dir):
        return False, "프로그램 폴더를 찾지 못했습니다."

    new_name = os.path.basename(new_exe_path)
    keep_names = _collect_exe_names(target_dir)

    # 1) 새 이름으로 넣기
    dest = os.path.join(target_dir, new_name)
    try:
        shutil.copy2(new_exe_path, dest)
    except Exception as e:
        return False, "새 파일을 넣지 못했습니다: {}".format(e)

    # 2) 옛 이름들도 같은 내용으로 덮어쓰기 (바로가기 유지)
    for name in keep_names:
        if name == new_name:
            continue
        try:
            shutil.copy2(new_exe_path, os.path.join(target_dir, name))
        except Exception:
            pass

    return True, "업데이트를 마쳤습니다."


# ── worker (교체는 별도 프로그램이 한다) ────────────────────
def _launch_worker_exe(new_exe_path, target_dir):
    """
    exe 교체용 worker 를 띄운다.

    worker 는 **현재 폴더를 %TEMP% 로 복사한 것**을 쓴다.
    (폴더 형태 exe 는 `_internal` 이 있어야 실행되므로, exe 하나만 복사하면 안 된다)
    """
    if not getattr(sys, "frozen", False):
        script = os.path.join(app_dir(), "common", "updater.py")
        cmd = [sys.executable, script, "--apply-exe", new_exe_path, target_dir]
        try:
            subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW
                             if os.name == "nt" else 0)
            return True
        except Exception:
            return False

    worker_root = os.path.join(tempfile.gettempdir(), "JE_automation_worker")
    _safe_remove(worker_root)
    try:
        # 설정·로그는 복사하지 않는다 (개인정보가 임시폴더에 남지 않게)
        shutil.copytree(target_dir, worker_root,
                        ignore=shutil.ignore_patterns("*.json", "*.log"))
    except Exception:
        return False

    worker_exe = None
    try:
        for name in _collect_exe_names(worker_root):
            worker_exe = os.path.join(worker_root, name)
            break
    except OSError:
        pass
    if not worker_exe:
        return False

    cmd = [worker_exe, "--apply-exe", new_exe_path, target_dir]
    try:
        subprocess.Popen(cmd, cwd=worker_root,
                         creationflags=subprocess.DETACHED_PROCESS
                         if os.name == "nt" else 0,
                         close_fds=True)
        return True
    except Exception:
        return False


def _launch_worker(zip_path, target_dir):
    """
    %TEMP% 에 worker 사본을 만들어 실행한다.
    프로그램 폴더에 두면 실행 중인 파일 잠금에 걸리기 때문이다.
    """
    if not getattr(sys, "frozen", False):
        # 소스로 실행하는 경우: 파이썬으로 직접 실행
        script = os.path.join(app_dir(), "common", "updater.py")
        cmd = [sys.executable, script, "--apply-update", zip_path, target_dir]
        try:
            subprocess.Popen(cmd, creationflags=subprocess.CREATE_NO_WINDOW
                             if os.name == "nt" else 0)
            return True
        except Exception:
            return False

    # ── worker 실행 ────────────────────────────────────────────
    # ★ 폴더 형태(onedir) exe 는 옆에 `_internal` 이 없으면 실행되지 않는다.
    #   ("Failed to load Python DLL ... python314.dll")
    #   exe 하나만 복사하면 안 되고, **zip 을 풀어서 그 안의 exe** 를 써야 한다.
    #   (zip 안에는 exe + _internal 이 통째로 들어 있다)
    worker_root = os.path.join(tempfile.gettempdir(), "JE_automation_worker")
    _safe_remove(worker_root)
    try:
        os.makedirs(worker_root, exist_ok=True)
        with zipfile.ZipFile(zip_path, "r") as zf:
            zf.extractall(worker_root)
    except Exception:
        return False

    # zip 안에 폴더가 한 겹 더 있을 수 있다 (JE_automation/)
    base = worker_root
    try:
        entries = [n for n in os.listdir(base) if not n.startswith("__MACOSX")]
        if len(entries) == 1 and os.path.isdir(os.path.join(base, entries[0])):
            base = os.path.join(base, entries[0])
    except OSError:
        pass

    worker_exe = None
    try:
        for name in os.listdir(base):
            if name.lower().endswith(".exe"):
                worker_exe = os.path.join(base, name)
                break
    except OSError:
        pass
    if not worker_exe:
        return False

    cmd = [worker_exe, "--apply-update", zip_path, target_dir]
    try:
        subprocess.Popen(cmd, cwd=base,
                         creationflags=subprocess.DETACHED_PROCESS
                         if os.name == "nt" else 0,
                         close_fds=True)
        return True
    except Exception:
        return False


def _newest_exe(folder):
    """폴더에서 가장 최근에 만들어진 exe. (방금 넣은 새 exe 를 고르기 위해)"""
    best = None
    best_time = -1.0
    for name in _collect_exe_names(folder):
        path = os.path.join(folder, name)
        try:
            t = os.path.getmtime(path)
        except OSError:
            continue
        if t > best_time:
            best_time = t
            best = path
    return best


def _relaunch(target_dir, prefer=None):
    """새 버전을 실행한다. 성공하면 True."""
    candidates = []
    if prefer:
        candidates.append(os.path.join(target_dir, prefer))
    newest = _newest_exe(target_dir)
    if newest:
        candidates.append(newest)
    for name in _collect_exe_names(target_dir):
        candidates.append(os.path.join(target_dir, name))

    seen = set()
    for path in candidates:
        key = os.path.abspath(path).lower()
        if key in seen:
            continue
        seen.add(key)
        if not os.path.exists(path):
            continue
        try:
            subprocess.Popen([path], cwd=target_dir,
                             creationflags=subprocess.DETACHED_PROCESS
                             if os.name == "nt" else 0,
                             close_fds=True)
            return True
        except Exception:
            continue
    return False


def run_worker(argv):
    """
    worker 로 실행되었을 때의 동작.

    argv[0] : "--apply-exe" (exe 교체) 또는 "--apply-update" (zip 교체)
    argv[1] : 받은 파일 경로
    argv[2] : 프로그램 폴더
    """
    if len(argv) < 3:
        return 1
    mode = argv[0]
    payload = argv[1]
    target_dir = argv[2]

    # 실행 중인 프로그램이 끝날 때까지 기다린다
    time.sleep(1.0)
    for name in _collect_exe_names(target_dir):
        _wait_until_unlocked(os.path.join(target_dir, name))

    if mode == "--apply-exe":
        ok, _msg = apply_exe_update(payload, target_dir)
        prefer = os.path.basename(payload)
    else:
        ok, _msg = apply_update(payload, target_dir)
        prefer = None

    # 새 버전 실행
    if ok:
        _relaunch(target_dir, prefer=prefer)

    # 받은 파일과 worker 자신은 정리
    try:
        if mode == "--apply-exe":
            _safe_remove(payload)
    except Exception:
        pass
    try:
        if getattr(sys, "frozen", False):
            _safe_remove(sys.executable)
    except Exception:
        pass
    return 0 if ok else 1


def _needs_full(info, target_dir):
    """
    zip(전체)으로 받아야 하는지 판단한다.

    exe 만 바꾸는 것은 `_internal` 폴더가 그대로일 때만 안전하다.
    라이브러리 구성이 바뀐 버전이면 zip 으로 전체를 받아야 한다.
    """
    full_url = (info.get("full_url") or "").strip()
    if not full_url:
        return False
    remote = (info.get("deps") or "").strip()
    local = deps_fingerprint(target_dir)
    if not remote or not local:
        # 비교할 수 없으면 안전한 쪽(전체)으로 간다
        return True
    return remote != local


def start_update(info, progress_cb=None):
    """
    업데이트를 시작한다. (다운로드 → worker 실행)
    반환: (성공여부, 메시지)
    """
    target_dir = app_dir()
    temp_base = tempfile.gettempdir()

    if _needs_full(info, target_dir):
        # ── 전체(zip) 교체 ──
        url = (info.get("full_url") or info.get("url") or "").strip()
        if not url:
            return False, "받을 주소를 찾지 못했습니다."
        zip_path = os.path.join(temp_base, "JE_automation_update.zip")
        try:
            if progress_cb:
                progress_cb(0)
            download(url, zip_path, progress_cb)
            if progress_cb:
                progress_cb(100)
        except Exception as e:
            return False, "다운로드에 실패했습니다: {}".format(e)
        if not _launch_worker(zip_path, target_dir):
            return False, "업데이트를 시작하지 못했습니다."
        return True, "업데이트를 시작합니다."

    # ── exe 하나만 교체 (기본) ──
    exe_url = (info.get("exe_url") or "").strip()
    if not exe_url:
        return False, "받을 주소를 찾지 못했습니다."

    name = (info.get("exe_file") or os.path.basename(exe_url)).strip()
    if not name.lower().endswith(".exe"):
        name = "JE_automation_new.exe"
    download_dir = os.path.join(temp_base, "JE_automation_download")
    _safe_remove(download_dir)
    try:
        os.makedirs(download_dir, exist_ok=True)
    except OSError:
        return False, "임시 폴더를 만들지 못했습니다."
    new_exe = os.path.join(download_dir, name)

    try:
        if progress_cb:
            progress_cb(0)
        download(exe_url, new_exe, progress_cb)
        if progress_cb:
            progress_cb(100)
    except Exception as e:
        return False, "다운로드에 실패했습니다: {}".format(e)

    # 받은 파일이 진짜 exe 인지 확인 (404 페이지를 받아오는 경우 방지)
    try:
        if os.path.getsize(new_exe) < 1024 * 1024:
            _safe_remove(new_exe)
            return False, "받은 파일이 올바르지 않습니다. 잠시 후 다시 시도해 주세요."
        with open(new_exe, "rb") as f:
            if f.read(2) != b"MZ":
                _safe_remove(new_exe)
                return False, "받은 파일이 프로그램이 아닙니다."
    except OSError:
        return False, "받은 파일을 확인하지 못했습니다."

    if not _launch_worker_exe(new_exe, target_dir):
        return False, "업데이트를 시작하지 못했습니다."
    return True, "업데이트를 시작합니다."


def is_worker_invocation():
    return ("--apply-exe" in sys.argv) or ("--apply-update" in sys.argv)


# ── 임시 폴더 정리 ──────────────────────────────────────────
def clean_temp():
    """
    지난 업데이트가 %TEMP% 에 남긴 것을 지운다.
    (worker 는 자기 폴더를 스스로 지울 수 없어서, 다음 실행 때 여기서 치운다)
    """
    base = tempfile.gettempdir()
    me = ""
    if getattr(sys, "frozen", False):
        me = os.path.abspath(sys.executable).lower()

    for name in ("JE_automation_worker", "JE_automation_extract",
                 "JE_automation_backup", "JE_automation_update.zip",
                 "JE_automation_download"):
        path = os.path.join(base, name)
        if not os.path.exists(path):
            continue
        # 지금 실행 중인 exe 가 이 안에 있으면 건드리지 않는다
        if me and me.startswith(os.path.abspath(path).lower()):
            continue
        _safe_remove(path)
