# -*- coding: utf-8 -*-
"""
exe 만들기
------------------------------------------------------------
build.bat 이 이 파일을 부릅니다.
배치 파일(.bat)은 한글·줄바꿈 문제가 잘 생기므로, 실제 작업은 파이썬이 합니다.

배포 형태
  - **폴더 형태(onedir)** 로 만듭니다. exe 하나가 스스로 압축을 풀어 실행하는 방식(onefile)은
    백신이 악성코드로 오해하는 경우가 많아서 바꿨습니다.
  - 결과물은 dist/JE_automation/ 폴더이고, 이 폴더를 zip 으로 묶어 전달합니다.
  - exe 에 회사명·제품명·버전·아이콘을 넣어 '누가 만든 프로그램인지' 보이게 했습니다.

직접 실행해도 됩니다:  python build.py
"""

import json
import os
import shutil
import subprocess
import sys
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))

# ── 빌드용 파이썬 찾기 ───────────────────────────────────────
# 이 PC 의 고정 경로를 소스에 적어 두면 저장소에 개인 경로가 남습니다.
# (공개 저장소로 올릴 수 없게 됩니다) 그래서 자동으로 찾습니다.
#
#   1) _dev/build_python.txt   ← 이 PC 전용 경로 (gitignore 대상이라 안 올라감)
#   2) 환경변수 JE_BUILD_PY
#   3) 지금 실행 중인 파이썬
#   4) py -3  /  python   (PATH 에 있는 것)
#   5) C:\Program Files\Python314\python.exe
#
# PyInstaller 와 tkinter 가 **둘 다** 있어야 빌드할 수 있습니다.
LOCAL_PY_FILE = os.path.join(BASE, "_dev", "build_python.txt")

# 버전을 올릴 때는 여기와 main.py 의 APP_VERSION, CHANGELOG.md 를 같이 고친다.
APP_VERSION = "1.1.8"

# exe 파일 이름은 **고정**입니다. (JE_automation.exe)
#  - 이름에 버전을 붙이면 자동 업데이트 때마다 exe 가 쌓이고, 바로가기가 깨진다
#  - 버전은 창 제목과 exe 속성(→ 자세히)에 표시된다
#  - 이름은 영어로 (한글 이름은 압축, 백신, 도구에서 깨지기 쉬움)
EXE_BASE = "JE_automation"
EXE_NAME = EXE_BASE

COMPANY_NAME = "JE"
PRODUCT_NAME = "JE automation"
DESCRIPTION = "JE automation (운용지시대장 · 자금판 · 합계검증 · BGC 레포)"
COPYRIGHT = "© 2026 JE"

GITHUB_OWNER = "doctorjay-code"
GITHUB_REPO = "JE_automation"

DIST_DIR = os.path.join(BASE, "dist")
BUILD_DIR = os.path.join(BASE, "build")

# 자동 업데이트용 zip 을 두는 저장소 폴더.
# GitHub Releases 를 쓰지 않고, zip 을 저장소에 함께 올려 git push 만으로 배포한다.
# 프로그램은 이 파일을 raw.githubusercontent.com 에서 내려받는다.
UPDATE_DIR = os.path.join(BASE, "update")

# 프로그램 안 '변경 이력' 에 보여줄 내용 (저장소에 함께 올라간다)
CHANGELOG_PATH = os.path.join(BASE, "changelog.json")
ICON_PATH = os.path.join(BASE, "assets", "icon.ico")
CONFIG_NAME = "JE_automation_settings.json"

# 예전 이름들 — 설정을 물려받고 옛 폴더를 찾을 때만 씁니다.
# ※ 새로 배포하는 파일 이름에는 쓰지 않습니다.
#   [0] 대소문자만 다르던 이름 (JE_Automation)
#   [1] 한글 이름 (JE_업무자동화)
LEGACY_EXE_NAMES = ["JE_Automation", "JE_업무자동화"]
LEGACY_CONFIG_NAMES = ["JE_Automation_settings.json", "JE_업무자동화_설정.json"]

VERSION_TEMPLATE = """VSVersionInfo(
  ffi=FixedFileInfo(
    filevers=({v0}, {v1}, {v2}, 0),
    prodvers=({v0}, {v1}, {v2}, 0),
    mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)
  ),
  kids=[
    StringFileInfo([
      StringTable('041204B0', [
        StringStruct('CompanyName', '{company}'),
        StringStruct('FileDescription', '{desc}'),
        StringStruct('FileVersion', '{version}.0'),
        StringStruct('InternalName', '{name}'),
        StringStruct('LegalCopyright', '{copyright}'),
        StringStruct('OriginalFilename', '{name}.exe'),
        StringStruct('ProductName', '{product}'),
        StringStruct('ProductVersion', '{version}.0')
      ])
    ]),
    VarFileInfo([VarStruct('Translation', [1042, 1200])])
  ]
)
"""


def fail(msg):
    print()
    print("[실패] " + msg)
    return 1


def _candidate_pythons():
    """빌드에 쓸 수 있는 파이썬 후보들 (순서대로 시도)."""
    cands = []

    # 1) 이 PC 전용 지정 파일 (저장소에는 올라가지 않는다)
    if os.path.exists(LOCAL_PY_FILE):
        try:
            with open(LOCAL_PY_FILE, "r", encoding="utf-8") as f:
                cands.append(f.read().strip())
        except Exception:
            pass

    # 2) 환경변수
    env = os.environ.get("JE_BUILD_PY", "").strip()
    if env:
        cands.append(env)

    # 3) 지금 실행 중인 파이썬
    cands.append(sys.executable)

    # 4) PATH 에 있는 것
    cands.append("py")
    cands.append("python")

    # 5) windows 기본 설치 위치
    cands.append(r"C:\Program Files\Python314\python.exe")

    seen = set()
    out = []
    for c in cands:
        if c and c not in seen:
            seen.add(c)
            out.append(c)
    return out


def _python_ok(path):
    """이 파이썬에 PyInstaller 와 tkinter 가 들어 있는지 확인한다."""
    args = []
    base = os.path.basename(path).lower()
    if base.startswith("py") and not base.startswith("python"):
        args = ["-3"]          # py 런처는 버전 지정이 필요
    try:
        result = subprocess.run(
            [path] + args + ["-c", "import PyInstaller, tkinter"],
            stdout=subprocess.PIPE, stderr=subprocess.PIPE)
        return result.returncode == 0
    except Exception:
        return False


def find_build_python():
    """빌드에 쓸 파이썬을 찾는다. 없으면 None."""
    for c in _candidate_pythons():
        if _python_ok(c):
            return c
    return None


def is_locked(path):
    """다른 프로그램이 이 파일을 붙잡고 있으면 True."""
    if not os.path.exists(path):
        return False
    try:
        with open(path, "a+b"):
            return False
    except OSError:
        return True


def write_version_file(path):
    parts = [int(x) for x in APP_VERSION.split(".")]
    while len(parts) < 3:
        parts.append(0)
    text = VERSION_TEMPLATE.format(
        v0=parts[0], v1=parts[1], v2=parts[2],
        version=APP_VERSION, name=EXE_NAME,
        company=COMPANY_NAME, product=PRODUCT_NAME,
        desc=DESCRIPTION, copyright=COPYRIGHT,
    )
    with open(path, "w", encoding="utf-8") as f:
        f.write(text)
    return path


def _old_version_dirs():
    """예전 버전 빌드 폴더 이름들 (JE_업무자동화_v1.0.1, JE_Automation_v1.0.0 등)"""
    names = []
    if not os.path.isdir(DIST_DIR):
        return names
    prefixes = tuple([EXE_BASE + "_v"] + [n + "_v" for n in LEGACY_EXE_NAMES])
    for name in os.listdir(DIST_DIR):
        if name.startswith(prefixes) and os.path.isdir(os.path.join(DIST_DIR, name)):
            names.append(name)
    return names


def _config_candidates(app_dir):
    """
    사용자 설정을 찾을 후보 경로 (우선순위 순).
    파일·폴더 이름이 한글 → JE_Automation → JE_automation 으로 바뀌었으므로,
    지금 이름이 없으면 예전 이름들을 차례로 찾아서 물려준다.
    """
    names = [CONFIG_NAME] + LEGACY_CONFIG_NAMES
    cands = [os.path.join(app_dir, n) for n in names]

    # 폴더 이름이 바뀐 예전 빌드
    #   JE_automation/      ← 버전 없는 고정 이름이던 시절
    #   JE_업무자동화/, JE_Automation/
    #   …_v1.0.1/ 등
    for folder in [EXE_BASE] + LEGACY_EXE_NAMES + _old_version_dirs():
        base = os.path.join(DIST_DIR, folder)
        for n in names:
            cands.append(os.path.join(base, n))
    return cands


def make_zip(folder, dest_path, include_config=False):
    """
    zip 을 만든다.

    include_config=True  → 사용자 설정 파일까지 넣는다. (사람에게 직접 줄 때)
    include_config=False → 설정 파일을 뺀다. (저장소에 올릴 때 — 개인 경로 노출 방지)

    zip 안쪽 최상위 폴더 이름은 exe 이름과 같게 고정한다 (JE_automation/).
    """
    if os.path.exists(dest_path):
        os.remove(dest_path)
    parent = os.path.dirname(dest_path)
    if parent:
        os.makedirs(parent, exist_ok=True)

    with zipfile.ZipFile(dest_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for root, _dirs, files in os.walk(folder):
            for name in files:
                if name == CONFIG_NAME and not include_config:
                    continue
                full = os.path.join(root, name)
                rel = os.path.relpath(full, folder)          # _internal/... 또는 exe
                zf.write(full, os.path.join(EXE_NAME, rel))   # JE_automation/...
    return dest_path


def _clean_dist_zips(keep=None):
    """dist 안의 zip 을 하나만 남긴다. (버전마다 29MB씩 쌓이는 것 방지)"""
    if not os.path.isdir(DIST_DIR):
        return
    for name in os.listdir(DIST_DIR):
        if not name.lower().endswith(".zip"):
            continue
        if keep and name == keep:
            continue
        try:
            os.remove(os.path.join(DIST_DIR, name))
            print("예전 zip 정리:", name)
        except OSError:
            pass


def _ver_key(text):
    parts = []
    for chunk in str(text).split("."):
        digits = "".join(c for c in chunk if c.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def load_changelog():
    """changelog.json — 프로그램 안 '변경 이력' 에 보여줄 내용."""
    try:
        with open(CHANGELOG_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
        if isinstance(data, dict):
            return {str(k).strip(): [str(x) for x in v if str(x).strip()]
                    for k, v in data.items() if isinstance(v, list)}
    except Exception:
        pass
    return {}


def save_changelog(data):
    ordered = {k: data[k] for k in sorted(data, key=_ver_key, reverse=True)}
    with open(CHANGELOG_PATH, "w", encoding="utf-8") as f:
        json.dump(ordered, f, ensure_ascii=False, indent=2)
    return CHANGELOG_PATH


def ask_notes(version):
    """
    이번 버전에 바뀐 내용을 한 줄씩 입력받는다.
    대화형(터미널)이 아니면 묻지 않고 빈 목록을 돌려준다.
    """
    try:
        if not sys.stdin or not sys.stdin.isatty():
            return []
    except Exception:
        return []

    print()
    print("이번 버전({})에 바뀐 내용을 한 줄씩 적어주세요. (그냥 Enter 면 끝)".format(version))
    notes = []
    while True:
        try:
            line = input("  > ").strip()
        except (EOFError, KeyboardInterrupt):
            break
        if not line:
            break
        notes.append(line)
    return notes


def _previous_version_json():
    """지난 빌드의 version.json (있으면). deps 지문을 비교하는 데 쓴다."""
    try:
        with open(os.path.join(BASE, "version.json"), "r", encoding="utf-8") as f:
            data = json.load(f)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def copy_to_update_folder(exe_path, zip_path=None):
    """
    저장소의 update/ 폴더에 배포 파일을 둔다. (git push 만으로 배포)

    - exe 는 **항상** 넣는다 (약 5MB). 자동 업데이트가 이 파일 하나만 받는다.
    - zip 은 `_internal` 이 바뀐 버전에서만 넣는다 (전체 교체가 필요할 때).
    - 예전 파일은 지운다. (저장소가 무거워지지 않게)

    반환: (exe 파일명, zip 파일명 또는 None)
    """
    os.makedirs(UPDATE_DIR, exist_ok=True)

    exe_name = os.path.basename(exe_path)
    shutil.copy2(exe_path, os.path.join(UPDATE_DIR, exe_name))
    keep = {exe_name}

    zip_name = None
    if zip_path:
        zip_name = os.path.basename(zip_path)
        dest = os.path.join(UPDATE_DIR, zip_name)
        if os.path.abspath(zip_path) != os.path.abspath(dest):
            shutil.copy2(zip_path, dest)
        keep.add(zip_name)

    for name in os.listdir(UPDATE_DIR):
        if name in keep:
            continue
        if name.lower().endswith((".exe", ".zip")):
            try:
                os.remove(os.path.join(UPDATE_DIR, name))
                print("예전 배포파일 정리:", name)
            except OSError:
                pass

    return exe_name, zip_name


def write_version_json(exe_filename, zip_filename, deps, notes=None):
    """
    프로그램이 버전 확인을 때 읽는 파일.

    `exe_url`  — 새 exe 하나만 받는 주소 (기본 업데이트)
    `deps`     — `_internal` 지문. 다르면 exe 만으로는 안 되므로 full(zip) 로 간다.
    `full_url` — zip 주소 (지문이 바뀐 버전에서만 채워진다)
    `notes`    — 이번 버전에 바뀐 내용 (업데이트 확인 창에 표시)
    """
    raw = "https://raw.githubusercontent.com/{}/{}/main/update".format(
        GITHUB_OWNER, GITHUB_REPO)

    data = {
        "version": APP_VERSION,
        "exe_url": "{}/{}".format(raw, exe_filename),
        "exe_file": exe_filename,
        "deps": deps,
        "full_url": ("{}/{}".format(raw, zip_filename) if zip_filename else ""),
        "full_file": zip_filename or "",
        "notes": list(notes or []),
    }
    path = os.path.join(BASE, "version.json")
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    return path


def main():
    python_path = find_build_python()
    if not python_path:
        return fail(
            "빌드용 파이썬을 찾지 못했습니다.\n"
            "  PyInstaller 와 tkinter 가 들어 있는 파이썬이 필요합니다.\n"
            "  설치:  pip install pyinstaller\n"
            "  또는 _dev\\build_python.txt 에 파이썬 경로를 적어주세요."
        )
    print("빌드용 파이썬:", python_path)

    # ── 이번 버전에 바뀐 내용 (프로그램 안 '변경 이력' + 업데이트 확인 창) ──
    #   --note "내용" 으로 넘길 수도 있고, 터미널에서 실행하면 물어본다.
    notes = []
    for i, arg in enumerate(sys.argv):
        if arg == "--note" and i + 1 < len(sys.argv):
            text = sys.argv[i + 1].strip()
            if text:
                notes.append(text)

    changelog = load_changelog()
    if not notes and APP_VERSION not in changelog:
        notes = ask_notes(APP_VERSION)
    if notes:
        changelog[APP_VERSION] = notes
        save_changelog(changelog)
    elif APP_VERSION not in changelog:
        print()
        print("[주의] changelog.json 에 {} 항목이 없습니다.".format(APP_VERSION))
        print("       업데이트 확인 창에 '바뀐 내용' 없이 표시됩니다.")
    notes_now = changelog.get(APP_VERSION, [])

    # 빌드 결과 폴더도 exe 이름을 따라 고정된다. (PyInstaller 가 --name 으로 폴더를 만든다)
    # (예전 버전 폴더는 정리해 준다)
    app_dir = os.path.join(DIST_DIR, EXE_NAME)
    exe_path = os.path.join(app_dir, EXE_NAME + ".exe")
    config_path = os.path.join(app_dir, CONFIG_NAME)
    before = os.path.getmtime(exe_path) if os.path.exists(exe_path) else 0

    # ★ 빌드는 결과 폴더를 통째로 새로 만든다 → 사용자 설정 파일이 지워진다.
    #   미리 읽어두었다가 빌드 후 되돌려 놓는다.
    #   (이름을 한글→영어로 바꿨으므로 예전 이름의 설정도 찾아본다)
    config_backup = None
    for cand in _config_candidates(app_dir):
        if not os.path.exists(cand):
            continue
        try:
            with open(cand, "r", encoding="utf-8") as f:
                config_backup = f.read()
        except Exception:
            config_backup = None
            continue
        if os.path.abspath(cand) != os.path.abspath(config_path):
            print("설정 파일 물려받기:", cand)
        break

    # ※ 예전 버전 폴더 정리는 **빌드가 성공한 뒤**에 한다.
    #   (중간에 실패하면 설정이 사라질 수 있으므로 — 실제로 겪음)

    if is_locked(exe_path):
        return fail(
            "지금 프로그램이 실행 중이라 파일을 덮어쓸 수 없습니다.\n"
            "  열려 있는 'JE automation' 창을 모두 닫고 다시 실행해 주세요."
        )

    os.makedirs(BUILD_DIR, exist_ok=True)
    version_file = write_version_file(os.path.join(BUILD_DIR, "version_info.txt"))

    cmd = [
        python_path, "-m", "PyInstaller",
        "--noconfirm", "--onedir", "--windowed", "--clean",
        # exe 이름은 버전을 빼고 고정한다. (자동 업데이트에서 이름이 바뀌면
        #  바로가가 깨지고 예전 exe 가 쌓인다)
        "--name", EXE_NAME,
        "--version-file", version_file,
    ]
    if os.path.exists(ICON_PATH):
        cmd += ["--icon", ICON_PATH]
    cmd += [
        # 합계검증의 '엑셀 열고 저장·닫기' 에 필요 (없으면 셀값이 비어 나온다)
        "--hidden-import", "pythoncom",
        "--hidden-import", "pywintypes",
        "--hidden-import", "win32com",
        "--hidden-import", "win32com.client",
        # 자동 업데이트: GitHub API 호출용 (표준 라이브러리지만 명시)
        "--hidden-import", "urllib.request",
        "--hidden-import", "urllib.error",
        "--hidden-import", "urllib.parse",
        "--exclude-module", "PIL",
        "--exclude-module", "numpy",
        "--exclude-module", "pytest",
        "main.py",
    ]

    print("빌드를 시작합니다... (1~2분)")
    print()
    result = subprocess.run(cmd, cwd=BASE)
    if result.returncode != 0:
        return fail("PyInstaller 가 오류로 끝났습니다. 위 메시지를 확인해 주세요.")

    if not os.path.exists(exe_path):
        return fail("exe 가 만들어지지 않았습니다.\n  " + exe_path)

    if os.path.getmtime(exe_path) <= before:
        return fail(
            "exe 가 갱신되지 않았습니다.\n"
            "  이미 실행 중인 창을 모두 닫고 다시 시도해 주세요.\n  " + exe_path
        )

    # 빌드가 지운 사용자 설정 파일을 되돌린다
    if config_backup is not None and not os.path.exists(config_path):
        try:
            with open(config_path, "w", encoding="utf-8") as f:
                f.write(config_backup)
            print("설정 파일 유지:", config_path)
        except Exception:
            pass

    # 빌드가 끝났으니 예전 버전 폴더를 정리한다.
    # (설정을 이미 새 폴더로 되돌려 놓았으므로 안전)
    if os.path.isdir(DIST_DIR):
        for name in _old_version_dirs():
            shutil.rmtree(os.path.join(DIST_DIR, name), ignore_errors=True)

    # ── _internal 이 바뀌었나? (exe 만으로 업데이트가 되는지 판단) ──
    from common.updater import deps_fingerprint
    deps_now = deps_fingerprint(app_dir)
    deps_before = str(_previous_version_json().get("deps", "") or "")
    deps_changed = (deps_now != deps_before)

    # zip 은 **필요할 때만** 만든다. (평소 빌드는 exe 만)
    #   - 처음 주는 사람이 생겼을 때 : python build.py --zip
    #     → dist/JE_automation.zip  (설정 포함, 이름 고정)
    #   - 라이브러리가 바뀐 버전 : 업데이트가 전체 교체를 해야 해서 필요
    #     → update/JE_automation_v1.1.7.zip (설정 없음 — 저장소에 올라가므로)
    want_zip = ("--zip" in sys.argv)

    delivery_zip = None
    if want_zip:
        delivery_zip = make_zip(
            app_dir,
            os.path.join(DIST_DIR, EXE_NAME + ".zip"),
            include_config=os.path.exists(config_path),
        )
    _clean_dist_zips(keep=os.path.basename(delivery_zip) if delivery_zip else None)

    upload_zip = None
    if deps_changed:
        upload_zip = make_zip(
            app_dir,
            os.path.join(UPDATE_DIR, "{}_v{}.zip".format(EXE_NAME, APP_VERSION)),
            include_config=False,
        )

    exe_name, zip_name = copy_to_update_folder(exe_path, upload_zip)
    version_json = write_version_json(exe_name, zip_name, deps_now, notes_now)

    total = 0
    for root, _dirs, files in os.walk(app_dir):
        for name in files:
            total += os.path.getsize(os.path.join(root, name))
    size_mb = total / (1024 * 1024)

    print()
    print("[완료] 폴더 {:.1f}MB".format(size_mb))
    print("  실행 파일  : " + exe_path)
    if delivery_zip:
        print("  배포용 zip : " + delivery_zip + "   ← 이 파일 하나만 주면 됩니다")
    print("  버전 정보  : " + version_json)
    print()
    print("=== 자동 업데이트 (저장소 update/ 폴더) ===")
    print("  올라가는 exe : update/" + exe_name + "   (약 {:.1f}MB)".format(
        os.path.getsize(os.path.join(UPDATE_DIR, exe_name)) / 1024 / 1024))
    if zip_name:
        print("  올라가는 zip : update/" + zip_name + "   (전체 교체용)")
        print("  → _internal 폴더가 바뀐 버전이라 zip 도 함께 올립니다.")
    else:
        print("  zip 은 만들지도, 올리지도 않습니다. exe 하나만 바뀌었습니다.")
    print()
    print("1) 저장소에 올리기:")
    print("     git add .")
    print("     git commit -m \".\"")
    print("     git push")
    print("2) 끝. 사용자가 프로그램을 켜면 새 버전을 찾아 자동으로 받습니다.")
    if not delivery_zip:
        print()
        print("※ 처음 주는 사람이 생기면:  python build.py --zip")

    if os.path.isdir(BUILD_DIR):
        shutil.rmtree(BUILD_DIR, ignore_errors=True)

    # PyInstaller 가 만든 .spec 은 이름이 exe 이름(한글)을 따라가고, 다음 빌드에서
    # 어차피 다시 만들어집니다. 폴더에 남겨두면 지저분하니 지웁니다.
    spec_path = os.path.join(BASE, EXE_NAME + ".spec")
    try:
        if os.path.exists(spec_path):
            os.remove(spec_path)
    except OSError:
        pass

    return 0


if __name__ == "__main__":
    sys.exit(main())
