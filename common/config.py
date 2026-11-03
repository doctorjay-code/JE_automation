# -*- coding: utf-8 -*-
"""
설정 파일 하나로 관리하기
------------------------------------------------------------
파일은 하나(JE_automation_settings.json)이고, 안에서 기능별로 구역을 나눕니다.

{
  "version": "1.0.0",
  "공통":        { "window_geometry": "1280x820+300+120", "last_tab": 0 },
  "운용지시대장": { ... },
  "자금판":      { ... },
  "합계검증":    { ... }
}

구역을 나눠 두면 한 기능의 설정만 초기화하기 쉽고,
다른 기능이 설정을 잘못 건드려도 피해가 번지지 않습니다.
"""

import json
import os
import sys
import threading

CONFIG_FILENAME = "JE_automation_settings.json"

# 예전 이름들. 읽기만 하고, 없으면 그냥 넘어갑니다.
# (처음 쓰는 사람에게는 아무 일도 일어나지 않음)
#   [0] 대소문자만 다르던 이름 (JE_Automation)
#   [1] 한글 이름 (JE_업무자동화)
LEGACY_CONFIG_FILENAMES = [
    "JE_Automation_settings.json",
    "JE_업무자동화_설정.json",
]

SECTION_COMMON = "공통"
SECTION_LEDGER = "운용지시대장"
SECTION_FUNDBOARD = "자금판"
SECTION_SUMCHECK = "합계검증"

_LOCK = threading.RLock()


def app_dir():
    """EXE로 실행하면 EXE가 있는 폴더, 소스로 실행하면 이 프로젝트 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


CONFIG_PATH = os.path.join(app_dir(), CONFIG_FILENAME)


def _read_all():
    """설정 파일 전체를 읽는다. 없거나 깨졌으면 빈 dict.
    새 이름이 없으면 예전 이름들도 차례로 찾아본다."""
    paths = [CONFIG_PATH] + [os.path.join(app_dir(), n) for n in LEGACY_CONFIG_FILENAMES]
    for path in paths:
        try:
            if os.path.exists(path):
                with open(path, "r", encoding="utf-8") as f:
                    data = json.load(f)
                if isinstance(data, dict):
                    return data
        except Exception:
            continue
    return {}


def _write_all(data):
    """통째로 저장. 임시 파일에 쓴 뒤 바꿔치기해서 저장 중 깨짐을 막는다."""
    tmp = CONFIG_PATH + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    os.replace(tmp, CONFIG_PATH)


def load_all():
    with _LOCK:
        return _read_all()


def save_all(data):
    with _LOCK:
        if not isinstance(data, dict):
            return False
        data.pop("version", None)   # 버전 표시는 남기지 않는다
        try:
            _write_all(data)
            return True
        except Exception:
            return False


def load_section(name, default=None):
    """구역 하나만 읽는다. 없으면 default."""
    with _LOCK:
        data = _read_all()
    value = data.get(name)
    if value is None:
        return default
    return value


def save_section(name, value):
    """구역 하나만 저장한다. 다른 구역은 그대로 둔다."""
    with _LOCK:
        data = _read_all()
        data[name] = value
        data.pop("version", None)   # 버전 표시는 남기지 않는다
        try:
            _write_all(data)
            return True
        except Exception:
            return False


def reset_section(name, default=None):
    """구역 하나만 초기화."""
    return save_section(name, default)


def config_location():
    """설정 파일 경로 (안내 문구용)."""
    return CONFIG_PATH


# 예전 프로그램이 각자 쓰던 설정 파일 이름 → 통합 설정의 구역 이름
LEGACY_FILES = [
    ("운용지시대장_자동반영도구_설정.json", SECTION_LEDGER),
    ("자금판_자동반영도구_설정.json", SECTION_FUNDBOARD),
    ("은대_합계_검증_설정.json", SECTION_SUMCHECK),
]


def _copy_old_config():
    """
    예전 이름의 설정 파일을 지금 이름으로 복사한다. 한 번만.
    예전 파일은 지우지 않는다. (혹시 모를 되돌리기용)
    """
    if os.path.exists(CONFIG_PATH):
        return False
    for legacy in LEGACY_CONFIG_FILENAMES:
        old = os.path.join(app_dir(), legacy)
        if not os.path.exists(old):
            continue
        try:
            with open(old, "r", encoding="utf-8") as f:
                data = json.load(f)
            if not isinstance(data, dict):
                continue
            _write_all(data)
            return True
        except Exception:
            continue
    return False


def migrate_legacy():
    """
    통합 설정 파일이 아직 없고, 예전 설정 파일이 exe 옆에 있으면 한 번만 가져온다.
    (없으면 아무 일도 하지 않는다. 처음 쓰는 사람에게는 그냥 빈 설정으로 시작)
    """
    # 1) 파일 이름을 한글→영어로 바꿨다. 예전 이름이 있으면 새 이름으로 복사해 둔다.
    _copy_old_config()

    if os.path.exists(CONFIG_PATH):
        return []

    imported = []
    data = {}
    for filename, section in LEGACY_FILES:
        path = os.path.join(app_dir(), filename)
        try:
            if not os.path.exists(path):
                continue
            with open(path, "r", encoding="utf-8") as f:
                value = json.load(f)
            if isinstance(value, dict) and value:
                value.pop("version", None)
                data[section] = value
                imported.append(section)
        except Exception:
            continue

    if imported:
        save_all(data)
    return imported

