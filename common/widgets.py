# -*- coding: utf-8 -*-
"""
공통 부품
------------------------------------------------------------
세 기능이 각자 다른 모양으로 상태를 표시하던 것을 하나로 맞춥니다.

StatusBar   : 화면 맨 아래 한 줄. 왼쪽=진행/안내, 오른쪽=건수·합계
StatusProxy : 기존 코드가 label.config(text=..., foreground=...) 로 쓰던 걸
              그대로 받아서 StatusBar 로 넘겨주는 연결 부품
              (원본 코드를 많이 고치지 않기 위한 장치)
"""

import os
import tkinter as tk

FONT = "맑은 고딕"

BG = "#f8fafc"
LINE = "#e2e8f0"
NORMAL = "#64748b"   # 평소 안내
BUSY = "#2563eb"     # 작업 중
WARN = "#b45309"     # 주의 (빨강 대신 눈에 덜 거슬리는 색)
ERROR = "#b91c1c"    # 실패


class StatusBar(tk.Frame):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("bg", BG)
        super().__init__(master, highlightthickness=1, highlightbackground=LINE,
                         bd=0, **kwargs)
        self._left = tk.Label(self, text="", bg=BG, fg=NORMAL,
                              font=(FONT, 9), anchor="w", justify="left")
        self._left.pack(side="left", padx=14, pady=6)
        self._right = tk.Label(self, text="", bg=BG, fg="#475569",
                               font=(FONT, 9, "bold"), anchor="e")
        self._right.pack(side="right", padx=14, pady=6)

    def set_status(self, text, kind="normal"):
        color = {"normal": NORMAL, "busy": BUSY, "warn": WARN, "error": ERROR}.get(kind, NORMAL)
        self._left.config(text=text or "", fg=color)

    def set_summary(self, text):
        self._right.config(text=text or "")

    def clear(self):
        self._left.config(text="")
        self._right.config(text="")


def normalize_dir(path):
    """
    경로 표시를 윈도우 형식으로 통일한다.

    설정 파일에 `//서버/공유/폴더` 처럼 **슬래시**로 저장된 경우가 있어서
    탭마다 경로 모양이 달라 보였다. (`\\\\서버\\공유\\폴더` 로 맞춘다)
    빈 값은 그대로 빈 값으로 돌려준다.
    """
    p = str(path or "").strip()
    if not p:
        return ""
    try:
        return os.path.normpath(p)
    except Exception:
        return p


def is_network_path(path):
    """
    네트워크 경로인가?

    - `\\\\서버\\공유`  (백슬래시)
    - `//서버/공유`     (설정 파일에 슬래시로 저장된 경우도 있다 — 실제로 겪음)
    - `Z:\\` 처럼 **매핑된 네트워크 드라이브**
    """
    p = str(path or "").strip().replace("/", "\\")
    if not p:
        return False
    if p.startswith("\\\\"):
        return True
    try:
        import ctypes
        drive = os.path.splitdrive(p)[0]
        if drive:
            # 4 = DRIVE_REMOTE (네트워크 드라이브)
            return ctypes.windll.kernel32.GetDriveTypeW(drive + "\\") == 4
    except Exception:
        pass
    return False


def safe_start_dir(*candidates):
    """폴더 선택창을 여는 위치.

    현재 경로를 그대로 넘기면, 그 경로가 네트워크 공유일 때
    윈도우가 접속을 시도하느라 창이 몇 초씩 멈춘다.
    **권한이 없는 공유면 훨씬 오래 멈춘다.**
    그래서 네트워크 경로와 없는 경로는 건너뛰고,
    쓸 만한 위치가 없으면 '문서' 폴더에서 연다.
    """
    for path in candidates:
        if not isinstance(path, str):
            continue
        path = path.strip()
        if not path or is_network_path(path):
            continue
        try:
            if not os.path.isdir(path):
                continue
        except Exception:
            continue
        return path
    home = os.path.expanduser("~")
    docs = os.path.join(home, "Documents")
    return docs if os.path.isdir(docs) else home


class StatusProxy:
    """label.config(...) 호출을 상태 표시줄로 넘긴다."""

    _WARN_COLORS = {"#dc2626", "#b91c1c", "#b45309", "#c2410c", "#d97706", "red"}

    def __init__(self, bar, target="status"):
        self._bar = bar
        self._target = target

    def config(self, text=None, foreground=None, **kwargs):
        if self._target == "summary":
            if text is not None:
                self._bar.set_summary(text)
            return
        kind = "warn" if foreground in self._WARN_COLORS else "normal"
        if text is not None:
            self._bar.set_status(text, kind)

    configure = config

    def cget(self, key):
        return ""

    def __getattr__(self, name):
        # pack/grid 등 나머지 호출은 조용히 무시 (원본 코드가 깨지지 않게)
        def _ignore(*args, **kwargs):
            return None
        return _ignore
