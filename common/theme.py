# -*- coding: utf-8 -*-
"""
화면 디자인 통일
------------------------------------------------------------
세 기능이 각자 만든 버튼·표·글꼴을 그대로 쓰면 들쑥날쑥해집니다.
이 파일에서 한 번에 정하고, 모든 탭에 같은 규칙을 적용합니다.

색 규칙
  본문 배경   #ffffff
  진한 머리글 #0f172a
  포인트(파랑) #2563eb
  보조 글자   #64748b
"""

import tkinter as tk
from tkinter import ttk

FONT = "맑은 고딕"

BG = "#ffffff"
BG_SOFT = "#f8fafc"
INK = "#0f172a"
TEXT = "#1f2937"
MUTED = "#64748b"
LINE = "#e2e8f0"
ACCENT = "#2563eb"
ACCENT_DARK = "#1d4ed8"
SUCCESS = "#16a34a"
SUCCESS_DARK = "#15803d"

FONT_UI = (FONT, 10)
FONT_UI_BOLD = (FONT, 10, "bold")
FONT_SMALL = (FONT, 9)
FONT_TITLE = (FONT, 16, "bold")


def apply(root):
    """앱 전체에 공통 디자인을 적용한다. 여러 번 불러도 안전."""
    try:
        root.configure(bg=BG)
    except tk.TclError:
        pass

    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except tk.TclError:
        pass

    style.configure(".", font=FONT_UI, background=BG, foreground=TEXT)

    style.configure("TFrame", background=BG)
    style.configure("Card.TFrame", background=BG_SOFT)
    style.configure("TLabel", font=FONT_UI, background=BG, foreground=TEXT)
    style.configure("Muted.TLabel", font=FONT_SMALL, background=BG, foreground=MUTED)
    style.configure("Title.TLabel", font=FONT_TITLE, background=BG, foreground=INK)

    style.configure("TLabelframe", background=BG, bordercolor=LINE,
                    relief="solid", borderwidth=1, padding=8)
    style.configure("TLabelframe.Label", font=FONT_UI_BOLD, background=BG, foreground=INK)

    style.configure("TButton", font=FONT_UI, padding=(14, 7))
    style.map("TButton",
              background=[("active", "#eef2f7")],
              foreground=[("disabled", "#9aa4b2")])

    style.configure("Accent.TButton", font=FONT_UI_BOLD, padding=(14, 7),
                    background=ACCENT, foreground="#ffffff", borderwidth=0)
    style.map("Accent.TButton",
              background=[("active", ACCENT_DARK), ("disabled", "#9db8ee")],
              foreground=[("disabled", "#ffffff")])

    style.configure("Primary.TButton", font=FONT_UI_BOLD, padding=(14, 7),
                    background=ACCENT, foreground="#ffffff", borderwidth=0)
    style.map("Primary.TButton",
              background=[("active", ACCENT_DARK), ("disabled", "#9db8ee")],
              foreground=[("disabled", "#ffffff")])

    # 초록은 쓰지 않는다. (예전 코드가 부르더라도 파랑으로 보이게)
    style.configure("Success.TButton", font=FONT_UI_BOLD, padding=(14, 7),
                    background=ACCENT, foreground="#ffffff", borderwidth=0)
    style.map("Success.TButton",
              background=[("active", ACCENT_DARK), ("disabled", "#9db8ee")],
              foreground=[("disabled", "#ffffff")])

    # 경로 영역 전용 — 흰색(찾아보기)과 파랑(스캔)이 **폭이 같아야** 한다.
    # 굵기·테두리 폭이 다르면 버튼 크기가 어긋나므로 둘 다 똑같이 맞춘다.
    style.configure("Path.TButton", font=FONT_UI, padding=(14, 7),
                    background="#ffffff", bordercolor=LINE,
                    borderwidth=1, relief="solid")
    style.map("Path.TButton",
              background=[("active", "#eef2f7")],
              foreground=[("disabled", "#9aa4b2")])

    style.configure("PathAccent.TButton", font=FONT_UI, padding=(14, 7),
                    background=ACCENT, foreground="#ffffff",
                    bordercolor=ACCENT, borderwidth=1, relief="solid")
    style.map("PathAccent.TButton",
              background=[("active", ACCENT_DARK), ("disabled", "#9db8ee")],
              foreground=[("disabled", "#ffffff")])

    style.configure("TEntry", fieldbackground="#ffffff", padding=4)
    style.configure("TCombobox", padding=4)
    style.configure("TCheckbutton", background=BG, font=FONT_UI)
    style.map("TCheckbutton", background=[("active", BG)])

    style.configure("Treeview", font=(FONT, 9), rowheight=28,
                    background="#ffffff", fieldbackground="#ffffff",
                    borderwidth=0, relief="flat")
    style.configure("Treeview.Heading", font=FONT_UI_BOLD,
                    background=LINE, foreground=INK, relief="flat", padding=(6, 6))
    style.map("Treeview",
              background=[("selected", ACCENT)],
              foreground=[("selected", "#ffffff")])
    style.map("Treeview.Heading", background=[("active", "#d7dfe9")])

    style.configure("Vertical.TScrollbar", background="#e2e8f0", troughcolor=BG_SOFT,
                    borderwidth=0, arrowsize=12)
    style.configure("Horizontal.TScrollbar", background="#e2e8f0", troughcolor=BG_SOFT,
                    borderwidth=0, arrowsize=12)

    # 탭: 선택 여부와 상관없이 크기가 같아야 한다.
    # (ttk 기본값은 '선택된 탭만 여백을 6 4 6 2 로 줄여서' 누르면 쪼그라든다)
    tab_pad = (20, 10)
    style.configure("TNotebook", background="#eef2f7", borderwidth=0, tabmargins=(10, 8, 10, 0))
    style.configure("TNotebook.Tab", font=FONT_UI, padding=tab_pad,
                    background="#dde5ee", foreground="#475569",
                    borderwidth=0, focuscolor="")
    style.map("TNotebook.Tab",
              padding=[("selected", tab_pad), ("!selected", tab_pad), ("active", tab_pad)],
              background=[("selected", BG), ("active", "#e8eef5")],
              foreground=[("selected", ACCENT_DARK), ("active", INK)])

    # 선택된 탭에 생기는 점선 테두리(포커스 표시)를 없앤다.
    try:
        style.layout("TNotebook.Tab", [
            ("Notebook.tab", {"sticky": "nswe", "children": [
                ("Notebook.padding", {"side": "top", "sticky": "nswe", "children": [
                    ("Notebook.label", {"side": "top", "sticky": ""}),
                ]}),
            ]}),
        ])
    except tk.TclError:
        pass
    return style
