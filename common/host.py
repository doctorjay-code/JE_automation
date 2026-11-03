# -*- coding: utf-8 -*-
"""
PanelHost — 기능 화면을 담는 칸
------------------------------------------------------------
원래 세 프로그램은 각자 '창(tk.Tk)'을 직접 만들었습니다.
통합하면 창은 하나뿐이므로, 각 기능은 '탭 안의 칸'에 들어가야 합니다.

이 클래스는 칸(Frame)이면서, 창에만 통하는 요청(제목 바꾸기, 창 크기 바꾸기 등)을
조용히 무시합니다. 그래서 원본 코드를 거의 고치지 않고 그대로 옮길 수 있습니다.

- 창 크기·제목은 통합 프로그램이 하나로 관리합니다.
- 마우스 커서(모래시계)는 창 전체에 걸리도록 위로 넘겨줍니다.
"""

import tkinter as tk

# 창(window)에만 존재하는 기능들. 칸에서는 아무 일도 하지 않는다.
_WINDOW_ONLY = (
    "title", "wm_title", "geometry", "wm_geometry",
    "minsize", "wm_minsize", "maxsize", "wm_maxsize",
    "resizable", "wm_resizable", "iconbitmap", "iconphoto",
    "protocol", "wm_protocol", "attributes", "wm_attributes",
    "state", "wm_state", "overrideredirect", "transient",
    "lift", "wm_iconbitmap",
)


class PanelHost(tk.Frame):
    def __init__(self, master, **kwargs):
        kwargs.setdefault("bg", "#ffffff")
        kwargs.setdefault("highlightthickness", 0)
        kwargs.setdefault("bd", 0)
        super().__init__(master, **kwargs)
        self._top = master.winfo_toplevel()

    def config(self, **kwargs):
        cursor = kwargs.pop("cursor", None)
        if cursor is not None:
            try:
                self._top.configure(cursor=cursor)
            except tk.TclError:
                pass
        if kwargs:
            try:
                super().config(**kwargs)
            except tk.TclError:
                pass

    configure = config

    def winfo_toplevel(self):
        return self._top


def _make_noop(name):
    def _noop(self, *args, **kwargs):
        return None
    _noop.__name__ = name
    return _noop


for _name in _WINDOW_ONLY:
    if not hasattr(PanelHost, _name):
        setattr(PanelHost, _name, _make_noop(_name))
