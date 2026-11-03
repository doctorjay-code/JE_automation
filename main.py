# -*- coding: utf-8 -*-
"""
JE automation
------------------------------------------------------------
창 하나에 탭으로 세 기능을 담습니다.

    [운용지시대장] [자금판] [합계검증]

- 탭을 누르면 아래 화면만 바뀝니다. 새 창은 열리지 않습니다.
- 탭을 좌우로 끌면 순서가 바뀌고, 그 순서가 저장됩니다.
- 오른쪽 위 X 하나로 프로그램 전체가 꺼집니다.
- 설정은 exe 옆의 JE_automation_settings.json 한 파일에 기능별로 나눠 저장됩니다.

새 기능을 추가하려면
  1) features/새기능/panel.py 를 만들고  build(parent)  함수를 넣습니다.
  2) 아래 FEATURES 목록에 (탭 이름, 모듈) 한 줄만 추가합니다.
  나머지(디자인·설정·종료 처리)는 자동으로 따라옵니다.
"""

import datetime
import os
import sys
import threading
import traceback

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR not in sys.path:
    sys.path.insert(0, BASE_DIR)

import tkinter as tk
from tkinter import ttk, messagebox

from common import config as appcfg
from common import theme
from common import updater as appupdater
from common.host import PanelHost

from features.opledger import panel as op_panel
from features.fundboard import panel as fund_panel
from features.sumcheck import panel as sum_panel
from features.bgcrepo import panel as bgc_panel

APP_TITLE = "JE automation"
APP_VERSION = "1.1.8"

DEFAULT_GEOMETRY = "1280x820"
MIN_WIDTH = 1100
MIN_HEIGHT = 680

LOG_FILENAME = "JE_automation_error.log"

# 탭 순서. 새 기능은 여기에 한 줄 추가하면 됩니다.
FEATURES = [
    ("운용지시대장", op_panel),
    ("자금판", fund_panel),
    ("합계검증", sum_panel),
    ("BGC 레포", bgc_panel),
]

DRAG_THRESHOLD = 8   # 이만큼 움직여야 '끌기'로 인정 (실수 클릭 방지)


def log_error(text):
    path = os.path.join(appcfg.app_dir(), LOG_FILENAME)
    try:
        with open(path, "a", encoding="utf-8") as f:
            f.write("\n[{}]\n{}\n".format(
                datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), text))
    except Exception:
        pass
    return path


class MainWindow:
    def __init__(self, root):
        self.root = root
        self.tabs = []        # [{"label":..., "module":..., "holder":...}]
        self.built = {}       # holder -> 기능 객체
        self._drag = None

        root.title("{}  v{}".format(APP_TITLE, APP_VERSION))
        root.minsize(MIN_WIDTH, MIN_HEIGHT)
        self._restore_geometry()
        theme.apply(root)

        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill="both", expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        for label, module in self._ordered_features():
            holder = tk.Frame(self.notebook, bg=theme.BG, highlightthickness=0, bd=0)
            self.notebook.add(holder, text=self._tab_text(label))
            self.tabs.append({"label": label, "module": module, "holder": holder})

        self._bind_tab_drag()
        self._make_log_button()

        root.protocol("WM_DELETE_WINDOW", self.on_close)
        root.bind("<Control-Tab>", lambda e: self._cycle(1))
        root.bind("<Control-Shift-Tab>", lambda e: self._cycle(-1))
        root.bind("<Control-Key-ISO_Left_Tab>", lambda e: self._cycle(-1))

        self._restore_tab()
        self.root.after(60, lambda: self._ensure(self.notebook.index("current")))

    @staticmethod
    def _tab_text(label):
        return "  {}  ".format(label)

    # ---------------- 변경 이력 ----------------
    def _make_log_button(self):
        """
        탭 줄 오른쪽 끝에 [변경 이력] 버튼을 얹는다.
        (Notebook 안에 넣을 수 없어서 위에 겹쳐 놓는다 — 레이아웃은 그대로)
        """
        self.btn_log = ttk.Button(self.root, text="변경 이력", width=11,
                                  command=self.show_changelog)
        self.btn_log.place(relx=1.0, x=-10, y=5, anchor="ne")

    def show_changelog(self):
        win = tk.Toplevel(self.root)
        win.title("변경 이력")
        win.geometry("470x440")
        win.configure(bg=theme.BG)
        win.transient(self.root)

        tk.Label(win, text="변경 이력", font=theme.FONT_TITLE, bg=theme.BG,
                 fg=theme.INK).pack(anchor="w", padx=16, pady=(14, 0))
        note = tk.Label(win, text="", font=theme.FONT_SMALL, bg=theme.BG,
                        fg=theme.MUTED)
        note.pack(anchor="w", padx=16, pady=(2, 6))

        # 닫기 버튼을 먼저 아래쪽에 고정한다. (나중에 pack 하면 밀려날 수 있음)
        btn_row = tk.Frame(win, bg=theme.BG)
        btn_row.pack(side="bottom", fill="x", padx=16, pady=(0, 14))
        ttk.Button(btn_row, text="닫기", width=10,
                   command=win.destroy).pack(side="right")

        box = tk.Frame(win, bg=theme.BG)
        box.pack(fill="both", expand=True, padx=16, pady=(0, 8))

        txt = tk.Text(box, wrap="word", font=theme.FONT_UI, bg="#ffffff",
                      fg=theme.INK, bd=0, highlightthickness=1,
                      highlightbackground=theme.LINE, padx=12, pady=10,
                      spacing1=1, spacing3=3)
        sb = ttk.Scrollbar(box, orient="vertical", command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        txt.tag_configure("ver", font=theme.FONT_UI_BOLD, foreground=theme.ACCENT)

        def render(data, state):
            try:
                txt.config(state="normal")
                txt.delete("1.0", "end")
                if state == "loading":
                    txt.insert("end", "불러오는 중...")
                    note.config(text="")
                elif not data:
                    txt.insert("end", "변경 이력을 불러오지 못했습니다.\n"
                                      "인터넷 연결을 확인한 뒤 다시 열어 주세요.")
                    note.config(text="")
                else:
                    for ver in sorted(data.keys(), key=appupdater.parse_version,
                                      reverse=True):
                        txt.insert("end", "v{}\n".format(ver), ("ver",))
                        for item in data[ver]:
                            txt.insert("end", "   · {}\n".format(item))
                        txt.insert("end", "\n")
                    note.config(text="")
                txt.config(state="disabled")
            except Exception:
                pass

        render(None, "loading")

        got = {}

        def _fetch():
            try:
                got["data"] = appupdater.fetch_changelog()
            except Exception:
                got["data"] = None

        def _poll():
            # tkinter 는 다른 스레드에서 직접 건드리면 안 되므로, 메인 스레드에서 확인한다
            if "data" in got:
                render(got["data"], "ok" if got["data"] else "fail")
                return
            try:
                if win.winfo_exists():
                    win.after(250, _poll)
            except Exception:
                pass

        threading.Thread(target=_fetch, daemon=True).start()
        win.after(250, _poll)

    # ---------------- 창 상태 ----------------
    def _common_section(self):
        data = appcfg.load_section(appcfg.SECTION_COMMON, {})
        return data if isinstance(data, dict) else {}

    def _restore_geometry(self):
        geo = self._common_section().get("window_geometry")
        if isinstance(geo, str) and "x" in geo:
            try:
                self.root.geometry(geo)
                return
            except tk.TclError:
                pass
        self._center_default()

    def _center_default(self):
        try:
            w, h = (int(v) for v in DEFAULT_GEOMETRY.split("x"))
            sw = self.root.winfo_screenwidth()
            sh = self.root.winfo_screenheight()
            x = max(0, (sw - w) // 2)
            y = max(0, (sh - h) // 3)
            self.root.geometry("{}x{}+{}+{}".format(w, h, x, y))
        except Exception:
            self.root.geometry(DEFAULT_GEOMETRY)

    def _ordered_features(self):
        """저장된 탭 순서를 반영해서 기능 목록을 돌려준다."""
        saved = self._common_section().get("tab_order")
        if not isinstance(saved, list):
            return list(FEATURES)
        order = [name for name in saved if isinstance(name, str)]
        ranked = []
        for pos, (label, module) in enumerate(FEATURES):
            rank = order.index(label) if label in order else len(order) + pos
            ranked.append((rank, label, module))
        ranked.sort(key=lambda item: item[0])
        return [(label, module) for _rank, label, module in ranked]

    def _restore_tab(self):
        last = self._common_section().get("last_tab")
        index = None
        if isinstance(last, str):
            for i, tab in enumerate(self.tabs):
                if tab["label"] == last:
                    index = i
                    break
        elif isinstance(last, int):
            index = last
        if index is not None and 0 <= index < len(self.tabs):
            try:
                self.notebook.select(index)
            except tk.TclError:
                pass

    def _cycle(self, step):
        try:
            count = len(self.tabs)
            current = self.notebook.index("current")
            self.notebook.select((current + step) % count)
        except Exception:
            pass
        return "break"

    def on_close(self):
        try:
            common = self._common_section()
            common["window_geometry"] = self.root.geometry()
            common["tab_order"] = [tab["label"] for tab in self.tabs]
            try:
                common["last_tab"] = self.tabs[self.notebook.index("current")]["label"]
            except Exception:
                pass
            appcfg.save_section(appcfg.SECTION_COMMON, common)
        except Exception:
            pass
        try:
            self.root.destroy()
        except Exception:
            pass

    # ---------------- 탭 끌어서 순서 바꾸기 ----------------
    def _bind_tab_drag(self):
        self.notebook.bind("<Button-1>", self._on_tab_press, add="+")
        self.notebook.bind("<B1-Motion>", self._on_tab_drag, add="+")
        self.notebook.bind("<ButtonRelease-1>", self._on_tab_release, add="+")

    def _tab_index_at(self, x, y):
        try:
            return self.notebook.index("@%d,%d" % (x, y))
        except tk.TclError:
            return None

    def _on_tab_press(self, event):
        self._drag = {"index": self._tab_index_at(event.x, event.y),
                      "x": event.x, "moved": False}

    def _on_tab_drag(self, event):
        d = self._drag
        if not d or d["index"] is None:
            return
        if not d["moved"] and abs(event.x - d["x"]) < DRAG_THRESHOLD:
            return
        d["moved"] = True
        target = self._tab_index_at(event.x, event.y)
        if target is None or target == d["index"]:
            return
        self._move_tab(d["index"], target)
        d["index"] = target

    def _on_tab_release(self, event):
        d = self._drag
        self._drag = None
        if not d or not d["moved"]:
            return
        # 바뀐 순서를 저장 (다음에 켜도 유지)
        try:
            common = self._common_section()
            common["tab_order"] = [tab["label"] for tab in self.tabs]
            appcfg.save_section(appcfg.SECTION_COMMON, common)
        except Exception:
            pass

    def _move_tab(self, src, dst):
        tab = self.tabs.pop(src)
        self.tabs.insert(dst, tab)
        try:
            # 이미 만들어진 화면을 그대로 옮긴다 (내용·입력값 유지)
            self.notebook.insert(dst, tab["holder"], text=self._tab_text(tab["label"]))
            self.notebook.select(tab["holder"])
        except tk.TclError:
            pass

    # ---------------- 탭 만들기 ----------------
    def _on_tab_changed(self, event=None):
        try:
            index = self.notebook.index("current")
        except Exception:
            return
        self._ensure(index)

    def _ensure(self, index):
        """아직 만들지 않은 탭이면 이때 만든다. (시작이 빨라진다)"""
        if not (0 <= index < len(self.tabs)):
            return
        tab = self.tabs[index]
        holder = tab["holder"]
        if holder in self.built:
            return
        host = PanelHost(holder)
        host.pack(fill="both", expand=True)
        try:
            self.built[holder] = tab["module"].build(host)
        except Exception:
            detail = traceback.format_exc()
            path = log_error("[{}] 화면 준비 중 오류\n{}".format(tab["label"], detail))
            host.destroy()
            self._show_error(holder, tab["label"], detail, path)
        finally:
            # 기능이 제각각 바꾼 디자인을 공통 디자인으로 되돌린다.
            theme.apply(self.root)

    def _show_error(self, holder, label, detail, log_path):
        box = tk.Frame(holder, bg=theme.BG, padx=24, pady=24)
        box.pack(fill="both", expand=True)
        tk.Label(box, text="'{}' 화면을 준비하지 못했습니다.".format(label),
                 font=theme.FONT_UI_BOLD, bg=theme.BG, fg="#b91c1c", anchor="w").pack(fill="x")
        tk.Label(box,
                 text="다른 탭은 그대로 사용할 수 있습니다.\n"
                      "아래 내용을 담당자에게 보여주시면 원인을 찾을 수 있습니다.",
                 font=theme.FONT_UI, bg=theme.BG, fg=theme.MUTED,
                 anchor="w", justify="left").pack(fill="x", pady=(8, 12))
        text = tk.Text(box, height=14, font=("맑은 고딕", 9), bg="#f8fafc",
                       fg="#334155", relief="solid", bd=1, wrap="word")
        text.pack(fill="both", expand=True)
        text.insert("1.0", detail)
        text.configure(state="disabled")
        tk.Label(box, text="오류 기록: {}".format(log_path),
                 font=theme.FONT_SMALL, bg=theme.BG, fg=theme.MUTED,
                 anchor="w").pack(fill="x", pady=(8, 0))


def _install_error_hook(root):
    def report(exc_type, exc, tb):
        detail = "".join(traceback.format_exception(exc_type, exc, tb))
        path = log_error(detail)
        try:
            messagebox.showerror(
                APP_TITLE,
                "예상치 못한 오류가 발생했습니다.\n\n{}\n\n기록 위치: {}".format(exc, path),
            )
        except Exception:
            pass

    root.report_callback_exception = report
    sys.excepthook = lambda t, e, tb: report(t, e, tb)


def _check_for_update_async(root):
    """
    백그라운드에서 GitHub의 최신 버전을 확인한다.
    새 버전이 있으면 메인 스레드에서 팝업을 띄운다.
    인터넷이 안 되거나 GitHub에 닿지 않으면 조용히 넘어간다.
    """
    def _worker():
        try:
            release = appupdater.check_latest(APP_VERSION)
        except Exception:
            return
        if release is None:
            return

        # 메인 스레드에서 팝업
        def _ask():
            try:
                tag = release.get("tag", "?")
                notes = release.get("notes") or []
                msg = "새 버전({})이 나왔습니다.".format(tag)
                if notes:
                    msg += "\n\n" + "\n".join("  · " + n for n in notes[:6])
                    if len(notes) > 6:
                        msg += "\n  · 외 {}건".format(len(notes) - 6)
                msg += "\n\n지금 업데이트할까요?"
                answer = messagebox.askyesno(APP_TITLE, msg)
                if not answer:
                    return
                _run_update_gui(root, release)
            except Exception:
                pass

        try:
            root.after(0, _ask)
        except Exception:
            pass

    t = threading.Thread(target=_worker, daemon=True)
    t.start()


def _run_update_gui(root, release):
    """
    업데이트 진행 창을 띄우고 백그라운드에서 다운로드 → 교체를 한다.
    실제 교체는 별도 worker 프로그램이 하므로, 여기서는 프로그램만 닫으면 된다.
    """
    win = tk.Toplevel(root)
    win.title("업데이트 중")
    win.geometry("380x130")
    win.resizable(False, False)
    win.transient(root)
    win.grab_set()

    try:
        rx = root.winfo_rootx() + root.winfo_width() // 2 - 190
        ry = root.winfo_rooty() + root.winfo_height() // 2 - 65
        win.geometry("+{}+{}".format(rx, ry))
    except Exception:
        pass

    tk.Label(win, text="새 버전을 내려받고 있습니다...", font=theme.FONT_UI_BOLD,
             bg=theme.BG, fg=theme.INK).pack(pady=(20, 10))

    pb = ttk.Progressbar(win, mode="determinate", length=320)
    pb.pack(pady=4, padx=30)
    pb["value"] = 0

    status_label = tk.Label(win, text="", font=theme.FONT_SMALL,
                            bg=theme.BG, fg=theme.MUTED)
    status_label.pack(pady=(2, 10))

    def _progress(pct):
        try:
            pb["value"] = pct
            status_label.config(text="내려받는 중... {}%".format(pct))
            win.update_idletasks()
        except Exception:
            pass

    def _do_update():
        try:
            ok, msg = appupdater.start_update(release, progress_cb=_progress)
        except Exception as e:
            ok, msg = False, str(e)

        def _finish():
            try:
                win.grab_release()
                win.destroy()
            except Exception:
                pass
            if ok:
                # worker 가 파일을 교체한 뒤 새 버전을 띄운다.
                messagebox.showinfo(APP_TITLE, "업데이트를 준비했습니다.\n잠시 뒤 새 버전이 실행됩니다.")
                try:
                    root.destroy()
                except Exception:
                    pass
            else:
                messagebox.showerror(APP_TITLE, "업데이트하지 못했습니다.\n\n{}".format(msg))

        try:
            root.after(0, _finish)
        except Exception:
            _finish()

    threading.Thread(target=_do_update, daemon=True).start()


def main():
    # worker 로 실행된 경우: 파일 교체만 하고 끝낸다.
    if appupdater.is_worker_invocation():
        try:
            for flag in ("--apply-exe", "--apply-update"):
                if flag in sys.argv:
                    return appupdater.run_worker(sys.argv[sys.argv.index(flag):])
            return 1
        except Exception:
            return 1

    root = tk.Tk()
    _install_error_hook(root)
    appcfg.migrate_legacy()
    MainWindow(root)

    # 프로그램이 뜬 직후 백그라운드에서 버전 확인 (UI를 막지 않음)
    # 지난 업데이트가 남긴 임시 폴더 정리 (화면 뜬 뒤, 조용히)
    threading.Thread(target=appupdater.clean_temp, daemon=True).start()

    root.after(1500, lambda: _check_for_update_async(root))

    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
