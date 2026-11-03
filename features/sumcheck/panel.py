# -*- coding: utf-8 -*-
"""
은대 합계 검증
------------------------------------------------------------
상위폴더를 스캔해서 '자금판' 파일들의 지정 셀(기본 F2) 값을 모으고,
수탁사별로 합계를 계산하는 프로그램.

- 마스터 엑셀 없이 단독으로 동작합니다.
- 펀드명/펀드설명/파일위치/시트명/셀위치/수탁사는 파일명 기준으로 기억합니다(JSON).
- 체크 상태는 저장하지 않습니다. (프로그램을 끄면 초기화)
"""

import os
import re
import sys
import json
import time
import decimal

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from common.config import load_section, save_section
from common.widgets import StatusBar, safe_start_dir

APP_TITLE = "은대 합계 검증"
APP_VERSION = "1.0.0"
CONFIG_FILENAME = "은대_합계_검증_설정.json"

FONT_UI = ("맑은 고딕", 10)
FONT_UI_BOLD = ("맑은 고딕", 10, "bold")
FONT_TITLE = ("맑은 고딕", 16, "bold")
FONT_SMALL = ("맑은 고딕", 9)

# 경로 영역 배치 규칙 (모든 탭 공통)
#   라벨(12자) - 입력칸(끝까지) - 버튼(10자 고정)
#   스캔은 맨 아래 줄, 입력칸 자리는 비워 둔다
LBL_W = 12
BTN_W = 12

DEFAULT_SHEET = "은대계좌"
DEFAULT_CELL = "F2"

# 기본 수탁사는 두지 않는다. **설정(json)에 저장된 값만** 쓴다.
# → 프로그램에 이름이 박혀 있으면 사용자가 지울 수가 없다.
#   추가·수정·삭제는 모두 화면의 [수탁사 관리]에서 한다. (json 에 바로 저장됨)
DEFAULT_TRUSTEES = []

TARGET_KEYWORD = "자금판"
EXCEL_EXTS = (".xlsx", ".xlsm", ".xls")


# ------------------------------------------------------------
# 기본 유틸
# ------------------------------------------------------------
def app_dir():
    """EXE로 실행하면 EXE가 있는 폴더, 소스로 실행하면 이 파일이 있는 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), CONFIG_FILENAME)


SECTION_NAME = "합계검증"


def load_config():
    """통합 설정 파일의 '합계검증' 구역만 읽는다."""
    cfg = load_section(SECTION_NAME, {})
    return cfg if isinstance(cfg, dict) else {}


def save_config(cfg):
    if not save_section(SECTION_NAME, cfg):
        messagebox.showerror(
            APP_TITLE,
            "설정을 저장하지 못했습니다.\n"
            "프로그램이 읽기 전용 폴더에 있으면 저장이 안 됩니다.\n"
            "바탕화면이나 문서 폴더로 옮긴 뒤 다시 실행해 주세요.",
        )


CELL_RE = re.compile(r"^\s*\$?([A-Za-z]{1,3})\s*\$?(\d+)\s*$")


def normalize_cell(text):
    """'f2' / 'F 2' / '$F$2' → 'F2'. 형식이 틀리면 None."""
    m = CELL_RE.match(text or "")
    if not m:
        return None
    return f"{m.group(1).upper()}{int(m.group(2))}"


def fund_name_from_filename(filename):
    """'03_샘플자산1호_자금판.xlsx' → '샘플자산1호'"""
    base = os.path.splitext(filename)[0]
    base = re.sub(r"[_\s]*" + TARGET_KEYWORD + r"[_\s]*$", "", base)
    base = re.sub(r"^\s*\d+\s*[._\-\s]\s*", "", base)
    return base.strip() or os.path.splitext(filename)[0]


def fmt_num(v):
    if v is None or v == "":
        return ""
    try:
        f = float(v)
    except (TypeError, ValueError):
        return str(v)
    if abs(f - round(f)) < 0.005:
        return f"{int(round(f)):,}"
    return f"{f:,.2f}"


def round_half_up(v):
    """소수점 첫째자리에서 '일반 반올림' (0.5 는 올림).
    파이썬 기본 round() 는 짝수 쪽으로 가서 500.5 → 500 이 되므로 쓰지 않는다."""
    try:
        return int(decimal.Decimal(str(v)).quantize(
            decimal.Decimal("1"), rounding=decimal.ROUND_HALF_UP))
    except Exception:
        return 0


def fmt_sum(v):
    """합계 표시용: 소수점 첫째자리에서 반올림해 정수로 보여준다."""
    if v is None or v == "":
        return ""
    try:
        float(v)
    except (TypeError, ValueError):
        return str(v)
    return f"{round_half_up(v):,}"


def is_foreign(trustee):
    return "외화" in (trustee or "")


def short_path(path, keep=46):
    """칸에 다 안 들어갈 때 파일명이 보이도록 앞쪽을 줄인다."""
    if not path or len(path) <= keep:
        return path
    tail = path[-(keep - 2):]
    cut = tail.find("\\")
    if 0 <= cut < 18:
        tail = tail[cut + 1:]
    return "…\\" + tail


# ------------------------------------------------------------
# 엑셀 읽기
# ------------------------------------------------------------
def read_cell_openpyxl(path, sheet, cell):
    """엑셀을 실행하지 않고 저장된 값을 읽는다. 실패하면 (None, 사유)."""
    try:
        import openpyxl
    except ImportError:
        return None, "openpyxl 없음"
    if path.lower().endswith(".xls"):
        return None, "구형 .xls 는 열고 저장 모드로만 읽을 수 있습니다"
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as e:
        return None, f"열기 실패: {e}"
    try:
        if sheet not in wb.sheetnames:
            return None, ""
        try:
            v = wb[sheet][cell].value
        except Exception:
            return None, ""
        return v, ""
    finally:
        try:
            wb.close()
        except Exception:
            pass


def process_with_excel(tasks, on_progress=None):
    """
    엑셀을 한 번만 띄워서 각 파일을 [열기 → 값 읽기 → 저장 → 닫기] 한다.
    tasks: [{'path':..., 'sheet':..., 'cell':...}, ...]
    반환: {path: (value, note)}
    """
    try:
        import pythoncom
        import win32com.client as win32
    except ImportError:
        return {t["path"]: (None, "pywin32 미설치") for t in tasks}

    pythoncom.CoInitialize()
    xl = None
    out = {}
    try:
        xl = win32.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        xl.ScreenUpdating = False
        xl.AskToUpdateLinks = False
        try:
            xl.AutomationSecurity = 3
        except Exception:
            pass

        for i, t in enumerate(tasks):
            if on_progress:
                on_progress(i, len(tasks), t["path"])
            note = ""
            val = None
            try:
                wb = xl.Workbooks.Open(t["path"])
            except Exception as e:
                out[t["path"]] = (None, f"열기 실패: {e}")
                continue
            try:
                try:
                    val = wb.Worksheets(t["sheet"]).Range(t["cell"]).Value
                except Exception:
                    val = None  # 시트나 셀이 없으면 비워둔다
                try:
                    wb.Save()
                except Exception as e:
                    note = f"저장 실패: {e}"
            finally:
                try:
                    wb.Close(SaveChanges=False)
                except Exception:
                    pass
            out[t["path"]] = (val, note)
    finally:
        if xl is not None:
            try:
                xl.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
    return out


def find_fund_files(root):
    """상위폴더 아래에서 이름에 '자금판'이 들어간 엑셀 파일을 모두 찾는다."""
    found = []
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames if not d.startswith(".")]
        for fn in filenames:
            if fn.startswith("~$"):
                continue
            if TARGET_KEYWORD not in fn:
                continue
            if not fn.lower().endswith(EXCEL_EXTS):
                continue
            found.append(os.path.join(dirpath, fn))
    found.sort(key=lambda p: os.path.basename(p))
    return found


# ------------------------------------------------------------
# 메인 창
# ------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()

        self.trustees = list(self.cfg.get("trustees") or DEFAULT_TRUSTEES)
        self.saved_files = dict(self.cfg.get("files") or {})
        self.rows = []                 # 화면에 표시되는 펀드 목록
        self.checked = {}              # 파일명 → 체크 여부 (저장 안 함)
        self.running = False
        self.last_scan = ""

        root.title(f"{APP_TITLE}  v{APP_VERSION}")
        root.geometry("1220x780")
        root.minsize(980, 600)

        self._setup_style()
        self._build_ui()
        self._restore_state()

    # ---------------- 스타일 ----------------
    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("TLabel", font=FONT_UI, background="#ffffff")
        style.configure("TButton", font=FONT_UI, padding=(10, 6))
        style.configure("Accent.TButton", font=FONT_UI_BOLD, padding=(14, 7))
        style.configure("Treeview.Heading", font=FONT_UI_BOLD,
                        background="#e2e8f0", foreground="#0f172a")
        style.configure("Treeview", font=("맑은 고딕", 9), rowheight=28,
                        background="#ffffff", fieldbackground="#ffffff")
        style.map("Treeview", background=[("selected", "#2563eb")],
                  foreground=[("selected", "#ffffff")])

    # ---------------- 화면 ----------------
    def _build_ui(self):
        self.root.configure(bg="#ffffff")

        # 화면 맨 아래 공통 상태 표시줄 (세 기능 동일) — 가장 먼저 자리를 잡아야 밀리지 않는다
        self.status_bar = StatusBar(self.root)
        self.status_bar.pack(side="bottom", fill="x")

        head = tk.Frame(self.root, bg="#0f172a")
        head.pack(fill="x")
        tk.Label(head, text=APP_TITLE, font=FONT_TITLE, bg="#0f172a", fg="#ffffff").pack(
            side="left", padx=18, pady=(12, 2))
        tk.Label(head, text="자금판 잔액을 모아 수탁사별로 합산합니다",
                 font=("맑은 고딕", 10), bg="#0f172a", fg="#cbd5e1").pack(
            side="right", padx=18, pady=(18, 0))

        pathbox = tk.Frame(self.root, bg="#ffffff", padx=18, pady=12)
        pathbox.pack(fill="x")
        pathbox.columnconfigure(1, weight=1)

        # 행0: 상위폴더
        tk.Label(pathbox, text="상위폴더", font=FONT_UI_BOLD, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        self.var_folder = tk.StringVar(value=self._initial_folder())
        ttk.Entry(pathbox, textvariable=self.var_folder, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(pathbox, text="① 찾아보기", width=BTN_W, style="Path.TButton",
                   command=self.browse_folder).grid(row=0, column=2, sticky="w", pady=3)

        # 행1: 체크박스는 비어 있는 입력칸 자리에, 스캔은 오른쪽 버튼 열에
        self.var_do_excel = tk.BooleanVar(value=bool(self.cfg.get("do_excel", True)))
        tk.Checkbutton(
            pathbox, text="엑셀 열고 저장·닫기", variable=self.var_do_excel,
            font=FONT_UI, bg="#ffffff", activebackground="#ffffff",
            anchor="w", command=self._on_toggle_excel,
        ).grid(row=1, column=1, sticky="w", padx=6, pady=3)

        self.btn_scan = ttk.Button(pathbox, text="② 스캔", width=BTN_W,
                                   style="PathAccent.TButton", command=self.do_scan)
        self.btn_scan.grid(row=1, column=2, sticky="w", pady=3)

        self.sum_outer = tk.Frame(self.root, bg="#f8fafc", padx=18, pady=10)
        self.sum_outer.pack(fill="x")
        self.sum_frame = tk.Frame(self.sum_outer, bg="#f8fafc")
        self.sum_frame.pack(fill="x")

        table = tk.Frame(self.root, bg="#ffffff", padx=18, pady=8)
        table.pack(fill="both", expand=True)

        cols = ("sel", "no", "desc", "folder", "file", "sheet", "cell", "value", "trustee")
        self.tree = ttk.Treeview(table, columns=cols, show="headings",
                                 selectmode="browse")
        heads = {
            "sel": ("선택", 46, "center"),
            "no": ("번호", 66, "center"),
            "desc": ("펀드설명", 96, "w"),
            "folder": ("폴더위치", 296, "w"),
            "file": ("파일명", 246, "w"),
            "sheet": ("시트명", 86, "w"),
            "cell": ("셀위치", 64, "center"),
            "value": ("셀값", 124, "e"),
            "trustee": ("수탁사", 140, "w"),
        }
        for c, (text, width, anchor) in heads.items():
            self.tree.heading(c, text=text, anchor=anchor)
            self.tree.column(c, width=width, anchor=anchor, stretch=(c == "path"))

        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        self.tree.tag_configure("off", background="#f1f5f9", foreground="#94a3b8")
        self.tree.tag_configure("err", background="#fff7ed", foreground="#9a3412")

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        bottom = tk.Frame(self.root, bg="#ffffff", padx=18, pady=10)
        bottom.pack(fill="x")
        ttk.Button(bottom, text="전체 선택", command=self.select_all).pack(side="left")
        ttk.Button(bottom, text="전체 해제", command=self.select_none).pack(side="left", padx=6)
        ttk.Button(bottom, text="수탁사 관리", command=self.open_trustee_manager).pack(
            side="left", padx=(18, 0))

        self.var_status = tk.StringVar(value="")
        self.var_status.trace_add(
            "write", lambda *a: self.status_bar.set_status(self.var_status.get()))

    def _restore_state(self):
        self.refresh_summary()

    def _initial_folder(self):
        """저장된 상위폴더는 '이 컴퓨터에 없더라도' 그대로 보여준다. (운용지시대장과 같은 방식)
        예전에는 없으면 빈칸으로 만들어서, 저장된 경로가 사라진 것처럼 보였다."""
        return (self.cfg.get("last_folder") or "").strip()

    def _on_toggle_excel(self):
        self.cfg["do_excel"] = bool(self.var_do_excel.get())
        save_config(self.cfg)

    # ---------------- 폴더 선택 ----------------
    def browse_folder(self):
        init = self.var_folder.get() or os.path.expanduser("~")
        path = filedialog.askdirectory(title="자금판 파일들이 들어있는 상위폴더",
                                       initialdir=safe_start_dir(init))
        if path:
            self.var_folder.set(os.path.normpath(path))
            self.cfg["last_folder"] = self.var_folder.get()
            save_config(self.cfg)

    # ---------------- 스캔 ----------------
    def do_scan(self):
        if self.running:
            return
        root_dir = self.var_folder.get().strip()
        if not root_dir or not os.path.isdir(root_dir):
            messagebox.showwarning(APP_TITLE, "상위폴더를 먼저 골라주세요.")
            return

        files = find_fund_files(root_dir)
        if not files:
            self.rows = []
            self.refresh_tree()
            self.var_status.set("'자금판' 파일을 찾지 못했습니다.")
            return

        self.running = True
        self.btn_scan.configure(state="disabled")
        self.cfg["last_folder"] = root_dir
        save_config(self.cfg)

        try:
            self.rows = self._build_rows(files)

            if self.var_do_excel.get():
                tasks = [{"path": r["path"], "sheet": r["sheet"], "cell": r["cell"]}
                         for r in self.rows if r["sheet"] and r["cell"]]

                def on_progress(i, total, path):
                    self.var_status.set(f"엑셀 처리 중 {i + 1}/{total}  {os.path.basename(path)}")
                    self.root.update_idletasks()
                    self.root.update()

                results = process_with_excel(tasks, on_progress)
                for r in self.rows:
                    if r["path"] in results:
                        val, note = results[r["path"]]
                        r["value"] = val
                        r["err"] = note
            else:
                for i, r in enumerate(self.rows):
                    self.var_status.set(f"값 읽는 중 {i + 1}/{len(self.rows)}")
                    self.root.update_idletasks()
                    if not r["sheet"] or not r["cell"]:
                        continue
                    val, note = read_cell_openpyxl(r["path"], r["sheet"], r["cell"])
                    r["value"] = val
                    r["err"] = note
        finally:
            self.running = False
            self.btn_scan.configure(state="normal")

        # 번호는 화면 순서일 뿐이라 저장하지 않는다.
        self.refresh_tree()
        self.last_scan = time.strftime("%H:%M:%S")
        self.refresh_summary()

        # 값을 못 읽은 이유가 있으면 상태줄에 같이 알려준다.
        # (예전에는 표만 주황색으로 변하고 이유가 안 보여서 '왜 비었지' 하게 됐다)
        notes = []
        for r in self.rows:
            note = (r.get("err") or "").strip()
            if note and note not in notes:
                notes.append(note)
        if notes:
            self.var_status.set(self.var_status.get() + "   ·   값을 못 읽음: " + notes[0])

    def _build_rows(self, files):
        """스캔한 파일에 화면용 번호를 01부터 차례로 붙인다.
        번호는 저장하지 않는다. 파일명 순서를 그대로 따르는 표시일 뿐이다.
        """
        rows = []
        for i, path in enumerate(files, start=1):
            name = os.path.basename(path)
            e = self.saved_files.get(name, {})
            rows.append({
                "no": f"{i:02d}",
                "name": name,
                "path": path,
                "folder": os.path.dirname(path),
                "desc": e.get("desc") or "",
                "sheet": e.get("sheet") or DEFAULT_SHEET,
                "cell": e.get("cell") or DEFAULT_CELL,
                "value": None,
                "trustee": e.get("trustee") or "",
                "err": "",
                "checked": self.checked.get(name, True),
            })
        return rows

    # ---------------- 표 그리기 ----------------
    def refresh_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self.rows.sort(key=lambda x: x["no"])
        for idx, r in enumerate(self.rows):
            mark = "☑" if r["checked"] else "☐"
            tags = []
            if r.get("err"):
                tags.append("err")
            elif not r["checked"]:
                tags.append("off")
            self.tree.insert(
                "", "end", iid=str(idx),
                values=(mark, r["no"], r["desc"], short_path(r["folder"], 34),
                        r["name"], r["sheet"], r["cell"], fmt_num(r["value"]),
                        r["trustee"]),
                tags=tuple(tags),
            )
        self.refresh_summary()

    def refresh_summary(self):
        for w in self.sum_frame.winfo_children():
            w.destroy()

        totals = {t: 0.0 for t in self.trustees}
        counts = {t: 0 for t in self.trustees}
        has_val = {t: False for t in self.trustees}
        won_total = 0.0
        won_count = 0
        picked = 0

        for r in self.rows:
            if not r["checked"]:
                continue
            picked += 1
            t = r["trustee"]
            f = None
            try:
                f = float(r["value"])
            except (TypeError, ValueError):
                f = None
            if t in totals:
                counts[t] += 1
                if f is not None:
                    totals[t] += f
                    has_val[t] = True
            if not is_foreign(t):
                if f is not None:
                    won_total += f
                    won_count += 1

        if not self.trustees:
            tk.Label(self.sum_frame,
                     text="수탁사가 없습니다  ·  [수탁사 관리] 에서 추가하세요",
                     font=FONT_SMALL, fg="#64748b", bg="#f8fafc").pack(
                         side="left", padx=(4, 12), pady=8)

        for t in self.trustees:
            if counts[t] and has_val[t]:
                # 화면에 보이는 그대로(3자리 콤마 포함) 복사한다
                val = fmt_sum(totals[t])
            else:
                val = "-"
            self._make_chip(t, val, sub=f"{counts[t]}개", accent=False, copy_text=val)

        won_text = fmt_sum(won_total)
        self._make_chip("원화 전체 합계", won_text,
                        sub=f"{won_count}개 (외화 제외)", accent=True,
                        copy_text=won_text)

        if self.rows:
            msg = f"선택 {picked}개 / 전체 {len(self.rows)}개"
            if self.last_scan:
                msg += f"  ·  {self.last_scan} 기준"
            self.var_status.set(msg)

    def _make_chip(self, title, value, sub="", accent=False, copy_text=""):
        bg = "#e0f2fe" if not accent else "#dcfce7"
        fg = "#0c4a6e" if not accent else "#166534"
        box = tk.Frame(self.sum_frame, bg=bg, padx=12, pady=8)
        box.pack(side="left", padx=(0, 8))

        parts = [box]
        parts.append(tk.Label(box, text=title, font=("맑은 고딕", 9), bg=bg, fg=fg,
                              anchor="w"))
        parts[-1].pack(anchor="w")
        parts.append(tk.Label(box, text=value, font=("맑은 고딕", 13, "bold"), bg=bg, fg=fg,
                              anchor="w"))
        parts[-1].pack(anchor="w")
        if sub:
            parts.append(tk.Label(box, text=sub, font=("맑은 고딕", 8), bg=bg, fg=fg,
                                  anchor="w"))
            parts[-1].pack(anchor="w")

        # 네모칸(글자 포함)을 누르면 값이 클립보드로 복사된다
        for w in parts:
            w.configure(cursor="hand2")
            w.bind("<Button-1>",
                   lambda e, t=copy_text, d=value: self._copy_chip(t, d))

    def _copy_chip(self, text, display):
        if not text:
            self.status_bar.set_status("복사할 값이 없습니다.", "warn")
            return
        try:
            self.root.clipboard_clear()
            self.root.clipboard_append(text)
            self.root.update_idletasks()
        except Exception:
            self.status_bar.set_status("복사하지 못했습니다.", "warn")
            return
        self.status_bar.set_status(f"복사됨: {display}")

    # ---------------- 클릭 / 수정 ----------------
    def _row_index(self, iid):
        try:
            return int(iid)
        except (TypeError, ValueError):
            return None

    def on_tree_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        idx = self._row_index(iid)
        if idx is None:
            return
        if col == "#1":
            r = self.rows[idx]
            r["checked"] = not r["checked"]
            self.checked[r["name"]] = r["checked"]
            self.refresh_tree()

    def on_tree_double_click(self, event):
        region = self.tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        idx = self._row_index(iid)
        if idx is None:
            return
        if col == "#1":
            return
        if col == "#2":
            messagebox.showinfo(APP_TITLE, "번호는 화면에 붙는 순번입니다.\n고칠 수 없습니다.")
            return
        if col == "#8":
            messagebox.showinfo(APP_TITLE, "셀값은 직접 고칠 수 없습니다.\n[스캔]으로 다시 읽어주세요.")
            return
        if col == "#9":
            self.edit_trustee(idx)
        elif col in ("#4", "#5"):
            self.edit_path(idx)
        elif col == "#3":
            self.edit_text(idx, "desc", "펀드설명")
        elif col == "#6":
            self.edit_text(idx, "sheet", "시트명")
        elif col == "#7":
            self.edit_text(idx, "cell", "셀위치 (예: F2)")

    def edit_text(self, idx, key, label):
        r = self.rows[idx]
        dlg = tk.Toplevel(self.root)
        dlg.title(label)
        dlg.configure(bg="#ffffff")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        tk.Label(dlg, text=label, font=FONT_UI_BOLD, bg="#ffffff").pack(
            anchor="w", padx=14, pady=(14, 4))
        tk.Label(dlg, text=r["name"], font=FONT_SMALL, fg="#64748b",
                 bg="#ffffff").pack(anchor="w", padx=14)

        var = tk.StringVar(value=str(r.get(key) or ""))
        ent = ttk.Entry(dlg, textvariable=var, font=FONT_UI, width=48)
        ent.pack(padx=14, pady=10, fill="x")
        ent.focus_set()
        ent.select_range(0, "end")

        def confirm(_=None):
            val = var.get().strip()
            old_no = r["no"]
            if key == "cell":
                norm = normalize_cell(val)
                if norm is None:
                    messagebox.showwarning(APP_TITLE, "셀 위치는 F2 처럼 적어주세요.", parent=dlg)
                    return
                val = norm
            if key == "sheet" and not val:
                messagebox.showwarning(APP_TITLE, "시트명을 적어주세요.", parent=dlg)
                return
            r[key] = val
            self._persist_row(r)
            dlg.destroy()
            self.refresh_tree()

        ent.bind("<Return>", confirm)
        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=(0, 14))
        ttk.Button(btn, text="확인", style="Accent.TButton", command=confirm).pack(side="right")
        ttk.Button(btn, text="취소", command=dlg.destroy).pack(side="right", padx=6)

    def edit_path(self, idx):
        r = self.rows[idx]
        dlg = tk.Toplevel(self.root)
        dlg.title("폴더위치 / 파일명")
        dlg.configure(bg="#ffffff")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        tk.Label(dlg, text="폴더위치 / 파일명", font=FONT_UI_BOLD, bg="#ffffff").pack(
            anchor="w", padx=14, pady=(14, 4))
        tk.Label(dlg, text="매칭이 이상하면 직접 고르세요.", font=FONT_SMALL,
                 fg="#64748b", bg="#ffffff").pack(anchor="w", padx=14)

        tk.Label(dlg, text="폴더위치", font=FONT_SMALL, fg="#334155",
                 bg="#ffffff").pack(anchor="w", padx=14, pady=(10, 2))
        var_dir = tk.StringVar(value=r["folder"])
        ttk.Entry(dlg, textvariable=var_dir, font=FONT_UI, width=70).pack(
            padx=14, fill="x")

        tk.Label(dlg, text="파일명", font=FONT_SMALL, fg="#334155",
                 bg="#ffffff").pack(anchor="w", padx=14, pady=(8, 2))
        var_name = tk.StringVar(value=r["name"])
        ttk.Entry(dlg, textvariable=var_name, font=FONT_UI, width=70).pack(
            padx=14, fill="x")

        def browse():
            p = filedialog.askopenfilename(
                title="자금판 파일 선택",
                initialdir=r["folder"] or os.path.expanduser("~"),
                filetypes=[("엑셀 파일", "*.xlsx *.xlsm *.xls"), ("모든 파일", "*.*")],
            )
            if p:
                var_dir.set(os.path.dirname(os.path.normpath(p)))
                var_name.set(os.path.basename(os.path.normpath(p)))

        def confirm():
            folder = var_dir.get().strip().strip('"')
            name = var_name.get().strip()
            if not folder or not name:
                messagebox.showwarning(APP_TITLE, "폴더와 파일명을 모두 적어주세요.", parent=dlg)
                return
            p = os.path.join(os.path.normpath(folder), name)
            if not os.path.isfile(p):
                messagebox.showwarning(APP_TITLE, "그 위치에 파일이 없습니다.\n" + p, parent=dlg)
                return
            old_name = r["name"]
            r["path"] = p
            r["folder"] = os.path.dirname(p)
            r["name"] = name
            if old_name != name:
                self.saved_files[name] = self.saved_files.pop(old_name, {})
                self.checked[name] = self.checked.pop(old_name, r["checked"])
            self._persist_row(r)
            dlg.destroy()
            self.refresh_tree()

        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=(12, 14))
        ttk.Button(btn, text="파일 선택", command=browse).pack(side="left")
        ttk.Button(btn, text="확인", style="Accent.TButton", command=confirm).pack(side="right")
        ttk.Button(btn, text="취소", command=dlg.destroy).pack(side="right", padx=6)

    def edit_trustee(self, idx):
        r = self.rows[idx]

        # 수탁사를 하나도 등록하지 않았으면 먼저 관리 창으로 안내한다
        if not self.trustees:
            if messagebox.askyesno(
                    APP_TITLE,
                    "등록된 수탁사가 없습니다.\n\n"
                    "[수탁사 관리] 에서 먼저 추가할까요?"):
                self.open_trustee_manager()
            return

        dlg = tk.Toplevel(self.root)
        dlg.title("수탁사")
        dlg.configure(bg="#ffffff")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        tk.Label(dlg, text="수탁사 선택", font=FONT_UI_BOLD, bg="#ffffff").pack(
            anchor="w", padx=14, pady=(14, 6))

        box = tk.Frame(dlg, bg="#ffffff")
        box.pack(fill="both", expand=True, padx=14)
        lb = tk.Listbox(box, font=FONT_UI, activestyle="none", exportselection=False,
                        height=min(10, max(4, len(self.trustees))), width=34)
        sb = ttk.Scrollbar(box, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def reload_list(select=None):
            lb.delete(0, "end")
            lb.insert("end", "(비움)")
            for t in self.trustees:
                lb.insert("end", t)
            if select in self.trustees:
                lb.selection_set(self.trustees.index(select) + 1)
                lb.see(self.trustees.index(select) + 1)
            else:
                lb.selection_set(0)

        reload_list(r["trustee"])

        addrow = tk.Frame(dlg, bg="#ffffff")
        addrow.pack(fill="x", padx=14, pady=(10, 0))
        var_new = tk.StringVar()
        ttk.Entry(addrow, textvariable=var_new, font=FONT_UI).pack(
            side="left", fill="x", expand=True)

        def add_new():
            name = var_new.get().strip()
            if not name:
                return
            if name in self.trustees:
                messagebox.showinfo(APP_TITLE, "이미 있는 수탁사입니다.", parent=dlg)
                return
            self.trustees.append(name)
            self.cfg["trustees"] = self.trustees
            save_config(self.cfg)
            var_new.set("")
            reload_list(name)
            self.refresh_summary()

        ttk.Button(addrow, text="추가", command=add_new).pack(side="left", padx=(6, 0))

        def confirm():
            sel = lb.curselection()
            if not sel:
                return
            val = lb.get(sel[0])
            r["trustee"] = "" if val == "(비움)" else val
            self._persist_row(r)
            dlg.destroy()
            self.refresh_tree()

        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=(12, 14))
        ttk.Button(btn, text="확인", style="Accent.TButton", command=confirm).pack(side="right")
        ttk.Button(btn, text="취소", command=dlg.destroy).pack(side="right", padx=6)

        lb.bind("<Double-1>", lambda e: confirm())

    # ---------------- 수탁사 관리 ----------------
    def open_trustee_manager(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("수탁사 관리")
        dlg.configure(bg="#ffffff")
        dlg.transient(self.root)
        dlg.resizable(False, False)

        tk.Label(dlg, text="수탁사 목록", font=FONT_UI_BOLD, bg="#ffffff").pack(
            anchor="w", padx=14, pady=(14, 6))
        tk.Label(dlg,
                 text="이름을 고치려면 목록에서 고르고, 아래 칸에서 고친 뒤 [수정] 을 누르세요.",
                 font=FONT_SMALL, fg="#64748b", bg="#ffffff").pack(anchor="w", padx=14)

        box = tk.Frame(dlg, bg="#ffffff")
        box.pack(fill="both", expand=True, padx=14, pady=10)
        lb = tk.Listbox(box, font=FONT_UI, activestyle="none", exportselection=False,
                        height=10, width=34)
        sb = ttk.Scrollbar(box, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        var_name = tk.StringVar()

        # 목록이 비었을 때만 보이는 안내 (아래에서 before= 로 자리를 잡는다)
        empty_hint = tk.Label(
            dlg,
            text="아직 등록된 수탁사가 없습니다.  아래 칸에 이름을 입력하고 [추가] 를 누르세요.",
            font=FONT_SMALL, fg="#b45309", bg="#ffffff")

        def reload_list(select_idx=None):
            """목록을 다시 그린다. select_idx 를 주면 그 항목만 골라둔다."""
            lb.delete(0, "end")
            for t in self.trustees:
                lb.insert("end", t)
            lb.selection_clear(0, "end")
            if select_idx is not None and 0 <= select_idx < len(self.trustees):
                lb.selection_set(select_idx)
                lb.see(select_idx)
            if self.trustees:
                empty_hint.pack_forget()
            else:
                empty_hint.pack(fill="x", padx=14, pady=(0, 6), before=addrow)

        def on_select(_event=None):
            sel = lb.curselection()
            if sel:
                var_name.set(self.trustees[sel[0]])

        lb.bind("<<ListboxSelect>>", on_select)

        def _save():
            """목록 + 파일별 지정을 json 에 저장하고 표를 다시 그린다."""
            self.cfg["trustees"] = self.trustees
            self.cfg["files"] = self.saved_files
            save_config(self.cfg)
            self.refresh_tree()

        def add_new():
            name = var_name.get().strip()
            if not name:
                messagebox.showwarning(APP_TITLE, "수탁사 이름을 입력해 주세요.", parent=dlg)
                return
            if name in self.trustees:
                messagebox.showwarning(
                    APP_TITLE, "'{}' 은(는) 이미 목록에 있습니다.".format(name), parent=dlg)
                return
            self.trustees.append(name)
            _save()
            var_name.set("")
            reload_list(len(self.trustees) - 1)
            self.status_bar.set_status("수탁사 '{}' 을(를) 추가했습니다.".format(name))

        def edit_selected():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(
                    APP_TITLE, "수정할 수탁사를 목록에서 먼저 골라주세요.", parent=dlg)
                return
            idx = sel[0]
            old = self.trustees[idx]
            new = var_name.get().strip()
            if not new:
                messagebox.showwarning(APP_TITLE, "수탁사 이름을 입력해 주세요.", parent=dlg)
                return
            if new == old:
                return
            if new in self.trustees:
                messagebox.showwarning(
                    APP_TITLE, "'{}' 은(는) 이미 목록에 있습니다.".format(new), parent=dlg)
                return

            # 이 수탁사를 쓰던 파일들도 새 이름으로 같이 바꿔준다
            changed = 0
            for info in self.saved_files.values():
                if (info or {}).get("trustee") == old:
                    info["trustee"] = new
                    changed += 1
            for r in self.rows:
                if r.get("trustee") == old:
                    r["trustee"] = new

            self.trustees[idx] = new
            _save()
            var_name.set(new)
            reload_list(idx)
            msg = "'{}' → '{}' 로 바꿨습니다.".format(old, new)
            if changed:
                msg += " (파일 {}개도 함께 바뀜)".format(changed)
            self.status_bar.set_status(msg)

        def delete_selected():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(
                    APP_TITLE, "삭제할 수탁사를 목록에서 먼저 골라주세요.", parent=dlg)
                return
            idx = sel[0]
            old = self.trustees[idx]
            used = [n for n, info in self.saved_files.items()
                    if (info or {}).get("trustee") == old]

            msg = "'{}' 을(를) 목록에서 지울까요?".format(old)
            if used:
                msg += ("\n\n이 수탁사를 쓰는 파일이 {}개 있습니다.\n"
                        "그 파일들의 수탁사는 '비움' 이 됩니다.").format(len(used))
            if not messagebox.askyesno(APP_TITLE, msg, parent=dlg):
                return

            for n in used:
                self.saved_files[n]["trustee"] = ""
            for r in self.rows:
                if r.get("trustee") == old:
                    r["trustee"] = ""

            del self.trustees[idx]
            _save()
            var_name.set("")
            reload_list(idx if idx < len(self.trustees) else None)
            self.status_bar.set_status("수탁사 '{}' 을(를) 지웠습니다.".format(old))

        lb.bind("<Delete>", lambda e: delete_selected())

        addrow = tk.Frame(dlg, bg="#ffffff")
        addrow.pack(fill="x", padx=14)
        entry = ttk.Entry(addrow, textvariable=var_name, font=FONT_UI)
        entry.pack(side="left", fill="x", expand=True)
        entry.bind("<Return>",
                   lambda e: edit_selected() if lb.curselection() else add_new())

        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=(10, 14))
        ttk.Button(btn, text="추가", style="Accent.TButton",
                   command=add_new).pack(side="left")
        ttk.Button(btn, text="수정", command=edit_selected).pack(side="left", padx=6)
        ttk.Button(btn, text="삭제", command=delete_selected).pack(side="left")
        ttk.Button(btn, text="닫기", command=dlg.destroy).pack(side="right")

        reload_list()
        entry.focus_set()

    # ---------------- 선택 ----------------
    def select_all(self):
        for r in self.rows:
            r["checked"] = True
            self.checked[r["name"]] = True
        self.refresh_tree()

    def select_none(self):
        for r in self.rows:
            r["checked"] = False
            self.checked[r["name"]] = False
        self.refresh_tree()

    # ---------------- 저장 ----------------
    def _persist_row(self, r):
        # 파일명이 열쇠. 번호는 저장하지 않는다.
        self.saved_files[r["name"]] = {
            "desc": r["desc"],
            "sheet": r["sheet"],
            "cell": r["cell"],
            "trustee": r["trustee"],
        }
        self.cfg["files"] = self.saved_files
        self.cfg["trustees"] = self.trustees
        save_config(self.cfg)


def build(parent):
    """통합 프로그램의 탭 안에 이 화면을 만든다."""
    return App(parent)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
