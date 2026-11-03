# -*- coding: utf-8 -*-
"""
운용지시대장 자동반영도구
------------------------------------------------------------
날인본 폴더에 쌓인 PDF 파일 이름을 읽어서,
마스터 엑셀(운용지시대장)의 펀드별 시트 C열(내용)에 자동으로 채워 넣는 프로그램.

- B열(운용지시일자), D열(기안문서번호)은 건드리지 않습니다. (엑셀 수식으로 처리)
"""

import os
import re
import sys
import json
import unicodedata
import datetime
import difflib
from copy import copy

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import openpyxl

from common.config import load_section, save_section
from common.widgets import StatusBar, safe_start_dir

APP_TITLE = "운용지시대장 자동반영도구"
APP_VERSION = "1.0.4"
CONFIG_FILENAME = "운용지시대장_자동반영도구_설정.json"

FONT_UI = ("맑은 고딕", 10)
FONT_UI_BOLD = ("맑은 고딕", 10, "bold")
FONT_TITLE = ("맑은 고딕", 16, "bold")

# 경로 영역 배치 규칙 (모든 탭 공통)
#   라벨(12자) - 입력칸(끝까지) - 버튼(10자 고정)
#   스캔은 맨 아래 줄, 입력칸 자리는 비워 둔다
LBL_W = 12
BTN_W = 12

# 데이터 시작 행 (5행은 머리글)
DATA_START_ROW = 6
CONTENT_COL = 3  # C열


# ------------------------------------------------------------
# 기본 유틸
# ------------------------------------------------------------
def app_dir():
    """EXE로 실행하면 EXE가 있는 폴더, 소스로 실행하면 이 파일이 있는 폴더."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


CONFIG_PATH = os.path.join(app_dir(), CONFIG_FILENAME)


def nfc(text):
    return unicodedata.normalize("NFC", str(text or ""))


def read_json(path, fallback):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return fallback


def write_json(path, payload):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


# ------------------------------------------------------------
# 설정
# ------------------------------------------------------------
def default_config():
    return {
        "version": APP_VERSION,
        "master_excel": "",
        "source_folder": "",
        "folder_sheet_map": {},   # 폴더명 -> 시트명 (사용자가 직접 고친 것)
        "period_from": "",        # 반영기간 시작 (YYYY-MM-DD)
        "period_to": "",          # 반영기간 끝
    }


SECTION_NAME = "운용지시대장"


def load_config():
    """통합 설정 파일의 '운용지시대장' 구역만 읽는다."""
    cfg = load_section(SECTION_NAME, None)
    if not isinstance(cfg, dict):
        cfg = default_config()
    base = default_config()
    base.update({k: v for k, v in cfg.items() if k in base})
    if not isinstance(base.get("folder_sheet_map"), dict):
        base["folder_sheet_map"] = {}
    return base


def save_config(cfg):
    if not save_section(SECTION_NAME, cfg):
        print("설정 저장 실패:", cfg and "저장 경로를 확인하세요")


# ------------------------------------------------------------
# 파일명 해석
# ------------------------------------------------------------
# 예: (058)_샘플자산1호_집합투자재산 평가 운용지시_261001.pdf
# 파일명이 조금씩 달라도(전각괄호, 뒤에 글자 붙음, 8자리 날짜 등) 최대한 읽어낸다.
SEQ_RE = re.compile(r"^[\(\[]\s*(\d+)\s*[\)\]]")
DATE8_RE = re.compile(r"(?<!\d)(\d{8})(?!\d)")
DATE6_RE = re.compile(r"(?<!\d)(\d{6})(?!\d)")

FULLWIDTH_MAP = {
    "（": "(", "）": ")", "［": "[", "］": "]",
    "０": "0", "１": "1", "２": "2", "３": "3", "４": "4",
    "５": "5", "６": "6", "７": "7", "８": "8", "９": "9",
}


def clean_filename_text(s):
    """파일명에서 눈에 안 보이는 문자를 없애고 전각문자를 반각으로 바꾼다."""
    s = nfc(s)
    for k, v in FULLWIDTH_MAP.items():
        s = s.replace(k, v)
    s = s.replace("\u00a0", " ").replace("\u200b", "").replace("\ufeff", "")
    s = re.sub(r"\s+", " ", s)
    return s.strip()


def _make_date(digits, year_len):
    try:
        if year_len == 4:
            return datetime.date(int(digits[:4]), int(digits[4:6]), int(digits[6:8]))
        return datetime.date(2000 + int(digits[:2]), int(digits[2:4]), int(digits[4:6]))
    except ValueError:
        return None


def find_date_in_name(s):
    """파일명에서 가장 마지막에 나오는 '진짜 날짜'를 찾는다."""
    for regex, ylen in ((DATE8_RE, 4), (DATE6_RE, 2)):
        for m in reversed(list(regex.finditer(s))):
            d = _make_date(m.group(1), ylen)
            if d and 2000 <= d.year <= 2049:
                return d
    return None


def parse_pdf_stem(stem):
    """PDF 파일명(확장자 제외)에서 번호 / 날짜를 뽑는다."""
    s = clean_filename_text(stem)

    seq = ""
    m = SEQ_RE.match(s)
    if m:
        seq = m.group(1)

    d = find_date_in_name(s)
    return {"seq": seq, "date": d, "ok": d is not None}


def parse_date_text(text):
    """사용자가 입력한 날짜 문자열 -> date (또는 None).

    '2026-07-01', '20260701', '2026.7.1', '26-07-01' 모두 허용.
    """
    t = clean_filename_text(text)
    if not t:
        return None
    digits = re.sub(r"[^0-9]", "", t)
    if len(digits) == 8:
        return _make_date(digits, 4)
    if len(digits) == 6:
        return _make_date(digits, 2)
    return None


# ------------------------------------------------------------
# 폴더 스캔
# ------------------------------------------------------------
FUND_DIR_RE = re.compile(r"^[A-Za-z]?\d+\.\s*\S")
NUMBER_RE = re.compile(r"(\d+)호")
# 펀드 폴더가 아닌 일반 폴더 이름들
SKIP_DIR_WORDS = ("운용지시", "날인본", "기타", "기준가", "결산")


def find_nalinbun_dirs(root):
    """root 아래에서 이름에 '날인본'이 들어간 폴더를 모두 찾는다."""
    found = []
    for cur, dirs, _files in os.walk(root):
        for d in list(dirs):
            if "날인본" in nfc(d):
                found.append(os.path.join(cur, d))
    # root 자체가 날인본인 경우
    if "날인본" in nfc(os.path.basename(root)):
        found.insert(0, root)
    return found


def is_fund_dir(name):
    if not FUND_DIR_RE.match(name):
        return False
    return not any(w in name for w in SKIP_DIR_WORDS)


def fund_dir_of(nalinbun_path, root):
    """날인본 폴더에서 위로 올라가며 '01. 샘플자산...' 같은 펀드 폴더를 찾는다."""
    cur = os.path.dirname(nalinbun_path)
    root_abs = os.path.abspath(root)
    while True:
        name = nfc(os.path.basename(cur))
        if is_fund_dir(name):
            return name
        parent = os.path.dirname(cur)
        if os.path.abspath(cur) == root_abs or parent == cur:
            break
        cur = parent
    return nfc(os.path.basename(os.path.dirname(nalinbun_path)))


def collect_pdfs(root):
    """날인본 폴더들의 PDF를 모아 목록으로 만든다."""
    results = []
    for nb in find_nalinbun_dirs(root):
        fund_dir = fund_dir_of(nb, root)
        for cur, _dirs, files in os.walk(nb):
            for fn in sorted(files):
                if not fn.lower().endswith(".pdf"):
                    continue
                if fn.startswith("~$"):
                    continue
                stem = os.path.splitext(fn)[0]
                info = parse_pdf_stem(stem)
                results.append({
                    "file": nfc(fn),
                    "stem": nfc(stem),
                    "path": os.path.join(cur, fn),
                    "fund_dir": fund_dir,
                    "sheet": "",
                    "score": 0.0,
                    "date": info["date"],
                    "seq": info["seq"],
                    "name_ok": info["ok"],
                    "status": "",
                    "selected": True,
                })
    results.sort(key=lambda r: (r["fund_dir"], r["date"] or datetime.date.min, r["file"]))
    return results


# ------------------------------------------------------------
# 엑셀 시트 매칭
# ------------------------------------------------------------
def norm_key(s):
    """비교용으로 문자열을 최대한 단순하게 만든다."""
    s = nfc(s)
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"\[.*?\]", "", s)
    s = re.sub(r"^[A-Za-z]?\d+\.\s*", "", s)      # 앞의 '01. ' 제거
    s = re.sub(r"[\s_\-\.\,·]+", "", s)
    return s.upper()


def split_name_num(s):
    """'샘플1호' -> ('샘플', '1')  (마지막 'N호' 기준)"""
    matches = list(NUMBER_RE.finditer(s))
    if not matches:
        return s, ""
    m = matches[-1]
    return s[:m.start()] + s[m.end():], m.group(1)


def _pair_score(a, cand):
    """정규화된 폴더명(a)과 후보 문자열(cand)의 유사도."""
    if not cand:
        return 0.0

    a_cat, a_num = split_name_num(a)
    cat, num = split_name_num(cand)

    # '1호' / '2호' 처럼 번호가 다르면 사실상 다른 펀드
    if a_num and num and a_num != num:
        return 0.0

    score = difflib.SequenceMatcher(None, a, cand).ratio()

    if cat:
        if a_cat.startswith(cat) or cat.startswith(a_cat):
            score = max(score, 0.90)
        elif a_cat in cat or cat in a_cat:
            score = max(score, 0.85)
    return score


def match_score(fund_dir, sheet_name, sheet_title):
    """(점수, 시트이름으로_번호까지_맞았는지) 를 돌려준다.

    시트 제목(B2)에 오타가 있어도 시트 '이름'이 맞으면 그쪽을 우선하기 위해
    두 번째 값을 함께 돌려준다.
    """
    a = norm_key(fund_dir)
    b = norm_key(sheet_name)
    c = norm_key(sheet_title)

    if not a:
        return 0.0, 0

    if a == b:
        return 1.0, 1
    if a == c:
        return 0.97, 0

    s_name = _pair_score(a, b)
    s_title = _pair_score(a, c) * 0.97   # 제목은 이름보다 약간 낮게

    _, a_num = split_name_num(a)
    _, b_num = split_name_num(b)
    name_num_ok = 1 if (a_num and a_num == b_num) else 0

    return max(s_name, s_title), name_num_ok


def read_excel_sheets(excel_path):
    """엑셀에서 펀드 시트 목록을 읽는다. [{name, title}]"""
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    sheets = []
    try:
        for ws in wb.worksheets:
            title = nfc(ws["B2"].value) if ws.max_row and ws.max_row >= 2 else ""
            header_b = nfc(ws["B5"].value)
            header_c = nfc(ws["C5"].value)
            if "운용지시일자" in header_b or "내" in header_c:
                sheets.append({"name": ws.title, "title": title})
    finally:
        wb.close()
    return sheets


def assign_sheets(items, sheets, folder_map, threshold=0.80):
    """항목들에 시트를 자동 배정한다."""
    for it in items:
        manual = folder_map.get(it["fund_dir"])
        if manual and any(s["name"] == manual for s in sheets):
            it["sheet"] = manual
            it["score"] = 1.0
            continue

        scored = []
        for s in sheets:
            score, name_num_ok = match_score(it["fund_dir"], s["name"], s["title"])
            scored.append((score, name_num_ok, s["name"]))
        scored.sort(reverse=True)

        if not scored:
            it["sheet"] = ""
            it["score"] = 0.0
            continue

        top = scored[0]
        second = scored[1] if len(scored) > 1 else (0.0, 0, "")

        clear_winner = False
        if top[0] >= threshold:
            if (top[0] - second[0]) >= 0.03:
                clear_winner = True
            elif abs(top[0] - second[0]) < 1e-9 and top[1] > second[1]:
                # 점수가 같으면 시트 '이름'이 맞는 쪽을 고른다
                clear_winner = True

        if clear_winner:
            it["sheet"] = top[2]
            it["score"] = top[0]
        else:
            it["sheet"] = ""
            it["score"] = top[0]


# ------------------------------------------------------------
# 엑셀 쓰기
# ------------------------------------------------------------
def excel_is_open(excel_path):
    folder = os.path.dirname(excel_path)
    base = os.path.basename(excel_path)
    return os.path.exists(os.path.join(folder, "~$" + base))


def existing_values(ws):
    """시트 C열에 이미 들어있는 값들(중복 확인용)."""
    vals = set()
    for r in range(DATA_START_ROW, ws.max_row + 1):
        v = ws.cell(row=r, column=CONTENT_COL).value
        if v not in (None, ""):
            vals.add(norm_key(v))
    return vals


def first_empty_row(ws):
    for r in range(DATA_START_ROW, ws.max_row + 1):
        if ws.cell(row=r, column=CONTENT_COL).value in (None, ""):
            return r
    return ws.max_row + 1


def read_existing_values(excel_path):
    """시트별로 C열에 이미 들어있는 값들을 한 번에 읽는다. (중복 확인용)

    {시트명: {정규화된 값, ...}} 형태로 돌려준다.
    """
    result = {}
    wb = openpyxl.load_workbook(excel_path, read_only=True, data_only=True)
    try:
        for ws in wb.worksheets:
            vals = set()
            for row in ws.iter_rows(min_row=DATA_START_ROW, min_col=CONTENT_COL,
                                    max_col=CONTENT_COL, values_only=True):
                v = row[0] if row else None
                if v not in (None, ""):
                    vals.add(norm_key(v))
            result[ws.title] = vals
    finally:
        wb.close()
    return result


def apply_items(excel_path, items):
    """items 를 시트별로 묶어 C열에 기입한다. 결과 요약을 돌려준다."""
    wb = openpyxl.load_workbook(excel_path)

    grouped = {}
    for it in items:
        grouped.setdefault(it["sheet"], []).append(it)

    written = 0
    skipped = 0
    details = []

    for sheet_name, rows in grouped.items():
        if sheet_name not in wb.sheetnames:
            details.append(f"[{sheet_name}] 시트가 없습니다 - 건너뜀")
            skipped += len(rows)
            continue

        ws = wb[sheet_name]
        seen = existing_values(ws)
        template = ws.cell(row=DATA_START_ROW, column=CONTENT_COL)
        next_row = first_empty_row(ws)

        rows.sort(key=lambda r: (r["date"] or datetime.date.min, r["seq"]))

        for it in rows:
            key = norm_key(it["stem"])
            if key in seen:
                skipped += 1
                continue

            if next_row > 52:
                # 표 범위를 넘으면 서식 복사해서 한 줄 늘린다
                src = ws.cell(row=52, column=CONTENT_COL)
                for col in (2, 3, 4):
                    tgt = ws.cell(row=next_row, column=col)
                    tgt._style = copy(ws.cell(row=52, column=col)._style)
                ws.row_dimensions[next_row].height = ws.row_dimensions[52].height or 16.5

            cell = ws.cell(row=next_row, column=CONTENT_COL)
            if cell.border.left.style is None:
                cell._style = copy(template._style)
            cell.value = it["stem"]
            cell.alignment = openpyxl.styles.Alignment(vertical="center", wrap_text=False)

            seen.add(key)
            next_row += 1
            written += 1

    wb.save(excel_path)
    return {"written": written, "skipped": skipped, "details": details}


# ------------------------------------------------------------
# 화면
# ------------------------------------------------------------
class App:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()
        self.items = []
        self.sheets = []
        self.existing = {}   # 시트명 -> 이미 C열에 들어있는 값 집합
        self.excel_path = self.cfg.get("master_excel", "")
        self.source_folder = self.cfg.get("source_folder", "")

        root.title(f"{APP_TITLE}  v{APP_VERSION}")
        root.geometry("1300x780")
        root.minsize(1100, 640)

        self._setup_style()
        self._build_ui()
        self._restore_state()

    # ---------- 공통 ----------
    def _setup_style(self):
        style = ttk.Style()
        try:
            style.theme_use("clam")
        except Exception:
            pass
        style.configure("TLabel", font=FONT_UI, background="#ffffff")
        style.configure("Hint.TLabel", font=("맑은 고딕", 9), foreground="#64748b", background="#ffffff")
        style.configure("TButton", font=FONT_UI, padding=(10, 6))
        style.configure("Accent.TButton", font=FONT_UI_BOLD, padding=(14, 7))
        style.configure("Treeview.Heading", font=FONT_UI_BOLD, background="#e2e8f0", foreground="#0f172a")
        style.configure("Treeview", font=("맑은 고딕", 9), rowheight=28,
                        background="#ffffff", fieldbackground="#ffffff")
        style.map("Treeview", background=[("selected", "#2563eb")],
                  foreground=[("selected", "#ffffff")])

    def _build_ui(self):
        self.root.configure(bg="#ffffff")

        # 화면 맨 아래 공통 상태 표시줄 (세 기능 동일) — 가장 먼저 자리를 잡아야 밀리지 않는다
        self.status_bar = StatusBar(self.root)
        self.status_bar.pack(side="bottom", fill="x")

        # ---- 제목 ----
        head = tk.Frame(self.root, bg="#0f172a")
        head.pack(fill="x")
        tk.Label(head, text=APP_TITLE, font=FONT_TITLE, bg="#0f172a", fg="#ffffff").pack(
            side="left", padx=18, pady=(12, 2))
        tk.Label(head, text="날인본 PDF 파일명을 읽어 운용지시대장 C열에 자동으로 기입합니다",
                 font=("맑은 고딕", 10), bg="#0f172a", fg="#cbd5e1").pack(
            side="right", padx=18, pady=(18, 0))

        # ---- 경로 영역 ----
        # 배치 규칙: 라벨(12자) - 입력칸(끝까지) - 버튼(10자 고정)
        #           스캔은 맨 아래 줄, 입력칸 자리는 비워 둔다
        pathbox = tk.Frame(self.root, bg="#ffffff", padx=18, pady=12)
        pathbox.pack(fill="x")
        pathbox.columnconfigure(1, weight=1)

        # 마스터 엑셀
        tk.Label(pathbox, text="마스터 엑셀", font=FONT_UI_BOLD, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        self.var_excel = tk.StringVar(value=self.excel_path)
        ttk.Entry(pathbox, textvariable=self.var_excel, font=FONT_UI).grid(
            row=0, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(pathbox, text="① 찾아보기", width=BTN_W, style="Path.TButton",
                   command=self.browse_excel).grid(row=0, column=2, sticky="w", pady=3)

        # 날인본 폴더
        tk.Label(pathbox, text="날인본 상위폴더", font=FONT_UI_BOLD, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
        self.var_folder = tk.StringVar(value=self.source_folder)
        ttk.Entry(pathbox, textvariable=self.var_folder, font=FONT_UI).grid(
            row=1, column=1, sticky="ew", padx=6, pady=3)
        ttk.Button(pathbox, text="② 찾아보기", width=BTN_W, style="Path.TButton",
                   command=self.browse_folder).grid(row=1, column=2, sticky="w", pady=3)

        # 반영기간
        #   [시작] ~ [끝] ......... [③ 기간 적용]  ← 기간 적용은 안쪽(입력칸 영역 오른쪽 끝)
        #                                          [③ 전체]      ← 전체는 버튼 열(찾아보기 아래)
        tk.Label(pathbox, text="반영기간", font=FONT_UI_BOLD, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=2, column=0, sticky="w", pady=3)
        self.var_from = tk.StringVar(value=self.cfg.get("period_from", ""))
        self.var_to = tk.StringVar(value=self.cfg.get("period_to", ""))
        period = tk.Frame(pathbox, bg="#ffffff")
        period.grid(row=2, column=1, sticky="ew", padx=6, pady=3)
        ttk.Entry(period, textvariable=self.var_from, font=FONT_UI, width=13).pack(side="left")
        tk.Label(period, text=" ~ ", bg="#ffffff", font=FONT_UI).pack(side="left")
        ttk.Entry(period, textvariable=self.var_to, font=FONT_UI, width=13).pack(side="left")
        ttk.Button(period, text="③ 기간 적용", width=BTN_W,
                   command=self.apply_period).pack(side="right")

        # 전체 — 버튼 열에 세로로 쌓는다 (찾아보기 / 찾아보기 / 전체 / 스캔)
        ttk.Button(pathbox, text="③ 전체", width=BTN_W,
                   command=self.clear_period).grid(row=2, column=2, sticky="w", pady=3)

        # 스캔 — 맨 아래 줄. 입력칸 자리는 비워 둔다.
        self.btn_scan = ttk.Button(pathbox, text="④ 스캔", width=BTN_W,
                                   style="PathAccent.TButton", command=self.do_scan)
        self.btn_scan.grid(row=3, column=2, sticky="w", pady=3)

        # ---- 요약 (화면 맨 아래 공통 상태 표시줄로 표시) ----
        self.var_summary = tk.StringVar(value="")

        # ---- 표 ----
        table = tk.Frame(self.root, bg="#ffffff", padx=18, pady=8)
        table.pack(fill="both", expand=True)

        cols = ("sel", "file", "folder", "sheet", "date", "status")
        self.tree = ttk.Treeview(table, columns=cols, show="headings", selectmode="browse")
        heads = {
            "sel": ("선택", 50, "center"),
            "file": ("PDF 파일명", 520, "w"),
            "folder": ("펀드 폴더", 240, "w"),
            "sheet": ("들어갈 시트", 170, "w"),
            "date": ("날짜", 95, "center"),
            "status": ("상태", 110, "center"),
        }
        for c, (text, width, anchor) in heads.items():
            self.tree.heading(c, text=text)
            self.tree.column(c, width=width, anchor=anchor, stretch=(c == "file"))

        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        self.tree.tag_configure("warn", background="#fff7ed")
        self.tree.tag_configure("done", background="#f0fdf4", foreground="#166534")
        self.tree.tag_configure("dim", background="#f1f5f9", foreground="#94a3b8")

        self.tree.bind("<Button-1>", self.on_tree_click)
        self.tree.bind("<Double-1>", self.on_tree_double_click)

        # ---- 아래 버튼 ----
        bottom = tk.Frame(self.root, bg="#ffffff", padx=18, pady=12)
        bottom.pack(fill="x")

        ttk.Button(bottom, text="전체 선택", command=self.select_all).pack(side="left")
        ttk.Button(bottom, text="전체 해제", command=self.select_none).pack(side="left", padx=6)
        ttk.Button(bottom, text="매핑 설정", command=self.open_mapping_dialog).pack(side="left", padx=(18, 0))

        self.btn_apply_all = ttk.Button(bottom, text="전체 반영", style="Accent.TButton",
                                        command=self.apply_all)
        self.btn_apply_all.pack(side="right")
        self.btn_apply_sel = ttk.Button(bottom, text="선택만 반영", command=self.apply_selected)
        self.btn_apply_sel.pack(side="right", padx=6)

        # ---- 화면 맨 아래 공통 상태 표시줄과 연결 ----
        self.var_summary.trace_add(
            "write", lambda *a: self.status_bar.set_status(self.var_summary.get()))

    def _restore_state(self):
        if self.excel_path and os.path.exists(self.excel_path):
            self.load_sheet_list(quiet=True)

    # ---------- 경로 선택 ----------
    def browse_excel(self):
        init = os.path.dirname(self.excel_path) if self.excel_path else ""
        path = filedialog.askopenfilename(
            title="마스터 엑셀 선택",
            initialdir=safe_start_dir(init),
            filetypes=[("엑셀 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
        )
        if not path:
            return
        self.excel_path = os.path.normpath(path)
        self.var_excel.set(self.excel_path)
        self.cfg["master_excel"] = self.excel_path
        save_config(self.cfg)
        self.load_sheet_list()

    def browse_folder(self):
        init = self.source_folder if self.source_folder else ""
        path = filedialog.askdirectory(title="날인본 상위폴더 선택",
                                       initialdir=safe_start_dir(init))
        if not path:
            return
        self.source_folder = os.path.normpath(path)
        self.var_folder.set(self.source_folder)
        self.cfg["source_folder"] = self.source_folder
        save_config(self.cfg)

    # ---------- 반영기간 ----------
    def current_period(self):
        """(시작일, 종료일) 반환. 비어 있으면 None. 잘못 입력하면 'error'."""
        raw_from = self.var_from.get().strip()
        raw_to = self.var_to.get().strip()
        d1 = parse_date_text(raw_from)
        d2 = parse_date_text(raw_to)
        if raw_from and d1 is None:
            return "error"
        if raw_to and d2 is None:
            return "error"
        if d1 and d2 and d1 > d2:
            return "error"
        if d1 is None and d2 is None:
            return None
        return (d1, d2)

    def in_period(self, item, period):
        if not period or period == "error":
            return True
        d = item.get("date")
        if d is None:
            return False
        d1, d2 = period
        if d1 and d < d1:
            return False
        if d2 and d > d2:
            return False
        return True

    def _reselect_all(self):
        for it in self.items:
            if it["status"] != "반영됨" and it["name_ok"]:
                it["selected"] = True

    def apply_period(self):
        period = self.current_period()
        if period == "error":
            messagebox.showwarning(
                APP_TITLE,
                "반영기간을 확인해 주세요.\n\n"
                "예:  2026-07-01   ~   2026-09-30\n"
                "(한쪽만 적어도 됩니다. 비우면 제한 없음)")
            return
        self.cfg["period_from"] = self.var_from.get().strip()
        self.cfg["period_to"] = self.var_to.get().strip()
        save_config(self.cfg)
        self._reselect_all()
        self.refresh_tree()

    def clear_period(self):
        self.var_from.set("")
        self.var_to.set("")
        self.cfg["period_from"] = ""
        self.cfg["period_to"] = ""
        save_config(self.cfg)
        self._reselect_all()
        self.refresh_tree()

    def load_sheet_list(self, quiet=False):
        if not self.excel_path or not os.path.exists(self.excel_path):
            if not quiet:
                messagebox.showwarning(APP_TITLE, "마스터 엑셀 파일을 먼저 선택해 주세요.")
            return False
        try:
            self.sheets = read_excel_sheets(self.excel_path)
        except Exception as e:
            messagebox.showerror(APP_TITLE, f"엑셀을 읽지 못했습니다.\n\n{e}")
            return False
        if not self.sheets:
            messagebox.showwarning(APP_TITLE, "엑셀에서 펀드 시트를 찾지 못했습니다.")
            return False
        return True

    def load_existing_values(self):
        """엑셀 C열에 이미 들어있는 값을 한 번만 읽어둔다."""
        try:
            self.existing = read_existing_values(self.excel_path)
        except Exception:
            self.existing = {}

    def is_already_applied(self, item):
        if not item["sheet"]:
            return False
        return norm_key(item["stem"]) in self.existing.get(item["sheet"], set())

    # ---------- 스캔 ----------
    def do_scan(self):
        if not self.load_sheet_list():
            return

        folder = self.var_folder.get().strip()
        if not folder or not os.path.isdir(folder):
            messagebox.showwarning(APP_TITLE, "날인본 상위폴더를 선택해 주세요.")
            return
        self.source_folder = os.path.normpath(folder)
        self.cfg["source_folder"] = self.source_folder
        save_config(self.cfg)

        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            items = collect_pdfs(self.source_folder)
            self.load_existing_values()
        finally:
            self.root.config(cursor="")

        if not items:
            self.items = []
            self.refresh_tree()
            messagebox.showinfo(APP_TITLE, "날인본 폴더에서 PDF를 찾지 못했습니다.\n\n"
                                           "· 폴더 이름에 '날인본'이 들어있는지 확인해 주세요.")
            return

        assign_sheets(items, self.sheets, self.cfg.get("folder_sheet_map", {}))
        self.items = items
        self.refresh_tree()

    def refresh_tree(self):
        self.tree.delete(*self.tree.get_children())

        period = self.current_period()
        if period == "error":
            period = None

        total = len(self.items)
        auto = 0
        need = 0
        out = 0
        bad = 0
        done = 0

        for idx, it in enumerate(self.items):
            if it["status"] == "반영됨":
                status = "반영됨"
                tag = "done"
                sel = "☑"
            elif not it["name_ok"]:
                status = "날짜 못 찾음"
                tag = "warn"
                sel = "☐"
                it["selected"] = False
                bad += 1
            elif self.is_already_applied(it):
                status = "이미 반영됨"
                tag = "done"
                sel = "☐"
                it["selected"] = False
                done += 1
            elif not self.in_period(it, period):
                status = "기간 외"
                tag = "dim"
                sel = "☐"
                it["selected"] = False
                out += 1
            elif not it["sheet"]:
                status = "확인요망"
                tag = "warn"
                sel = "☑" if it["selected"] else "☐"
                need += 1
            else:
                status = "준비됨"
                tag = ""
                sel = "☑" if it["selected"] else "☐"
                auto += 1

            date_txt = it["date"].strftime("%Y-%m-%d") if it["date"] else "-"
            self.tree.insert("", "end", iid=str(idx), tags=(tag,) if tag else (),
                             values=(sel, it["file"], it["fund_dir"],
                                     it["sheet"] or "-", date_txt, status))

        if total:
            parts = [f"총 {total}건"]
            if period:
                parts.append(f"기간 내 {auto + need}건")
            parts.append(f"자동 매칭 {auto}건")
            parts.append(f"확인 필요 {need}건")
            if done:
                parts.append(f"이미 반영됨 {done}건")
            if out:
                parts.append(f"기간 외 {out}건")
            if bad:
                parts.append(f"날짜 못 찾음 {bad}건")
            self.var_summary.set("   ·   ".join(parts))
        else:
            self.var_summary.set("")

    # ---------- 표 클릭 ----------
    def on_tree_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        it = self.items[int(iid)]
        if it["status"] == "반영됨" or not it["name_ok"]:
            return
        if self.is_already_applied(it):
            return
        if not self.in_period(it, self.current_period()):
            return
        it["selected"] = not it["selected"]
        self.tree.set(iid, "sel", "☑" if it["selected"] else "☐")

    def on_tree_double_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        it = self.items[int(iid)]
        if it["status"] == "반영됨":
            return
        self.open_sheet_picker(it)

    def open_sheet_picker(self, item):
        dlg = tk.Toplevel(self.root)
        dlg.title("시트 지정")
        dlg.geometry("460x520")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg="#ffffff")

        tk.Label(dlg, text="이 폴더의 PDF를 어느 시트에 넣을까요?", font=FONT_UI_BOLD,
                 bg="#ffffff").pack(anchor="w", padx=14, pady=(14, 2))
        tk.Label(dlg, text=f"폴더 : {item['fund_dir']}", font=("맑은 고딕", 9),
                 fg="#64748b", bg="#ffffff").pack(anchor="w", padx=14, pady=(0, 8))

        box = tk.Frame(dlg, bg="#ffffff")
        box.pack(fill="both", expand=True, padx=14)
        lb = tk.Listbox(box, font=FONT_UI, activestyle="none", exportselection=False)
        sb = ttk.Scrollbar(box, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        lb.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        for s in self.sheets:
            lb.insert("end", f"{s['name']}   ·   {s['title']}")
        for i, s in enumerate(self.sheets):
            if s["name"] == item["sheet"]:
                lb.selection_set(i)
                lb.see(i)
                break

        var_all = tk.BooleanVar(value=True)
        tk.Checkbutton(dlg, text="같은 폴더의 PDF 전체에 적용", variable=var_all,
                       font=FONT_UI, bg="#ffffff", anchor="w").pack(fill="x", padx=14, pady=(8, 0))

        def confirm():
            sel = lb.curselection()
            if not sel:
                messagebox.showwarning(APP_TITLE, "시트를 하나 선택해 주세요.", parent=dlg)
                return
            sheet_name = self.sheets[sel[0]]["name"]
            targets = [item]
            if var_all.get():
                targets = [x for x in self.items if x["fund_dir"] == item["fund_dir"]]
            for x in targets:
                x["sheet"] = sheet_name
                x["score"] = 1.0
                if x["name_ok"] and x["status"] != "반영됨":
                    x["selected"] = True

            # 폴더 단위로 기억
            self.cfg.setdefault("folder_sheet_map", {})[item["fund_dir"]] = sheet_name
            save_config(self.cfg)

            self.refresh_tree()
            dlg.destroy()

        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=12)
        ttk.Button(btn, text="확인", style="Accent.TButton", command=confirm).pack(side="right")
        ttk.Button(btn, text="취소", command=dlg.destroy).pack(side="right", padx=6)

    # ---------- 선택 ----------
    def select_all(self):
        for it in self.items:
            if it["name_ok"] and it["status"] != "반영됨":
                it["selected"] = True
        self.refresh_tree()

    def select_none(self):
        for it in self.items:
            it["selected"] = False
        self.refresh_tree()

    # ---------- 매핑 설정 ----------
    def open_mapping_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("폴더 ↔ 시트 매핑 설정")
        dlg.geometry("760x480")
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.configure(bg="#ffffff")

        tk.Label(dlg, text="한 번 지정한 폴더는 다음 스캔 때 자동으로 이 시트에 들어갑니다.",
                 font=FONT_UI_BOLD, bg="#ffffff").pack(anchor="w", padx=14, pady=(14, 8))

        frame = tk.Frame(dlg, bg="#ffffff")
        frame.pack(fill="both", expand=True, padx=14)
        cols = ("folder", "sheet")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        tree.heading("folder", text="펀드 폴더")
        tree.heading("sheet", text="시트")
        tree.column("folder", width=430, anchor="w")
        tree.column("sheet", width=240, anchor="w")
        sb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=sb.set)
        tree.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")

        def reload_list():
            tree.delete(*tree.get_children())
            for k, v in sorted(self.cfg.get("folder_sheet_map", {}).items()):
                tree.insert("", "end", values=(k, v))

        reload_list()

        def delete_selected():
            sel = tree.selection()
            if not sel:
                return
            folder = tree.item(sel[0], "values")[0]
            if folder in self.cfg.get("folder_sheet_map", {}):
                del self.cfg["folder_sheet_map"][folder]
                save_config(self.cfg)
                reload_list()

        def clear_all():
            if not messagebox.askyesno(APP_TITLE, "저장된 폴더 매핑을 모두 지울까요?", parent=dlg):
                return
            self.cfg["folder_sheet_map"] = {}
            save_config(self.cfg)
            reload_list()

        btn = tk.Frame(dlg, bg="#ffffff")
        btn.pack(fill="x", padx=14, pady=12)
        ttk.Button(btn, text="선택 삭제", command=delete_selected).pack(side="left")
        ttk.Button(btn, text="전체 초기화", command=clear_all).pack(side="left", padx=6)
        ttk.Button(btn, text="닫기", command=dlg.destroy).pack(side="right")

    # ---------- 반영 ----------
    def _targets(self, only_selected):
        period = self.current_period()
        targets = []
        for it in self.items:
            if it["status"] == "반영됨":
                continue
            if not it["name_ok"]:
                continue
            if not it["sheet"]:
                continue
            if self.is_already_applied(it):
                continue
            if not self.in_period(it, period):
                continue
            if only_selected and not it["selected"]:
                continue
            targets.append(it)
        return targets

    def apply_selected(self):
        self._run_apply(self._targets(True), "선택한")

    def apply_all(self):
        self._run_apply(self._targets(False), "전체")

    def _run_apply(self, targets, label):
        if not targets:
            messagebox.showinfo(APP_TITLE, "반영할 항목이 없습니다.\n\n"
                                           "· '확인요망' 항목은 더블클릭해서 시트를 먼저 지정해 주세요.")
            return

        excel = self.var_excel.get().strip()
        if not excel or not os.path.exists(excel):
            messagebox.showwarning(APP_TITLE, "마스터 엑셀 파일을 확인해 주세요.")
            return

        if excel_is_open(excel):
            messagebox.showwarning(APP_TITLE,
                                   "엑셀 파일이 열려 있습니다.\n\n"
                                   "엑셀을 모두 닫은 뒤 다시 시도해 주세요.")
            return

        # 확인
        by_sheet = {}
        for it in targets:
            by_sheet.setdefault(it["sheet"], 0)
            by_sheet[it["sheet"]] += 1
        lines = [f"{s} : {n}건" for s, n in sorted(by_sheet.items())]
        preview = "\n".join(lines[:12])
        if len(lines) > 12:
            preview += f"\n... 외 {len(lines) - 12}개 시트"

        period = self.current_period()
        period_txt = ""
        if period and period != "error":
            d1, d2 = period
            a = d1.strftime("%Y-%m-%d") if d1 else "처음"
            b = d2.strftime("%Y-%m-%d") if d2 else "끝"
            period_txt = f"반영기간 : {a} ~ {b}\n\n"

        if not messagebox.askyesno(APP_TITLE,
                                   f"{label} {len(targets)}건을 아래 시트 C열에 넣습니다.\n\n"
                                   f"{period_txt}"
                                   f"{preview}\n\n"
                                   f"진행할까요?"):
            return

        self.root.config(cursor="watch")
        self.root.update_idletasks()
        try:
            result = apply_items(excel, targets)
        except PermissionError:
            self.root.config(cursor="")
            messagebox.showerror(APP_TITLE, "엑셀에 저장하지 못했습니다.\n엑셀이 열려 있는지 확인해 주세요.")
            return
        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror(APP_TITLE, f"반영 중 오류가 났습니다.\n\n{e}")
            return
        finally:
            self.root.config(cursor="")

        for it in targets:
            it["status"] = "반영됨"
            if it["sheet"]:
                self.existing.setdefault(it["sheet"], set()).add(norm_key(it["stem"]))

        self.refresh_tree()

        msg = (f"반영 완료\n\n"
               f"· 새로 넣은 항목 : {result['written']}건\n"
               f"· 이미 있어서 건너뜀 : {result['skipped']}건")
        if result["details"]:
            msg += "\n\n" + "\n".join(result["details"])
        messagebox.showinfo(APP_TITLE, msg)


def build(parent):
    """통합 프로그램의 탭 안에 이 화면을 만든다."""
    return App(parent)


def main():
    root = tk.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
