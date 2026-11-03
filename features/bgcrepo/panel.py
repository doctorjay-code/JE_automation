import os
import sys
import re
import datetime
import glob
import threading
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from copy import copy

import pdfplumber
import openpyxl

from common.config import load_section, save_section
from common.widgets import StatusBar, StatusProxy, safe_start_dir

SECTION_NAME = "BGC Repo"


def load_config():
    """통합 설정 파일의 'BGC Repo' 구역만 읽는다."""
    cfg = load_section(SECTION_NAME, {})
    base = {"target_folder": "", "last_pdf_folder": ""}
    if isinstance(cfg, dict):
        base.update({k: v for k, v in cfg.items() if k in base})
    return base


def save_config(cfg):
    save_section(SECTION_NAME, cfg)


# ---------- PDF Parsing ----------
def parse_repo_pdf(pdf_path):
    """PDF에서 결제금액 표를 찾아 개시금액/계, 매수일/환매일, 펀드코드 추출"""
    with pdfplumber.open(pdf_path) as doc:
        page = doc.pages[0]
        tables = page.extract_tables() or []
        text = page.extract_text() or ""

        # dates: [작성일, 담보만기1, 만기2, 매수일, 환매일] -> 뒤 2개 사용
        dates = re.findall(r"(\d{4}-\d{2}-\d{2})", text)
        if len(dates) >= 5:
            buy_date = dates[3]
            sell_date = dates[4]
        elif len(dates) >= 2:
            buy_date = dates[-2]
            sell_date = dates[-1]
        else:
            buy_date = sell_date = None

        amt_start = amt_end = None
        # Table 1 is 결제금액: row 2 col1=개시금액, col3=계
        if len(tables) >= 2:
            t1 = tables[1]
            # find row where col1 looks like number with comma
            for row in t1:
                # skip header rows
                if row and len(row) >= 4 and row[1] and row[3]:
                    # row[1] and row[3] both numeric strings with comma
                    if re.match(r"[\d,]+", str(row[1])) and re.match(r"[\d,]+", str(row[3])):
                        # this is 거래금액 row (first numeric row after header)
                        # t1[2] is expected
                        pass
            if len(t1) >= 3:
                # t1[2] = ['거래금액', '1,100,000,000', '90,410', '1,100,090,410']
                try:
                    amt_start = int(str(t1[2][1]).replace(",", "").strip())
                    amt_end = int(str(t1[2][3]).replace(",", "").strip())
                except Exception:
                    pass

        # fund code from Table 2
        fund_code = None
        if len(tables) >= 3:
            t2 = tables[2]
            if len(t2) >= 2 and t2[1][0]:
                fund_code = str(t2[1][0]).strip()

        # fallback regex if table parse failed
        if amt_start is None:
            # find numbers in text like "1,100,000,000"
            nums = re.findall(r"(\d{1,3}(?:,\d{3})+)", text)
            if len(nums) >= 2:
                try:
                    amt_start = int(nums[0].replace(",", ""))
                    amt_end = int(nums[2].replace(",", "")) if len(nums) >= 4 else int(nums[1].replace(",", ""))
                except Exception:
                    pass

        return {
            "buy_date": buy_date,  # YYYY-MM-DD
            "sell_date": sell_date,
            "amt_start": amt_start,
            "amt_end": amt_end,
            "fund_code": fund_code,
            "pdf_name": os.path.basename(pdf_path),
        }

# ---------- Excel helpers ----------
def find_fund_sheet(wb):
    for s in wb.sheetnames:
        if "은대계좌" in s or "계좌" in s or "예금" in s:
            return s
    # fallback second sheet if exists
    if len(wb.sheetnames) > 1:
        return wb.sheetnames[1]
    return wb.sheetnames[0] if wb.sheetnames else None

def scan_fund_files(target_dir):
    """자금판 파일들을 스캔해 자산코드 suffix 맵 생성"""
    fund_map = {}  # suffix -> filename, also full code -> filename
    file_list = []
    if not os.path.exists(target_dir):
        return fund_map, file_list
    for f in os.listdir(target_dir):
        if f.endswith("_자금판.xlsx") and not f.startswith("~$"):
            file_list.append(f)
            path = os.path.join(target_dir, f)
            try:
                wb = openpyxl.load_workbook(path, data_only=True, read_only=False)
                ws = wb.worksheets[0]
                # B3: fund name, B5? 자산코드: row 5 col3
                asset_code = None
                # try common positions: B5/C5, B6 etc. Scan small area
                for r in range(1, 10):
                    for c in range(1, 5):
                        v = ws.cell(r, c).value
                        if v and isinstance(v, str) and "ZZ" in v:
                            # asset code like 00001ZZ, 00002ZZ
                            code = v.strip()
                            if len(code) >= 4:
                                fund_map[code] = f
                                # also store suffix variants
                                # store last 7 chars maybe
                                fund_map[code[-7:]] = f
                wb.close()
            except Exception as e:
                print(f"scan fail {f}: {e}")
    # also map filename without suffix
    for f in file_list:
        base = f.replace("_자금판.xlsx", "")
        fund_map[base] = f
    return fund_map, sorted(file_list)

def find_matching_excel(fund_code, fund_map):
    if not fund_code:
        return None
    fund_code = fund_code.strip()
    # direct
    if fund_code in fund_map:
        return fund_map[fund_code]
    # suffix match: last 7 chars (e.g., 00001ZZ, 00002ZZ) or last 6
    for key in [fund_code[-7:], fund_code[-6:], fund_code[-5:]]:
        if key in fund_map:
            return fund_map[key]
    # contains
    for k, v in fund_map.items():
        if k and fund_code.endswith(k):
            return v
        if k and k in fund_code:
            return v
    return None

def check_exists(filepath, date_str, amount, desc):
    """이미 같은 일자/금액/적요가 있는지 체크 (중복 방지)"""
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True, read_only=False)
        sheet = find_fund_sheet(wb)
        if not sheet:
            wb.close()
            return False
        ws = wb[sheet]
        # date_str: YYYY-MM-DD -> compare as date object or string
        try:
            y, m, d = map(int, date_str.split("-"))
            target = datetime.date(y, m, d)
        except Exception:
            target = None
        # 최근 60일치만 검사하면 충분 (전체 스캔 시 read_only=False여도 70초 걸리므로 범위 제한)
        start_r = max(4, ws.max_row - 200)
        for r in range(start_r, ws.max_row + 1):
            c_date = ws.cell(r, 2).value
            c_desc = str(ws.cell(r, 3).value or "").strip()
            c_in = ws.cell(r, 4).value
            c_out = ws.cell(r, 5).value
            # date match
            date_match = False
            if isinstance(c_date, datetime.datetime):
                date_match = (c_date.date() == target)
            elif isinstance(c_date, datetime.date):
                date_match = (c_date == target)
            elif c_date and str(c_date).strip() == date_str:
                date_match = True
            if not date_match:
                continue
            if c_desc != desc:
                continue
            # amount match: desc에 따라 입금/출금
            amt = None
            if desc == "REPO매수":
                amt = c_out
            elif desc == "REPO환매":
                amt = c_in
            if amt is not None:
                try:
                    if abs(float(amt) - float(amount)) < 0.01:
                        wb.close()
                        return True
                except Exception:
                    pass
        wb.close()
    except Exception as e:
        print(f"check_exists error {filepath}: {e}")
    return False

def append_repo_rows(filepath, buy_date, sell_date, amt_start, amt_end):
    """엑셀에 REPO매수/환매 2행 추가 (서식 복사)"""
    wb = openpyxl.load_workbook(filepath)
    sheet = find_fund_sheet(wb)
    ws = wb[sheet]

    # find last data row
    last_r = ws.max_row
    while last_r > 4 and not any(ws.cell(last_r, c).value is not None for c in range(1, 8)):
        last_r -= 1

    results = []

    # prepare 2 items
    items = [
        {"date": buy_date, "desc": "REPO매수", "amount": amt_start, "type": "출금"},
        {"date": sell_date, "desc": "REPO환매", "amount": amt_end, "type": "입금"},
    ]

    # inline duplicate check using already-opened ws (파일 잠금 회피)
    def _exists_in_ws(date_str, amount, desc):
        try:
            y, m, d = map(int, date_str.split("-"))
            target = datetime.date(y, m, d)
        except Exception:
            target = None
        for r in range(4, ws.max_row + 1):
            c_date = ws.cell(r, 2).value
            c_desc = str(ws.cell(r, 3).value or "").strip()
            if c_desc != desc:
                continue
            date_match = False
            if isinstance(c_date, datetime.datetime):
                date_match = (c_date.date() == target)
            elif isinstance(c_date, datetime.date):
                date_match = (c_date == target)
            elif c_date and str(c_date).strip() == date_str:
                date_match = True
            if not date_match:
                continue
            amt = ws.cell(r, 4).value if desc == "REPO환매" else ws.cell(r, 5).value
            if amt is not None:
                try:
                    if abs(float(amt) - float(amount)) < 0.01:
                        return True
                except Exception:
                    pass
        return False

    for item in items:
        date_str = item["date"]  # YYYY-MM-DD
        try:
            y, m, d = map(int, date_str.split("-"))
            date_obj = datetime.date(y, m, d)
        except Exception:
            date_obj = date_str

        # duplicate check (ws 기반)
        if _exists_in_ws(date_str, item["amount"], item["desc"]):
            results.append((item, "skip_duplicate"))
            continue

        last_r += 1
        new_r = last_r
        prev_r = new_r - 1

        # copy row height
        if prev_r in ws.row_dimensions and ws.row_dimensions[prev_r].height:
            ws.row_dimensions[new_r].height = ws.row_dimensions[prev_r].height
        else:
            ws.row_dimensions[new_r].height = 18.0

        ws.cell(new_r, 2).value = date_obj
        ws.cell(new_r, 3).value = item["desc"]
        if item["type"] == "입금":
            ws.cell(new_r, 4).value = item["amount"]
            ws.cell(new_r, 5).value = None
        else:
            ws.cell(new_r, 4).value = None
            ws.cell(new_r, 5).value = item["amount"]
        ws.cell(new_r, 6).value = f"=F{prev_r}+D{new_r}-E{new_r}"

        max_c = max(7, ws.max_column)
        for col_idx in range(1, max_c + 1):
            ref_cell = None
            for r in range(prev_r, 4, -1):
                c = ws.cell(r, col_idx)
                if c.value is not None:
                    ref_cell = c
                    break
            if not ref_cell:
                for r in range(prev_r, 4, -1):
                    c = ws.cell(r, col_idx)
                    if c.has_style:
                        ref_cell = c
                        break
            if not ref_cell:
                ref_cell = ws.cell(prev_r, col_idx)
            tgt = ws.cell(new_r, col_idx)
            if ref_cell and ref_cell.has_style:
                tgt.font = copy(ref_cell.font)
                tgt.fill = copy(ref_cell.fill)
                tgt.alignment = copy(ref_cell.alignment)
                if col_idx == 1:
                    b = ref_cell.border
                    if not any([b.left, b.right, b.top, b.bottom]):
                        from openpyxl.styles import Border
                        tgt.border = Border()
                    else:
                        tgt.border = copy(ref_cell.border)
                else:
                    tgt.border = copy(ref_cell.border)
                if col_idx in (4, 5) and ref_cell.number_format == 'General':
                    bal_ref = None
                    for r in range(prev_r, 4, -1):
                        c = ws.cell(r, 6)
                        if c.value is not None and c.number_format != 'General':
                            bal_ref = c
                            break
                    if bal_ref:
                        tgt.number_format = bal_ref.number_format
                    else:
                        tgt.number_format = ref_cell.number_format
                else:
                    tgt.number_format = ref_cell.number_format

        results.append((item, "added"))

    wb.save(filepath)
    wb.close()
    return results

# ---------- Main GUI Application ----------
class RepoApp:
    def __init__(self, root):
        self.root = root
        self.cfg = load_config()

        # 저장된 경로는 이 컴퓨터에 없더라도 그대로 둔다 (다른 탭과 같은 규칙)
        self.target_dir = (self.cfg.get("target_folder") or "").strip()
        self.pdf_dir = (self.cfg.get("last_pdf_folder") or "").strip()

        self.fund_map = {}
        self.file_list = []
        self.pdf_infos = []

        self.create_widgets()
        self.status_bar.set_status("스캔을 눌러주세요.")

    def create_widgets(self):
        # 화면 맨 아래 공통 상태 표시줄 (세 기능 동일) — 가장 먼저 자리를 잡아야 밀리지 않는다
        self.status_bar = StatusBar(self.root)
        self.status_bar.pack(side="bottom", fill="x")
        self.lbl_status = StatusProxy(self.status_bar, "status")

        # 머리글 (다른 탭과 같은 모양)
        head = tk.Frame(self.root, bg="#0f172a")
        head.pack(fill="x")
        tk.Label(head, text="BGC 레포 자동반영도구", font=("맑은 고딕", 16, "bold"),
                 bg="#0f172a", fg="#ffffff").pack(side="left", padx=18, pady=(12, 2))
        tk.Label(head, text="레포 거래내역서 PDF를 읽어 자금판에 매수·환매를 자동 반영합니다",
                 font=("맑은 고딕", 10), bg="#0f172a", fg="#cbd5e1").pack(
            side="right", padx=18, pady=(18, 0))

        # 폴더 영역 (모든 탭 공통 규칙: 라벨 - 입력칸 - 버튼)
        top = tk.Frame(self.root, bg="#ffffff", padx=18, pady=12)
        top.pack(fill="x")
        top.columnconfigure(1, weight=1)
        lbl_font = ("맑은 고딕", 10, "bold")
        # 경로 영역 배치 규칙: 라벨(12자) - 입력칸(끝까지) - 버튼(10자 고정)
        LBL_W = 12
        BTN_W = 12

        tk.Label(top, text="자금판 폴더", font=lbl_font, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=0, column=0, sticky="w", pady=3)
        self.entry_target = ttk.Entry(top, font=("맑은 고딕", 10))
        self.entry_target.insert(0, self.target_dir)
        self.entry_target.grid(row=0, column=1, sticky="ew", padx=6, pady=3)
        self.entry_target.bind("<KeyRelease>", self._show_scan_hint)
        ttk.Button(top, text="① 찾아보기", width=BTN_W, style="Path.TButton",
                   command=self.browse_target).grid(row=0, column=2, sticky="w", pady=3)

        tk.Label(top, text="PDF 폴더", font=lbl_font, bg="#ffffff",
                 width=LBL_W, anchor="w").grid(row=1, column=0, sticky="w", pady=3)
        self.entry_pdf = ttk.Entry(top, font=("맑은 고딕", 10))
        self.entry_pdf.insert(0, self.pdf_dir)
        self.entry_pdf.grid(row=1, column=1, sticky="ew", padx=6, pady=3)
        self.entry_pdf.bind("<KeyRelease>", self._show_scan_hint)
        ttk.Button(top, text="② 찾아보기", width=BTN_W, style="Path.TButton",
                   command=self.browse_pdf).grid(row=1, column=2, sticky="w", pady=3)

        # 스캔 — 맨 아래 줄. 입력칸 자리는 비워 둔다.
        self.btn_scan = ttk.Button(top, text="③ 스캔", width=BTN_W,
                                   style="PathAccent.TButton", command=self.scan_pdfs)
        self.btn_scan.grid(row=2, column=2, sticky="w", pady=3)

        # 결과 표
        table = ttk.Frame(self.root, padding=(18, 6))
        table.pack(fill="both", expand=True)

        cols = ("pdf", "fund_code", "buy", "sell", "amt_start", "amt_end", "matched", "status")
        self.tree = ttk.Treeview(table, columns=cols, show="headings")
        heads = {
            "pdf": ("PDF 파일", 240, "w"),
            "fund_code": ("펀드코드(PDF)", 120, "center"),
            "buy": ("매수일", 100, "center"),
            "sell": ("환매일", 100, "center"),
            "amt_start": ("개시금액(매수)", 130, "e"),
            "amt_end": ("계(환매)", 130, "e"),
            "matched": ("매칭 자금판", 220, "w"),
            "status": ("상태", 110, "center"),
        }
        for c, (text, width, anchor) in heads.items():
            self.tree.heading(c, text=text, anchor=anchor)
            self.tree.column(c, width=width, anchor=anchor)

        vsb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(table, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        table.rowconfigure(0, weight=1)
        table.columnconfigure(0, weight=1)

        # 반영 버튼
        bottom = tk.Frame(self.root, bg="#ffffff", padx=18, pady=10)
        bottom.pack(fill="x")
        self.btn_apply = ttk.Button(bottom, text="엑셀에 반영", style="Primary.TButton",
                                    command=self.apply_all)
        self.btn_apply.pack(side="right")

    def _show_scan_hint(self, event=None):
        """경로가 바뀌었을 때: 스캔을 눌러야 한다고 알려준다. (파일 접근 없음)"""
        self.status_bar.set_status("폴더가 바뀌었습니다. 스캔을 눌러주세요.", "warn")

    def log_msg(self, msg):
        """진행 상황은 화면 맨 아래 상태 표시줄에만 표시한다.
        (백그라운드 스레드에서 불리므로 after 로 메인 스레드에 넘긴다)"""
        text = str(msg)
        try:
            self.root.after(0, lambda: self.status_bar.set_status(text))
        except Exception:
            pass

    def browse_target(self):
        d = filedialog.askdirectory(
            initialdir=safe_start_dir(self.entry_target.get().strip()),
            title="자금판 폴더 선택")
        if d:
            self.entry_target.delete(0, tk.END)
            self.entry_target.insert(0, d)
            self.target_dir = d
            self.cfg["target_folder"] = d
            save_config(self.cfg)
            self._show_scan_hint()      # 스캔은 사용자가 누를 때만

    def browse_pdf(self):
        d = filedialog.askdirectory(
            initialdir=safe_start_dir(self.entry_pdf.get().strip()),
            title="PDF 폴더 선택")
        if d:
            self.entry_pdf.delete(0, tk.END)
            self.entry_pdf.insert(0, d)
            self.pdf_dir = d
            self.cfg["last_pdf_folder"] = d
            save_config(self.cfg)
            self._show_scan_hint()

    def refresh_funds(self):
        self.target_dir = self.entry_target.get().strip()
        self.fund_map, self.file_list = scan_fund_files(self.target_dir)
        self.log_msg(f"자금판 {len(self.file_list)}개 발견: {', '.join(self.file_list[:3])}")

    def scan_pdfs(self):
        def _work():
            try:
                self.root.after(0, lambda: self.lbl_status.config(text="PDF 폴더 스캔 중..."))
                self.refresh_funds()
                pdf_dir = self.entry_pdf.get().strip()
                if not os.path.exists(pdf_dir):
                    self.root.after(0, lambda: messagebox.showwarning("경로 오류", "PDF 폴더가 존재하지 않습니다."))
                    self.root.after(0, lambda: self.lbl_status.config(text="경로 오류"))
                    return

                pdfs = glob.glob(os.path.join(pdf_dir, "*.pdf"))
                if not pdfs:
                    pdfs = glob.glob(os.path.join(pdf_dir, "*.PDF"))

                def clear_ui():
                    self.pdf_infos = []
                    for w in self.tree.get_children():
                        self.tree.delete(w)
                self.root.after(0, clear_ui)
                import time; time.sleep(0.05)

                if not pdfs:
                    self.root.after(0, lambda: self.lbl_status.config(text="PDF 없음"))
                    self.log_msg("PDF 파일 없음")
                    return

                infos = []
                for pdf in sorted(pdfs):
                    try:
                        info = parse_repo_pdf(pdf)
                        info["path"] = pdf
                        matched = find_matching_excel(info.get("fund_code"), self.fund_map)
                        info["matched"] = matched
                        if matched:
                            fp = os.path.join(self.target_dir, matched)
                            exists_buy = check_exists(fp, info["buy_date"], info["amt_start"], "REPO매수") if info["buy_date"] and info["amt_start"] else False
                            exists_sell = check_exists(fp, info["sell_date"], info["amt_end"], "REPO환매") if info["sell_date"] and info["amt_end"] else False
                            status = "이미 반영됨" if (exists_buy and exists_sell) else ("일부 반영됨" if (exists_buy or exists_sell) else "미반영")
                        else:
                            status = "매칭 실패"
                        info["status"] = status
                        infos.append(info)

                        self.root.after(0, lambda i=info, m=matched, s=status, p=pdf: self.tree.insert("", "end", values=(
                            os.path.basename(p),
                            i.get("fund_code") or "-",
                            i.get("buy_date") or "-",
                            i.get("sell_date") or "-",
                            f"{i.get('amt_start'):,}" if i.get("amt_start") else "-",
                            f"{i.get('amt_end'):,}" if i.get("amt_end") else "-",
                            m or "(없음)",
                            s
                        )))
                        self.log_msg(f"{os.path.basename(pdf)} -> 자금판: {matched or '매칭실패'} (상태: {status})")
                    except Exception as e:
                        self.log_msg(f"PDF 파싱 실패 {pdf}: {e}")
                        self.root.after(0, lambda p=pdf, err=str(e): self.tree.insert("", "end", values=(os.path.basename(p), "-", "-", "-", "-", "-", "-", f"파싱 실패: {err}")))

                def finish():
                    self.pdf_infos = infos
                    self.lbl_status.config(text=f"{len(infos)}개 PDF 스캔 완료")
                    counts = {}
                    for i in infos:
                        key = i.get("status") or ""
                        if key:
                            counts[key] = counts.get(key, 0) + 1
                    self.status_bar.set_summary(
                        " · ".join(f"{k} {v}건" for k, v in counts.items()))
                    self.cfg["last_pdf_folder"] = pdf_dir
                    self.cfg["target_folder"] = self.target_dir
                    save_config(self.cfg)

                self.root.after(0, finish)
            except Exception as e:
                self.log_msg(f"스캔 오류: {e}")
                self.root.after(0, lambda: messagebox.showerror("오류", str(e)))
                self.root.after(0, lambda: self.lbl_status.config(text="오류"))

        threading.Thread(target=_work, daemon=True).start()

    def apply_all(self):
        if not self.pdf_infos:
            messagebox.showwarning("경고", "먼저 PDF 스캔을 하세요.")
            return

        def _work():
            ok = 0
            skip = 0
            fail = 0
            self.root.after(0, lambda: self.lbl_status.config(text="엑셀 반영 중..."))
            for info in list(self.pdf_infos):
                if not info.get("matched"):
                    self.log_msg(f"건너뜀: {info['pdf_name']} (매칭 실패)")
                    skip += 1
                    continue
                if info.get("status") == "이미 반영됨":
                    self.log_msg(f"건너뜀: {info['pdf_name']} (이미 반영됨)")
                    skip += 1
                    continue

                fp = os.path.join(self.target_dir, info["matched"])
                try:
                    res = append_repo_rows(fp, info["buy_date"], info["sell_date"], info["amt_start"], info["amt_end"])
                    for item, st in res:
                        self.log_msg(f"{info['matched']} {item['date']} {item['desc']} {item['amount']:,} -> {st}")
                        if st == "added":
                            ok += 1
                        else:
                            skip += 1
                except Exception as e:
                    self.log_msg(f"반영 실패 {info['pdf_name']}: {e}")
                    fail += 1

            self.root.after(0, lambda: messagebox.showinfo("완료", f"추가 성공: {ok}건\n건너뜀: {skip}건\n실패: {fail}건"))
            self.root.after(0, self.scan_pdfs)

        threading.Thread(target=_work, daemon=True).start()

def build(parent):
    """통합 프로그램의 탭 안에 이 화면을 만든다."""
    return RepoApp(parent)


def main():
    root = tk.Tk()
    RepoApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
