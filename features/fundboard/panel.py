import os
import sys
import re
import json
import datetime
import unicodedata
from copy import copy
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import xlrd
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill

from common.config import load_section, save_section
from common.widgets import StatusBar, StatusProxy, safe_start_dir

SECTION_NAME = "자금판"

CONFIG_FILENAME = "mapping_config.json"
FUND_MAP_FILENAME = "fund_name_mapping.json"
SETTINGS_FILENAME = "app_settings.json"
UNIFIED_CONFIG_FILENAME = "자금판_자동반영도구_설정.json"


def get_app_directory():
    # PyInstaller EXE에서는 EXE가 있는 폴더, 소스 실행에서는 app.py 폴더를 사용합니다.
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.abspath(__file__))


APP_DIRECTORY = get_app_directory()
UNIFIED_CONFIG_PATH = os.path.join(APP_DIRECTORY, UNIFIED_CONFIG_FILENAME)
# 이전 버전의 파일을 읽어 한 번만 통합하기 위한 경로입니다.
LEGACY_CONFIG_PATH = os.path.join(APP_DIRECTORY, CONFIG_FILENAME)
LEGACY_FUND_MAP_PATH = os.path.join(APP_DIRECTORY, FUND_MAP_FILENAME)
LEGACY_SETTINGS_PATH = os.path.join(APP_DIRECTORY, SETTINGS_FILENAME)


DEFAULT_MAPPING = {
    "보통예금이자": "은대이자수령",
    "운용보수": "운용보수",
    "수탁보수": "수탁보수",
    "사무관리보수": "사무관리보수",
    "판매보수": "판매보수"
}

def clean_path(path_str):
    if not path_str:
        return ""
    return os.path.normpath(path_str.strip())

def normalize_text(text):
    if not text:
        return ""
    text = unicodedata.normalize("NFC", str(text))
    return re.sub(r'\s+', '', text)

def _read_json_file(path, fallback):
    try:
        if os.path.exists(path):
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
    except Exception:
        pass
    return fallback


def _write_json_file(path, payload):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def _archive_legacy_json(path):
    if not os.path.exists(path):
        return
    backup_path = path + ".legacy.bak"
    try:
        if not os.path.exists(backup_path):
            os.replace(path, backup_path)
        else:
            os.remove(path)
    except Exception:
        # 권한 문제 등으로 정리하지 못한 경우에도 통합 설정은 정상 사용합니다.
        pass


def _coerce_mapping_rules(raw_rules):
    if isinstance(raw_rules, dict) and isinstance(raw_rules.get("rules"), list):
        return normalize_mapping_rules(raw_rules["rules"])
    if isinstance(raw_rules, list):
        return normalize_mapping_rules(raw_rules)
    if isinstance(raw_rules, dict):
        return normalize_mapping_rules([
            {"reason": str(reason), "target_item": "", "description": str(description)}
            for reason, description in raw_rules.items()
        ])
    return []


def _new_shared_config():
    return {
        "app_settings": {},
        "mapping_rules": [],
        "fund_name_mapping": {},
    }


def _save_shared_config(config):
    config.pop("version", None)
    config["app_settings"] = config.get("app_settings") if isinstance(config.get("app_settings"), dict) else {}
    config["mapping_rules"] = normalize_mapping_rules(config.get("mapping_rules", []))
    config["fund_name_mapping"] = config.get("fund_name_mapping") if isinstance(config.get("fund_name_mapping"), dict) else {}
    save_section(SECTION_NAME, config)


def _load_shared_config():
    data = load_section(SECTION_NAME, None)
    if isinstance(data, dict):
        config = _new_shared_config()
        config.update(data)
        if not config.get("mapping_rules"):
            config["mapping_rules"] = default_mapping_rules()
        config["mapping_rules"] = normalize_mapping_rules(config["mapping_rules"])
        return config

    # 기존 3개 JSON이 있을 때 첫 실행에서 통합 설정으로 옮깁니다.
    config = _new_shared_config()
    legacy_settings = _read_json_file(LEGACY_SETTINGS_PATH, {})
    if isinstance(legacy_settings, dict):
        config["app_settings"] = legacy_settings
    legacy_mapping = _read_json_file(LEGACY_CONFIG_PATH, {})
    config["mapping_rules"] = _coerce_mapping_rules(legacy_mapping)
    existing_rule_keys = {
        (normalize_text(rule["reason"]), normalize_text(rule["target_item"]))
        for rule in config["mapping_rules"]
    }
    for default_rule in default_mapping_rules():
        default_key = (normalize_text(default_rule["reason"]), normalize_text(default_rule["target_item"]))
        if default_key not in existing_rule_keys:
            config["mapping_rules"].append(default_rule)
    config["mapping_rules"] = normalize_mapping_rules(config["mapping_rules"])
    legacy_fund_mapping = _read_json_file(LEGACY_FUND_MAP_PATH, {})
    if isinstance(legacy_fund_mapping, dict):
        config["fund_name_mapping"] = legacy_fund_mapping
    _save_shared_config(config)
    _archive_legacy_json(LEGACY_SETTINGS_PATH)
    _archive_legacy_json(LEGACY_CONFIG_PATH)
    _archive_legacy_json(LEGACY_FUND_MAP_PATH)
    return config


def load_app_settings():
    return _load_shared_config()["app_settings"]


def save_app_settings(settings_dict):
    config = _load_shared_config()
    config["app_settings"] = settings_dict if isinstance(settings_dict, dict) else {}
    _save_shared_config(config)


def _clean_mapping_rule(rule):
    if not isinstance(rule, dict):
        return None
    reason = str(rule.get("reason", "")).strip()
    target_item = str(rule.get("target_item", "")).strip()
    description = str(rule.get("description", "")).strip()
    if not reason or not description:
        return None
    return {
        "reason": reason,
        "target_item": target_item,
        "description": description,
    }


def normalize_mapping_rules(rules):
    result = []
    seen = set()
    for rule in rules or []:
        cleaned = _clean_mapping_rule(rule)
        if not cleaned:
            continue
        key = (normalize_text(cleaned["reason"]), normalize_text(cleaned["target_item"]))
        if key in seen:
            continue
        seen.add(key)
        result.append(cleaned)
    return result


def default_mapping_rules():
    """
    처음 쓸 때 넣어두는 기본 규칙.

    ※ 특정 회사(거래처) 이름은 넣지 않습니다. 필요하면 [사유 매핑 설정] 에서 직접 추가하세요.
    """
    rules = [
        {"reason": str(reason), "target_item": "", "description": str(description)}
        for reason, description in DEFAULT_MAPPING.items()
    ]
    return normalize_mapping_rules(rules)


def load_mapping_config(folder_path=None):
    # folder_path 매개변수는 이전 호출 호환성을 위해 유지하지만 사용하지 않습니다.
    config = _load_shared_config()
    return normalize_mapping_rules(config.get("mapping_rules", []))


def save_mapping_config(folder_path, rules):
    # folder_path 매개변수는 이전 호출 호환성을 위해 유지하지만 사용하지 않습니다.
    config = _load_shared_config()
    config["mapping_rules"] = normalize_mapping_rules(rules)
    _save_shared_config(config)


def resolve_mapping_description(rules, reason, target_item):
    normalized_reason = normalize_text(reason)
    normalized_target_item = normalize_text(target_item)
    for rule in rules:
        if (
            normalize_text(rule.get("reason", "")) == normalized_reason
            and rule.get("target_item", "").strip()
            and normalize_text(rule.get("target_item", "")) == normalized_target_item
        ):
            return rule.get("description", reason)
    for rule in rules:
        if (
            normalize_text(rule.get("reason", "")) == normalized_reason
            and not rule.get("target_item", "").strip()
        ):
            return rule.get("description", reason)
    return reason


def find_mapping_rule_index(rules, reason, target_item):
    key = (normalize_text(reason), normalize_text(target_item))
    for index, rule in enumerate(rules):
        if key == (normalize_text(rule.get("reason", "")), normalize_text(rule.get("target_item", ""))):
            return index
    return None


def load_fund_name_mapping(folder_path=None):
    # folder_path 매개변수는 이전 호출 호환성을 위해 유지하지만 사용하지 않습니다.
    config = _load_shared_config()
    mapping = config.get("fund_name_mapping", {})
    return mapping if isinstance(mapping, dict) else {}


def save_fund_name_mapping(folder_path, mapping_dict):
    # folder_path 매개변수는 이전 호출 호환성을 위해 유지하지만 사용하지 않습니다.
    config = _load_shared_config()
    config["fund_name_mapping"] = mapping_dict if isinstance(mapping_dict, dict) else {}
    _save_shared_config(config)


def check_is_applied(filepath, target_date_str, target_amt, is_inout, target_desc):
    try:
        wb = openpyxl.load_workbook(filepath, data_only=True)
        acc_sheet = None
        for s in wb.sheetnames:
            if "은대계좌" in s or "계좌" in s or "예금" in s:
                acc_sheet = s
                break
        if not acc_sheet and len(wb.sheetnames) > 1:
            acc_sheet = wb.sheetnames[1]
            
        if not acc_sheet:
            return False
            
        ws = wb[acc_sheet]
        
        if len(target_date_str) == 8 and target_date_str.isdigit():
            t_dt = datetime.date(int(target_date_str[:4]), int(target_date_str[4:6]), int(target_date_str[6:8]))
        else:
            t_dt = None

        max_r = ws.max_row
        for r in range(max_r, max(4, max_r - 30), -1):
            row_date = ws.cell(r, 2).value
            row_desc = str(ws.cell(r, 3).value or "").strip()
            row_in = ws.cell(r, 4).value
            row_out = ws.cell(r, 5).value
            
            date_match = False
            if isinstance(row_date, datetime.datetime):
                row_dt = row_date.date()
                date_match = (row_dt == t_dt)
            elif isinstance(row_date, datetime.date):
                date_match = (row_date == t_dt)
            elif str(row_date).replace("-", "").strip() == target_date_str:
                date_match = True

            amt_match = False
            if is_inout == "환입" and row_in is not None:
                amt_match = abs(float(row_in) - float(target_amt)) < 0.01
            elif is_inout == "인출" and row_out is not None:
                amt_match = abs(float(row_out) - float(target_amt)) < 0.01

            if date_match and amt_match and row_desc == str(target_desc).strip():
                wb.close()
                return True

        wb.close()
    except Exception as e:
        print(f"Error checking applied status in {filepath}: {e}")
    return False

def apply_single_item(filepath, item):
    wb = openpyxl.load_workbook(filepath)
    acc_sheet = None
    for s in wb.sheetnames:
        if "은대계좌" in s or "계좌" in s or "예금" in s:
            acc_sheet = s
            break
    if not acc_sheet and len(wb.sheetnames) > 1:
        acc_sheet = wb.sheetnames[1]

    ws = wb[acc_sheet]

    last_r = ws.max_row
    while last_r > 4 and not any(ws.cell(last_r, c).value is not None for c in range(1, 8)):
        last_r -= 1

    new_r = last_r + 1

    # 1. Copy exact Row Height from previous row
    if last_r in ws.row_dimensions and ws.row_dimensions[last_r].height:
        ws.row_dimensions[new_r].height = ws.row_dimensions[last_r].height
    else:
        ws.row_dimensions[new_r].height = 18.0

    date_str = item["date"]
    if len(date_str) == 8 and date_str.isdigit():
        date_obj = datetime.date(int(date_str[:4]), int(date_str[4:6]), int(date_str[6:8]))
    else:
        date_obj = date_str

    desc = str(item["target_desc"]).strip()
    amt = item["amount"]
    type_inout = item["type"]

    ws.cell(new_r, 2).value = date_obj
    ws.cell(new_r, 3).value = desc

    if type_inout == "환입":
        ws.cell(new_r, 4).value = amt
        ws.cell(new_r, 5).value = None
    else:
        ws.cell(new_r, 4).value = None
        ws.cell(new_r, 5).value = amt

    ws.cell(new_r, 6).value = f"=F{last_r}+D{new_r}-E{new_r}"

    max_c = max(7, ws.max_column)

    # 2. Copy EXACT cell styles column-by-column without cross-column format leakage!
    for col_idx in range(1, max_c + 1):
        ref_cell = None
        for r in range(last_r, 4, -1):
            c_candidate = ws.cell(r, col_idx)
            if c_candidate.value is not None:
                ref_cell = c_candidate
                break
        
        if not ref_cell:
            for r in range(last_r, 4, -1):
                c_candidate = ws.cell(r, col_idx)
                if c_candidate.has_style:
                    ref_cell = c_candidate
                    break
        if not ref_cell:
            ref_cell = ws.cell(last_r, col_idx)

        tgt_cell = ws.cell(new_r, col_idx)
        if ref_cell and ref_cell.has_style:
            tgt_cell.font = copy(ref_cell.font)
            tgt_cell.fill = copy(ref_cell.fill)
            tgt_cell.alignment = copy(ref_cell.alignment)
            
            if col_idx == 1:
                b = ref_cell.border
                if not any([b.left, b.right, b.top, b.bottom]):
                    tgt_cell.border = Border()
                else:
                    tgt_cell.border = copy(ref_cell.border)
            else:
                tgt_cell.border = copy(ref_cell.border)

            if col_idx in (4, 5) and ref_cell.number_format == 'General':
                bal_ref = None
                for r in range(last_r, 4, -1):
                    c_bal = ws.cell(r, 6)
                    if c_bal.value is not None and c_bal.number_format != 'General':
                        bal_ref = c_bal
                        break
                if bal_ref:
                    tgt_cell.number_format = bal_ref.number_format
                else:
                    tgt_cell.number_format = ref_cell.number_format
            else:
                tgt_cell.number_format = ref_cell.number_format

    wb.save(filepath)
    wb.close()

def rollback_single_item(filepath, target_date_str, target_amt, is_inout, target_desc):
    wb = openpyxl.load_workbook(filepath)
    acc_sheet = None
    for s in wb.sheetnames:
        if "은대계좌" in s or "계좌" in s or "예금" in s:
            acc_sheet = s
            break
    if not acc_sheet and len(wb.sheetnames) > 1:
        acc_sheet = wb.sheetnames[1]

    if not acc_sheet:
        wb.close()
        return False

    ws = wb[acc_sheet]

    if len(target_date_str) == 8 and target_date_str.isdigit():
        t_dt = datetime.date(int(target_date_str[:4]), int(target_date_str[4:6]), int(target_date_str[6:8]))
    else:
        t_dt = None

    max_r = ws.max_row
    deleted_row = False

    for r in range(max_r, 4, -1):
        row_date = ws.cell(r, 2).value
        row_desc = str(ws.cell(r, 3).value or "").strip()
        row_in = ws.cell(r, 4).value
        row_out = ws.cell(r, 5).value

        date_match = False
        if isinstance(row_date, datetime.datetime):
            row_dt = row_date.date()
            date_match = (row_dt == t_dt)
        elif isinstance(row_date, datetime.date):
            date_match = (row_date == t_dt)
        elif str(row_date).replace("-", "").strip() == target_date_str:
            date_match = True

        amt_match = False
        if is_inout == "환입" and row_in is not None:
            amt_match = abs(float(row_in) - float(target_amt)) < 0.01
        elif is_inout == "인출" and row_out is not None:
            amt_match = abs(float(row_out) - float(target_amt)) < 0.01

        if date_match and amt_match and row_desc == str(target_desc).strip():
            ws.delete_rows(r, 1)
            deleted_row = True
            break

    if deleted_row:
        wb.save(filepath)

    wb.close()
    return deleted_row


class FundAutoUpdaterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("자금판 이자/환입인출 자동 반영 도구")
        self.root.geometry("1280x760")
        self.root.minsize(1050, 620)

        self.style = ttk.Style()
        self.style.theme_use("clam")
        
        self.style.configure(".", font=("Malgun Gothic", 10))
        self.style.configure("Treeview.Heading", font=("Malgun Gothic", 10, "bold"), background="#e2e8f0", foreground="#0f172a")
        self.style.configure("Treeview", font=("Malgun Gothic", 9), rowheight=28, background="#ffffff", fieldbackground="#ffffff")
        self.style.map("Treeview", background=[("selected", "#2563eb")], foreground=[("selected", "#ffffff")])
        
        self.style.configure("Primary.TButton", font=("Malgun Gothic", 10, "bold"), background="#2563eb", foreground="white")
        self.style.map("Primary.TButton", background=[("active", "#1d4ed8")])
        self.style.configure("Success.TButton", font=("Malgun Gothic", 10, "bold"), background="#16a34a", foreground="white")
        self.style.map("Success.TButton", background=[("active", "#15803d")])

        # Load persistent app settings
        self.app_settings = load_app_settings()
        
        # 1. Row 1: 작업 폴더
        # 저장된 경로는 '이 컴퓨터에 없더라도' 그대로 보여준다. (운용지시대장과 같은 방식)
        # 예전에는 없으면 현재 폴더로 바꿔서 저장해버려서, 다른 PC/공유 경로가 지워졌다.
        saved_target_folder = (self.app_settings.get("last_target_folder") or "").strip()
        if saved_target_folder:
            self.target_dir = clean_path(saved_target_folder)
        else:
            self.target_dir = clean_path(os.getcwd())

        # 2. Row 2: 지시내역 폴더 (저장된 값이 있으면 그대로 유지)
        saved_instr_folder = (self.app_settings.get("last_instr_folder") or "").strip()
        if saved_instr_folder:
            self.instr_dir = clean_path(saved_instr_folder)
        else:
            sub_default = os.path.join(self.target_dir, "지시내역")
            if os.path.exists(sub_default):
                self.instr_dir = clean_path(sub_default)
            else:
                self.instr_dir = clean_path(self.target_dir)

        self.mapping_rules = load_mapping_config(self.target_dir)
        self.custom_fund_mappings = load_fund_name_mapping(self.target_dir)
        
        self.instruction_files_info = {}
        self.selected_instr_file = tk.StringVar()
        self.fund_files_map = {}
        self.all_fund_file_names = []
        self.items_data = []
        self.checked_vars = {}
        self.badge_widgets = []
        self.select_all_var = tk.BooleanVar(value=True)

        self.create_widgets()
        # 탭을 열 때는 아무 파일도 건드리지 않는다 (렉 방지). 스캔은 사용자가 누를 때만.
        self.status_bar.set_status("스캔을 눌러주세요.")

    def create_widgets(self):
        # 화면 맨 아래 공통 상태 표시줄 (세 기능 동일) — 가장 먼저 자리를 잡아야 밀리지 않는다
        self.status_bar = StatusBar(self.root)
        self.status_bar.pack(side="bottom", fill="x")

        # 통합 프로그램 공통 머리글 (다른 탭과 같은 모양으로 맞춤)
        head = tk.Frame(self.root, bg="#0f172a")
        head.pack(fill="x")
        tk.Label(head, text="자금판", font=("맑은 고딕", 16, "bold"),
                 bg="#0f172a", fg="#ffffff").pack(side="left", padx=18, pady=(12, 2))
        tk.Label(head, text="지시내역을 자금판에 자동으로 반영합니다",
                 font=("맑은 고딕", 10), bg="#0f172a", fg="#cbd5e1").pack(
            side="right", padx=18, pady=(18, 0))

        # 1. 폴더 선택 영역 (운용지시대장과 같은 모양: 테두리 상자 없이 줄 3개)
        top_frame = tk.Frame(self.root, bg="#ffffff", padx=18, pady=12)
        top_frame.pack(fill="x")
        top_frame.columnconfigure(1, weight=1)

        lbl_font = ("맑은 고딕", 10, "bold")
        # 경로 영역 배치 규칙 (모든 탭 공통)
        #   라벨(12자) - 입력칸(끝까지) - 버튼(10자 고정)
        #   스캔은 맨 아래 줄, 입력칸 자리는 비워 둔다
        LBL_W = 12
        BTN_W = 12

        # Row 1: 작업 폴더
        lbl_target = tk.Label(top_frame, text="자금판 폴더", font=lbl_font, bg="#ffffff",
                              width=LBL_W, anchor="w")
        lbl_target.grid(row=0, column=0, sticky="w", pady=3)

        self.folder_path_entry = ttk.Entry(top_frame, font=("맑은 고딕", 10))
        self.folder_path_entry.insert(0, self.target_dir)
        self.folder_path_entry.grid(row=0, column=1, sticky="ew", padx=6, pady=3)
        self.folder_path_entry.bind("<KeyRelease>", self._show_scan_hint)
        self.folder_path_entry.bind("<FocusOut>", self._show_scan_hint)

        btn_browse_target = ttk.Button(top_frame, text="① 찾아보기", width=BTN_W, style="Path.TButton",
                                       command=self.browse_target_folder)
        btn_browse_target.grid(row=0, column=2, sticky="w", pady=3)

        # Row 2: 지시내역 폴더
        lbl_instr_dir = tk.Label(top_frame, text="지시내역 폴더", font=lbl_font, bg="#ffffff",
                                 width=LBL_W, anchor="w")
        lbl_instr_dir.grid(row=1, column=0, sticky="w", pady=3)

        self.instr_folder_entry = ttk.Entry(top_frame, font=("맑은 고딕", 10))
        self.instr_folder_entry.insert(0, self.instr_dir)
        self.instr_folder_entry.grid(row=1, column=1, sticky="ew", padx=6, pady=3)
        self.instr_folder_entry.bind("<KeyRelease>", self._show_scan_hint)
        self.instr_folder_entry.bind("<FocusOut>", self._show_scan_hint)

        btn_browse_instr = ttk.Button(top_frame, text="② 찾아보기", width=BTN_W, style="Path.TButton",
                                      command=self.browse_instr_folder)
        btn_browse_instr.grid(row=1, column=2, sticky="w", pady=3)

        # Row 3: 스캔 — 입력칸 자리는 비워 둔다
        self.btn_scan = ttk.Button(top_frame, text="③ 스캔", width=BTN_W,
                                   style="PathAccent.TButton", command=self.scan_folder)
        self.btn_scan.grid(row=2, column=2, sticky="w", pady=3)

        # Row 4: 지시내역 파일 (스캔해야 목록이 채워진다)
        lbl_instr_file = tk.Label(top_frame, text="지시내역 파일", font=lbl_font, bg="#ffffff",
                                  width=LBL_W, anchor="w")
        lbl_instr_file.grid(row=3, column=0, sticky="w", pady=3)

        self.file_combo = ttk.Combobox(top_frame, textvariable=self.selected_instr_file, state="readonly")
        self.file_combo.grid(row=3, column=1, sticky="ew", padx=6, pady=3)
        self.file_combo.bind("<<ComboboxSelected>>", lambda e: self.load_instruction_data())

        btn_rule = ttk.Button(top_frame, text="사유 매핑 설정", width=BTN_W,
                              command=self.open_mapping_dialog)
        btn_rule.grid(row=3, column=2, sticky="w", pady=3)

        # 2. 상태 표시는 화면 맨 아래 공통 표시줄로 이동 (create_widgets 끝에서 만든다)

        # 3. Main Treeview Table Frame
        table_frame = ttk.Frame(self.root, padding=(18, 5))
        table_frame.pack(fill="both", expand=True)

        cols = ("chk", "seq", "fund_name", "type", "reason", "target_desc", "amount", "matched_file", "match_status", "apply_status")
        self.tree = ttk.Treeview(table_frame, columns=cols, show="headings", selectmode="browse")

        self.tree.heading("chk", text="☑ 선택", command=self.toggle_select_all)
        self.tree.heading("seq", text="순번")
        self.tree.heading("fund_name", text="지시내역 펀드명")
        self.tree.heading("type", text="구분")
        self.tree.heading("reason", text="지시 사유")
        self.tree.heading("target_desc", text="자금판 작성 내용")
        self.tree.heading("amount", text="금액(원)")
        self.tree.heading("matched_file", text="매칭된 자금판 파일")
        self.tree.heading("match_status", text="매칭 여부")
        self.tree.heading("apply_status", text="반영 여부")

        self.tree.column("chk", width=55, anchor="center")
        self.tree.column("seq", width=40, anchor="center")
        self.tree.column("fund_name", width=220, anchor="w")
        self.tree.column("type", width=50, anchor="center")
        self.tree.column("reason", width=95, anchor="center")
        self.tree.column("target_desc", width=145, anchor="w")
        self.tree.column("amount", width=105, anchor="e")
        self.tree.column("matched_file", width=200, anchor="w")
        self.tree.column("match_status", width=95, anchor="center")
        self.tree.column("apply_status", width=180, anchor="center")

        # Scrollbars
        self.vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.on_scroll_v)
        self.hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.on_scroll_h)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)

        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")

        table_frame.rowconfigure(0, weight=1)
        table_frame.columnconfigure(0, weight=1)

        # Bindings
        self.tree.bind("<Double-1>", self.on_double_click_cell)
        self.tree.bind("<Button-1>", self.on_single_click_cell)
        self.tree.bind("<Configure>", lambda e: self.root.after(50, self.render_cell_badges))

        # 4. Bottom Action Frame with TWO Buttons
        bottom_frame = ttk.Frame(self.root, padding=18)
        bottom_frame.pack(fill="x")

        btn_box = ttk.Frame(bottom_frame)
        btn_box.pack(side="right")

        self.btn_execute_selected = ttk.Button(btn_box, text="선택만 반영", command=self.execute_update_selected)
        self.btn_execute_selected.pack(side="left", padx=5)

        self.btn_execute_all = ttk.Button(btn_box, text="전체 반영", style="Primary.TButton", command=self.execute_update_all)
        self.btn_execute_all.pack(side="left", padx=5)

        # 5. 상태 표시줄 연결 (표시줄 자체는 create_widgets 맨 위에서 만들었다)
        self.lbl_status = StatusProxy(self.status_bar, "status")
        self.lbl_summary = StatusProxy(self.status_bar, "summary")

    def on_scroll_v(self, *args):
        self.tree.yview(*args)
        self.root.after(10, self.render_cell_badges)

    def on_scroll_h(self, *args):
        self.tree.xview(*args)
        self.root.after(10, self.render_cell_badges)

    def toggle_select_all(self):
        new_val = not self.select_all_var.get()
        self.select_all_var.set(new_val)
        for iid in self.tree.get_children():
            try:
                idx = int(iid)
                if idx in self.checked_vars:
                    self.checked_vars[idx].set(new_val)
                    vals = list(self.tree.item(iid, "values"))
                    vals[0] = "☑" if new_val else "☐"
                    self.tree.item(iid, values=vals)
            except ValueError:
                pass

    def on_single_click_cell(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        
        column = self.tree.identify_column(event.x)
        item_id = self.tree.identify_row(event.y)

        if column == "#1" and item_id:
            try:
                idx = int(item_id)
                current_chk = self.checked_vars[idx].get()
                new_chk = not current_chk
                self.checked_vars[idx].set(new_chk)
                
                vals = list(self.tree.item(item_id, "values"))
                vals[0] = "☑" if new_chk else "☐"
                self.tree.item(item_id, values=vals)
            except ValueError:
                pass

    def open_manual_match_dialog(self, idx):
        if idx >= len(self.items_data):
            return
        item = self.items_data[idx]

        chosen_filepath = filedialog.askopenfilename(
            initialdir=safe_start_dir(self.target_dir,
                                      self.app_settings.get("last_picker_dir")),
            title=f"[{item['fund_name']}] 연결할 자금판 엑셀 파일 선택",
            filetypes=[("자금판 엑셀 파일", "*_자금판.xlsx"), ("모든 엑셀 파일", "*.xlsx")]
        )

        if chosen_filepath:
            chosen_file = os.path.basename(chosen_filepath)
            norm_fund = normalize_text(item["fund_name"])
            self.custom_fund_mappings[norm_fund] = chosen_file
            save_fund_name_mapping(self.target_dir, self.custom_fund_mappings)
            self.load_instruction_data()
            messagebox.showinfo("매칭 완료", f"수동 매칭이 완료되었습니다!\n[{item['fund_name']}] ➔ [{chosen_file}]")

    def toggle_single_item_apply(self, idx):
        if idx >= len(self.items_data):
            return
        item = self.items_data[idx]

        if not item["matched_file"]:
            messagebox.showerror("오류", f"'{item['fund_name']}' 펀드에 해당되는 자금판 엑셀 파일이 없습니다.\n[매칭 여부] 칸을 더블클릭하여 자금판 파일을 직접 선택하세요.")
            return

        filepath = os.path.join(self.target_dir, item["matched_file"])

        if not item["is_applied"]:
            try:
                apply_single_item(filepath, item)
                self.load_instruction_data()
            except Exception as err:
                messagebox.showerror("오류", f"반영 실패:\n{err}")
        else:
            try:
                rollback_single_item(filepath, item["date"], item["amount"], item["type"], item["target_desc"])
                self.load_instruction_data()
            except Exception as err:
                messagebox.showerror("오류", f"반영 취소 실패:\n{err}")

    def render_cell_badges(self):
        for w in self.badge_widgets:
            w.destroy()
        self.badge_widgets = []

        for item_id in self.tree.get_children():
            try:
                idx = int(item_id)
            except ValueError:
                continue

            if idx >= len(self.items_data):
                continue

            item = self.items_data[idx]

            # 1. Col 9 (match_status) badge overlay
            bbox_match = self.tree.bbox(item_id, "#9")
            if bbox_match and len(bbox_match) == 4:
                x, y, w, h = bbox_match
                if w > 10 and h > 5:
                    bg = "#dcfce7" if item["matched_file"] else "#fecaca"
                    fg = "#14532d" if item["matched_file"] else "#7f1d1d"
                    lbl = tk.Label(self.tree, text=item["match_status"], bg=bg, fg=fg, font=("Malgun Gothic", 9, "bold"), relief="flat", cursor="hand2")
                    lbl.place(x=x+2, y=y+2, width=w-4, height=h-4)
                    
                    lbl.bind("<Double-1>", lambda e, item_idx=idx: self.open_manual_match_dialog(item_idx))
                    self.badge_widgets.append(lbl)

            # 2. Col 10 (apply_status) badge overlay
            bbox_apply = self.tree.bbox(item_id, "#10")
            if bbox_apply and len(bbox_apply) == 4:
                x, y, w, h = bbox_apply
                if w > 10 and h > 5:
                    if item["matched_file"]:
                        bg = "#dcfce7" if item["is_applied"] else "#fecaca"
                        fg = "#14532d" if item["is_applied"] else "#7f1d1d"
                    else:
                        bg = "#e2e8f0"
                        fg = "#475569"
                    lbl = tk.Label(self.tree, text=item["apply_status"], bg=bg, fg=fg, font=("Malgun Gothic", 9, "bold"), relief="flat", cursor="hand2")
                    lbl.place(x=x+2, y=y+2, width=w-4, height=h-4)
                    
                    lbl.bind("<Double-1>", lambda e, item_idx=idx: self.toggle_single_item_apply(item_idx))
                    self.badge_widgets.append(lbl)

    def _picker_start(self, current=""):
        """폴더 선택창을 여는 위치.
        현재 경로를 그대로 쓰면, 그 경로가 네트워크 공유일 때 윈도우가 접속을 시도하느라
        창이 멈춘다. 그래서 '마지막으로 실제 선택했던 폴더'를 우선 쓰고, 네트워크 경로는 건너뛴다."""
        return safe_start_dir(self.app_settings.get("last_picker_dir"), current)

    def browse_target_folder(self):
        selected = filedialog.askdirectory(initialdir=self._picker_start(self.target_dir),
                                           title="자금판 폴더 선택")
        if selected:
            self.target_dir = clean_path(selected)
            self.folder_path_entry.delete(0, tk.END)
            self.folder_path_entry.insert(0, self.target_dir)

            # Keep the independently selected instruction folder unchanged.
            self.app_settings["last_target_folder"] = self.target_dir
            self.app_settings["last_instr_folder"] = self.instr_dir
            self.app_settings["last_picker_dir"] = self.target_dir
            save_app_settings(self.app_settings)

            self.mapping_rules = load_mapping_config(self.target_dir)
            self.custom_fund_mappings = load_fund_name_mapping(self.target_dir)
            self._show_scan_hint()   # 스캔은 사용자가 누를 때만

    def browse_instr_folder(self):
        selected = filedialog.askdirectory(initialdir=self._picker_start(self.instr_dir),
                                           title="지시내역 폴더 선택")
        if selected:
            self.instr_dir = clean_path(selected)
            self.instr_folder_entry.delete(0, tk.END)
            self.instr_folder_entry.insert(0, self.instr_dir)

            self.app_settings["last_instr_folder"] = self.instr_dir
            self.app_settings["last_picker_dir"] = self.instr_dir
            save_app_settings(self.app_settings)
            self._show_scan_hint()   # 스캔은 사용자가 누를 때만

    def _show_scan_hint(self, event=None):
        """경로가 바뀌었을 때: 스캔을 눌러야 목록이 갱신된다고 알려준다. (파일 접근 없음)"""
        self.status_bar.set_status("폴더가 바뀌었습니다. 스캔을 눌러주세요.", "warn")

    def scan_folder(self, save=True):
        """save=False 이면 설정을 저장하지 않는다.
        프로그램을 켜는 것만으로 저장된 경로가 바뀌는 사고를 막기 위한 장치."""
        self.target_dir = clean_path(self.folder_path_entry.get())
        self.instr_dir = clean_path(self.instr_folder_entry.get())

        if not os.path.exists(self.target_dir):
            self.lbl_status.config(text="작업 폴더 경로가 존재하지 않습니다.", foreground="#dc2626")
            self.instruction_files_info = {}
            self.file_combo["values"] = []
            self.selected_instr_file.set("")
            self.items_data = []
            self.clear_tree()
            self.lbl_summary.config(text="")
            return

        if not os.path.exists(self.instr_dir):
            self.lbl_status.config(text="지시내역 폴더 경로가 존재하지 않습니다.", foreground="#dc2626")
            self.instruction_files_info = {}
            self.file_combo["values"] = []
            self.selected_instr_file.set("")
            self.items_data = []
            self.clear_tree()
            self.lbl_summary.config(text="")
            return

        # Save settings (사용자가 직접 바꾼 경우에만)
        if save:
            self.app_settings["last_target_folder"] = self.target_dir
            self.app_settings["last_instr_folder"] = self.instr_dir
            save_app_settings(self.app_settings)

        self.mapping_rules = load_mapping_config(self.target_dir)
        self.custom_fund_mappings = load_fund_name_mapping(self.target_dir)
        
        # Scan fund files (*_자금판.xlsx) ONLY in target_dir
        self.fund_files_map = {}
        self.all_fund_file_names = []
        if os.path.exists(self.target_dir):
            for f in os.listdir(self.target_dir):
                if f.endswith("_자금판.xlsx"):
                    self.all_fund_file_names.append(f)
                    filepath = os.path.join(self.target_dir, f)
                    try:
                        wb = openpyxl.load_workbook(filepath, data_only=True)
                        ws_meta = wb.worksheets[0]
                        full_name = ws_meta.cell(3, 2).value or ""
                        code = ws_meta.cell(5, 3).value or ""
                        
                        norm_file = normalize_text(f.replace("_자금판.xlsx", ""))
                        norm_full = normalize_text(full_name)
                        norm_code = normalize_text(code)
                        
                        self.fund_files_map[norm_file] = f
                        if norm_full:
                            self.fund_files_map[norm_full] = f
                            self.fund_files_map[norm_full.replace("부동산", "")] = f
                        if norm_code:
                            self.fund_files_map[norm_code] = f
                    except Exception as e:
                        print(f"Error reading metadata from {f}: {e}")

        self.all_fund_file_names.sort()

        # Scan instruction files STRICTLY & EXCLUSIVELY in instr_dir! (No fallback leakage!)
        self.instruction_files_info = {}
        if os.path.exists(self.instr_dir):
            for f in os.listdir(self.instr_dir):
                if ("환입인출운용지시내역" in f or "운용지시" in f) and (f.endswith(".xls") or f.endswith(".xlsx")):
                    full_p = os.path.join(self.instr_dir, f)
                    if os.path.isfile(full_p):
                        self.instruction_files_info[f] = full_p

        sorted_file_keys = sorted(list(self.instruction_files_info.keys()), reverse=True)
        self.file_combo["values"] = sorted_file_keys
        
        if sorted_file_keys:
            self.selected_instr_file.set(sorted_file_keys[0])
            self.load_instruction_data()
        else:
            self.selected_instr_file.set("")
            self.items_data = []
            self.clear_tree()
            self.lbl_status.config(text="선택하신 지시내역 폴더에 운용지시내역 엑셀 파일이 없습니다.", foreground="#dc2626")
            self.lbl_summary.config(text="")

    def clear_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for w in self.badge_widgets:
            w.destroy()
        self.badge_widgets = []
        self.checked_vars = {}

    def load_instruction_data(self):
        disp_name = self.selected_instr_file.get()
        if not disp_name or disp_name not in self.instruction_files_info:
            self.items_data = []
            self.clear_tree()
            self.lbl_summary.config(text="")
            return

        filepath = self.instruction_files_info[disp_name]
        self.clear_tree()
        self.items_data = []

        try:
            raw_rows = []
            if filepath.endswith(".xls"):
                wb = xlrd.open_workbook(filepath)
                ws = wb.sheet_by_index(0)
                for r in range(3, ws.nrows):
                    row = [ws.cell_value(r, c) for c in range(ws.ncols)]
                    if any(row[1:]):
                        raw_rows.append(row)
            else:
                wb = openpyxl.load_workbook(filepath, data_only=True)
                ws = wb.active
                for r in range(4, ws.max_row + 1):
                    row = [ws.cell(r, c).value for c in range(1, ws.max_column + 1)]
                    if any(row[1:]):
                        raw_rows.append(row)

            matched_ok_count = 0
            applied_count = 0
            pending_count = 0
            fail_count = 0
            total_amt = 0

            for idx, row in enumerate(raw_rows):
                if len(row) < 17:
                    continue
                seq_val = row[1]
                if not seq_val and seq_val != 0:
                    continue
                try:
                    seq = int(float(seq_val))
                except ValueError:
                    continue

                date_val = str(row[2]).strip()
                type_inout = str(row[5]).strip()
                reason = str(row[6]).strip()
                target_item = str(row[13]).strip() if len(row) > 13 and row[13] is not None else ""
                fund_code = str(row[7]).strip()
                fund_name = str(row[8]).strip()
                
                try:
                    amt = float(row[16])
                except (ValueError, TypeError):
                    amt = 0.0

                norm_fund = normalize_text(fund_name)
                norm_code = normalize_text(fund_code)
                
                # Check user custom manual fund mapping first!
                matched_file = self.custom_fund_mappings.get(norm_fund)
                if matched_file:
                    configured_path = os.path.join(self.target_dir, matched_file)
                    if not os.path.isfile(configured_path):
                        matched_file = None
                
                if not matched_file:
                    matched_file = self.fund_files_map.get(norm_fund)
                if not matched_file:
                    matched_file = self.fund_files_map.get(norm_fund.replace("부동산", ""))
                if not matched_file:
                    matched_file = self.fund_files_map.get(norm_code)
                    
                target_desc = resolve_mapping_description(self.mapping_rules, reason, target_item)
                match_status = "🟢 매칭 성공" if matched_file else "🔴 매칭 실패"

                is_applied = False
                if matched_file:
                    matched_ok_count += 1
                    target_file_path = os.path.join(self.target_dir, matched_file)
                    is_applied = check_is_applied(target_file_path, date_val, amt, type_inout, target_desc)

                if matched_file:
                    if is_applied:
                        apply_status = "🟢 반영 완료"
                        applied_count += 1
                    else:
                        apply_status = "🔴 미반영"
                        pending_count += 1
                else:
                    apply_status = "⚪ 대상 없음"
                    fail_count += 1

                total_amt += amt

                self.checked_vars[idx] = tk.BooleanVar(value=True)

                item_info = {
                    "seq": seq,
                    "date": date_val,
                    "type": type_inout,
                    "reason": reason,
                    "target_item": target_item,
                    "target_desc": target_desc,
                    "fund_code": fund_code,
                    "fund_name": fund_name,
                    "amount": amt,
                    "matched_file": matched_file,
                    "match_status": match_status,
                    "apply_status": apply_status,
                    "is_applied": is_applied
                }
                self.items_data.append(item_info)

                amt_formatted = f"{int(amt):,}"
                self.tree.insert("", "end", iid=str(idx), values=(
                    "☑", seq, fund_name, type_inout, reason, target_desc, amt_formatted, matched_file or "(자금판 파일 없음)", match_status, apply_status
                ))

            self.lbl_status.config(text=f"{disp_name} 로드 완료", foreground="#1f2937")
            self.lbl_summary.config(text=f"총 {len(self.items_data)}원 | 매칭: {matched_ok_count}건 | 미반영: {pending_count}건 | 반영완료: {applied_count}건 | 합계: {int(total_amt):,}원".replace("총 " + str(len(self.items_data)) + "원", f"총 {len(self.items_data)}건"))

            self.root.after(100, self.render_cell_badges)

        except Exception as e:
            messagebox.showerror("오류", f"지시내역 파일을 읽는 중 오류가 발생했습니다:\n{e}")
            self.lbl_status.config(text="파일 로드 실패", foreground="#dc2626")

    def on_double_click_cell(self, event):
        item_id = self.tree.identify_row(event.y)
        if not item_id:
            return
        
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return

        column = self.tree.identify_column(event.x)
        idx = int(item_id)
        item = self.items_data[idx]

        # Column #6 is target_desc
        if column == "#6":
            x, y, w, h = self.tree.bbox(item_id, column)
            current_val = self.tree.item(item_id, "values")[5]

            entry = ttk.Entry(self.tree)
            entry.place(x=x, y=y, width=w, height=h)
            entry.insert(0, current_val)
            entry.focus()

            def save_edit(event=None):
                new_val = entry.get().strip()
                if new_val:
                    vals = list(self.tree.item(item_id, "values"))
                    vals[5] = new_val
                    self.tree.item(item_id, values=vals)
                    self.items_data[idx]["target_desc"] = new_val
                entry.destroy()
                self.root.after(50, self.render_cell_badges)

            entry.bind("<Return>", save_edit)
            entry.bind("<FocusOut>", lambda e: (entry.destroy(), self.root.after(50, self.render_cell_badges)))

        # Column #8 is matched_file
        elif column == "#8":
            if item["matched_file"]:
                filepath = os.path.join(self.target_dir, item["matched_file"])
                if os.path.exists(filepath):
                    try:
                        os.startfile(filepath)
                    except Exception as err:
                        messagebox.showerror("오류", f"엑셀 파일을 열 수 없습니다:\n{err}")
                else:
                    messagebox.showerror("오류", f"파일을 찾을 수 없습니다:\n{filepath}")
            else:
                self.open_manual_match_dialog(idx)

        # Column #9 is match_status
        elif column == "#9":
            self.open_manual_match_dialog(idx)

        # Column #10 is apply_status
        elif column == "#10":
            self.toggle_single_item_apply(idx)

    def open_mapping_dialog(self):
        dlg = tk.Toplevel(self.root)
        dlg.title("사유 매핑 설정 (사유 · 대상기관/항목 · 적요)")
        dlg.geometry("820x570")
        dlg.transient(self.root)
        dlg.grab_set()
        explanation = (
            "사유는 필수이고 대상기관/항목은 선택입니다.\n"
            "대상기관/항목까지 일치하는 규칙을 먼저 적용하며, 없으면 같은 사유의 기본 규칙을 적용합니다."
        )
        ttk.Label(dlg, text=explanation, font=("Malgun Gothic", 10)).pack(padx=15, pady=10, anchor="w")

        frame_table = ttk.Frame(dlg, padding=(15, 0, 15, 10))
        frame_table.pack(fill="both", expand=True)
        tree_map = ttk.Treeview(
            frame_table,
            columns=("reason", "target_item", "description"),
            show="headings",
            selectmode="browse",
        )
        tree_map.heading("reason", text="사유")
        tree_map.heading("target_item", text="대상기관/항목 (선택)")
        tree_map.heading("description", text="자금판 적요")
        tree_map.column("reason", width=210)
        tree_map.column("target_item", width=235)
        tree_map.column("description", width=285)
        vsb_map = ttk.Scrollbar(frame_table, orient="vertical", command=tree_map.yview)
        tree_map.configure(yscrollcommand=vsb_map.set)
        tree_map.pack(side="left", fill="both", expand=True)
        vsb_map.pack(side="right", fill="y")

        def refresh_tree_map():
            for item_id in tree_map.get_children():
                tree_map.delete(item_id)
            for index, rule in enumerate(self.mapping_rules):
                tree_map.insert(
                    "",
                    "end",
                    iid=str(index),
                    values=(rule["reason"], rule["target_item"], rule["description"]),
                )

        edit_frame = ttk.LabelFrame(dlg, text=" 규칙 추가 / 수정 / 삭제 ", padding=15)
        edit_frame.pack(fill="x", padx=15, pady=(0, 15))
        ttk.Label(edit_frame, text="사유:").grid(row=0, column=0, sticky="e", padx=(0, 5), pady=5)
        entry_reason = ttk.Entry(edit_frame, width=26)
        entry_reason.grid(row=0, column=1, sticky="ew", padx=(0, 15), pady=5)
        ttk.Label(edit_frame, text="대상기관/항목:").grid(row=0, column=2, sticky="e", padx=(0, 5), pady=5)
        entry_target_item = ttk.Entry(edit_frame, width=26)
        entry_target_item.grid(row=0, column=3, sticky="ew", padx=(0, 5), pady=5)
        ttk.Label(edit_frame, text="자금판 적요:").grid(row=1, column=0, sticky="e", padx=(0, 5), pady=5)
        entry_description = ttk.Entry(edit_frame, width=60)
        entry_description.grid(row=1, column=1, columnspan=3, sticky="ew", padx=(0, 5), pady=5)
        edit_frame.columnconfigure(1, weight=1)
        edit_frame.columnconfigure(3, weight=1)

        def clear_inputs():
            entry_reason.delete(0, tk.END)
            entry_target_item.delete(0, tk.END)
            entry_description.delete(0, tk.END)

        def on_tree_select(event):
            selected_id = tree_map.focus()
            if not selected_id:
                return
            values = tree_map.item(selected_id, "values")
            if not values:
                return
            clear_inputs()
            entry_reason.insert(0, values[0])
            entry_target_item.insert(0, values[1])
            entry_description.insert(0, values[2])

        def save_rule():
            reason = entry_reason.get().strip()
            target_item = entry_target_item.get().strip()
            description = entry_description.get().strip()
            if not reason or not description:
                messagebox.showwarning("입력 필요", "사유와 자금판 적요는 반드시 입력하세요.", parent=dlg)
                return
            index = find_mapping_rule_index(self.mapping_rules, reason, target_item)
            rule = {"reason": reason, "target_item": target_item, "description": description}
            if index is None:
                self.mapping_rules.append(rule)
            else:
                self.mapping_rules[index] = rule
            self.mapping_rules = normalize_mapping_rules(self.mapping_rules)
            save_mapping_config(self.target_dir, self.mapping_rules)
            refresh_tree_map()
            clear_inputs()
            self.load_instruction_data()
            messagebox.showinfo("저장 완료", "매핑 규칙을 저장했습니다.", parent=dlg)

        def delete_rule():
            reason = entry_reason.get().strip()
            target_item = entry_target_item.get().strip()
            index = find_mapping_rule_index(self.mapping_rules, reason, target_item)
            if index is None:
                messagebox.showwarning("선택 필요", "삭제할 규칙을 표에서 선택하세요.", parent=dlg)
                return
            rule = self.mapping_rules[index]
            label = rule["reason"]
            if rule["target_item"]:
                label += f" / {rule['target_item']}"
            if not messagebox.askyesno("삭제 확인", f"'{label}' 규칙을 삭제할까요?", parent=dlg):
                return
            del self.mapping_rules[index]
            save_mapping_config(self.target_dir, self.mapping_rules)
            refresh_tree_map()
            clear_inputs()
            self.load_instruction_data()

        tree_map.bind("<<TreeviewSelect>>", on_tree_select)
        button_frame = ttk.Frame(edit_frame)
        button_frame.grid(row=2, column=0, columnspan=4, sticky="ew", pady=(10, 0))
        ttk.Button(button_frame, text="저장 / 수정", style="Primary.TButton", command=save_rule).pack(side="left", padx=5)
        ttk.Button(button_frame, text="선택 규칙 삭제", command=delete_rule).pack(side="left", padx=10)
        ttk.Button(button_frame, text="입력 지우기", command=clear_inputs).pack(side="left", padx=5)
        ttk.Button(button_frame, text="닫기", command=dlg.destroy).pack(side="right", padx=5)
        refresh_tree_map()

    def execute_update_selected(self):
        if not self.items_data:
            messagebox.showwarning("경고", "반영할 지시내역 데이터가 없습니다.")
            return

        selected_pending = []
        for idx, item in enumerate(self.items_data):
            if self.checked_vars.get(idx, tk.BooleanVar()).get():
                if item["matched_file"] and not item["is_applied"]:
                    selected_pending.append(item)

        if not selected_pending:
            messagebox.showwarning("알림", "체크 선택된 항목 중 미반영 상태인 항목이 없습니다.")
            return

        msg = f"체크 선택된 미반영 내역 총 {len(selected_pending)}건을 각 자금판 엑셀 파일에 일괄 반영하시겠습니까?"
        if not messagebox.askyesno("선택 항목 반영 확인", msg):
            return

        self._process_apply_items(selected_pending)

    def execute_update_all(self):
        if not self.items_data:
            messagebox.showwarning("경고", "반영할 지시내역 데이터가 없습니다.")
            return

        pending_items = [item for item in self.items_data if item["matched_file"] and not item["is_applied"]]

        if not pending_items:
            already_done = sum(1 for item in self.items_data if item["is_applied"])
            if already_done > 0:
                messagebox.showinfo("알림", f"모든 항목이 이미 자금판 엑셀에 반영 완료되었습니다. ({already_done}건)")
            else:
                messagebox.showerror("오류", "반영 가능한 자금판 엑셀 파일이 없습니다.")
            return

        msg = f"미반영 전체 항목 총 {len(pending_items)}건을 각 자금판 엑셀 파일에 일괄 반영하시겠습니까?"
        if not messagebox.askyesno("전체 반영 확인", msg):
            return

        self._process_apply_items(pending_items)

    def _process_apply_items(self, target_items):
        success_count = 0
        error_logs = []

        for item in target_items:
            filename = item["matched_file"]
            filepath = os.path.join(self.target_dir, filename)
            
            try:
                apply_single_item(filepath, item)
                success_count += 1
            except Exception as e:
                error_logs.append(f"{filename}: {e}")

        if error_logs:
            err_msg = "\n".join(error_logs)
            messagebox.showwarning("완료 (일부 오류)", f"총 {success_count}개 파일 성공적으로 반영 완료!\n\n[오류 목록]\n{err_msg}")
        else:
            messagebox.showinfo("성공", f"🎉 총 {success_count}개 자금판 엑셀 파일에 성공적으로 반영되었습니다!")

        self.load_instruction_data()

def build(parent):
    """통합 프로그램의 탭 안에 이 화면을 만든다."""
    return FundAutoUpdaterApp(parent)


def main():
    root = tk.Tk()
    app = FundAutoUpdaterApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
