# -*- coding: utf-8 -*-
"""
Pump 판가 계산기 (Python 이식판)
원본: 대외비_Pump_판가_계산기_이력추가_ver4_0.xlsm (판가계산기 / Sheet3 / Rawdata / 이력)

이식 원칙
---------
1. 원본 엑셀의 계산 로직(수식)을 그대로 유지한다.
     D3(재료비)          = 판가계산기!B5
     D4(노무비+경비)      = D3 * E4         (E4 = 노무비+경비 비율, 원본 하드코딩 0.2 -> 직접 입력 가능)
     D5(판관비)           = (D3+D4) * E5    (E5 = 판관비 비율,      원본 하드코딩 0.3 -> 직접 입력 가능)
     D6(영업이익)         = (D3+D4+D5) * E6 (E6 = 영업이익률,       원본부터 직접 입력 가능 = 판가계산기!C5)
     D7(적정 판가)        = D3+D4+D5+D6
2. '이력등록' 버튼 클릭 시 원본 VBA(Module1.bas RegisterHistory)와 동일한 항목을 입력받아
   '이력' 데이터에 동일한 규칙(No. 자동증가, 오늘 날짜, 계산금액 자동기록, 채택 현황='대기')으로 기록한다.
   (단, 입력 UX는 팝업 15개 연속 대신 한 화면 입력폼으로 개선)
3. Rawdata(과거 실적)는 조회 전용으로 제공한다 (수정 불가).
4. UI는 알파버전 대비 더 직관적이고 세련되게 재구성한다 (요청사항 반영).
"""

import os
import re
import csv
import json
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter import font as tkfont

# --------------------------------------------------------------------------------------
# 경로 / 상수
# --------------------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
RAWDATA_CSV = os.path.join(DATA_DIR, "rawdata.csv")
HISTORY_CSV = os.path.join(DATA_DIR, "history.csv")
SETTINGS_JSON = os.path.join(DATA_DIR, "settings.json")
HISTORY_EDIT_LOG = os.path.join(DATA_DIR, "history_edit_log.txt")

HISTORY_COLUMNS = [
    "No.", "일시", "제출여부", "지역", "고객구분", "사업부", "대공정", "세부공정",
    "장비사", "장비 모델", "Option", "재료비(원가)", "영업이익률", "계산금액", "제출금액",
    "채택 현황", "PUMP Model", "Pump FSC (As Is)", "Pump FSC (To Be)",
    "SEC 담당", "LOT 담당", "비고",
]

# 원본 VBA(Module1.bas)의 팝업 선택지와 동일
CHOICE_제출여부 = ["미제출", "제출예정", "제출완료"]
CHOICE_지역 = ["기흥", "화성", "평택", "미국", "중국", "ALL"]
CHOICE_고객구분 = ["연구소", "PCS", "기술팀", "그 외"]
CHOICE_사업부 = ["메모리", "파운드리", "그 외"]
CHOICE_대공정 = ["ALL", "GCS", "CLEAN", "CVD", "METAL", "IMP", "ETCH", "DIFF", "그 외"]
CHOICE_장비사 = ["GCS", "TES", "WONIK_IPS", "LAM", "ASM", "AMAT", "ULVAC",
               "EUGENETECH", "TEL", "SEMES", "KOKUSAI", "AXCELIS", "그 외"]

# 원본 VBA RegisterHistory 의 입력 순서와 동일한 순서로 정의
# type: "choice"(목록에서 선택) / "text"(자유 입력)
WIZARD_FIELDS = [
    ("v_progress", "제출여부", "choice", CHOICE_제출여부),
    ("v_region", "지역", "choice", CHOICE_지역),
    ("v_cust", "고객구분", "choice", CHOICE_고객구분),
    ("v_div", "사업부", "choice", CHOICE_사업부),
    ("v_process", "대공정", "choice", CHOICE_대공정),
    ("v_detail", "세부공정", "text", None),
    ("v_maker", "장비사", "choice", CHOICE_장비사),
    ("v_model", "장비 모델", "text", None),
    ("v_content", "Option", "text", None),
    ("v_submit", "제출금액", "text", None),
    ("v_modle", "PUMP Model", "text", None),
    ("v_fscAsIs", "Pump FSC (As Is)", "text", None),
    ("v_fscToBe", "Pump FSC (To Be)", "text", None),
    ("v_sec", "SEC 담당", "text", None),
    ("v_lot", "LOT 담당", "text", None),
    ("v_note", "비고", "text", None),
]

# 이력등록 마법사에서 "목록 선택"으로 입력되는 열 -> 더블클릭 수정 시에도 동일한 선택지를 보여준다.
HISTORY_CHOICE_MAP = {
    "제출여부": CHOICE_제출여부,
    "지역": CHOICE_지역,
    "고객구분": CHOICE_고객구분,
    "사업부": CHOICE_사업부,
    "대공정": CHOICE_대공정,
    "장비사": CHOICE_장비사,
}

DEFAULT_SETTINGS = {
    "material_cost": 1000000,      # 판가계산기!B5  재료비(원가)
    "profit_margin": 0.5,          # 판가계산기!C5  영업이익률
    "labor_expense_ratio": 0.2,    # Sheet3!E4      노무비+경비 비율  (직접 입력 가능하도록 변경)
    "sga_ratio": 0.3,              # Sheet3!E5      판관비 비율      (직접 입력 가능하도록 변경)
}

WON_FMT = "{:,.0f}"

# --------------------------------------------------------------------------------------
# 색상 / 폰트 (UI 테마)
# --------------------------------------------------------------------------------------
COLOR_BG = "#F3F5F9"
COLOR_CARD = "#FFFFFF"
COLOR_BORDER = "#E1E5EE"
COLOR_TEXT = "#20242C"
COLOR_SUBTEXT = "#6B7280"
COLOR_ACCENT = "#3B6FE0"
COLOR_ACCENT_DARK = "#2E58B8"
COLOR_ACCENT_SOFT = "#EAF1FF"
COLOR_HEADER_BG = "#EEF1F8"
COLOR_INPUT_BG = "#FFFDE9"

FONT_BASE = ("Malgun Gothic", 10)
FONT_BOLD = ("Malgun Gothic", 10, "bold")
FONT_TITLE = ("Malgun Gothic", 14, "bold")
FONT_SUBTITLE = ("Malgun Gothic", 11, "bold")
FONT_MONO = ("Consolas", 11)


def fmt_won(v):
    try:
        return WON_FMT.format(float(v))
    except (TypeError, ValueError):
        return str(v)


def fmt_pct(v):
    try:
        return f"{float(v) * 100:.1f}%"
    except (TypeError, ValueError):
        return str(v)


def _num_to_cell(v):
    """엑셀 셀처럼 정수는 정수로, 소수는 그대로 저장 (콤마 등 서식 문자는 넣지 않음)."""
    if v is None:
        return ""
    if float(v).is_integer():
        return int(v)
    return round(float(v), 6)


def parse_number(s, default=0.0):
    """콤마, 공백, 줄바꿈, '원', '%' 등이 섞여 있어도 숫자만 파싱. 실패 시 default 반환."""
    if s is None:
        return default
    text = str(s).replace("\xa0", " ").strip()
    text = text.replace(",", "").replace("원", "").replace("%", "").replace(" ", "")
    if text == "":
        return default
    try:
        return float(text)
    except ValueError:
        return default


# --------------------------------------------------------------------------------------
# 텍스트 정리 (줄바꿈 -> 공백)
# --------------------------------------------------------------------------------------
_NEWLINE_RE = re.compile(r"[\r\n]+")
_MULTI_SPACE_RE = re.compile(r"[ \t ]+")


def clean_text(value):
    """셀 안에 줄바꿈이 있으면 공백으로 치환하고 연속 공백을 정리한다."""
    if value is None:
        return value
    s = str(value).replace("\xa0", " ")
    s = _NEWLINE_RE.sub(" ", s)
    s = _MULTI_SPACE_RE.sub(" ", s)
    return s.strip()


# --------------------------------------------------------------------------------------
# 열(컬럼) 종류 판별: 금액 / 비율(%) / 일반 텍스트
# --------------------------------------------------------------------------------------
MONEY_KEYWORDS = ("재료비", "노무비", "경비", "원가", "판관비", "판가", "금액")
PERCENT_KEYWORDS = ("퍼센트", "이익률")


_PAREN_RE = re.compile(r"\([^)]*\)")


def classify_column(header):
    h = (header or "").replace("\n", " ")
    # 괄호 안 설명("BEP 판가 (영업이익률 20%)"의 "영업이익률 20%" 등)은 열 자체의 성격이
    # 아니라 산정 근거 설명이므로 분류 판단에서 제외한다.
    core = _PAREN_RE.sub(" ", h)
    if any(k in core for k in PERCENT_KEYWORDS):
        return "percent"
    if any(k in core for k in MONEY_KEYWORDS):
        return "money"
    return "text"


def compute_default_widths(headers, raw_rows, sample=300, min_w=70, max_w=260, base_pad=28, char_w=9):
    """헤더/내용 길이를 기준으로 열 기본 너비를 계산한다 (새로고침 시 이 값으로 초기화)."""
    widths = []
    for i, h in enumerate(headers):
        header_text = (h or "").replace("\n", " ")
        max_len = len(header_text)
        for row in raw_rows[:sample]:
            if i < len(row) and row[i]:
                max_len = max(max_len, len(str(row[i])))
        widths.append(min(max_w, max(min_w, base_pad + max_len * char_w)))
    return widths


# --------------------------------------------------------------------------------------
# 데이터 계층
# --------------------------------------------------------------------------------------
def load_settings():
    settings = dict(DEFAULT_SETTINGS)
    if os.path.exists(SETTINGS_JSON):
        try:
            with open(SETTINGS_JSON, "r", encoding="utf-8") as f:
                settings.update(json.load(f))
        except (json.JSONDecodeError, OSError):
            pass
    return settings


def save_settings(settings):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(SETTINGS_JSON, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)



# CSV_ENCODING: "utf-8-sig"(BOM 포함)로 읽고 쓴다. BOM이 없는 순수 UTF-8 CSV를 엑셀에서
# 직접 열면 한글이 깨져 보이는데(엑셀이 시스템 기본 코드페이지로 잘못 해석), BOM을 붙이면
# 엑셀이 UTF-8로 정확히 인식한다. utf-8-sig로 읽으면 BOM 유무와 상관없이 항상 올바르게 읽힌다.
CSV_ENCODING = "utf-8-sig"


def load_csv_header(path):
    if not os.path.exists(path):
        return []
    with open(path, "r", newline="", encoding=CSV_ENCODING) as f:
        reader = csv.reader(f)
        try:
            return next(reader)
        except StopIteration:
            return []


def load_csv_rows(path, columns):
    if not os.path.exists(path):
        return []
    with open(path, "r", newline="", encoding=CSV_ENCODING) as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return []
    return rows[1:]  # 헤더 제외


def save_csv_rows(path, columns, rows):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(path, "w", newline="", encoding=CSV_ENCODING) as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        writer.writerows(rows)


def ensure_utf8_bom(path):
    """기존에 BOM 없이 저장된 CSV 파일을 엑셀에서 열어도 한글이 깨지지 않도록,
    내용은 그대로 두고 UTF-8 BOM만 추가해 둔다."""
    if not os.path.exists(path):
        return
    with open(path, "rb") as f:
        raw = f.read()
    if raw.startswith(b"\xef\xbb\xbf") or not raw:
        return
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError:
        return
    with open(path, "w", newline="", encoding=CSV_ENCODING) as f:
        f.write(text)


def clean_history_csv_if_needed():
    """이력 CSV에 남아있는 줄바꿈을 공백으로 정리해서 저장해 둔다 (요청사항 8)."""
    rows = load_csv_rows(HISTORY_CSV, HISTORY_COLUMNS)
    if not rows:
        return
    changed = False
    cleaned_rows = []
    for row in rows:
        new_row = [clean_text(v) for v in row]
        if new_row != row:
            changed = True
        cleaned_rows.append(new_row)
    if changed:
        save_csv_rows(HISTORY_CSV, HISTORY_COLUMNS, cleaned_rows)


def append_history_edit_log(no_value, context_label, header, old_display, new_display):
    """이력 셀 수정 이력을 텍스트 파일로 남긴다: 시간 + 셀위치(No./열) + 값변경(A→B)."""
    os.makedirs(DATA_DIR, exist_ok=True)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    location = f"No.{no_value}" if no_value not in (None, "") else "No.(미확인)"
    if context_label:
        location += f" ({context_label})"
    old_text = old_display if old_display not in (None, "") else "(빈 값)"
    new_text = new_display if new_display not in (None, "") else "(빈 값)"
    line = f"{timestamp} | {location} | {header}: {old_text} → {new_text}\n"
    with open(HISTORY_EDIT_LOG, "a", encoding="utf-8-sig") as f:
        f.write(line)


# --------------------------------------------------------------------------------------
# 계산 로직 (Sheet3 수식과 1:1 대응)
# --------------------------------------------------------------------------------------
def compute_price(material_cost, labor_ratio, sga_ratio, profit_margin):
    d3 = material_cost                              # 재료비
    d4 = d3 * labor_ratio                            # 노무비+경비
    d5 = (d3 + d4) * sga_ratio                        # 판관비
    d6 = (d3 + d4 + d5) * profit_margin               # 영업이익
    d7 = d3 + d4 + d5 + d6                            # 적정 판가
    return {"재료비": d3, "노무비+경비": d4, "판관비": d5, "영업이익": d6, "적정판가": d7}


# --------------------------------------------------------------------------------------
# 창 유틸
# --------------------------------------------------------------------------------------
def maximize_window(win):
    """새 창을 화면 전체 크기로 띄운다 (요청사항 2)."""
    win.update_idletasks()
    # 창관리자가 없거나 "zoomed"/"-zoomed"가 무시되는 환경에서도 항상 화면 크기를
    # 보장하도록, 실제 화면 크기로 우선 지정해 둔 뒤 OS 기본 최대화를 추가로 시도한다.
    sw = win.winfo_screenwidth()
    sh = win.winfo_screenheight()
    win.geometry(f"{sw}x{sh}+0+0")
    try:
        win.state("zoomed")
        return
    except tk.TclError:
        pass
    try:
        win.attributes("-zoomed", True)
    except tk.TclError:
        pass


# --------------------------------------------------------------------------------------
# 이력등록 입력폼 (원본 VBA 팝업 15개 연속 -> 한 화면 입력폼으로 개선, 요청사항 3)
# --------------------------------------------------------------------------------------
class HistoryRegisterDialog(tk.Toplevel):
    def __init__(self, parent, summary_text):
        super().__init__(parent)
        self.title("이력등록")
        self.configure(bg=COLOR_BG)
        self.resizable(False, False)
        self.result = None
        self.transient(parent)
        self.grab_set()

        tk.Label(self, text="이력등록 정보 입력", font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_TEXT).pack(
            anchor="w", padx=24, pady=(22, 4))
        tk.Label(self, text=summary_text, font=FONT_BASE, bg=COLOR_BG, fg=COLOR_SUBTEXT,
                 justify="left").pack(anchor="w", padx=24, pady=(0, 14))

        grid = tk.Frame(self, bg=COLOR_BG)
        grid.pack(padx=24)

        self._vars = {}
        pairs = [WIZARD_FIELDS[i:i + 2] for i in range(0, len(WIZARD_FIELDS), 2)]
        for r, pair in enumerate(pairs):
            for c, field in enumerate(pair):
                key, label, kind, choices = field
                cell = tk.Frame(grid, bg=COLOR_BG)
                cell.grid(row=r, column=c, padx=10, pady=6, sticky="w")
                tk.Label(cell, text=label, font=FONT_BOLD, bg=COLOR_BG, fg=COLOR_TEXT).pack(anchor="w")
                if kind == "choice":
                    var = tk.StringVar(value=choices[0])
                    combo = ttk.Combobox(cell, textvariable=var, values=choices, state="readonly",
                                          width=21, font=FONT_BASE)
                    combo.pack(anchor="w", pady=(3, 0))
                else:
                    var = tk.StringVar()
                    entry = tk.Entry(cell, textvariable=var, width=24, font=FONT_BASE,
                                      relief="solid", bd=1, highlightthickness=0)
                    entry.pack(anchor="w", pady=(3, 0), ipady=3)
                self._vars[key] = var

        btns = tk.Frame(self, bg=COLOR_BG)
        btns.pack(pady=(6, 22))
        ttk.Button(btns, text="등록", style="Accent.TButton", command=self._ok).pack(side="left", padx=6)
        ttk.Button(btns, text="취소", command=self._cancel).pack(side="left", padx=6)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        x = parent.winfo_rootx() + 60
        y = parent.winfo_rooty() + 40
        self.geometry(f"+{x}+{y}")
        self.wait_window(self)

    def _ok(self):
        self.result = {k: clean_text(v.get()).strip() for k, v in self._vars.items()}
        self.destroy()

    def _cancel(self):
        self.result = None
        self.destroy()


# --------------------------------------------------------------------------------------
# 셀 내용이 열 너비를 넘칠 때 마우스를 올리면 전체 내용을 보여주는 툴팁 (요청사항)
# --------------------------------------------------------------------------------------
class _CellTooltip:
    def __init__(self):
        self._win = None

    def show(self, widget, x, y, text):
        self.hide()
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        try:
            win.attributes("-topmost", True)
        except tk.TclError:
            pass
        win.wm_geometry(f"+{x}+{y}")
        tk.Label(
            win, text=text, justify="left", bg="#333333", fg="white",
            font=("Malgun Gothic", 9), padx=8, pady=5, wraplength=420,
            relief="solid", bd=1,
        ).pack()
        self._win = win

    def hide(self):
        if self._win is not None:
            try:
                self._win.destroy()
            except tk.TclError:
                pass
            self._win = None


# --------------------------------------------------------------------------------------
# 셀 표시(금액/비율 서식) + 더블클릭 편집(선택형/텍스트)이 가능한 Treeview
# --------------------------------------------------------------------------------------
class DataTreeview(ttk.Treeview):
    def __init__(self, master, headers, *, editable=False, choice_map=None,
                 on_change=None, on_cell_change=None, id_prefix="c", **kwargs):
        self.headers = list(headers)
        col_ids = [f"{id_prefix}{i}" for i in range(len(self.headers))]
        super().__init__(master, columns=col_ids, show="headings", **kwargs)
        self.col_ids = col_ids
        self.editable = editable
        self.choice_map = choice_map or {}
        self.on_change = on_change
        self.on_cell_change = on_cell_change  # (row_index, header, old_value, new_value)
        self._raw_rows = []
        self._default_widths = [110] * len(self.headers)

        for cid, header in zip(self.col_ids, self.headers):
            self.heading(cid, text=header.replace("\n", " "))

        self._edit_widget = None
        if self.editable:
            self.bind("<Double-1>", self._begin_edit)

        # 열 너비를 넘치는 셀에 마우스를 올리면 전체 내용을 툴팁으로 보여준다 (요청사항).
        self._tooltip = _CellTooltip()
        self._hover_cell = None
        self._hover_after_id = None
        self._tooltip_font = tkfont.Font(font=("Malgun Gothic", 9))
        self.bind("<Motion>", self._on_motion)
        self.bind("<Leave>", self._on_leave)
        self.bind("<ButtonPress>", lambda e: self._cancel_hover())
        self.bind("<MouseWheel>", lambda e: self._cancel_hover(), add="+")

    # ---------------------------------------------------------------- 데이터 적재
    def set_rows(self, rows):
        width = len(self.headers)
        self._raw_rows = [(list(r) + [""] * width)[:width] for r in rows]
        for row in self._raw_rows:
            for i, v in enumerate(row):
                row[i] = clean_text(v)
        self._default_widths = compute_default_widths(self.headers, self._raw_rows)
        self.apply_default_column_widths()
        self.apply_filter("")

    def get_raw_rows(self):
        return [list(r) for r in self._raw_rows]

    # ---------------------------------------------------------------- 열 너비 (요청사항 9, 10)
    def apply_default_column_widths(self):
        for cid, header, width in zip(self.col_ids, self.headers, self._default_widths):
            kind = classify_column(header)
            anchor = "e" if kind in ("money", "percent") else "w"
            # stretch=False: 한 열을 늘리거나 줄여도 옆 열이 밀리거나 당겨지지 않고
            # 가로 스크롤로 처리된다.
            self.column(cid, width=width, minwidth=40, stretch=False, anchor=anchor)

    # ---------------------------------------------------------------- 서식(콤마/퍼센트)
    def _format_cell(self, header, value):
        if value in (None, ""):
            return ""
        kind = classify_column(header)
        if kind == "money":
            n = parse_number(value, None)
            return fmt_won(n) if n is not None else value
        if kind == "percent":
            n = parse_number(value, None)
            return fmt_pct(n) if n is not None else value
        return value

    def format_cell(self, header, value):
        """_format_cell의 공개 버전 (수정 로그 등 외부에서 표시용 서식이 필요할 때 사용)."""
        return self._format_cell(header, value)

    # ---------------------------------------------------------------- 검색 필터 (요청사항 7)
    def apply_filter(self, keyword):
        self._close_editor()
        self.delete(*self.get_children())
        kw = keyword.strip().lower()
        for idx, row in enumerate(self._raw_rows):
            if kw:
                display_row = [self._format_cell(h, v) for h, v in zip(self.headers, row)]
                haystack = " ".join(str(v).lower() for v in row) + " " + " ".join(str(v).lower() for v in display_row)
                if kw not in haystack:
                    continue
            else:
                display_row = [self._format_cell(h, v) for h, v in zip(self.headers, row)]
            self.insert("", tk.END, iid=str(idx), values=display_row)
        return len(self.get_children())

    # ---------------------------------------------------------------- 셀 툴팁 (요청사항)
    def _on_motion(self, event):
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            self._cancel_hover()
            return
        row_id = self.identify_row(event.y)
        col_id = self.identify_column(event.x)
        if not row_id or not col_id:
            self._cancel_hover()
            return
        cell_key = (row_id, col_id)
        if cell_key == self._hover_cell:
            return
        self._cancel_hover()
        self._hover_cell = cell_key
        x_root, y_root = event.x_root, event.y_root
        self._hover_after_id = self.after(450, lambda: self._maybe_show_tooltip(cell_key, x_root, y_root))

    def _on_leave(self, _event=None):
        self._cancel_hover()

    def _cancel_hover(self):
        if self._hover_after_id is not None:
            try:
                self.after_cancel(self._hover_after_id)
            except tk.TclError:
                pass
            self._hover_after_id = None
        self._hover_cell = None
        self._tooltip.hide()

    def _maybe_show_tooltip(self, cell_key, x_root, y_root):
        if cell_key != self._hover_cell:
            return
        row_id, col_id = cell_key
        if not self.exists(row_id):
            return
        col_index = int(col_id.replace("#", "")) - 1
        if col_index < 0 or col_index >= len(self.headers):
            return
        row_index = int(row_id)
        if row_index >= len(self._raw_rows):
            return
        header = self.headers[col_index]
        display_text = self._format_cell(header, self._raw_rows[row_index][col_index])
        if not display_text:
            return
        col_width = self.column(col_id, "width")
        text_width = self._tooltip_font.measure(str(display_text)) + 14
        if text_width <= col_width:
            return  # 잘리지 않고 다 보이는 셀은 툴팁을 띄우지 않는다.
        self._tooltip.show(self, x_root + 12, y_root + 18, str(display_text))

    # ---------------------------------------------------------------- 편집 (요청사항 4, 5, 6, 11)
    def _close_editor(self):
        if self._edit_widget is not None:
            try:
                self._edit_widget.destroy()
            except tk.TclError:
                pass
            self._edit_widget = None

    def _begin_edit(self, event):
        self._cancel_hover()
        if not self.editable:
            return
        if self.identify("region", event.x, event.y) != "cell":
            return
        row_id = self.identify_row(event.y)
        col_id = self.identify_column(event.x)
        if not row_id or not col_id:
            return
        self._close_editor()
        col_index = int(col_id.replace("#", "")) - 1
        header = self.headers[col_index]
        row_index = int(row_id)
        raw_value = self._raw_rows[row_index][col_index]
        bbox = self.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox

        if header in self.choice_map:
            self._edit_choice(row_id, col_id, row_index, col_index, header, raw_value, x, y, w, h)
        else:
            self._edit_text(row_id, col_id, row_index, col_index, header, raw_value, x, y, w, h)

    def _commit(self, row_id, col_id, row_index, col_index, header, stored_value):
        old_value = self._raw_rows[row_index][col_index]
        self._raw_rows[row_index][col_index] = stored_value
        self.set(row_id, col_id, self._format_cell(header, stored_value))
        if self.on_cell_change and str(old_value) != str(stored_value):
            self.on_cell_change(row_index, header, old_value, stored_value)
        if self.on_change:
            self.on_change()

    def _edit_choice(self, row_id, col_id, row_index, col_index, header, raw_value, x, y, w, h):
        # 이력등록 CLICK 시 목록에서 선택했던 열은, 더블클릭 수정 시에도 동일한 선택지를 보여준다.
        choices = self.choice_map[header]
        var = tk.StringVar(value=str(raw_value) if raw_value else choices[0])
        combo = ttk.Combobox(self, textvariable=var, values=choices, state="readonly", font=("Malgun Gothic", 9))
        combo.place(x=x, y=y, width=w, height=h)
        self._edit_widget = combo
        combo.focus_set()

        def commit(_e=None):
            self._commit(row_id, col_id, row_index, col_index, header, var.get())
            self._close_editor()

        combo.bind("<<ComboboxSelected>>", commit)
        combo.bind("<FocusOut>", lambda e: self._close_editor())
        combo.bind("<Escape>", lambda e: self._close_editor())

    def _edit_text(self, row_id, col_id, row_index, col_index, header, raw_value, x, y, w, h):
        kind = classify_column(header)
        if kind == "percent":
            n = parse_number(raw_value, None)
            edit_value = f"{n * 100:.0f}" if n is not None else ""
        else:
            edit_value = "" if raw_value is None else str(raw_value)

        entry = tk.Entry(self, font=("Malgun Gothic", 9))
        entry.insert(0, edit_value)
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        self._edit_widget = entry

        def commit(_e=None):
            new_text = clean_text(entry.get())
            if kind == "percent":
                n = parse_number(new_text, None)
                stored = _num_to_cell(n / 100.0) if n is not None else new_text
            elif kind == "money":
                n = parse_number(new_text, None)
                stored = _num_to_cell(n) if n is not None else new_text
            else:
                stored = new_text
            self._commit(row_id, col_id, row_index, col_index, header, stored)
            self._close_editor()

        def cancel(_e=None):
            self._close_editor()

        entry.bind("<Return>", commit)
        entry.bind("<FocusOut>", commit)
        entry.bind("<Escape>", cancel)


# --------------------------------------------------------------------------------------
# 메인 애플리케이션
# --------------------------------------------------------------------------------------
class PumpPriceApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pump 판가 계산기 (Python 이식판) - 대외비")
        self.geometry("1360x660")
        self.minsize(1000, 480)
        self.configure(bg=COLOR_BG)

        ensure_utf8_bom(RAWDATA_CSV)
        ensure_utf8_bom(HISTORY_CSV)
        clean_history_csv_if_needed()

        self.settings = load_settings()

        self.material_cost_var = tk.StringVar(value=fmt_won(self.settings["material_cost"]))
        self.profit_margin_var = tk.StringVar(value=f"{self.settings['profit_margin']*100:.0f}")
        self.labor_ratio_var = tk.StringVar(value=f"{self.settings['labor_expense_ratio']*100:.0f}")
        self.sga_ratio_var = tk.StringVar(value=f"{self.settings['sga_ratio']*100:.0f}")

        self.calc_labels = {}  # 원가구조(구 Sheet3) 계산결과 표시용 라벨 저장

        self.rawdata_headers = load_csv_header(RAWDATA_CSV)

        self._rawdata_win = None
        self._rawdata_tree = None
        self._rawdata_search_var = None
        self._rawdata_count_label = None
        self._history_win = None
        self._history_tree = None
        self._history_search_var = None
        self._history_count_label = None

        self._setup_style()
        self._build_main()
        self._recalc()

    # ---------------------------------------------------------------- 스타일 (요청사항 12)
    def _setup_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(".", background=COLOR_BG, font=FONT_BASE)
        style.configure("TFrame", background=COLOR_BG)
        style.configure("Card.TFrame", background=COLOR_CARD)

        style.configure("Treeview", font=("Malgun Gothic", 9), rowheight=24,
                         background=COLOR_CARD, fieldbackground=COLOR_CARD, borderwidth=0)
        style.configure("Treeview.Heading", font=("Malgun Gothic", 9, "bold"),
                         background=COLOR_HEADER_BG, foreground=COLOR_TEXT, relief="flat")
        style.map("Treeview", background=[("selected", COLOR_ACCENT_SOFT)],
                  foreground=[("selected", COLOR_TEXT)])

        style.configure("TButton", font=FONT_BASE, padding=(12, 7))
        style.map("TButton", background=[("active", "#E4E8F1")])

        style.configure("Accent.TButton", font=FONT_BOLD, padding=(16, 9),
                         background=COLOR_ACCENT, foreground="white")
        style.map("Accent.TButton",
                  background=[("active", COLOR_ACCENT_DARK), ("pressed", COLOR_ACCENT_DARK)],
                  foreground=[("disabled", "#AAAAAA")])

        style.configure("TCombobox", padding=4)
        style.configure("TNotebook", background=COLOR_BG, borderwidth=0)

    def _card(self, master, width=None, **pack_kwargs):
        outer = tk.Frame(master, bg=COLOR_BORDER)
        if width is not None:
            outer.configure(width=width)
            outer.pack_propagate(False)
        card = tk.Frame(outer, bg=COLOR_CARD)
        card.pack(fill="both", expand=True, padx=1, pady=1)
        if pack_kwargs:
            outer.pack(**pack_kwargs)
        return outer, card

    # ---------------------------------------------------------------- 전체 레이아웃
    def _build_main(self):
        # 상단 안내 배너
        banner = tk.Frame(self, bg=COLOR_ACCENT_SOFT)
        banner.pack(fill="x", padx=16, pady=(16, 10))
        tk.Label(
            banner, text="재료비(원가) / 영업이익률을 입력하면 자동 계산됩니다  →  이력등록 CLICK!",
            bg=COLOR_ACCENT_SOFT, fg=COLOR_ACCENT_DARK, font=FONT_SUBTITLE, pady=12, padx=16,
            anchor="w",
        ).pack(fill="x")

        # 상단 툴바 (Rawdata / 이력 새창 버튼) - 요청사항 2. 아래 본문이 내용에 맞춰
        # 왼쪽으로 타이트하게 배치되므로, 버튼도 오른쪽이 아닌 왼쪽에 맞춘다 (요청사항).
        toolbar_wrap = tk.Frame(self, bg=COLOR_BG)
        toolbar_wrap.pack(fill="x", padx=16, pady=(0, 12))

        toolbar = tk.Frame(toolbar_wrap, bg=COLOR_BG)
        toolbar.pack(fill="x")
        ttk.Button(toolbar, text="📄  PUMP 판가 DATA 보기", command=self._open_rawdata_window).pack(side="left")
        ttk.Button(toolbar, text="🕒  이력 보기", command=self._open_history_window).pack(side="left", padx=(8, 0))

        tk.Label(
            toolbar_wrap, text="※ PUMP 판가 DATA / 이력은 위 버튼을 눌러 새 창에서 확인합니다.",
            font=("Malgun Gothic", 9), fg=COLOR_SUBTEXT, bg=COLOR_BG, anchor="w",
        ).pack(fill="x", pady=(4, 0))

        body = tk.Frame(self, bg=COLOR_BG)
        body.pack(fill="both", expand=True, padx=16, pady=(0, 16))

        # 두 패널 모두 고정폭으로 늘리지 않고 내용에 맞춰 타이트하게 감싸서, 원가 구조
        # 계산 패널 오른쪽에 불필요한 흰 여백이 남지 않게 한다 (요청사항).
        body.columnconfigure(0, weight=0)
        body.columnconfigure(1, weight=0)
        body.rowconfigure(0, weight=1)

        self._build_calc_panel(body).grid(row=0, column=0, sticky="ns", padx=(0, 8))
        self._build_cost_panel(body).grid(row=0, column=1, sticky="ns", padx=(8, 0))

    # ---------------------------------------------------------------- 좌측: 판가계산기
    def _build_calc_panel(self, master):
        # 고정 너비를 주지 않고 내용(표+버튼)에 맞춰 패널이 저절로 타이트하게 감싸도록 한다
        # (요청사항: 표~버튼 너비와 계산 결과 상자 너비를 맞추고 남는 여백을 없앤다).
        outer, card = self._card(master)

        tk.Label(card, text="판가계산기", font=FONT_TITLE, bg=COLOR_CARD, fg=COLOR_TEXT).pack(
            anchor="w", padx=20, pady=(20, 14))

        grid = tk.Frame(card, bg=COLOR_CARD)
        grid.pack(padx=20, pady=(0, 20), anchor="w")

        header_style = dict(font=FONT_BOLD, bg=COLOR_HEADER_BG, fg=COLOR_TEXT,
                             relief="flat", width=12, height=2)
        tk.Label(grid, text="재료비(원가)", **header_style).grid(row=0, column=0, padx=(0, 1), pady=(0, 1), sticky="nsew")
        tk.Label(grid, text="영업이익률", **header_style).grid(row=0, column=1, pady=(0, 1), sticky="nsew")

        e1 = tk.Entry(grid, textvariable=self.material_cost_var, font=("Malgun Gothic", 12),
                      relief="solid", bd=1, highlightthickness=0, width=12, justify="right")
        e1.grid(row=1, column=0, ipady=8, padx=(0, 1))
        e1.bind("<FocusOut>", lambda e: self._on_material_cost_changed())
        e1.bind("<Return>", lambda e: self._on_material_cost_changed())
        # 입력하는 동안에도(글자마다) 즉시 재계산되도록 trace를 건다. 콤마 서식은
        # 포커스를 벗어날 때만(_on_material_cost_changed) 적용해 타이핑을 방해하지 않는다.
        self.material_cost_var.trace_add("write", lambda *a: self._on_material_cost_live())

        pct_frame = tk.Frame(grid, bg=COLOR_CARD, highlightbackground="#C9CEDA",
                              highlightthickness=1, bd=0)
        pct_frame.grid(row=1, column=1, ipady=8, sticky="nsew")
        e2 = tk.Entry(pct_frame, textvariable=self.profit_margin_var, font=("Malgun Gothic", 12),
                      relief="flat", bd=0, width=7, justify="right")
        e2.pack(side="left", padx=(8, 0), fill="y")
        tk.Label(pct_frame, text="%", font=("Malgun Gothic", 12), bg=COLOR_CARD).pack(side="left", padx=(2, 8))
        e2.bind("<FocusOut>", lambda e: self._on_profit_margin_changed())
        e2.bind("<Return>", lambda e: self._on_profit_margin_changed())
        self.profit_margin_var.trace_add("write", lambda *a: self._on_profit_margin_live())

        # 이력등록 버튼을 재료비/영업이익률 표 오른쪽에 붙여 배치한다 (요청사항).
        register_btn = ttk.Button(
            grid, text="이력등록\nCLICK!", style="Accent.TButton", command=self._on_register_history,
        )
        register_btn.grid(row=0, column=2, rowspan=2, sticky="nsew", padx=(12, 0))

        # 계산 결과 상자를 표/버튼과 같은 grid의 3개 열에 걸쳐 배치해서, 위쪽 표~버튼의
        # 총 너비와 정확히 같은 너비를 갖도록 한다 (요청사항).
        result_outer = tk.Frame(grid, bg=COLOR_HEADER_BG)
        result_outer.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(20, 0))
        tk.Label(result_outer, text="계산 결과", font=FONT_SUBTITLE, bg=COLOR_HEADER_BG,
                 fg=COLOR_TEXT, anchor="w").pack(fill="x", padx=16, pady=(14, 4))
        self.summary_label = tk.Label(result_outer, text="", justify="left", anchor="w",
                                       font=FONT_MONO, bg=COLOR_HEADER_BG, fg=COLOR_TEXT)
        self.summary_label.pack(fill="x", padx=16, pady=(0, 16))

        return outer

    def _on_material_cost_live(self):
        # 타이핑 중(엔터/포커스아웃 전)에도 거의 실시간으로 재계산만 반영한다 (요청사항).
        val = parse_number(self.material_cost_var.get(), None)
        if val is None:
            return
        self.settings["material_cost"] = val
        self._recalc()

    def _on_material_cost_changed(self):
        val = parse_number(self.material_cost_var.get(), self.settings["material_cost"])
        self.settings["material_cost"] = val
        self.material_cost_var.set(fmt_won(val))
        self._recalc()

    def _on_profit_margin_live(self):
        val = parse_number(self.profit_margin_var.get(), None)
        if val is None:
            return
        self.settings["profit_margin"] = val / 100.0
        self._recalc()

    def _on_profit_margin_changed(self):
        val = parse_number(self.profit_margin_var.get(), self.settings["profit_margin"] * 100) / 100.0
        self.settings["profit_margin"] = val
        self.profit_margin_var.set(f"{val*100:.0f}")
        self._recalc()

    # ---------------------------------------------------------------- 우측: 원가 구조 (구 Sheet3, 요청사항 1)
    def _build_cost_panel(self, master):
        outer, card = self._card(master)

        tk.Label(card, text="원가 구조 계산", font=FONT_TITLE, bg=COLOR_CARD, fg=COLOR_TEXT).pack(
            anchor="w", padx=20, pady=(20, 4))
        tk.Label(card, text="비율(노란 칸)은 직접 입력할 수 있습니다.", font=("Malgun Gothic", 9),
                 bg=COLOR_CARD, fg=COLOR_SUBTEXT).pack(anchor="w", padx=20, pady=(0, 14))

        grid = tk.Frame(card, bg=COLOR_CARD)
        grid.pack(padx=20, pady=(0, 8), fill="x")

        # 초기화 버튼을 "기본값 산정 근거" 열(마지막 열) 바로 위에 배치한다 (요청사항).
        ttk.Button(grid, text="초기화", command=self._on_reset_cost_ratios).grid(
            row=0, column=5, sticky="e", pady=(0, 4))

        headers = ["분류", "구분", "금액", "비율", "비율 상세", "기본값 산정 근거"]
        widths = [7, 9, 12, 7, 14, 15]
        for c, h in enumerate(headers):
            tk.Label(grid, text=h, font=("Malgun Gothic", 9, "bold"), bg=COLOR_HEADER_BG, fg=COLOR_TEXT,
                     relief="flat", width=widths[c], height=2).grid(
                row=1, column=c, sticky="nsew", padx=(0 if c == 0 else 1, 0), pady=(0, 1))

        def cell(row, col, text, editable_var=None, bold=False):
            width = widths[col]
            if editable_var is not None:
                e = tk.Entry(grid, textvariable=editable_var, width=width, justify="center",
                             relief="solid", bd=1, highlightthickness=0, font=("Malgun Gothic", 10),
                             bg=COLOR_INPUT_BG)
                e.grid(row=row, column=col, sticky="nsew", ipady=5, padx=(0 if col == 0 else 1, 0), pady=1)
                return e
            # 분류/구분/비율처럼 짧은 열은 줄바꿈 없이 한 줄로, 비율 상세/근거처럼 긴
            # 설명이 들어가는 열만 줄바꿈해 어색하게 단어 중간에서 잘리지 않게 한다.
            wrap = width * 9 if col in (4, 5) else 0
            # 표의 모든 값은 가운데 정렬한다 (요청사항).
            lbl = tk.Label(grid, text=text, width=width, relief="flat",
                            font=("Malgun Gothic", 10, "bold" if bold else "normal"),
                            bg=COLOR_CARD if not bold else COLOR_HEADER_BG, fg=COLOR_TEXT,
                            wraplength=wrap, justify="center", anchor="center", padx=6)
            lbl.grid(row=row, column=col, sticky="nsew", ipady=5, padx=(0 if col == 0 else 1, 0), pady=1)
            return lbl

        # Row2: 원가 / 재료비
        cell(2, 0, "원가")
        cell(2, 1, "재료비")
        self.calc_labels["재료비_금액"] = cell(2, 2, "-")
        cell(2, 3, "-")
        cell(2, 4, "연구소 및 구매팀")
        cell(2, 5, "-")

        # Row3: 노무비+경비 (비율 EDITABLE)
        cell(3, 0, "")
        cell(3, 1, "노무비+경비")
        self.calc_labels["노무비_금액"] = cell(3, 2, "-")
        cell(3, 3, "", editable_var=self.labor_ratio_var)
        self.calc_labels["노무비_상세"] = cell(3, 4, "-")
        cell(3, 5, "PUMP 판가 DATA 기준 평균값")

        # Row4: 판관비 (비율 EDITABLE)
        cell(4, 0, "판매관리비")
        cell(4, 1, "판관비")
        self.calc_labels["판관비_금액"] = cell(4, 2, "-")
        cell(4, 3, "", editable_var=self.sga_ratio_var)
        self.calc_labels["판관비_상세"] = cell(4, 4, "-")
        cell(4, 5, "제조업 평균값")

        # Row5: 영업이익 (비율은 판가계산기 입력값과 연동, 읽기전용)
        cell(5, 0, "이익")
        cell(5, 1, "영업이익")
        self.calc_labels["영업이익_금액"] = cell(5, 2, "-")
        self.calc_labels["영업이익_비율"] = cell(5, 3, "-")
        self.calc_labels["영업이익_상세"] = cell(5, 4, "-")
        cell(5, 5, "-")

        # Row6: 적정 판가
        cell(6, 0, "적정 판가", bold=True)
        cell(6, 1, "", bold=True)
        self.calc_labels["적정판가"] = cell(6, 2, "-", bold=True)
        cell(6, 3, "", bold=True)
        cell(6, 4, "", bold=True)
        cell(6, 5, "", bold=True)

        self.labor_ratio_var.trace_add(
            "write", lambda *a: self._on_ratio_var_changed("labor_expense_ratio", self.labor_ratio_var))
        self.sga_ratio_var.trace_add(
            "write", lambda *a: self._on_ratio_var_changed("sga_ratio", self.sga_ratio_var))

        tk.Label(
            card, text="※ 노무비+경비 비율, 판관비 비율 칸(노란색)에 숫자를 입력하면 즉시 재계산됩니다. (예: 20 → 20%)",
            font=("Malgun Gothic", 9), fg=COLOR_SUBTEXT, bg=COLOR_CARD, anchor="w", justify="left",
        ).pack(anchor="w", padx=20, pady=(10, 20), fill="x")

        return outer

    def _on_ratio_var_changed(self, settings_key, var):
        val = parse_number(var.get(), self.settings[settings_key] * 100)
        val = max(0.0, val) / 100.0
        self.settings[settings_key] = val
        self._recalc()

    def _on_reset_cost_ratios(self):
        # 값을 바꾸면 기존 trace(_on_ratio_var_changed)가 자동으로 재계산·저장까지 처리한다.
        self.labor_ratio_var.set(f"{DEFAULT_SETTINGS['labor_expense_ratio']*100:.0f}")
        self.sga_ratio_var.set(f"{DEFAULT_SETTINGS['sga_ratio']*100:.0f}")

    # ---------------------------------------------------------------- Rawdata 새창 (요청사항 2, 4)
    def _open_rawdata_window(self):
        if self._rawdata_win is not None and self._rawdata_win.winfo_exists():
            self._rawdata_win.lift()
            self._rawdata_win.focus_force()
            return

        win = tk.Toplevel(self)
        win.title("PUMP 판가 DATA - 조회 전용")
        win.configure(bg=COLOR_BG)
        maximize_window(win)
        self._rawdata_win = win

        top = tk.Frame(win, bg=COLOR_BG)
        top.pack(fill="x", padx=18, pady=14)
        tk.Label(top, text="PUMP 판가 DATA", font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_TEXT).pack(side="left")
        tk.Label(top, text="  조회 전용 (수정 불가)", font=("Malgun Gothic", 10), bg=COLOR_BG,
                 fg=COLOR_SUBTEXT).pack(side="left")
        ttk.Button(top, text="닫기", command=win.destroy).pack(side="right")
        ttk.Button(top, text="새로고침", command=self._reload_rawdata_window).pack(side="right", padx=(0, 8))

        search_bar = tk.Frame(win, bg=COLOR_BG)
        search_bar.pack(fill="x", padx=18, pady=(0, 12))
        tk.Label(search_bar, text="🔍 검색", font=FONT_BASE, bg=COLOR_BG, fg=COLOR_TEXT).pack(side="left")
        self._rawdata_search_var = tk.StringVar()
        search_entry = tk.Entry(search_bar, textvariable=self._rawdata_search_var, font=FONT_BASE,
                                 relief="solid", bd=1, highlightthickness=0, width=32)
        search_entry.pack(side="left", padx=8, ipady=4)
        self._rawdata_search_var.trace_add("write", lambda *a: self._apply_rawdata_filter())
        self._rawdata_count_label = tk.Label(search_bar, text="", font=("Malgun Gothic", 9),
                                              bg=COLOR_BG, fg=COLOR_SUBTEXT)
        self._rawdata_count_label.pack(side="left", padx=6)

        container_outer, container = self._card(win)
        container_outer.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        tree_frame = tk.Frame(container, bg=COLOR_CARD)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=12)

        tree = DataTreeview(tree_frame, self.rawdata_headers, editable=False)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._rawdata_tree = tree
        self._reload_rawdata_window()

    def _reload_rawdata_window(self):
        if self._rawdata_tree is None:
            return
        rows = load_csv_rows(RAWDATA_CSV, self.rawdata_headers)
        self._rawdata_tree.set_rows(rows)
        if self._rawdata_search_var is not None:
            self._rawdata_search_var.set("")
        self._update_rawdata_count()

    def _apply_rawdata_filter(self):
        if self._rawdata_tree is None:
            return
        self._rawdata_tree.apply_filter(self._rawdata_search_var.get())
        self._update_rawdata_count()

    def _update_rawdata_count(self):
        if self._rawdata_tree is None or self._rawdata_count_label is None:
            return
        shown = len(self._rawdata_tree.get_children())
        total = len(self._rawdata_tree.get_raw_rows())
        self._rawdata_count_label.config(text=f"{shown} / {total}건 표시")

    # ---------------------------------------------------------------- 이력 새창 (요청사항 2, 5, 6, 7, 11)
    def _open_history_window(self, select_no=None):
        if self._history_win is not None and self._history_win.winfo_exists():
            self._history_win.lift()
            self._history_win.focus_force()
            self._reload_history_window(select_no=select_no)
            return

        win = tk.Toplevel(self)
        win.title("이력")
        win.configure(bg=COLOR_BG)
        maximize_window(win)
        self._history_win = win

        top = tk.Frame(win, bg=COLOR_BG)
        top.pack(fill="x", padx=18, pady=14)
        tk.Label(top, text="이력", font=FONT_TITLE, bg=COLOR_BG, fg=COLOR_TEXT).pack(side="left")
        tk.Label(top, text="  셀 더블클릭으로 수정", font=("Malgun Gothic", 10), bg=COLOR_BG,
                 fg=COLOR_SUBTEXT).pack(side="left")
        ttk.Button(top, text="닫기", command=win.destroy).pack(side="right")
        ttk.Button(top, text="새로고침", command=self._reload_history_window).pack(side="right", padx=(0, 8))

        search_bar = tk.Frame(win, bg=COLOR_BG)
        search_bar.pack(fill="x", padx=18, pady=(0, 12))
        tk.Label(search_bar, text="🔍 검색", font=FONT_BASE, bg=COLOR_BG, fg=COLOR_TEXT).pack(side="left")
        self._history_search_var = tk.StringVar()
        search_entry = tk.Entry(search_bar, textvariable=self._history_search_var, font=FONT_BASE,
                                 relief="solid", bd=1, highlightthickness=0, width=32)
        search_entry.pack(side="left", padx=8, ipady=4)
        self._history_search_var.trace_add("write", lambda *a: self._apply_history_filter())
        self._history_count_label = tk.Label(search_bar, text="", font=("Malgun Gothic", 9),
                                              bg=COLOR_BG, fg=COLOR_SUBTEXT)
        self._history_count_label.pack(side="left", padx=6)

        container_outer, container = self._card(win)
        container_outer.pack(fill="both", expand=True, padx=18, pady=(0, 18))

        tree_frame = tk.Frame(container, bg=COLOR_CARD)
        tree_frame.pack(fill="both", expand=True, padx=12, pady=12)

        tree = DataTreeview(tree_frame, HISTORY_COLUMNS, editable=True, choice_map=HISTORY_CHOICE_MAP,
                             on_change=self._on_history_changed, on_cell_change=self._on_history_cell_edited)
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        self._history_tree = tree
        self._reload_history_window(select_no=select_no)

    def _reload_history_window(self, select_no=None):
        if self._history_tree is None:
            return
        rows = load_csv_rows(HISTORY_CSV, HISTORY_COLUMNS)
        self._history_tree.set_rows(rows)
        if self._history_search_var is not None:
            self._history_search_var.set("")
        self._update_history_count()
        if select_no is not None:
            self._select_history_no(select_no)

    def _apply_history_filter(self):
        if self._history_tree is None:
            return
        self._history_tree.apply_filter(self._history_search_var.get())
        self._update_history_count()

    def _update_history_count(self):
        if self._history_tree is None or self._history_count_label is None:
            return
        shown = len(self._history_tree.get_children())
        total = len(self._history_tree.get_raw_rows())
        self._history_count_label.config(text=f"{shown} / {total}건 표시")

    def _select_history_no(self, no_value):
        if self._history_tree is None:
            return
        for idx, row in enumerate(self._history_tree.get_raw_rows()):
            if row and str(row[0]).strip() == str(no_value):
                iid = str(idx)
                if self._history_tree.exists(iid):
                    self._history_tree.selection_set(iid)
                    self._history_tree.see(iid)
                break

    def _on_history_changed(self):
        if self._history_tree is None:
            return
        save_csv_rows(HISTORY_CSV, HISTORY_COLUMNS, self._history_tree.get_raw_rows())
        self._update_history_count()

    def _on_history_cell_edited(self, row_index, header, old_value, new_value):
        """이력 셀 수정 시 시간/셀위치/변경내용을 텍스트 로그로 남긴다 (요청사항)."""
        tree = self._history_tree
        if tree is None:
            return
        raw_rows = tree.get_raw_rows()
        row = raw_rows[row_index] if 0 <= row_index < len(raw_rows) else None
        no_value = row[HISTORY_COLUMNS.index("No.")] if row else None

        # No.만으로는 어떤 항목인지 바로 떠올리기 어려우므로, 장비 모델을 함께 남겨
        # 어떤 이력인지 더 쉽게 알아볼 수 있게 한다.
        context_label = ""
        if row:
            model_value = row[HISTORY_COLUMNS.index("장비 모델")]
            if model_value:
                context_label = f"장비 모델: {model_value}"

        old_display = tree.format_cell(header, old_value)
        new_display = tree.format_cell(header, new_value)
        append_history_edit_log(no_value, context_label, header, old_display, new_display)

    # ---------------------------------------------------------------- 계산 / 재계산
    def _recalc(self):
        result = compute_price(
            self.settings["material_cost"],
            self.settings["labor_expense_ratio"],
            self.settings["sga_ratio"],
            self.settings["profit_margin"],
        )
        self._last_result = result

        self.summary_label.config(text=(
            f"재료비(원가)      : {fmt_won(result['재료비'])} 원\n"
            f"노무비+경비       : {fmt_won(result['노무비+경비'])} 원\n"
            f"판관비            : {fmt_won(result['판관비'])} 원\n"
            f"영업이익          : {fmt_won(result['영업이익'])} 원\n"
            f"{'-'*40}\n"
            f"적정 판가         : {fmt_won(result['적정판가'])} 원"
        ))

        if self.calc_labels:
            self.calc_labels["재료비_금액"].config(text=fmt_won(result["재료비"]))
            self.calc_labels["노무비_금액"].config(text=fmt_won(result["노무비+경비"]))
            self.calc_labels["노무비_상세"].config(text=f"재료비의 {self.settings['labor_expense_ratio']*100:.0f}%")
            self.calc_labels["판관비_금액"].config(text=fmt_won(result["판관비"]))
            self.calc_labels["판관비_상세"].config(text=f"원가의 {self.settings['sga_ratio']*100:.0f}%")
            self.calc_labels["영업이익_금액"].config(text=fmt_won(result["영업이익"]))
            self.calc_labels["영업이익_비율"].config(text=f"{self.settings['profit_margin']*100:.0f}%")
            self.calc_labels["영업이익_상세"].config(text=f"(원가+판관비)의 {self.settings['profit_margin']*100:.0f}%")
            self.calc_labels["적정판가"].config(text=fmt_won(result["적정판가"]))

        save_settings(self.settings)

    # ---------------------------------------------------------------- 이력등록 (요청사항 3)
    def _on_register_history(self):
        summary = (
            f"재료비(원가) {fmt_won(self.settings['material_cost'])}원 · "
            f"영업이익률 {self.settings['profit_margin']*100:.0f}% · "
            f"적정 판가 {fmt_won(self._last_result['적정판가'])}원 으로 등록합니다."
        )
        dlg = HistoryRegisterDialog(self, summary)
        values = dlg.result
        if values is None:
            return  # 사용자가 취소 -> 원본과 동일하게 아무 것도 등록하지 않음

        existing_rows = load_csv_rows(HISTORY_CSV, HISTORY_COLUMNS)
        if existing_rows:
            try:
                next_no = int(float(existing_rows[-1][0])) + 1
            except (ValueError, IndexError):
                next_no = len(existing_rows) + 1
        else:
            next_no = 1

        today = datetime.date.today().strftime("%Y-%m-%d")
        material_cost_raw = self.settings["material_cost"]
        profit_margin_raw = self.settings["profit_margin"]
        calc_amount_raw = self._last_result["적정판가"]
        note_text = self._build_note_with_ratio_changes(values["v_note"])
        new_row = [
            next_no,
            today,
            values["v_progress"],
            values["v_region"],
            values["v_cust"],
            values["v_div"],
            values["v_process"],
            values["v_detail"],
            values["v_maker"],
            values["v_model"],
            values["v_content"],
            _num_to_cell(material_cost_raw),
            _num_to_cell(profit_margin_raw),
            _num_to_cell(calc_amount_raw),
            values["v_submit"],
            "대기",
            values["v_modle"],
            values["v_fscAsIs"],
            values["v_fscToBe"],
            values["v_sec"],
            values["v_lot"],
            note_text,
        ]
        existing_rows.append(new_row)
        save_csv_rows(HISTORY_CSV, HISTORY_COLUMNS, existing_rows)

        messagebox.showinfo("이력등록", f"이력이 등록되었습니다. (No. {next_no})")
        # 팝업이 계속 뜨는 대신, 등록된 내용을 바로 이력 창에서 확인할 수 있도록 연다.
        self._open_history_window(select_no=next_no)

    def _build_note_with_ratio_changes(self, manual_note):
        """노무비+경비/판관비 비율이 기본값(20%/30%)이 아니면 비고에 자동으로 남긴다 (요청사항).
        예: 노무비+경비 10% 반영, 판관비 5% 반영"""
        auto_notes = []
        labor_ratio = self.settings["labor_expense_ratio"]
        sga_ratio = self.settings["sga_ratio"]
        if abs(labor_ratio - DEFAULT_SETTINGS["labor_expense_ratio"]) > 1e-9:
            auto_notes.append(f"노무비+경비 {labor_ratio*100:.0f}% 반영")
        if abs(sga_ratio - DEFAULT_SETTINGS["sga_ratio"]) > 1e-9:
            auto_notes.append(f"판관비 {sga_ratio*100:.0f}% 반영")
        if not auto_notes:
            return manual_note
        auto_text = ", ".join(auto_notes)
        return f"{manual_note} ({auto_text})" if manual_note else auto_text


if __name__ == "__main__":
    app = PumpPriceApp()
    app.mainloop()
