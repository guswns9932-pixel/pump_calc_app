# -*- coding: utf-8 -*-
"""
Pump 판가 계산기 (Python 이식판)
원본: 대외비_Pump_판가_계산기_이력추가_ver4_0.xlsm (판가계산기 / Sheet3 / Rawdata / 이력)

이식 원칙
---------
1. 원본 엑셀의 계산 로직(수식)을 그대로 유지한다.
     D3(재료비)          = 판가계산기!B5
     D4(노무비+경비)      = D3 * E4         (E4 = 노무비+경비 비율, 원본 하드코딩 0.2 -> 이번 버전에서 직접 입력 가능)
     D5(판관비)           = (D3+D4) * E5    (E5 = 판관비 비율,      원본 하드코딩 0.3 -> 이번 버전에서 직접 입력 가능)
     D6(영업이익)         = (D3+D4+D5) * E6 (E6 = 영업이익률,       원본부터 직접 입력 가능 = 판가계산기!C5)
     D7(적정 판가)        = D3+D4+D5+D6
2. '이력등록' 버튼 클릭 시 원본 VBA(Module1.bas RegisterHistory)와 동일한 순서 / 동일한 항목으로
   정보를 입력받아 '이력' 시트에 동일한 규칙(No. 자동증가, 오늘 날짜, 계산금액 자동기록, 채택 현황='대기')으로 기록한다.
3. Rawdata(과거 실적 73건)는 원본과 동일하게 조회용으로 제공한다.
4. UI는 엑셀의 탭(시트) 구조를 그대로 재현한다: 판가계산기 / Sheet3 / Rawdata / 이력
"""

import os
import csv
import json
import datetime
import tkinter as tk
from tkinter import ttk, messagebox

# --------------------------------------------------------------------------------------
# 경로 / 상수
# --------------------------------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
RAWDATA_CSV = os.path.join(DATA_DIR, "rawdata.csv")
HISTORY_CSV = os.path.join(DATA_DIR, "history.csv")
SETTINGS_JSON = os.path.join(DATA_DIR, "settings.json")

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
    ("v_progress", "제출여부를 선택하세요", "choice", CHOICE_제출여부),
    ("v_region", "지역을 선택하세요", "choice", CHOICE_지역),
    ("v_cust", "고객구분을 선택하세요", "choice", CHOICE_고객구분),
    ("v_div", "사업부를 선택하세요", "choice", CHOICE_사업부),
    ("v_process", "대공정을 선택하세요", "choice", CHOICE_대공정),
    ("v_detail", "세부공정을 입력하세요", "text", None),
    ("v_maker", "장비사를 선택하세요", "choice", CHOICE_장비사),
    ("v_model", "장비 모델을 입력하세요", "text", None),
    ("v_content", "Option을 입력하세요", "text", None),
    ("v_submit", "제출금액을 입력하세요", "text", None),
    ("v_modle", "Pump Model을 입력하세요", "text", None),
    ("v_fscAsIs", "Pump FSC (As Is)를 입력하세요", "text", None),
    ("v_fscToBe", "Pump FSC (To Be)를 입력하세요", "text", None),
    ("v_sec", "SEC 담당자를 입력하세요", "text", None),
    ("v_lot", "LOT 담당자를 입력하세요", "text", None),
    ("v_note", "비고를 입력하세요", "text", None),
]

DEFAULT_SETTINGS = {
    "material_cost": 1000000,      # 판가계산기!B5  재료비(원가)
    "profit_margin": 0.5,          # 판가계산기!C5  영업이익률
    "labor_expense_ratio": 0.2,    # Sheet3!E4      노무비+경비 비율  (직접 입력 가능하도록 변경)
    "sga_ratio": 0.3,              # Sheet3!E5      판관비 비율      (직접 입력 가능하도록 변경)
}

WON_FMT = "{:,.0f}"


def fmt_won(v):
    try:
        return WON_FMT.format(v)
    except (TypeError, ValueError):
        return str(v)


def fmt_pct(v):
    try:
        return f"{v * 100:.1f}%"
    except (TypeError, ValueError):
        return str(v)


def _num_to_cell(v):
    """엑셀 셀처럼 정수는 정수로, 소수는 그대로 저장 (콤마 등 서식 문자는 넣지 않음)."""
    if float(v).is_integer():
        return int(v)
    return round(float(v), 6)


def parse_number(s, default=0.0):
    """콤마, 공백, '원', '%' 등이 섞여 있어도 숫자만 파싱."""
    if s is None:
        return default
    s = str(s).strip().replace(",", "").replace("원", "").replace("%", "")
    if s == "":
        return default
    try:
        return float(s)
    except ValueError:
        return default


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


def load_csv_rows(path, columns):
    if not os.path.exists(path):
        return []
    with open(path, "r", newline="", encoding="utf-8") as f:
        reader = csv.reader(f)
        rows = list(reader)
    if not rows:
        return []
    return rows[1:]  # 헤더 제외


def save_csv_rows(path, columns, rows):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(columns)
        writer.writerows(rows)


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
# 이력등록 마법사 (원본 VBA 팝업 시퀀스 재현)
# --------------------------------------------------------------------------------------
class WizardStepDialog(tk.Toplevel):
    """한 번에 한 항목씩 입력받는 팝업. 원본 frmSelect / frmTextInput 과 동일한 흐름."""

    def __init__(self, parent, prompt, kind, choices=None):
        super().__init__(parent)
        self.title("이력등록")
        self.resizable(False, False)
        self.result = None
        self.cancelled = True
        self.transient(parent)
        self.grab_set()

        tk.Label(self, text=prompt, font=("Malgun Gothic", 11)).pack(anchor="w", padx=16, pady=(16, 8))

        if kind == "choice":
            self.listbox = tk.Listbox(self, height=min(8, len(choices)), exportselection=False, font=("Malgun Gothic", 10))
            for item in choices:
                self.listbox.insert(tk.END, item)
            self.listbox.selection_set(0)
            self.listbox.pack(padx=16, fill="x")
            self.listbox.focus_set()
            self.listbox.bind("<Return>", lambda e: self._ok())
            self.listbox.bind("<Double-Button-1>", lambda e: self._ok())
            self.entry = None
        else:
            self.entry_var = tk.StringVar()
            self.entry = tk.Entry(self, textvariable=self.entry_var, width=42, font=("Malgun Gothic", 10))
            self.entry.pack(padx=16, fill="x")
            self.entry.focus_set()
            self.entry.bind("<Return>", lambda e: self._ok())
            self.listbox = None

        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=12)
        tk.Button(btn_frame, text="확인", width=10, command=self._ok).pack(side="left", padx=4)
        tk.Button(btn_frame, text="취소", width=10, command=self._cancel).pack(side="left", padx=4)

        self.protocol("WM_DELETE_WINDOW", self._cancel)
        self.update_idletasks()
        x = parent.winfo_rootx() + 60
        y = parent.winfo_rooty() + 60
        self.geometry(f"+{x}+{y}")
        self.wait_window(self)

    def _ok(self):
        if self.listbox is not None:
            sel = self.listbox.curselection()
            self.result = self.listbox.get(sel[0]) if sel else ""
        else:
            self.result = self.entry_var.get().strip()
        self.cancelled = False
        self.destroy()

    def _cancel(self):
        self.cancelled = True
        self.destroy()


def run_register_history_wizard(parent, current_price):
    """원본 VBA RegisterHistory 서브루틴 재현. 취소 시 None 반환(전체 취소)."""
    values = {}
    for key, prompt, kind, choices in WIZARD_FIELDS:
        dlg = WizardStepDialog(parent, prompt, kind, choices)
        if dlg.cancelled:
            return None  # If gLastCancelled Then Exit Sub 와 동일
        values[key] = dlg.result
    return values


# --------------------------------------------------------------------------------------
# 셀 더블클릭 편집이 가능한 Treeview (엑셀처럼 셀을 직접 수정)
# --------------------------------------------------------------------------------------
class EditableTreeview(ttk.Treeview):
    def __init__(self, master, columns, on_cell_edited=None, **kwargs):
        super().__init__(master, columns=columns, show="headings", **kwargs)
        self._columns = columns
        self._on_cell_edited = on_cell_edited
        self.bind("<Double-1>", self._begin_edit)
        self._edit_entry = None

    def _begin_edit(self, event):
        region = self.identify("region", event.x, event.y)
        if region != "cell":
            return
        row_id = self.identify_row(event.y)
        col_id = self.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_index = int(col_id.replace("#", "")) - 1
        x, y, w, h = self.bbox(row_id, col_id)
        value = self.set(row_id, self._columns[col_index])

        if self._edit_entry is not None:
            self._edit_entry.destroy()

        entry = tk.Entry(self, font=("Malgun Gothic", 9))
        entry.insert(0, value)
        entry.select_range(0, tk.END)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        self._edit_entry = entry

        def commit(_event=None):
            new_value = entry.get()
            self.set(row_id, self._columns[col_index], new_value)
            entry.destroy()
            self._edit_entry = None
            if self._on_cell_edited:
                self._on_cell_edited(row_id, self._columns[col_index], new_value)

        def cancel(_event=None):
            entry.destroy()
            self._edit_entry = None

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
        self.geometry("1180x680")

        self.settings = load_settings()

        self.material_cost_var = tk.StringVar(value=fmt_won(self.settings["material_cost"]))
        self.profit_margin_var = tk.StringVar(value=f"{self.settings['profit_margin']*100:.1f}")
        self.labor_ratio_var = tk.StringVar(value=f"{self.settings['labor_expense_ratio']*100:.1f}")
        self.sga_ratio_var = tk.StringVar(value=f"{self.settings['sga_ratio']*100:.1f}")

        self.calc_labels = {}  # Sheet3 계산결과 표시용 라벨 저장

        self._build_notebook()
        self._recalc()
        self._load_rawdata()
        self._load_history()

    # ---------------------------------------------------------------- UI 골격
    def _build_notebook(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass
        style.configure("Treeview.Heading", font=("Malgun Gothic", 9, "bold"))
        style.configure("Treeview", font=("Malgun Gothic", 9), rowheight=22)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True)

        self.tab_calc = tk.Frame(self.notebook, bg="white")
        self.tab_sheet3 = tk.Frame(self.notebook, bg="white")
        self.tab_rawdata = tk.Frame(self.notebook, bg="white")
        self.tab_history = tk.Frame(self.notebook, bg="white")

        self.notebook.add(self.tab_calc, text="판가계산기")
        self.notebook.add(self.tab_sheet3, text="Sheet3")
        self.notebook.add(self.tab_rawdata, text="Rawdata")
        self.notebook.add(self.tab_history, text="이력")

        self._build_tab_calc()
        self._build_tab_sheet3()
        self._build_tab_rawdata()
        self._build_tab_history()

    # ---------------------------------------------------------------- 판가계산기 탭
    def _build_tab_calc(self):
        f = self.tab_calc

        banner = tk.Label(
            f, text="재료비(원가) / 영업이익률을 입력 → 이력등록 CLICK!",
            bg="#FFFF00", fg="#FF0000", font=("Malgun Gothic", 12, "bold"), pady=10,
        )
        banner.pack(fill="x", padx=12, pady=(12, 20))

        grid = tk.Frame(f, bg="white")
        grid.pack(padx=12, anchor="w")

        header_style = dict(font=("Malgun Gothic", 10, "bold"), bg="#D9D9D9",
                             relief="solid", bd=1, width=16, height=2)
        entry_style = dict(font=("Malgun Gothic", 11), relief="solid", bd=1, width=16, justify="right")

        tk.Label(grid, text="재료비(원가)", **header_style).grid(row=0, column=0)
        tk.Label(grid, text="영업이익률", **header_style).grid(row=0, column=1)

        e1 = tk.Entry(grid, textvariable=self.material_cost_var, **entry_style)
        e1.grid(row=1, column=0, ipady=6)
        e1.bind("<FocusOut>", lambda e: self._on_material_cost_changed())
        e1.bind("<Return>", lambda e: self._on_material_cost_changed())

        pct_frame = tk.Frame(grid)
        pct_frame.grid(row=1, column=1, ipady=6)
        e2 = tk.Entry(pct_frame, textvariable=self.profit_margin_var, font=("Malgun Gothic", 11),
                      relief="solid", bd=1, width=13, justify="right")
        e2.pack(side="left")
        tk.Label(pct_frame, text="%", font=("Malgun Gothic", 11)).pack(side="left")
        e2.bind("<FocusOut>", lambda e: self._on_profit_margin_changed())
        e2.bind("<Return>", lambda e: self._on_profit_margin_changed())

        result_frame = tk.LabelFrame(f, text="계산 결과 (Sheet3 연동)", font=("Malgun Gothic", 10),
                                      bg="white", padx=16, pady=12)
        result_frame.pack(fill="x", padx=12, pady=20)
        self.summary_label = tk.Label(result_frame, text="", justify="left",
                                       font=("Consolas", 11), bg="white", anchor="w")
        self.summary_label.pack(fill="x")

        tk.Button(
            f, text="이력등록 CLICK!", font=("Malgun Gothic", 12, "bold"),
            bg="#FFC000", activebackground="#FFD966", padx=20, pady=10,
            command=self._on_register_history,
        ).pack(padx=12, pady=10, anchor="w")

    def _on_material_cost_changed(self):
        val = parse_number(self.material_cost_var.get(), self.settings["material_cost"])
        self.settings["material_cost"] = val
        self.material_cost_var.set(fmt_won(val))
        self._recalc()

    def _on_profit_margin_changed(self):
        val = parse_number(self.profit_margin_var.get(), self.settings["profit_margin"] * 100) / 100.0
        self.settings["profit_margin"] = val
        self.profit_margin_var.set(f"{val*100:.1f}")
        self._recalc()

    # ---------------------------------------------------------------- Sheet3 탭
    def _build_tab_sheet3(self):
        f = self.tab_sheet3
        tk.Label(f, text="원가 구조 계산 (비율은 직접 입력 가능)", font=("Malgun Gothic", 12, "bold"),
                 bg="white").pack(anchor="w", padx=12, pady=(12, 6))

        grid = tk.Frame(f, bg="white")
        grid.pack(padx=12, pady=8, anchor="w")

        headers = ["분류", "구분", "금액", "비율", "비율 상세", "기본값 산정 근거"]
        for c, h in enumerate(headers):
            tk.Label(grid, text=h, font=("Malgun Gothic", 9, "bold"), bg="#D9D9D9",
                     relief="solid", bd=1, width=16, height=2).grid(row=0, column=c, sticky="nsew")

        def cell(row, col, text, editable_var=None, width=16, bold=False):
            if editable_var is not None:
                e = tk.Entry(grid, textvariable=editable_var, width=width, justify="center",
                             relief="solid", bd=1, font=("Malgun Gothic", 10), bg="#FFFFCC")
                e.grid(row=row, column=col, sticky="nsew", ipady=4)
                return e
            lbl = tk.Label(grid, text=text, width=width, relief="solid", bd=1,
                            font=("Malgun Gothic", 10, "bold" if bold else "normal"), bg="white")
            lbl.grid(row=row, column=col, sticky="nsew", ipady=4)
            return lbl

        # Row1: 원가 / 재료비
        cell(1, 0, "원가")
        cell(1, 1, "재료비")
        self.calc_labels["재료비_금액"] = cell(1, 2, "-")
        cell(1, 3, "-")
        cell(1, 4, "연구소 및 구매팀")
        cell(1, 5, "-")

        # Row2: 노무비+경비 (비율 EDITABLE)
        cell(2, 0, "")
        cell(2, 1, "노무비+경비")
        self.calc_labels["노무비_금액"] = cell(2, 2, "-")
        cell(2, 3, "", editable_var=self.labor_ratio_var)
        self.calc_labels["노무비_상세"] = cell(2, 4, "-")
        cell(2, 5, "Rawdata 기준 평균값")

        # Row3: 판관비 (비율 EDITABLE)
        cell(3, 0, "판매관리비")
        cell(3, 1, "판관비")
        self.calc_labels["판관비_금액"] = cell(3, 2, "-")
        cell(3, 3, "", editable_var=self.sga_ratio_var)
        self.calc_labels["판관비_상세"] = cell(3, 4, "-")
        cell(3, 5, "제조업 평균값")

        # Row4: 영업이익 (비율은 판가계산기 탭과 연동, 읽기전용)
        cell(4, 0, "이익")
        cell(4, 1, "영업이익")
        self.calc_labels["영업이익_금액"] = cell(4, 2, "-")
        self.calc_labels["영업이익_비율"] = cell(4, 3, "-")
        self.calc_labels["영업이익_상세"] = cell(4, 4, "-")
        cell(4, 5, "-")

        # Row5: 적정 판가
        cell(5, 0, "적정 판가", bold=True)
        cell(5, 1, "")
        self.calc_labels["적정판가"] = cell(5, 2, "-", bold=True)
        cell(5, 3, "")
        cell(5, 4, "")
        cell(5, 5, "")

        self.labor_ratio_var.trace_add("write", lambda *a: self._on_ratio_var_changed("labor_expense_ratio", self.labor_ratio_var))
        self.sga_ratio_var.trace_add("write", lambda *a: self._on_ratio_var_changed("sga_ratio", self.sga_ratio_var))

        tk.Label(
            f, text="※ 노무비+경비 비율, 판관비 비율 칸(노란색)에 직접 숫자를 입력하면 즉시 재계산됩니다. (예: 20 → 20%)",
            font=("Malgun Gothic", 9), fg="#555555", bg="white",
        ).pack(anchor="w", padx=12, pady=(8, 0))

    def _on_ratio_var_changed(self, settings_key, var):
        val = parse_number(var.get(), self.settings[settings_key] * 100)
        val = max(0.0, val) / 100.0
        self.settings[settings_key] = val
        self._recalc()

    # ---------------------------------------------------------------- Rawdata 탭
    def _build_tab_rawdata(self):
        f = self.tab_rawdata
        tk.Label(f, text="Rawdata (과거 실적 조회용 - 셀 더블클릭으로 수정 가능)",
                 font=("Malgun Gothic", 11, "bold"), bg="white").pack(anchor="w", padx=12, pady=(12, 6))

        container = tk.Frame(f)
        container.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.rawdata_columns = []
        if os.path.exists(RAWDATA_CSV):
            with open(RAWDATA_CSV, encoding="utf-8") as fp:
                self.rawdata_columns = next(csv.reader(fp))

        self.rawdata_tree = EditableTreeview(
            container, columns=self.rawdata_columns, on_cell_edited=self._on_rawdata_edited,
        )
        for col in self.rawdata_columns:
            self.rawdata_tree.heading(col, text=col)
            self.rawdata_tree.column(col, width=100, anchor="center")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.rawdata_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.rawdata_tree.xview)
        self.rawdata_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.rawdata_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def _load_rawdata(self):
        for item in self.rawdata_tree.get_children():
            self.rawdata_tree.delete(item)
        rows = load_csv_rows(RAWDATA_CSV, self.rawdata_columns)
        for row in rows:
            self.rawdata_tree.insert("", tk.END, values=row)

    def _on_rawdata_edited(self, row_id, col, value):
        self._persist_treeview(self.rawdata_tree, self.rawdata_columns, RAWDATA_CSV)

    # ---------------------------------------------------------------- 이력 탭
    def _build_tab_history(self):
        f = self.tab_history
        top = tk.Frame(f, bg="white")
        top.pack(fill="x", padx=12, pady=(12, 6))
        tk.Label(top, text="이력 (셀 더블클릭으로 직접 수정 가능)",
                 font=("Malgun Gothic", 11, "bold"), bg="white").pack(side="left")
        tk.Button(top, text="새로고침", command=self._load_history).pack(side="right")

        container = tk.Frame(f)
        container.pack(fill="both", expand=True, padx=12, pady=(0, 12))

        self.history_tree = EditableTreeview(
            container, columns=HISTORY_COLUMNS, on_cell_edited=self._on_history_edited,
        )
        for col in HISTORY_COLUMNS:
            self.history_tree.heading(col, text=col)
            width = 60 if col == "No." else 110
            self.history_tree.column(col, width=width, anchor="center")

        vsb = ttk.Scrollbar(container, orient="vertical", command=self.history_tree.yview)
        hsb = ttk.Scrollbar(container, orient="horizontal", command=self.history_tree.xview)
        self.history_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        self.history_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

    def _load_history(self):
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        rows = load_csv_rows(HISTORY_CSV, HISTORY_COLUMNS)
        for row in rows:
            self.history_tree.insert("", tk.END, values=row)

    def _on_history_edited(self, row_id, col, value):
        self._persist_treeview(self.history_tree, HISTORY_COLUMNS, HISTORY_CSV)

    def _persist_treeview(self, tree, columns, path):
        rows = [tree.item(item)["values"] for item in tree.get_children()]
        save_csv_rows(path, columns, rows)

    # ---------------------------------------------------------------- 계산 / 재계산
    def _recalc(self):
        result = compute_price(
            self.settings["material_cost"],
            self.settings["labor_expense_ratio"],
            self.settings["sga_ratio"],
            self.settings["profit_margin"],
        )
        self._last_result = result

        # 판가계산기 탭 요약
        self.summary_label.config(text=(
            f"재료비(원가)      : {fmt_won(result['재료비'])} 원\n"
            f"노무비+경비       : {fmt_won(result['노무비+경비'])} 원\n"
            f"판관비            : {fmt_won(result['판관비'])} 원\n"
            f"영업이익          : {fmt_won(result['영업이익'])} 원\n"
            f"{'-'*40}\n"
            f"적정 판가         : {fmt_won(result['적정판가'])} 원"
        ))

        # Sheet3 탭 갱신
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

    # ---------------------------------------------------------------- 이력등록
    def _on_register_history(self):
        values = run_register_history_wizard(self, self._last_result["적정판가"])
        if values is None:
            return  # 사용자가 중간에 취소 -> 원본과 동일하게 아무 것도 등록하지 않음

        existing_rows = load_csv_rows(HISTORY_CSV, HISTORY_COLUMNS)
        if existing_rows:
            try:
                next_no = int(float(existing_rows[-1][0])) + 1
            except (ValueError, IndexError):
                next_no = len(existing_rows) + 1
        else:
            next_no = 1

        today = datetime.date.today().strftime("%Y-%m-%d")
        # 원본 VBA와 동일하게 서식 없는 raw 숫자값으로 기록한다.
        # (L=재료비(원가) raw, M=영업이익률 raw fraction, N=계산금액 raw)
        material_cost_raw = self.settings["material_cost"]
        profit_margin_raw = self.settings["profit_margin"]
        calc_amount_raw = self._last_result["적정판가"]
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
            values["v_note"],
        ]
        existing_rows.append(new_row)
        save_csv_rows(HISTORY_CSV, HISTORY_COLUMNS, existing_rows)
        self._load_history()
        self.notebook.select(self.tab_history)
        messagebox.showinfo("이력등록", f"이력이 등록되었습니다. (No. {next_no})")


if __name__ == "__main__":
    app = PumpPriceApp()
    app.mainloop()
