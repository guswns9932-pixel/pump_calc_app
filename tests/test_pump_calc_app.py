# -*- coding: utf-8 -*-
"""
pump_calc_app.py의 핵심 로직(계산/파싱/서식/CSV 입출력/백업/로그/엑셀 내보내기)에 대한
자동 테스트. GUI(Tk 위젯) 자체는 다루지 않는다 — 디스플레이 없는 환경(CI 등)에서도
그대로 실행되도록, tkinter 위젯을 생성하지 않는 순수 함수만 검증한다.

실행 방법:
    pip install pytest
    pytest tests/

주의: 아래 fixture(patch_data_paths)가 모든 파일 입출력 테스트에서 pump_calc_app의
데이터 경로 상수를 임시 폴더로 바꿔치기하므로, 실제 저장소의 data/ 폴더는 절대
건드리지 않는다.
"""
import os
import zipfile
import xml.etree.ElementTree as ET

import pytest

import pump_calc_app as m


# ----------------------------------------------------------------------------------
# fixture: 모든 데이터 경로를 임시 폴더로 리다이렉트 (실제 data/ 폴더 보호)
# ----------------------------------------------------------------------------------
@pytest.fixture
def patch_data_paths(tmp_path, monkeypatch):
    data_dir = tmp_path / "data"
    data_dir.mkdir()
    monkeypatch.setattr(m, "DATA_DIR", str(data_dir))
    monkeypatch.setattr(m, "RAWDATA_CSV", str(data_dir / "rawdata.csv"))
    monkeypatch.setattr(m, "HISTORY_CSV", str(data_dir / "history.csv"))
    monkeypatch.setattr(m, "SETTINGS_JSON", str(data_dir / "settings.json"))
    monkeypatch.setattr(m, "HISTORY_EDIT_LOG", str(data_dir / "history_edit_log.txt"))
    monkeypatch.setattr(m, "BACKUP_DIR", str(data_dir / "backups"))
    return data_dir


# ----------------------------------------------------------------------------------
# 계산 로직
# ----------------------------------------------------------------------------------
def test_compute_price_matches_original_excel_sample():
    # README에 명시된 원본 Sheet3 샘플값: 재료비 1,000,000 / 20% / 30% / 50% -> 2,340,000
    result = m.compute_price(1_000_000, 0.2, 0.3, 0.5)
    assert result["재료비"] == 1_000_000
    assert result["노무비+경비"] == 200_000
    assert result["판관비"] == 360_000
    assert result["영업이익"] == 780_000
    assert result["적정판가"] == 2_340_000


def test_compute_price_zero_ratios():
    result = m.compute_price(500_000, 0, 0, 0)
    assert result["적정판가"] == 500_000


# ----------------------------------------------------------------------------------
# 숫자/텍스트 파싱 및 서식
# ----------------------------------------------------------------------------------
@pytest.mark.parametrize("raw,expected", [
    ("1,000,000", 1_000_000.0),
    ("50%", 50.0),
    ("1,950,000원", 1_950_000.0),
    ("  1,234 ", 1234.0),
    ("\xa0\xa0 1,350,000", 1_350_000.0),
    ("-7.7%", -7.7),
    ("", None),
    (None, None),
    ("텍스트", None),
])
def test_parse_number(raw, expected):
    assert m.parse_number(raw, None) == expected


def test_parse_number_default_fallback():
    assert m.parse_number("", 42) == 42
    assert m.parse_number("abc", 7) == 7


def test_fmt_won():
    assert m.fmt_won(1000000) == "1,000,000"
    assert m.fmt_won(0) == "0"
    assert m.fmt_won(1234.6) == "1,235"  # 반올림


def test_fmt_pct():
    assert m.fmt_pct(0.5) == "50.0%"
    assert m.fmt_pct(0.0366985869233388) == "3.7%"
    assert m.fmt_pct(-0.0767) == "-7.7%"


def test_num_to_cell_integer_vs_float():
    assert m._num_to_cell(1000000.0) == 1000000
    assert isinstance(m._num_to_cell(1000000.0), int)
    assert m._num_to_cell(0.5) == 0.5
    assert m._num_to_cell(None) == ""


def test_clean_text_normalizes_newlines_and_spaces():
    assert m.clean_text("Heat Exchanger(19F)+End Cell Purge\n+ Flush Purge") == \
        "Heat Exchanger(19F)+End Cell Purge + Flush Purge"
    assert m.clean_text("a\r\n\r\nb") == "a b"
    assert m.clean_text("  a   b  ") == "a b"
    assert m.clean_text("\xa0시작\xa0") == "시작"
    assert m.clean_text(None) is None


# ----------------------------------------------------------------------------------
# 열 종류 판별 (금액/비율/텍스트)
# ----------------------------------------------------------------------------------
@pytest.mark.parametrize("header,expected", [
    ("재료비(원가)", "money"),
    ("영업이익률", "percent"),
    ("계산금액", "money"),
    ("제출금액", "money"),
    ("실제 영업이익률", "percent"),
    ("BEP 판가\n(영업이익률 20%)", "money"),  # 괄호 안 설명 때문에 percent로 오분류되면 안 됨
    ("판관비 30%", "money"),
    ("퍼센트", "percent"),
    ("No.", "text"),
    ("장비 모델", "text"),
    ("채택 현황", "text"),
])
def test_classify_column(header, expected):
    assert m.classify_column(header) == expected


# ----------------------------------------------------------------------------------
# CSV 입출력 (임시 폴더 사용)
# ----------------------------------------------------------------------------------
def test_save_and_load_csv_roundtrip(tmp_path):
    path = str(tmp_path / "sample.csv")
    columns = ["a", "b", "c"]
    rows = [["1", "2", "3"], ["가", "나", "다"]]
    m.save_csv_rows(path, columns, rows)
    assert m.load_csv_header(path) == columns
    assert m.load_csv_rows(path, columns) == rows


def test_save_csv_rows_is_atomic_no_tmp_left_behind(tmp_path):
    path = str(tmp_path / "sample.csv")
    m.save_csv_rows(path, ["a"], [["1"]])
    assert os.path.exists(path)
    assert not os.path.exists(path + ".tmp")


def test_ensure_utf8_bom_adds_bom_and_preserves_content(tmp_path):
    path = tmp_path / "plain.csv"
    path.write_text("a,b\n가,나\n", encoding="utf-8", newline="")
    m.ensure_utf8_bom(str(path))
    raw = path.read_bytes()
    assert raw.startswith(b"\xef\xbb\xbf")
    assert raw[3:].decode("utf-8") == "a,b\n가,나\n"

    # 이미 BOM이 있으면 다시 손대지 않는다 (멱등성)
    before = path.read_bytes()
    m.ensure_utf8_bom(str(path))
    assert path.read_bytes() == before


def test_ensure_utf8_bom_reading_with_bom_has_no_leaked_prefix(tmp_path):
    path = tmp_path / "plain.csv"
    path.write_text("FSC,No.\nADD05N413177,1\n", encoding="utf-8", newline="")
    m.ensure_utf8_bom(str(path))
    header = m.load_csv_header(str(path))
    assert header[0] == "FSC"  # "﻿FSC"로 깨지면 안 됨


# ----------------------------------------------------------------------------------
# 저장 전 백업
# ----------------------------------------------------------------------------------
def test_backup_history_csv_creates_timestamped_copy(patch_data_paths):
    m.save_csv_rows(m.HISTORY_CSV, m.HISTORY_COLUMNS, [["1"] + [""] * (len(m.HISTORY_COLUMNS) - 1)])
    m.backup_history_csv()
    backups = os.listdir(m.BACKUP_DIR)
    assert len(backups) == 1
    assert backups[0].startswith("history_") and backups[0].endswith(".csv")


def test_backup_history_csv_noop_when_no_history_file_yet(patch_data_paths):
    m.backup_history_csv()  # history.csv가 아직 없으면 조용히 아무 것도 하지 않는다
    assert not os.path.exists(m.BACKUP_DIR)


def test_backup_history_csv_prunes_old_backups(patch_data_paths, monkeypatch):
    monkeypatch.setattr(m, "MAX_HISTORY_BACKUPS", 3)
    m.save_csv_rows(m.HISTORY_CSV, m.HISTORY_COLUMNS, [])
    for _ in range(5):
        m.backup_history_csv()
    backups = os.listdir(m.BACKUP_DIR)
    assert len(backups) == 3


def test_save_history_csv_backs_up_then_saves(patch_data_paths):
    row1 = ["1"] + [""] * (len(m.HISTORY_COLUMNS) - 1)
    m.save_history_csv([row1])
    assert not os.path.isdir(m.BACKUP_DIR)  # 최초 저장 시엔 백업할 기존 파일이 없음

    row2 = ["2"] + [""] * (len(m.HISTORY_COLUMNS) - 1)
    m.save_history_csv([row1, row2])
    assert os.path.isdir(m.BACKUP_DIR)
    assert len(os.listdir(m.BACKUP_DIR)) == 1
    assert m.load_csv_rows(m.HISTORY_CSV, m.HISTORY_COLUMNS) == [row1, row2]


# ----------------------------------------------------------------------------------
# 이력 등록/수정/삭제 로그
# ----------------------------------------------------------------------------------
def test_append_history_edit_log_format(patch_data_paths):
    m.append_history_edit_log("12", "장비 모델: DD605P", "제출여부", "제출예정", "제출완료")
    text = open(m.HISTORY_EDIT_LOG, encoding="utf-8-sig").read()
    assert "No.12 (장비 모델: DD605P)" in text
    assert "제출여부: 제출예정 → 제출완료" in text


def test_append_history_edit_log_blank_value_shown_clearly(patch_data_paths):
    m.append_history_edit_log("5", "", "비고", "", "특이사항 있음")
    text = open(m.HISTORY_EDIT_LOG, encoding="utf-8-sig").read()
    assert "(빈 값) → 특이사항 있음" in text


def test_append_history_register_log_format(patch_data_paths):
    m.append_history_register_log("52", "장비 모델: DD605P-TEST", "재료비(원가) 1,000,000원 · 영업이익률 50%")
    text = open(m.HISTORY_EDIT_LOG, encoding="utf-8-sig").read()
    assert "신규 등록: 재료비(원가) 1,000,000원 · 영업이익률 50%" in text


def test_append_history_delete_log_format(patch_data_paths):
    m.append_history_delete_log("3", "", "(재료비 정보 없음)")
    text = open(m.HISTORY_EDIT_LOG, encoding="utf-8-sig").read()
    assert "No.3" in text and "삭제됨: (재료비 정보 없음)" in text


def test_history_log_entries_append_not_overwrite(patch_data_paths):
    m.append_history_edit_log("1", "", "A", "1", "2")
    m.append_history_edit_log("1", "", "A", "2", "3")
    lines = open(m.HISTORY_EDIT_LOG, encoding="utf-8-sig").readlines()
    assert len(lines) == 2


# ----------------------------------------------------------------------------------
# 엑셀(.xlsx) 내보내기 — 외부 라이브러리 없이 직접 생성한 zip/XML의 유효성 검증
# ----------------------------------------------------------------------------------
def test_xlsx_col_letter():
    assert m._xlsx_col_letter(0) == "A"
    assert m._xlsx_col_letter(25) == "Z"
    assert m._xlsx_col_letter(26) == "AA"
    assert m._xlsx_col_letter(27) == "AB"


def test_export_rows_to_xlsx_produces_wellformed_package(tmp_path):
    path = str(tmp_path / "out.xlsx")
    headers = ["이름", "금액", "비고"]
    rows = [
        ["항목A", "1,000,000", "일반 텍스트"],
        ["항목B & C", "2,000,000", "특수문자 <테스트> \"인용\""],
    ]
    m.export_rows_to_xlsx(path, "테스트시트", headers, rows)

    assert os.path.exists(path)
    with zipfile.ZipFile(path) as zf:
        names = set(zf.namelist())
        assert {
            "[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels", "xl/styles.xml", "xl/worksheets/sheet1.xml",
        } <= names

        # 모든 XML 파트가 well-formed 해야 한다 (엑셀/오피스가 열 수 있는 최소 조건)
        for name in names:
            data = zf.read(name)
            ET.fromstring(data)  # 파싱 실패 시 예외 발생 -> 테스트 실패

        sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        assert "항목A" in sheet_xml
        # &, <, > 는 반드시 XML 엔티티로 escape되어 있어야 한다 (그렇지 않으면 XML 자체가 깨짐)
        assert "항목B &amp; C" in sheet_xml
        assert "&lt;테스트&gt;" in sheet_xml
        assert "항목B & C" not in sheet_xml  # 이스케이프 안 된 원문 & 이 그대로 남아있으면 안 됨


def test_export_rows_to_xlsx_empty_rows(tmp_path):
    path = str(tmp_path / "empty.xlsx")
    m.export_rows_to_xlsx(path, "빈 표", ["A", "B"], [])
    with zipfile.ZipFile(path) as zf:
        sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")
        ET.fromstring(sheet_xml)
        assert "<row r=\"1\">" in sheet_xml  # 헤더 행만 존재
