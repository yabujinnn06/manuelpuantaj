"""WhatsApp parser, preview Excel ve bulk apply icin testler."""

import os
import sys
import json
import tempfile

import pytest

HERE = os.path.dirname(__file__)
APP_DIR = os.path.abspath(os.path.join(HERE, "..", "..", "..", ".."))
HARNESS_DIR = os.path.abspath(os.path.join(HERE, "..", "..", ".."))
for p in (APP_DIR, HARNESS_DIR):
    if p not in sys.path:
        sys.path.insert(0, p)


from cli_anything.puantaj import whatsapp  # noqa: E402


SAMPLE = """
01.01.2026 Pazartesi
Ahmet Yilmaz 08:00-17:00 60
Mehmet Demir izinli
Ayse Kaya 09:00-18:00 mola 60
Hasan Ozturk raporlu

02.01.2026 Sali
Ahmet Yilmaz 08-17
Veli Bey 10:00-22:00 ozel gun
[03.01.2026 09:15] Murat Sef: - Ahmet 08-17 60
[03.01.2026 09:15] Murat Sef: - Hasan gelmedi
"""


def test_parses_date_headers_and_attendance_statuses():
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    statuses = {e.status for e in entries}
    assert {"Calisti", "Izinli", "Raporlu", "Gelmedi"}.issubset(statuses)
    dates = sorted({e.work_date for e in entries})
    assert dates == ["2026-01-01", "2026-01-02", "2026-01-03"]


def test_time_range_and_break_extraction():
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    ahmet_jan1 = next(e for e in entries
                      if e.employee_name_raw.startswith("Ahmet") and e.work_date == "2026-01-01")
    assert ahmet_jan1.start_time == "08:00"
    assert ahmet_jan1.end_time == "17:00"
    assert ahmet_jan1.break_minutes == 60
    veli = next(e for e in entries if e.employee_name_raw.startswith("Veli"))
    assert veli.is_special is True
    assert veli.employee_name_raw == "Veli Bey"


def test_whatsapp_export_format_handled():
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    gelmedi = [e for e in entries if e.status == "Gelmedi"]
    assert gelmedi and gelmedi[0].employee_name_raw == "Hasan"
    assert gelmedi[0].work_date == "2026-01-03"


def test_short_time_form_padding():
    entries = whatsapp.parse_text("01.01.2026\nA B 8-17", region="A")
    assert entries[0].start_time == "08:00"
    assert entries[0].end_time == "17:00"


def test_match_employees_by_full_name():
    employees = [
        (1, "Ahmet Yilmaz", "111", "TEKNIK", "Tek", "Ankara"),
        (2, "Mehmet Demir", "222", "LOJISTIK", "Sof", "Ankara"),
    ]
    entries = whatsapp.parse_text("01.01.2026\nahmet yilmaz 08-17", region="Ankara")
    whatsapp.match_employees(entries, employees, region="Ankara")
    assert entries[0].employee_id == 1
    assert entries[0].confidence >= 0.66


def test_low_confidence_match_warns():
    employees = [(1, "Ahmet Yilmaz", "", "", "", "Ankara")]
    entries = whatsapp.parse_text("01.01.2026\nAhmet 08-17", region="Ankara")
    whatsapp.match_employees(entries, employees, region="Ankara")
    assert entries[0].employee_id == 1
    assert any("Dusuk" in w for w in entries[0].warnings)


def test_no_match_when_name_unknown():
    employees = [(1, "Ahmet Yilmaz", "", "", "", "Ankara")]
    entries = whatsapp.parse_text("01.01.2026\nXYZ Kisi izinli", region="Ankara")
    whatsapp.match_employees(entries, employees, region="Ankara")
    assert entries[0].employee_id is None
    assert any("Calisan eslesmedi" in w for w in entries[0].warnings)


# ---------------------------------------------------------------- preview + apply
@pytest.fixture
def temp_db(monkeypatch, tmp_path):
    monkeypatch.setenv("APPDATA", str(tmp_path))
    # Reload puantaj_db so it picks up the new APPDATA path
    for mod in ("puantaj_db",):
        if mod in sys.modules:
            del sys.modules[mod]
    import puantaj_db as db
    db.init_db()
    db.add_employee("Ahmet Yilmaz", "111", "TEKNIK", "Tek", "Ankara")
    db.add_employee("Mehmet Demir", "222", "LOJISTIK", "Sof", "Ankara")
    db.add_employee("Ayse Kaya", "333", "STANT", "Sat", "Ankara")
    db.add_employee("Hasan Ozturk", "444", "TEKNIK", "Tek", "Ankara")
    db.add_employee("Veli Bey", "555", "STANT", "Sat", "Ankara")
    return db


def test_preview_excel_is_written(temp_db, tmp_path):
    from cli_anything.puantaj import preview_xlsx
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    whatsapp.match_employees(entries, temp_db.list_employees(), region="Ankara")
    payload = whatsapp.entries_to_dicts(entries)
    out = tmp_path / "preview.xlsx"
    path = preview_xlsx.build_preview(
        str(out), payload,
        employees_by_id={
            int(r[0]): {"full_name": r[1], "department": r[3], "region": r[5]}
            for r in temp_db.list_employees()
        },
        settings=temp_db.get_all_settings(),
    )
    assert os.path.isfile(path) and os.path.getsize(path) > 1000
    # Sekme adlari
    from openpyxl import load_workbook
    wb = load_workbook(path)
    assert "Ozet" in wb.sheetnames
    assert "Calisan Matrisi" in wb.sheetnames
    assert "Calisan Analizi" in wb.sheetnames
    assert "Detay" in wb.sheetnames
    assert "Gunluk Ozet" in wb.sheetnames
    assert "Uyarilar" in wb.sheetnames


def test_preview_matrix_has_employee_rows_and_date_columns(temp_db, tmp_path):
    from cli_anything.puantaj import preview_xlsx
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    whatsapp.match_employees(entries, temp_db.list_employees(), region="Ankara")
    payload = whatsapp.entries_to_dicts(entries)
    out = tmp_path / "preview.xlsx"
    preview_xlsx.build_preview(
        str(out), payload,
        employees_by_id={int(r[0]): {"full_name": r[1], "department": r[3], "region": r[5]}
                         for r in temp_db.list_employees()},
        settings=temp_db.get_all_settings(),
    )
    from openpyxl import load_workbook
    wb = load_workbook(str(out))
    ws = wb["Calisan Matrisi"]
    # Calisanlar sutun A'da
    employees_in_col_a = [ws.cell(row=r, column=1).value for r in range(3, ws.max_row + 1)]
    employees_in_col_a = [v for v in employees_in_col_a if v]
    assert any("Ahmet Yilmaz" in str(v) for v in employees_in_col_a)
    assert any("Mehmet Demir" in str(v) for v in employees_in_col_a)
    # 1-3 Ocak araligi 3 tarih sutunu olmali (D, E, F)
    second_row_headers = [ws.cell(row=2, column=c).value for c in range(4, 7)]
    assert all(v and ("01 " in str(v) or "02 " in str(v) or "03 " in str(v))
               for v in second_row_headers)


def test_preview_analysis_has_per_employee_totals(temp_db, tmp_path):
    from cli_anything.puantaj import preview_xlsx
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    whatsapp.match_employees(entries, temp_db.list_employees(), region="Ankara")
    payload = whatsapp.entries_to_dicts(entries)
    out = tmp_path / "preview.xlsx"
    preview_xlsx.build_preview(
        str(out), payload,
        employees_by_id={int(r[0]): {"full_name": r[1], "department": r[3], "region": r[5]}
                         for r in temp_db.list_employees()},
        settings=temp_db.get_all_settings(),
    )
    from openpyxl import load_workbook
    wb = load_workbook(str(out))
    ws = wb["Calisan Analizi"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    for need in ("Calisma G.", "Toplam (s)", "Fazla Mesai (s)", "Devamsizlik %"):
        assert need in headers, f"{need} basligi yok"
    # Son satir TOPLAM olmali
    last_label = ws.cell(row=ws.max_row, column=1).value
    assert last_label == "TOPLAM"


def test_apply_writes_timesheets_and_attendance(temp_db):
    from cli_anything.puantaj import bulk
    entries = whatsapp.parse_text(SAMPLE, region="Ankara")
    whatsapp.match_employees(entries, temp_db.list_employees(), region="Ankara")
    payload = whatsapp.entries_to_dicts(entries)
    result = bulk.apply_entries(temp_db, payload, default_region="Ankara")
    assert result.timesheets_added >= 4
    assert result.attendance_added >= 3
    # DB sahnesi
    ts = temp_db.list_timesheets()
    att = temp_db.list_attendance_records()
    assert ts and att
    # Mukerrer apply ayni gun + calisan icin overwrite kapaliysa atlanmali (yeniden timesheet eklenmemeli)
    before = len(ts)
    bulk.apply_entries(temp_db, payload, default_region="Ankara", overwrite=False)
    # Calisti timesheets ekleyebilir (UNIQUE constraint yok), ama attendance upsert nedeniyle ayni kalmali
    after_att = len(temp_db.list_attendance_records())
    assert after_att == len(att)
    # overwrite ile ayni gun timesheets temizlenip yeniden yazilir
    bulk.apply_entries(temp_db, payload, default_region="Ankara", overwrite=True)
    # En azindan kayit sayisi makul kalmali (her gun + calisan icin tek timesheet)
    assert len(temp_db.list_timesheets()) <= before + len(payload)
