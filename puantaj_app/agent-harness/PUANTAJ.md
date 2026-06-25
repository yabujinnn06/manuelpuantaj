# PUANTAJ - cli-anything analiz ve SOP

## Phase 1: Kod analizi

- Uygulama: Rainstaff puantaj, Tkinter masaustu GUI (`app.py`, 8249 satir).
- Backend ayri modullerde, GUI sadece sunum:
  - `puantaj_db.py` - SQLite veri katmani (employees, timesheets,
    attendance_records, leave_records, vehicles, drivers, users, settings).
  - `calc.py` - saf hesaplama (calc_day_hours: calisma/plan/fazla/gece).
  - `report.py` - Excel/PDF cikti (export_report, openpyxl).
- Veri modeli: SQLite. Windows'ta gercek yol `%APPDATA%\Rainstaff\data\puantaj.db`
  (puantaj_db.py satir 21-23). `puantaj_app/data/` lokal kopya, GUI bunu kullanmaz.
- Harici binary yok: "gercek yazilim" burada Python modullerinin kendisi.
  CLI bu modulleri import edip cagirir, hicbir mantik yeniden yazilmaz.

## Phase 2: CLI mimarisi

- Etkilesim: hem tek-atim subcommand hem REPL (argumansiz REPL default).
- Durum modeli: SQLite DB'nin kendisi kalici durum; ayri session JSON yok.
- Cikti: insan-okunur tablo + `--json` (UTF-8) makine-okunur.
- Komut gruplari uygulamanin domainlerine birebir: employee, timesheet,
  attendance, leave, vehicle, driver, template, settings, users, report.

## Phase 3: Uygulama

- `cli_anything/puantaj/puantaj_cli.py` Click tabanli. `_find_app_dir()` ile
  kaynak klasoru bulur (PUANTAJ_APP_DIR override), sys.path'e ekler, import eder.
- `timesheet add/calc --template` vardiya sablonunu ad veya ID ile bulur; eksik
  giris/cikis/mola alanlarini sablondan doldurur.
- `--template` verilmezse CLI departman/gun/saat kuralina gore otomatik sablon
  secer. Bilinen resmi tatillerde `is_special` otomatik isaretlenir ve nota
  ozel gun calismasi eklenir. 2026 Kurban Bayrami Arefesi (2026-05-26) ve
  Kurban Bayrami (2026-05-27 - 2026-05-30) yil bazli ozel gun olarak tanimli.
- Rapor modulu lazy import (openpyxl olmadan da temel komutlar calisir).
- stdout/stderr UTF-8'e reconfigure edilir (Windows konsol codepage bagimsiz).

## Phase 4-6: Test

- `tests/test_core.py` + `tests/TEST.md`. calc saf testleri + CLI subprocess
  salt-okunur testleri. 8/8 gecti. Yazma komutlari uretim DB'sini korumak
  icin otomatik test disinda.

## Phase 7: Paketleme

- `setup.py`, PEP 420 namespace: `cli_anything.puantaj`,
  paket adi `cli-anything-puantaj`, console_scripts entry point.
- `pip install -e .` ile kuruldu, PATH'te dogrulandi.
