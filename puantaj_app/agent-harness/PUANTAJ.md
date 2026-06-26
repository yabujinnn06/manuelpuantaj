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

## Phase 8: WhatsApp toplu puantaj (yeni)

Modul: `cli_anything/puantaj/whatsapp.py`, `preview_xlsx.py`, `bulk.py`.

Akis (parse -> preview -> apply):

1. `cli-anything-puantaj whatsapp parse --input msg.txt --region Ankara --out-json p.json`
   - WhatsApp disa aktarim (`[01.01.2026 09:15] Sender: ...`) ve serbest grup
     metnini ayristirir, tarih basliklarini takip eder, calisanlari fuzzy
     (token-Jaccard) ile DB'deki kayitlarla eslestirir.
   - status: Calisti / Izinli / Raporlu / Gelmedi / Mazeret / Tatil / Diger.
   - "08-17", "08:00-17:00", "mola 60", "60 dk", "ozel gun" yakalanir.

2. `cli-anything-puantaj whatsapp preview --input msg.txt -o onay.xlsx`
   - Onay icin renkli Excel; sekmeler: Ozet, Detay, Gunluk Ozet, Uyarilar, JSON.
   - Calisti = yesil, Izinli/Mazeret = sari, Raporlu/Gelmedi = kirmizi, Tatil = mor.
   - calc.calc_day_hours ile calisilan/plan/fazla/gece sutunlari onceden hesaplanir.

3. `cli-anything-puantaj whatsapp apply --xlsx onay.xlsx --region Ankara --yes`
   - Onaylanmis Excel'in Detay sekmesindeki son hali okunur ve DB'ye yazilir.
   - `--records p.json` ile JSON dosyasi da kabul edilir.
   - `--overwrite` ayni gun + calisan icin onceki timesheet'leri siler.
   - Calisti -> timesheets + attendance_records; digerleri -> sadece attendance.

4. `cli-anything-puantaj whatsapp ingest --input msg.txt --region Ankara`
   - Tek komutla parse + preview olusturur; interaktif TTY'de kullaniciya
     onay sorar, `--yes` ile dogrudan apply eder.

## Phase 9: Yonetim komutlari (yeni)

- `department list/employees/rename/assign` - departman alanini (employees.department)
  yonetir. Departman ayri tablo olmadigi icin tum islemler bu TEXT alani uzerinden
  yapilir.
- `timesheet bulk --records p.json` - whatsapp apply ile ayni motor; JSON'dan
  toplu puantaj girisi yapar.
- `report --month YYYY-MM` - tek argumanla ay raporu uretir.
