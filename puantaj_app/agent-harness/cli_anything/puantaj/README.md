# cli-anything-puantaj

Rainstaff puantaj uygulamasinin agent-kullanilabilir komut satiri arayuzu.

GUI'yi (Tkinter) taklit etmez. Uygulamanin kendi `puantaj_db`, `calc`, `report`
modullerini import edip cagirir; GUI ile **ayni** SQLite veritabanini kullanir
(`%APPDATA%\Rainstaff\data\puantaj.db`). CLI'dan girilen veri GUI'de gorunur.

## Kurulum

```powershell
cd C:\Users\canor\OneDrive\Masaüstü\puantaj\puantaj_app\agent-harness
pip install -e .
# Excel raporu icin opsiyonel:
pip install -e ".[report]"
```

`cli-anything-puantaj` PATH'e kurulur.

Kaynak klasor otomatik bulunur (harness, puantaj_app icinde). Gerekirse:
`set PUANTAJ_APP_DIR=C:\...\puantaj_app`

## Kullanim

```powershell
cli-anything-puantaj --help
cli-anything-puantaj info
cli-anything-puantaj                       # argumansiz: REPL modu

cli-anything-puantaj --json employee list
cli-anything-puantaj employee add --full-name "AHMET YILMAZ" --department TEKNIK --title TSG --region Ankara
cli-anything-puantaj timesheet add --employee-id 5 --work-date 2026-05-19 --start 10:00 --end 18:00 --region Ankara
cli-anything-puantaj --json template list
cli-anything-puantaj timesheet add --employee-id 5 --work-date 2026-05-19 --template "Hafta Ici 8.30-17.30" --region Ankara
cli-anything-puantaj timesheet add --employee-id 5 --work-date 2026-05-19 --template "Hafta Ici 8.30-17.30" --end 18:30 --region Ankara
cli-anything-puantaj --json timesheet calc --work-date 2026-05-19 --start 08:00 --end 19:00 --break-minutes 60
cli-anything-puantaj --json timesheet calc --work-date 2026-05-19 --template "Hafta Ici 8.30-17.30" --end 18:30
cli-anything-puantaj --json timesheet list --start-date 2026-05-01 --end-date 2026-05-31
cli-anything-puantaj report -o rapor.xlsx --start-date 2026-05-01 --end-date 2026-05-31
```

`--json` bayragi tum komutlarda makine okunur UTF-8 cikti verir (agent tuketimi icin).
`--template` verilirse eksik `--start`, `--end` ve `--break-minutes` degerleri
vardiya sablonundan doldurulur. Saatleri elle verip sadece sablon molasini
kullanmak icin `--template` ile birlikte `--start/--end` verilebilir.
`--template` verilmezse CLI varsayilan olarak otomatik sablon secer:
LOJISTIK hafta ici `sofor`, TEKNIK hafta ici `Hafta Ici 8.30-17.30`,
LOJISTIK/TEKNIK cumartesi `Cumartesi 09-14`; STANT icin saatlere gore
`stant 10-18`, `stant 14-22` veya `stant full`. Bilinen resmi tatillerde
`--special` otomatik isaretlenir ve nota ozel gun calismasi eklenir.
2026 icin Kurban Bayrami Arefesi `2026-05-26`, Kurban Bayrami
`2026-05-27` - `2026-05-30` ozel gun listesine dahildir.

## Komut gruplari

- `info` - DB yolu ve ozet sayilar
- `employee` - list / add / update / delete
- `timesheet` - list / add / delete / calc (calc DB'ye yazmaz)
- `attendance` - set / list
- `leave` - list / add
- `vehicle` - list
- `driver` - list
- `template` - list / set / delete
- `settings` - list / set
- `users` - listele
- `report` - Excel rapor (openpyxl gerekir)

## Uyari

Komutlar gercek uretim veritabanina yazar. Test/deneme oncesi yedek alin:
`Copy-Item "$env:APPDATA\Rainstaff\data\puantaj.db" "$env:APPDATA\Rainstaff\data\puantaj.db.bak"`
