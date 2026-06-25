---
name: "cli-anything-puantaj"
description: "Rainstaff puantaj uygulamasini komut satirindan kontrol et: personel, puantaj/mesai, devamsizlik, izin, arac, surucu, Excel rapor. GUI ile ayni SQLite veritabanini kullanir."
---

# cli-anything-puantaj

Rainstaff puantaj (timesheet/bordro) uygulamasinin agent-kullanilabilir CLI'i.
Uygulamanin kendi Python modullerini (`puantaj_db`, `calc`, `report`) cagirir,
GUI ile **ayni** veritabanina yazar: `%APPDATA%\Rainstaff\data\puantaj.db`.

## Kurulum

```
cd C:\Users\canor\OneDrive\Masaüstü\puantaj\puantaj_app\agent-harness
pip install -e .            # rapor icin: pip install -e ".[report]"
```

## Agent kullanimi

Her komut `--json` ile UTF-8 makine okunur cikti verir. Once oku, sonra yaz:

```
cli-anything-puantaj --json info
cli-anything-puantaj --json employee list
cli-anything-puantaj --json template list
cli-anything-puantaj --json timesheet list --start-date 2026-05-01 --end-date 2026-05-31 --employee-id 5
cli-anything-puantaj --json timesheet calc --work-date 2026-05-19 --start 08:00 --end 19:00 --break-minutes 60
cli-anything-puantaj --json timesheet calc --work-date 2026-05-19 --template "Hafta Ici 8.30-17.30" --end 18:30
```

Yazma:

```
cli-anything-puantaj employee add --full-name "AD SOYAD" --department TEKNIK --title TSG --region Ankara
cli-anything-puantaj timesheet add --employee-id 5 --work-date 2026-05-19 --start 10:00 --end 18:00 --region Ankara
cli-anything-puantaj timesheet add --employee-id 5 --work-date 2026-05-19 --template "Hafta Ici 8.30-17.30" --end 18:30 --region Ankara
cli-anything-puantaj attendance set --employee-id 5 --work-date 2026-05-19 --status izinli --region Ankara
cli-anything-puantaj leave add --employee-id 5 --start-date 2026-06-01 --end-date 2026-06-05 --leave-type yillik --region Ankara
cli-anything-puantaj report -o C:\rapor.xlsx --start-date 2026-05-01 --end-date 2026-05-31
```

## Komut gruplari

| Grup | Komutlar |
|------|----------|
| `info` | DB yolu + ozet sayilar |
| `employee` | list, add, update <id>, delete <id> |
| `timesheet` | list, add, delete <id>, calc (yazmaz) |
| `attendance` | set, list |
| `leave` | list, add |
| `vehicle` | list |
| `driver` | list |
| `template` | list, set, delete |
| `settings` | list, set <key> <value> |
| `users` | listele |
| `report` | Excel uretir (openpyxl gerekir) |

## Agent notlari

- Bolge degerleri: Ankara, Izmir, Bursa, Istanbul. Cogu yazma komutu `--region` ister.
- Puantaj girerken once `template list` oku. `--template` verilirse eksik `--start`, `--end`
  ve `--break-minutes` degerleri sablondan gelir; saatler elle verilirse sablon molasi kullanilir.
- `--template` verilmezse otomatik sablon secilir: LOJISTIK hafta ici `sofor`,
  TEKNIK hafta ici `Hafta Ici 8.30-17.30`, LOJISTIK/TEKNIK cumartesi
  `Cumartesi 09-14`; STANT icin saatlere gore `stant 10-18`, `stant 14-22`,
  `stant full`. STANT icin hafta sonu sablonu uygulanmaz.
- `--special` ozel gun; saatler `special_*` alanlarina yazilir.
- Bilinen resmi tatillerde `--special` otomatik isaretlenir ve nota ozel gun calismasi eklenir.
  2026 Kurban Bayrami Arefesi `2026-05-26`, Kurban Bayrami `2026-05-27` - `2026-05-30`.
- `timesheet calc` sadece hesaplar, DB'ye dokunmaz - guvenli on kontrol.
- Komutlar gercek uretim verisine yazar. Toplu islemden once DB yedegi al.
- Hata durumunda komut sifirdan farkli kod ve net mesaj doner; agent kendini duzeltmeli.
