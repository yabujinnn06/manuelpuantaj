# -*- coding: utf-8 -*-
"""Excel'deki (mayis_kontrol.rows) NORMAL kayitlari puantaj DB'sine isler."""
import sys, os

sys.path.insert(0, r"C:\Users\canor\OneDrive\Masaüstü\puantaj")
import mayis_kontrol  # ayni veri kaynagi; rows uretir

sys.path.insert(0, r"C:\Users\canor\OneDrive\Masaüstü\puantaj\puantaj_app")
import puantaj_db as db

NAME2ID = {
    "Hasan Tontur": 3, "Ata Türkbey": 8, "Uğur Ertürk": 1, "Ercüment Çalışkan": 2,
    "Eda Nur Yılmaz": 5, "Kübracan Gündoğdu": 9, "Başak Çelik": 6,
    "Leyla Karayağız": 10,
}

db.init_db()

# Mukerrer korumasi: mevcut (employee_id, work_date) seti
existing = set()
for r in db.list_timesheets(start_date="2026-05-01", end_date="2026-05-31"):
    existing.add((r[1], r[4]))  # employee_id, work_date

added, dup, skipped, errors = 0, 0, [], []
for ds, person, g, c, sure, durum, note in mayis_kontrol.rows:
    if durum != "NORMAL":
        skipped.append((ds, person, durum))
        continue
    eid = NAME2ID[person]
    if (eid, ds) in existing:
        dup += 1
        continue
    notes = "WhatsApp Mayis 2026" + (" | " + note if note else "")
    try:
        db.add_timesheet(eid, ds, g, c, 0, 0, notes, "Ankara")
        existing.add((eid, ds))
        added += 1
    except Exception as e:
        errors.append((ds, person, str(e)))

print("Eklenen puantaj:", added)
print("Mukerrer (atlandi):", dup)
print("Izin/calismadi (puantaj degil, atlandi):", len(skipped))
if errors:
    print("HATALAR:", errors)
