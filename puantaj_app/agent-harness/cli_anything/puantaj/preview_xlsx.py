"""WhatsApp puantaj on izleme Excel uretici.

Parser ciktisini renkli, sekmeli, analitik bir Excel olarak verir. Kullanici
bu dosyayi inceler; uygunsa cli'a `apply` ile geri verir.

Sekmeler:
  - Ozet            : Toplam sayilar, durum dagilimi, bolge dagilimi, onay notu
  - Calisan Matrisi : Satir=Calisan, Sutun=Tarih. Hucre=kisaltma + saat
  - Calisan Analizi : Kisi basi calisma gunu, toplam/fazla/gece saat, izin/rapor
  - Detay           : Tum kayitlar + hesaplanmis saatler
  - Gunluk Ozet     : Gune gore durum dagilimi ve saatler
  - Uyarilar        : Eslemeyen calisan / eksik saat / dusuk guven uyarilari
  - JSON            : Apply icin makine okunur kayitlar
"""

from __future__ import annotations

import os
import re
from collections import Counter, defaultdict
from datetime import datetime, date, timedelta

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter


HEADER_FILL = PatternFill("solid", fgColor="1F4E78")
SECTION_FILL = PatternFill("solid", fgColor="D9E1F2")
GOOD_FILL = PatternFill("solid", fgColor="E2EFDA")
WARN_FILL = PatternFill("solid", fgColor="FFF2CC")
BAD_FILL = PatternFill("solid", fgColor="F8CBAD")
SPECIAL_FILL = PatternFill("solid", fgColor="D9D2E9")
HEADER_FONT = Font(color="FFFFFF", bold=True, size=11)
TITLE_FONT = Font(size=14, bold=True, color="1F4E78")
THIN = Side(border_style="thin", color="B0B0B0")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


STATUS_FILL = {
    "Calisti": GOOD_FILL,
    "Izinli": WARN_FILL,
    "Mazeret": WARN_FILL,
    "Tatil": SPECIAL_FILL,
    "Raporlu": BAD_FILL,
    "Gelmedi": BAD_FILL,
    "Diger": SECTION_FILL,
}

# Matris sekmesinde hucreye yazilacak kisaltma
STATUS_CODE = {
    "Calisti": "C",
    "Izinli": "I",
    "Mazeret": "M",
    "Tatil": "T",
    "Raporlu": "R",
    "Gelmedi": "G",
    "Diger": "D",
}

WEEKDAY_TR = ["Pzt", "Sal", "Cr", "Per", "Cum", "Cmt", "Paz"]

EMPTY_DAY_FILL = PatternFill("solid", fgColor="F2F2F2")
WEEKEND_HEADER_FILL = PatternFill("solid", fgColor="9DC3E6")


def _autosize(ws, max_col: int, sample_rows: int = 200):
    for col_idx in range(1, max_col + 1):
        col = get_column_letter(col_idx)
        max_len = 0
        for r in range(1, min(ws.max_row, sample_rows) + 1):
            v = ws.cell(row=r, column=col_idx).value
            if v is None:
                continue
            l = len(str(v))
            if l > max_len:
                max_len = l
        ws.column_dimensions[col].width = min(max(10, max_len + 2), 38)


def _write_headers(ws, row: int, headers: list[str]):
    for c, h in enumerate(headers, start=1):
        cell = ws.cell(row=row, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
    ws.row_dimensions[row].height = 22


_CALC_FN = None
_CALC_TRIED = False


def _load_calc_fn():
    """calc.calc_day_hours'i bulur. Caller sys.path'i ayarlamamis olsa bile
    dosya konumundan yukari dogru calc.py arayip yukler (puantaj_app/calc.py)."""
    global _CALC_FN, _CALC_TRIED
    if _CALC_TRIED:
        return _CALC_FN
    _CALC_TRIED = True
    # 1) Hazirsa dogrudan import
    try:
        from calc import calc_day_hours  # type: ignore
        _CALC_FN = calc_day_hours
        return _CALC_FN
    except Exception:
        pass
    # 2) Dosya konumundan yukari dogru calc.py ara
    import importlib.util
    here = os.path.dirname(os.path.abspath(__file__))
    cur = here
    for _ in range(8):
        for cand in (os.path.join(cur, "calc.py"),
                     os.path.join(cur, "puantaj_app", "calc.py"),
                     os.path.join(cur, "server", "calc.py")):
            if os.path.isfile(cand):
                try:
                    spec = importlib.util.spec_from_file_location("_puantaj_calc", cand)
                    mod = importlib.util.module_from_spec(spec)
                    spec.loader.exec_module(mod)  # type: ignore
                    _CALC_FN = getattr(mod, "calc_day_hours", None)
                    if _CALC_FN:
                        return _CALC_FN
                except Exception:
                    pass
        parent = os.path.dirname(cur)
        if parent == cur:
            break
        cur = parent
    return _CALC_FN


def _safe_calc(work_date, start_time, end_time, break_minutes, settings, is_special, department):
    if not start_time or not end_time:
        return None
    fn = _load_calc_fn()
    if fn is None:
        return None
    try:
        return fn(
            work_date, start_time, end_time, int(break_minutes or 0),
            settings or {}, 1 if is_special else 0, department,
        )
    except Exception:
        return None


def build_preview(
    output_path: str,
    entries: list[dict],
    employees_by_id: dict[int, dict] | None = None,
    settings: dict | None = None,
    title: str = "WhatsApp Puantaj On Izleme",
    shift_templates: list[tuple] | None = None,
) -> str:
    """entries: whatsapp.entries_to_dicts ciktisi.

    employees_by_id: {id: {full_name, department, region}} sozlugu.
    settings: db.get_all_settings() (calc icin).
    shift_templates: db.list_shift_templates() ciktisi. Vardiya tanima icin.

    Return: yazilan dosyanin tam yolu.
    """
    employees_by_id = employees_by_id or {}
    settings = settings or {}
    shift_templates = shift_templates or []
    wb = Workbook()

    # 1) Ozet sekmesi
    ws = wb.active
    ws.title = "Ozet"

    ws.cell(row=1, column=1, value=title).font = TITLE_FONT
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=6)
    ws.row_dimensions[1].height = 28

    created_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ws.cell(row=2, column=1, value=f"Olusturuldu: {created_at}").font = Font(italic=True, color="555555")
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=6)

    total = len(entries)
    matched = sum(1 for e in entries if e.get("employee_id"))
    unmatched = total - matched
    warned = sum(1 for e in entries if e.get("warnings"))
    status_counts = Counter(e.get("status", "Diger") for e in entries)
    region_counts = Counter(e.get("region") or "(yok)" for e in entries)
    dates = sorted({e.get("work_date") for e in entries if e.get("work_date")})
    date_range = f"{dates[0]} – {dates[-1]}" if dates else "(tarih yok)"

    summary_rows = [
        ("Toplam kayit", total),
        ("Eslemeyen calisan", unmatched),
        ("Uyarili kayit", warned),
        ("Tarih araligi", date_range),
        ("Farkli gun sayisi", len(dates)),
    ]
    for status, count in status_counts.most_common():
        summary_rows.append((f"Durum: {status}", count))
    for region, count in region_counts.most_common():
        summary_rows.append((f"Bolge: {region}", count))

    row = 4
    ws.cell(row=row, column=1, value="OZET").font = Font(bold=True, color="FFFFFF")
    ws.cell(row=row, column=1).fill = HEADER_FILL
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=2)
    row += 1
    for label, value in summary_rows:
        ws.cell(row=row, column=1, value=label).font = Font(bold=True)
        ws.cell(row=row, column=1).fill = SECTION_FILL
        ws.cell(row=row, column=2, value=value)
        row += 1
    row += 1

    # Onay kutusu acikIamasi
    ws.cell(row=row, column=1, value="Onay:").font = Font(bold=True)
    ws.cell(row=row, column=2, value="Bu dosyayi inceleyip kabul ediyorsaniz CLI'a `whatsapp apply` ile geri verin.")
    row += 1
    ws.cell(row=row, column=2, value="(Onceden olusturulmus JSON ile: `cli-anything-puantaj whatsapp apply --records records.json`)")
    row += 1
    ws.cell(row=row, column=2, value="Kayitlari duzeltmeniz gerekirse Detay sekmesindeki degerleri editleyip --xlsx ile gonderebilirsiniz.")

    _autosize(ws, 6)

    # 2) Calisan Matrisi sekmesi (satir=calisan, sutun=tarih)
    _build_matrix_sheet(wb, entries, employees_by_id, settings)

    # 3) Calisan Analizi sekmesi
    _build_analysis_sheet(wb, entries, employees_by_id, settings)

    # 3a) Her calisan icin ayri detay sekmesi
    _build_per_employee_sheets(wb, entries, employees_by_id, settings, shift_templates)

    ws = wb.create_sheet("Detay")
    headers = [
        "#", "Tarih", "Calisan (mesaj)", "Eslesen Calisan", "Departman", "Bolge",
        "Durum", "Giris", "Cikis", "Mola (dk)", "Ozel Gun",
        "Calisilan (s)", "Plan (s)", "Fazla Mesai (s)", "Gece (s)",
        "Guven", "Notlar", "Uyarilar",
    ]
    _write_headers(ws, 1, headers)
    out_row = 2
    for idx, e in enumerate(entries, start=1):
        emp = employees_by_id.get(e.get("employee_id") or -1) or {}
        department = emp.get("department") or ""
        region = e.get("region") or emp.get("region") or ""
        calc_res = _safe_calc(
            e.get("work_date"), e.get("start_time"), e.get("end_time"),
            e.get("break_minutes") or 0, settings,
            e.get("is_special"), department,
        )
        worked = calc_res[0] if calc_res else ""
        scheduled = calc_res[1] if calc_res else ""
        overtime = calc_res[2] if calc_res else ""
        night = calc_res[3] if calc_res else ""

        values = [
            idx,
            e.get("work_date") or "",
            e.get("employee_name_raw") or "",
            e.get("matched_name") or "",
            department,
            region,
            e.get("status") or "",
            e.get("start_time") or "",
            e.get("end_time") or "",
            e.get("break_minutes") or 0,
            "Evet" if e.get("is_special") else "Hayir",
            worked, scheduled, overtime, night,
            e.get("confidence") or "",
            e.get("notes") or "",
            " | ".join(e.get("warnings") or []),
        ]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=out_row, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center")
        # Renklendir
        status_fill = STATUS_FILL.get(e.get("status"), None)
        if not e.get("employee_id"):
            status_fill = BAD_FILL
        elif e.get("warnings"):
            status_fill = WARN_FILL
        if status_fill:
            for c in range(1, len(headers) + 1):
                ws.cell(row=out_row, column=c).fill = status_fill
        out_row += 1

    ws.freeze_panes = "A2"
    _autosize(ws, len(headers))

    # 3) Tarih bazli ozet
    ws = wb.create_sheet("Gunluk Ozet")
    day_buckets: dict[str, dict] = defaultdict(lambda: {
        "calisti": 0, "izinli": 0, "raporlu": 0, "gelmedi": 0,
        "mazeret": 0, "tatil": 0, "diger": 0, "toplam_s": 0.0,
        "fazla_s": 0.0,
    })
    for e in entries:
        d = e.get("work_date") or "(tarih yok)"
        b = day_buckets[d]
        key = (e.get("status") or "Diger").lower()
        if key not in b:
            key = "diger"
        b[key] = b[key] + 1
        emp = employees_by_id.get(e.get("employee_id") or -1) or {}
        calc_res = _safe_calc(
            e.get("work_date"), e.get("start_time"), e.get("end_time"),
            e.get("break_minutes") or 0, settings,
            e.get("is_special"), emp.get("department"),
        )
        if calc_res:
            b["toplam_s"] += calc_res[0]
            b["fazla_s"] += calc_res[2]

    headers = ["Tarih", "Calisti", "Izinli", "Raporlu", "Gelmedi", "Mazeret",
               "Tatil", "Diger", "Toplam (s)", "Fazla Mesai (s)"]
    _write_headers(ws, 1, headers)
    r = 2
    for d in sorted(day_buckets.keys()):
        b = day_buckets[d]
        values = [d, b["calisti"], b["izinli"], b["raporlu"], b["gelmedi"],
                  b["mazeret"], b["tatil"], b["diger"],
                  round(b["toplam_s"], 2), round(b["fazla_s"], 2)]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.border = BORDER
        r += 1
    ws.freeze_panes = "A2"
    _autosize(ws, len(headers))

    # 4) Uyarilar
    ws = wb.create_sheet("Uyarilar")
    headers = ["#", "Tarih", "Mesajdaki Isim", "Eslesen", "Durum", "Uyari"]
    _write_headers(ws, 1, headers)
    r = 2
    for idx, e in enumerate(entries, start=1):
        for w in (e.get("warnings") or []):
            ws.cell(row=r, column=1, value=idx)
            ws.cell(row=r, column=2, value=e.get("work_date") or "")
            ws.cell(row=r, column=3, value=e.get("employee_name_raw") or "")
            ws.cell(row=r, column=4, value=e.get("matched_name") or "")
            ws.cell(row=r, column=5, value=e.get("status") or "")
            ws.cell(row=r, column=6, value=w)
            for c in range(1, len(headers) + 1):
                ws.cell(row=r, column=c).border = BORDER
                ws.cell(row=r, column=c).fill = WARN_FILL
            r += 1
    if r == 2:
        ws.cell(row=2, column=1, value="(Uyari yok)").font = Font(italic=True, color="888888")
    _autosize(ws, len(headers))

    # 5) JSON (apply icin)
    ws = wb.create_sheet("JSON")
    ws.cell(row=1, column=1, value="Apply icin JSON kopyalanabilir veya --records ile dosyadan verilir.").font = Font(italic=True)
    import json
    payload = json.dumps(entries, ensure_ascii=False, indent=2, default=str)
    chunks = [payload[i:i + 4000] for i in range(0, len(payload), 4000)] or [""]
    for i, ch in enumerate(chunks, start=2):
        ws.cell(row=i, column=1, value=ch)
    ws.column_dimensions["A"].width = 90

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir and not os.path.isdir(out_dir):
        os.makedirs(out_dir, exist_ok=True)
    wb.save(output_path)
    return os.path.abspath(output_path)


def read_entries_from_xlsx(path: str) -> list[dict]:
    """Onaylanmis xlsx Detay sekmesinden duzenlenmis entries okur.

    Kullanici Detay sekmesindeki "Eslesen Calisan" sutununu doldurmus veya
    "Durum/Giris/Cikis/Mola" alanlarini guncellemis olabilir. Bu fonksiyon
    bu duzeltmeleri JSON sekmesindeki bazi alanlari kaybetmeden geri okur.
    """
    from openpyxl import load_workbook
    wb = load_workbook(path, data_only=True)
    if "Detay" not in wb.sheetnames:
        raise ValueError("Onizleme dosyasinda 'Detay' sekmesi yok.")
    ws = wb["Detay"]
    headers = [ws.cell(row=1, column=c).value for c in range(1, ws.max_column + 1)]
    h = {str(v).strip(): i + 1 for i, v in enumerate(headers) if v is not None}

    def get(row: int, name: str):
        col = h.get(name)
        if not col:
            return None
        return ws.cell(row=row, column=col).value

    out: list[dict] = []
    for r in range(2, ws.max_row + 1):
        if get(r, "Tarih") is None and get(r, "Calisan (mesaj)") is None:
            continue
        rec = {
            "work_date": str(get(r, "Tarih") or "").strip()[:10] or None,
            "employee_name_raw": str(get(r, "Calisan (mesaj)") or "").strip(),
            "matched_name": str(get(r, "Eslesen Calisan") or "").strip() or None,
            "region": str(get(r, "Bolge") or "").strip() or None,
            "status": str(get(r, "Durum") or "").strip() or None,
            "start_time": _normalize_time(get(r, "Giris")),
            "end_time": _normalize_time(get(r, "Cikis")),
            "break_minutes": int(get(r, "Mola (dk)") or 0),
            "is_special": str(get(r, "Ozel Gun") or "").strip().lower() in ("evet", "yes", "true", "1"),
            "notes": str(get(r, "Notlar") or "").strip(),
            "warnings": [w for w in str(get(r, "Uyarilar") or "").split(" | ") if w],
            "employee_id": None,
            "confidence": float(get(r, "Guven") or 0.0) if get(r, "Guven") else 0.0,
        }
        out.append(rec)
    return out


# ============================================================================
# Calisan x Tarih matrisi
# ============================================================================

def _iter_dates(start: date, end: date):
    d = start
    while d <= end:
        yield d
        d += timedelta(days=1)


def _employee_label(entry: dict, employees_by_id: dict) -> tuple[str, str]:
    """Return (key, display_name). Eslemeyen kayitlar 'raw' ismiyle gruplanir."""
    eid = entry.get("employee_id")
    if eid:
        info = employees_by_id.get(int(eid)) or {}
        name = info.get("full_name") or entry.get("matched_name") or entry.get("employee_name_raw") or f"#{eid}"
        return f"id:{eid}", name
    raw = (entry.get("employee_name_raw") or "?").strip()
    return f"raw:{raw.lower()}", f"{raw} (eslemeyen)"


def _parse_date_str(s: str | None) -> date | None:
    if not s or len(s) < 10:
        return None
    try:
        return date.fromisoformat(s[:10])
    except ValueError:
        return None


def _build_matrix_sheet(wb, entries: list[dict], employees_by_id: dict, settings: dict):
    """Satir=calisan, sutun=tarih. Her hucre durum kisaltmasi + (varsa) calisilan saat.

    Sonda her satira: Calisma, Izin, Rapor, Gelmedi gun sayilari ve toplam saat.
    """
    ws = wb.create_sheet("Calisan Matrisi", index=1)

    dates = sorted({_parse_date_str(e.get("work_date"))
                    for e in entries if _parse_date_str(e.get("work_date"))})
    if not dates:
        ws.cell(row=1, column=1, value="(Tarih bulunamadi)").font = Font(italic=True, color="888888")
        return
    span = list(_iter_dates(dates[0], dates[-1]))

    # Satir basliklari (calisan listesi) - kararli sirali
    rows: dict[str, dict] = {}
    for e in entries:
        key, label = _employee_label(e, employees_by_id)
        if key not in rows:
            info = employees_by_id.get(int(e["employee_id"])) if e.get("employee_id") else {}
            rows[key] = {
                "label": label,
                "department": (info or {}).get("department") or "",
                "region": (info or {}).get("region") or e.get("region") or "",
                "days": {},  # date -> {"status": ..., "hours": ..., "is_special": ...}
                "calisma": 0, "izin": 0, "rapor": 0, "gelmedi": 0,
                "mazeret": 0, "tatil": 0, "diger": 0,
                "toplam_s": 0.0, "fazla_s": 0.0, "gece_s": 0.0,
            }
        d = _parse_date_str(e.get("work_date"))
        if not d:
            continue
        info = employees_by_id.get(int(e["employee_id"])) if e.get("employee_id") else {}
        calc_res = _safe_calc(
            e.get("work_date"), e.get("start_time"), e.get("end_time"),
            e.get("break_minutes") or 0, settings,
            e.get("is_special"), (info or {}).get("department"),
        )
        worked = calc_res[0] if calc_res else 0.0
        overtime = calc_res[2] if calc_res else 0.0
        night = calc_res[3] if calc_res else 0.0
        status = e.get("status") or "Diger"
        rows[key]["days"][d] = {
            "status": status,
            "hours": worked,
            "overtime": overtime,
            "is_special": bool(e.get("is_special")),
        }
        if status == "Calisti":
            rows[key]["calisma"] += 1
            rows[key]["toplam_s"] += worked
            rows[key]["fazla_s"] += overtime
            rows[key]["gece_s"] += night
        elif status == "Izinli":
            rows[key]["izin"] += 1
        elif status == "Raporlu":
            rows[key]["rapor"] += 1
        elif status == "Gelmedi":
            rows[key]["gelmedi"] += 1
        elif status == "Mazeret":
            rows[key]["mazeret"] += 1
        elif status == "Tatil":
            rows[key]["tatil"] += 1
        else:
            rows[key]["diger"] += 1

    # Basliklar
    fixed_left = ["Calisan", "Departman", "Bolge"]
    fixed_right = ["Calisma", "Izin", "Rapor", "Gelmedi", "Mazeret",
                   "Tatil", "Diger", "Toplam (s)", "Fazla (s)", "Gece (s)"]

    # Iki satirli baslik: ust satir = ay/yil (her ayin ilk gunune ad yazilir), alt = "dd Gun"
    for c, h in enumerate(fixed_left, start=1):
        cell = ws.cell(row=1, column=c, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.merge_cells(start_row=1, start_column=c, end_row=2, end_column=c)

    date_col_start = len(fixed_left) + 1
    last_month_label = None
    month_start_col = date_col_start
    for i, d in enumerate(span):
        col = date_col_start + i
        # Alt satir: gun + kisa weekday
        sub = ws.cell(row=2, column=col, value=f"{d.day:02d} {WEEKDAY_TR[d.weekday()]}")
        sub.alignment = Alignment(horizontal="center")
        sub.font = Font(size=9, bold=True)
        sub.border = BORDER
        if d.weekday() >= 5:
            sub.fill = WEEKEND_HEADER_FILL
        else:
            sub.fill = SECTION_FILL
        # Ust satir: ay/yil etiketi, ardisik gunler birlestirilir
        month_label = d.strftime("%Y-%m")
        if month_label != last_month_label:
            if last_month_label is not None and month_start_col < col:
                _merge_month_header(ws, last_month_label, month_start_col, col - 1)
            last_month_label = month_label
            month_start_col = col
    # Son ay icin merge
    if last_month_label is not None and month_start_col <= date_col_start + len(span) - 1:
        _merge_month_header(ws, last_month_label, month_start_col, date_col_start + len(span) - 1)

    right_col_start = date_col_start + len(span)
    for i, h in enumerate(fixed_right):
        col = right_col_start + i
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BORDER
        ws.merge_cells(start_row=1, start_column=col, end_row=2, end_column=col)

    ws.row_dimensions[1].height = 18
    ws.row_dimensions[2].height = 22

    # Veri satirlari
    out_row = 3
    sorted_keys = sorted(rows.keys(), key=lambda k: rows[k]["label"].casefold())
    for key in sorted_keys:
        r = rows[key]
        ws.cell(row=out_row, column=1, value=r["label"]).border = BORDER
        ws.cell(row=out_row, column=2, value=r["department"]).border = BORDER
        ws.cell(row=out_row, column=3, value=r["region"]).border = BORDER
        if key.startswith("raw:"):
            for c in range(1, 4):
                ws.cell(row=out_row, column=c).fill = BAD_FILL

        for i, d in enumerate(span):
            col = date_col_start + i
            day = r["days"].get(d)
            cell = ws.cell(row=out_row, column=col)
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BORDER
            if not day:
                cell.value = ""
                if d.weekday() >= 5:
                    cell.fill = SECTION_FILL
                else:
                    cell.fill = EMPTY_DAY_FILL
                continue
            code = STATUS_CODE.get(day["status"], "?")
            hours = day["hours"] or 0.0
            if day["status"] == "Calisti" and hours:
                cell.value = f"{code} {hours:g}"
            else:
                cell.value = code
            fill = STATUS_FILL.get(day["status"]) or SECTION_FILL
            if day["is_special"]:
                fill = SPECIAL_FILL
            cell.fill = fill
            cell.font = Font(size=9, bold=(day["status"] != "Calisti"))

        # Sag toplam sutunlari
        right_values = [
            r["calisma"], r["izin"], r["rapor"], r["gelmedi"], r["mazeret"],
            r["tatil"], r["diger"],
            round(r["toplam_s"], 2), round(r["fazla_s"], 2), round(r["gece_s"], 2),
        ]
        for i, v in enumerate(right_values):
            cell = ws.cell(row=out_row, column=right_col_start + i, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center")
            cell.font = Font(bold=True)
        out_row += 1

    # Genis sutun: ilk 3 ve sag tarafta totaller; tarih sutunlari dar
    ws.column_dimensions["A"].width = 28
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 10
    for i in range(len(span)):
        ws.column_dimensions[get_column_letter(date_col_start + i)].width = 7
    for i in range(len(fixed_right)):
        ws.column_dimensions[get_column_letter(right_col_start + i)].width = 10

    # Dondurma + filtre
    ws.freeze_panes = ws.cell(row=3, column=date_col_start).coordinate

    # Legend (en alta)
    legend_row = out_row + 1
    ws.cell(row=legend_row, column=1, value="Kisaltma:").font = Font(bold=True)
    legend = [("C", "Calisti", GOOD_FILL), ("I", "Izinli", WARN_FILL),
              ("M", "Mazeret", WARN_FILL), ("T", "Tatil", SPECIAL_FILL),
              ("R", "Raporlu", BAD_FILL), ("G", "Gelmedi", BAD_FILL),
              ("D", "Diger", SECTION_FILL)]
    for i, (code, name, fill) in enumerate(legend):
        cell = ws.cell(row=legend_row, column=2 + i,
                       value=f"{code} = {name}")
        cell.fill = fill
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER


def _merge_month_header(ws, month_label: str, col_start: int, col_end: int):
    cell = ws.cell(row=1, column=col_start, value=month_label)
    cell.font = HEADER_FONT
    cell.fill = HEADER_FILL
    cell.alignment = Alignment(horizontal="center", vertical="center")
    cell.border = BORDER
    if col_end > col_start:
        ws.merge_cells(start_row=1, start_column=col_start, end_row=1, end_column=col_end)


# ============================================================================
# Her calisan icin ayri sekme (gun gun)
# ============================================================================

_SHEET_INVALID = re.compile(r"[\\/?*\[\]:]")


def _sanitize_sheet_name(name: str, used: set[str]) -> str:
    base = _SHEET_INVALID.sub("", name).strip() or "Kayit"
    base = base[:31]
    final = base
    i = 1
    while final.casefold() in {u.casefold() for u in used}:
        suffix = f" ({i})"
        final = (base[: 31 - len(suffix)]) + suffix
        i += 1
    used.add(final)
    return final


def _detect_shift(start_time, end_time, break_minutes, templates) -> str:
    if not start_time or not end_time:
        return ""
    try:
        bm = int(break_minutes or 0)
    except (TypeError, ValueError):
        bm = 0
    for tpl in templates:
        try:
            _id, name, tpl_start, tpl_end, tpl_break = tpl
        except ValueError:
            continue
        if str(tpl_start) == str(start_time) and str(tpl_end) == str(end_time) \
                and int(tpl_break or 0) == bm:
            return str(name)
    # Esit mola yoksa saat eslesmesi
    for tpl in templates:
        try:
            _id, name, tpl_start, tpl_end, _tpl_break = tpl
        except ValueError:
            continue
        if str(tpl_start) == str(start_time) and str(tpl_end) == str(end_time):
            return f"{name} (mola farkli)"
    return "Ozel"


def _sunday_hours(work_date, department, is_special, worked_hours):
    if is_special or not worked_hours:
        return 0.0
    try:
        from calc import is_sunday_non_stand  # type: ignore
        if is_sunday_non_stand(work_date, department):
            return float(worked_hours)
    except Exception:
        return 0.0
    return 0.0


def _build_per_employee_sheets(wb, entries, employees_by_id, settings, shift_templates):
    """Her calisan icin gun gun puantaj detay sekmesi yaratir.

    Sutunlar: Tarih, Gun, Durum, Vardiya, Giris, Cikis, Mola, Calisilan,
    Plan, Fazla Mesai, Gece, Geceye Tasan, Pazar Mesaisi, Ozel Gun, Not, Uyari.
    """
    # Tarih araligini belirle
    dates = sorted({_parse_date_str(e.get("work_date"))
                    for e in entries if _parse_date_str(e.get("work_date"))})
    if not dates:
        return
    global_start, global_end = dates[0], dates[-1]
    # Global donem makulse (<=45 gun ~ 1 ay) tum calisanlar ayni takvimde
    # hizalanir. Cok genisse (orn dagik tarihler) her calisan kendi araligini
    # kullanir; aksi halde sayfa yuzlerce bos "(kayit yok)" satiriyla dolar.
    use_global_span = (global_end - global_start).days <= 45

    # Calisan bazli grupla
    grouped: dict[str, dict] = {}
    for e in entries:
        key, label = _employee_label(e, employees_by_id)
        if key not in grouped:
            info = employees_by_id.get(int(e["employee_id"])) if e.get("employee_id") else {}
            grouped[key] = {
                "label": label,
                "department": (info or {}).get("department") or "",
                "region": (info or {}).get("region") or e.get("region") or "",
                "matched": e.get("employee_id") is not None,
                "days": {},  # date -> entry dict
            }
        d = _parse_date_str(e.get("work_date"))
        if d:
            grouped[key]["days"][d] = e

    used_sheet_names: set[str] = set()
    sorted_keys = sorted(grouped.keys(), key=lambda k: grouped[k]["label"].casefold())
    for key in sorted_keys:
        g = grouped[key]
        # Bu calisanin gosterim araligi
        emp_dates = sorted(d for d in g["days"].keys() if d)
        if not emp_dates:
            continue
        if use_global_span:
            span_start, span_end = global_start, global_end
        else:
            span_start, span_end = emp_dates[0], emp_dates[-1]
        sheet_name = _sanitize_sheet_name(g["label"][:28], used_sheet_names)
        ws = wb.create_sheet(sheet_name)

        # Baslik bloku
        ws.cell(row=1, column=1, value=g["label"]).font = TITLE_FONT
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=16)
        ws.row_dimensions[1].height = 26
        meta = (f"Departman: {g['department'] or '-'}    "
                f"Bolge: {g['region'] or '-'}    "
                f"Esleme: {'OK' if g['matched'] else 'ESLEMEDI'}    "
                f"Donem: {span_start.isoformat()} - {span_end.isoformat()}")
        info_cell = ws.cell(row=2, column=1, value=meta)
        info_cell.font = Font(italic=True, color="555555")
        ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=16)

        if not g["matched"]:
            warn_cell = ws.cell(row=3, column=1,
                                value="DIKKAT: Bu calisan DB'de eslesmedi; apply asamasinda atlanir.")
            warn_cell.font = Font(bold=True, color="9C0006")
            warn_cell.fill = BAD_FILL
            ws.merge_cells(start_row=3, start_column=1, end_row=3, end_column=16)

        # Sutun basliklari
        headers = [
            "Tarih", "Gun", "Durum", "Vardiya", "Giris", "Cikis", "Mola (dk)",
            "Calisilan (s)", "Plan (s)", "Fazla Mesai (s)", "Gece (s)",
            "Geceye Tasan (s)", "Pazar Mesaisi (s)", "Ozel Gun", "Not", "Uyari",
        ]
        header_row = 4 if not g["matched"] else 4
        _write_headers(ws, header_row, headers)

        # Veri satirlari (donem boyunca her gun)
        out_row = header_row + 1
        totals = {
            "worked": 0.0, "plan": 0.0, "overtime": 0.0, "night": 0.0,
            "overnight": 0.0, "sunday": 0.0,
            "calisma": 0, "izin": 0, "rapor": 0, "gelmedi": 0,
            "mazeret": 0, "tatil": 0, "ozel": 0,
        }
        for d in _iter_dates(span_start, span_end):
            entry = g["days"].get(d)
            weekday_label = WEEKDAY_TR[d.weekday()]
            row_fill = None
            if d.weekday() == 6:
                row_fill = SPECIAL_FILL  # pazar
            elif d.weekday() == 5:
                row_fill = SECTION_FILL  # cumartesi
            if entry:
                status = entry.get("status") or "Diger"
                start_t = entry.get("start_time") or ""
                end_t = entry.get("end_time") or ""
                break_min = int(entry.get("break_minutes") or 0)
                is_special = bool(entry.get("is_special"))
                shift_name = _detect_shift(start_t, end_t, break_min, shift_templates)
                calc_res = _safe_calc(
                    entry.get("work_date"), start_t, end_t, break_min,
                    settings, is_special, g["department"],
                )
                if calc_res:
                    worked, plan, overtime, night, overnight, *_ = calc_res
                else:
                    worked = plan = overtime = night = overnight = ""
                sunday_hrs = _sunday_hours(
                    entry.get("work_date"), g["department"], is_special,
                    worked if isinstance(worked, (int, float)) else 0.0,
                )
                # Totals
                if isinstance(worked, (int, float)):
                    totals["worked"] += worked
                if isinstance(plan, (int, float)):
                    totals["plan"] += plan
                if isinstance(overtime, (int, float)):
                    totals["overtime"] += overtime
                if isinstance(night, (int, float)):
                    totals["night"] += night
                if isinstance(overnight, (int, float)):
                    totals["overnight"] += overnight
                totals["sunday"] += sunday_hrs
                status_counter = {
                    "Calisti": "calisma", "Izinli": "izin", "Raporlu": "rapor",
                    "Gelmedi": "gelmedi", "Mazeret": "mazeret", "Tatil": "tatil",
                }.get(status)
                if status_counter:
                    totals[status_counter] += 1
                if is_special:
                    totals["ozel"] += 1

                values = [
                    d.isoformat(), weekday_label, status, shift_name,
                    start_t, end_t, break_min,
                    worked, plan, overtime, night, overnight, sunday_hrs,
                    "Evet" if is_special else "Hayir",
                    entry.get("notes") or "",
                    " | ".join(entry.get("warnings") or []),
                ]
                base_fill = STATUS_FILL.get(status)
                if is_special:
                    base_fill = SPECIAL_FILL
                elif entry.get("warnings"):
                    base_fill = WARN_FILL
            else:
                values = [d.isoformat(), weekday_label, "-", "", "", "", "",
                          "", "", "", "", "", "", "", "(kayit yok)", ""]
                base_fill = EMPTY_DAY_FILL

            for c, v in enumerate(values, start=1):
                cell = ws.cell(row=out_row, column=c, value=v)
                cell.border = BORDER
                cell.alignment = Alignment(horizontal="center" if c not in (15, 16) else "left",
                                           vertical="center")
                if base_fill:
                    cell.fill = base_fill
                elif row_fill and not entry:
                    cell.fill = row_fill
            out_row += 1

        # Toplam satiri
        ws.cell(row=out_row, column=1, value="TOPLAM").font = Font(bold=True)
        ws.cell(row=out_row, column=2, value="").fill = SECTION_FILL
        total_values = [
            "TOPLAM",
            f"{(span_end - span_start).days + 1} gun",
            f"C:{totals['calisma']} I:{totals['izin']} R:{totals['rapor']} "
            f"G:{totals['gelmedi']} M:{totals['mazeret']} T:{totals['tatil']}",
            f"Ozel: {totals['ozel']}",
            "", "", "",
            round(totals["worked"], 2), round(totals["plan"], 2),
            round(totals["overtime"], 2), round(totals["night"], 2),
            round(totals["overnight"], 2), round(totals["sunday"], 2),
            "", "", "",
        ]
        for c, v in enumerate(total_values, start=1):
            cell = ws.cell(row=out_row, column=c, value=v)
            cell.fill = SECTION_FILL
            cell.font = Font(bold=True)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center", vertical="center")

        ws.freeze_panes = ws.cell(row=header_row + 1, column=1).coordinate
        # Sutun genislikleri
        widths = [12, 8, 10, 18, 8, 8, 9, 11, 9, 13, 9, 13, 14, 10, 22, 22]
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w


# ============================================================================
# Calisan Analizi
# ============================================================================

def _build_analysis_sheet(wb, entries: list[dict], employees_by_id: dict, settings: dict):
    """Kisi basi ozet: calisma gunu, toplam/fazla/gece saat, izin/rapor/gelmedi,
    ortalama gunluk saat, ozel gun sayisi, devamsizlik orani, en uzun gun.
    """
    ws = wb.create_sheet("Calisan Analizi", index=2)

    dates = sorted({_parse_date_str(e.get("work_date"))
                    for e in entries if _parse_date_str(e.get("work_date"))})
    period_days = (dates[-1] - dates[0]).days + 1 if dates else 0

    # Aggrege et
    agg: dict[str, dict] = {}
    for e in entries:
        key, label = _employee_label(e, employees_by_id)
        info = employees_by_id.get(int(e["employee_id"])) if e.get("employee_id") else {}
        if key not in agg:
            agg[key] = {
                "label": label,
                "department": (info or {}).get("department") or "",
                "region": (info or {}).get("region") or e.get("region") or "",
                "calisma": 0, "izin": 0, "rapor": 0, "gelmedi": 0,
                "mazeret": 0, "tatil": 0, "diger": 0,
                "ozel_gun": 0, "matched": e.get("employee_id") is not None,
                "toplam_s": 0.0, "plan_s": 0.0, "fazla_s": 0.0,
                "gece_s": 0.0, "en_uzun": 0.0, "first_day": None, "last_day": None,
            }
        r = agg[key]
        d = _parse_date_str(e.get("work_date"))
        if d:
            r["first_day"] = d if not r["first_day"] or d < r["first_day"] else r["first_day"]
            r["last_day"] = d if not r["last_day"] or d > r["last_day"] else r["last_day"]
        status = e.get("status") or "Diger"
        status_key = {
            "Calisti": "calisma", "Izinli": "izin", "Raporlu": "rapor",
            "Gelmedi": "gelmedi", "Mazeret": "mazeret", "Tatil": "tatil",
        }.get(status, "diger")
        r[status_key] += 1
        if e.get("is_special"):
            r["ozel_gun"] += 1
        calc_res = _safe_calc(
            e.get("work_date"), e.get("start_time"), e.get("end_time"),
            e.get("break_minutes") or 0, settings,
            e.get("is_special"), (info or {}).get("department"),
        )
        if calc_res:
            worked, plan, overtime, night, *_ = calc_res
            r["toplam_s"] += worked
            r["plan_s"] += plan
            r["fazla_s"] += overtime
            r["gece_s"] += night
            if worked > r["en_uzun"]:
                r["en_uzun"] = worked

    # Basliklar
    headers = [
        "Calisan", "Departman", "Bolge", "Esleme",
        "Kayit Gunu", "Calisma G.", "Izin G.", "Rapor G.", "Gelmedi G.",
        "Mazeret G.", "Tatil G.", "Ozel Gun G.",
        "Toplam (s)", "Plan (s)", "Fazla Mesai (s)", "Gece (s)",
        "Ortalama (s/gun)", "En Uzun Gun (s)",
        "Devamsizlik %", "Aralik",
    ]
    _write_headers(ws, 1, headers)

    out_row = 2
    sorted_keys = sorted(agg.keys(), key=lambda k: (-agg[k]["calisma"], agg[k]["label"].casefold()))
    totals = defaultdict(float)
    for key in sorted_keys:
        r = agg[key]
        kayit_gun = r["calisma"] + r["izin"] + r["rapor"] + r["gelmedi"] + r["mazeret"] + r["tatil"] + r["diger"]
        ort = round(r["toplam_s"] / r["calisma"], 2) if r["calisma"] else 0
        absent = r["izin"] + r["rapor"] + r["gelmedi"] + r["mazeret"]
        devamsizlik = round(absent / kayit_gun * 100, 1) if kayit_gun else 0
        aralik = ""
        if r["first_day"] and r["last_day"]:
            aralik = f"{r['first_day'].isoformat()} – {r['last_day'].isoformat()}"
        values = [
            r["label"], r["department"], r["region"],
            "Eslesti" if r["matched"] else "ESLEMEDI",
            kayit_gun, r["calisma"], r["izin"], r["rapor"], r["gelmedi"],
            r["mazeret"], r["tatil"], r["ozel_gun"],
            round(r["toplam_s"], 2), round(r["plan_s"], 2),
            round(r["fazla_s"], 2), round(r["gece_s"], 2),
            ort, round(r["en_uzun"], 2),
            devamsizlik, aralik,
        ]
        for c, v in enumerate(values, start=1):
            cell = ws.cell(row=out_row, column=c, value=v)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center" if c > 3 else "left", vertical="center")
        # Renklendirme
        if not r["matched"]:
            for c in range(1, len(headers) + 1):
                ws.cell(row=out_row, column=c).fill = BAD_FILL
        elif absent > r["calisma"]:
            for c in range(1, len(headers) + 1):
                ws.cell(row=out_row, column=c).fill = WARN_FILL
        elif r["fazla_s"] >= 10:
            ws.cell(row=out_row, column=15).fill = SPECIAL_FILL
        # Totals
        totals["calisma"] += r["calisma"]
        totals["izin"] += r["izin"]
        totals["rapor"] += r["rapor"]
        totals["gelmedi"] += r["gelmedi"]
        totals["mazeret"] += r["mazeret"]
        totals["tatil"] += r["tatil"]
        totals["ozel_gun"] += r["ozel_gun"]
        totals["toplam_s"] += r["toplam_s"]
        totals["plan_s"] += r["plan_s"]
        totals["fazla_s"] += r["fazla_s"]
        totals["gece_s"] += r["gece_s"]
        out_row += 1

    # Toplam satiri
    if agg:
        total_values = [
            "TOPLAM", "", "",
            "",  # Esleme
            int(sum([totals["calisma"], totals["izin"], totals["rapor"],
                     totals["gelmedi"], totals["mazeret"], totals["tatil"]])),
            int(totals["calisma"]), int(totals["izin"]), int(totals["rapor"]),
            int(totals["gelmedi"]), int(totals["mazeret"]), int(totals["tatil"]),
            int(totals["ozel_gun"]),
            round(totals["toplam_s"], 2), round(totals["plan_s"], 2),
            round(totals["fazla_s"], 2), round(totals["gece_s"], 2),
            "", "", "",
            f"{period_days} gun" if period_days else "",
        ]
        for c, v in enumerate(total_values, start=1):
            cell = ws.cell(row=out_row, column=c, value=v)
            cell.fill = SECTION_FILL
            cell.font = Font(bold=True)
            cell.border = BORDER
            cell.alignment = Alignment(horizontal="center" if c > 3 else "left")

    ws.freeze_panes = "E2"
    _autosize(ws, len(headers))


def _normalize_time(value) -> str | None:
    if value is None or value == "":
        return None
    if hasattr(value, "strftime"):
        try:
            return value.strftime("%H:%M")
        except Exception:
            pass
    s = str(value).strip()
    if not s:
        return None
    s = s.replace(".", ":")
    if len(s) == 4 and ":" not in s:
        s = s[:2] + ":" + s[2:]
    if len(s) == 5 and s[2] == ":":
        return s
    return s
