"""WhatsApp puantaj on izleme Excel uretici.

Parser ciktisini renkli, sekmeli, ozetli bir Excel olarak verir. Kullanici
bu dosyayi inceler; uygunsa cli'a `apply` ile geri verir.
"""

from __future__ import annotations

import os
from collections import Counter, defaultdict
from datetime import datetime

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


def _safe_calc(work_date, start_time, end_time, break_minutes, settings, is_special, department):
    if not start_time or not end_time:
        return None
    try:
        from calc import calc_day_hours  # type: ignore
        return calc_day_hours(
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
) -> str:
    """entries: whatsapp.entries_to_dicts ciktisi.

    employees_by_id: {id: {full_name, department, region}} sozlugu.
    settings: db.get_all_settings() (calc icin).

    Return: yazilan dosyanin tam yolu.
    """
    employees_by_id = employees_by_id or {}
    settings = settings or {}
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

    # 2) Detay sekmesi
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
