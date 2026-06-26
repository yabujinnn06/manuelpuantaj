"""Onaylanmis WhatsApp puantaj kayitlarini DB'ye toplu yazan modul."""

from __future__ import annotations

import unicodedata
from dataclasses import dataclass


@dataclass
class ApplyResult:
    timesheets_added: int = 0
    attendance_added: int = 0
    skipped: int = 0
    errors: list[str] = None
    skipped_details: list[dict] = None

    def __post_init__(self):
        if self.errors is None:
            self.errors = []
        if self.skipped_details is None:
            self.skipped_details = []

    def to_dict(self):
        return {
            "timesheets_added": self.timesheets_added,
            "attendance_added": self.attendance_added,
            "skipped": self.skipped,
            "errors": self.errors,
            "skipped_details": self.skipped_details,
        }


def _norm(text: str) -> str:
    t = (text or "").strip().casefold().replace("ı", "i")
    t = unicodedata.normalize("NFKD", t)
    return "".join(ch for ch in t if not unicodedata.combining(ch))


def _resolve_employee_id(record: dict, employees: list[tuple]) -> int | None:
    if record.get("employee_id"):
        return int(record["employee_id"])
    target = _norm(record.get("matched_name") or record.get("employee_name_raw") or "")
    if not target:
        return None
    for row in employees:
        if _norm(row[1]) == target:
            return int(row[0])
    # Token-level fallback
    target_tokens = set(target.split())
    if not target_tokens:
        return None
    best = (None, 0.0)
    for row in employees:
        ct = set(_norm(row[1]).split())
        if not ct:
            continue
        inter = target_tokens & ct
        union = target_tokens | ct
        score = len(inter) / len(union)
        if score > best[1]:
            best = (int(row[0]), score)
    if best[1] >= 0.66:
        return best[0]
    return None


def apply_entries(
    db,
    entries: list[dict],
    default_region: str | None = None,
    source: str = "whatsapp",
    overwrite: bool = False,
) -> ApplyResult:
    """Parsed entries'i DB'ye yazar.

    - status == 'Calisti' ve start/end varsa -> timesheets
    - aksi halde attendance_records (upsert)
    overwrite=True ise ayni gun + calisana ait timesheets ilk once silinir.
    """
    employees = db.list_employees()
    result = ApplyResult()
    seen: set[tuple[int, str]] = set()

    for raw in entries:
        rec = dict(raw)
        emp_id = _resolve_employee_id(rec, employees)
        work_date = rec.get("work_date")
        if not emp_id or not work_date:
            result.skipped += 1
            result.skipped_details.append({
                "reason": "Calisan veya tarih cozulemedi",
                "name": rec.get("employee_name_raw"),
                "work_date": work_date,
            })
            continue
        region = rec.get("region") or default_region
        if not region:
            # En azindan calisanin bolgesini kullan
            for row in employees:
                if int(row[0]) == emp_id:
                    region = row[5] if len(row) > 5 else None
                    break
        key = (emp_id, work_date)
        if key in seen:
            result.skipped += 1
            result.skipped_details.append({
                "reason": "Ayni mesajda mukerrer kayit",
                "employee_id": emp_id, "work_date": work_date,
            })
            continue
        seen.add(key)

        status = (rec.get("status") or "").strip() or "Diger"
        notes = rec.get("notes") or "WhatsApp puantaj"

        try:
            if status == "Calisti" and rec.get("start_time") and rec.get("end_time"):
                if overwrite:
                    existing = db.list_timesheets(
                        employee_id=emp_id, start_date=work_date, end_date=work_date,
                    )
                    for ex in existing:
                        try:
                            db.delete_timesheet(int(ex[0]))
                        except Exception as e:
                            result.errors.append(f"Eski timesheet silinemedi #{ex[0]}: {e}")
                db.add_timesheet(
                    emp_id, work_date,
                    rec.get("start_time"), rec.get("end_time"),
                    int(rec.get("break_minutes") or 0),
                    1 if rec.get("is_special") else 0,
                    notes, region,
                )
                # Devamsizlik tablosuna da 'Calisti' isaretle (rapor butunluk)
                db.upsert_attendance_record(
                    emp_id, work_date, "Calisti",
                    rec.get("notes") or "", region, source,
                )
                result.timesheets_added += 1
            else:
                db.upsert_attendance_record(
                    emp_id, work_date, status,
                    rec.get("notes") or "", region, source,
                )
                result.attendance_added += 1
        except Exception as e:
            result.errors.append(
                f"Kayit yazilamadi (calisan={emp_id}, tarih={work_date}): {e}"
            )
            result.skipped += 1
            result.skipped_details.append({
                "reason": str(e), "employee_id": emp_id, "work_date": work_date,
            })

    return result
