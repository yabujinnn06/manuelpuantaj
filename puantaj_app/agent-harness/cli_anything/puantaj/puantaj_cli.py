"""cli-anything-puantaj: Rainstaff puantaj uygulamasinin agent-kullanilabilir CLI'i.

Uygulamanin kendi modullerini (puantaj_db, calc, report) import edip cagirir.
Mantik yeniden yazilmaz; ayni SQLite veritabani (%APPDATA%\\Rainstaff\\data) kullanilir.
"""

import os
import sys
import json
import shlex
import unicodedata

import click

for _s in (sys.stdout, sys.stderr):
    try:
        _s.reconfigure(encoding="utf-8")
    except (AttributeError, ValueError):
        pass


def _find_app_dir():
    env = os.environ.get("PUANTAJ_APP_DIR")
    if env and os.path.isfile(os.path.join(env, "puantaj_db.py")):
        return env
    here = os.path.abspath(__file__)
    # .../puantaj_app/agent-harness/cli_anything/puantaj/puantaj_cli.py -> .../puantaj_app
    cand = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(here))))
    if os.path.isfile(os.path.join(cand, "puantaj_db.py")):
        return cand
    raise click.ClickException(
        "puantaj kaynak klasoru bulunamadi. PUANTAJ_APP_DIR ortam degiskenini "
        "puantaj_db.py iceren klasore ayarlayin."
    )


APP_DIR = _find_app_dir()
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

import puantaj_db as db  # noqa: E402
import calc as calc_mod  # noqa: E402

from cli_anything.puantaj import whatsapp as wa_mod  # noqa: E402
from cli_anything.puantaj import bulk as bulk_mod  # noqa: E402


def _load_preview_mod():
    """Lazy import: openpyxl olmadan da temel komutlar calissin."""
    try:
        from cli_anything.puantaj import preview_xlsx as preview_mod  # noqa: WPS433
        return preview_mod
    except ImportError as e:
        raise click.ClickException(
            f"Onizleme modulu yuklenemedi ({e}). Gerekli: pip install openpyxl"
        )

EMP_COLS = ["id", "full_name", "identity_no", "department", "title", "region"]
TS_COLS = ["id", "employee_id", "full_name", "department", "work_date",
           "start_time", "end_time", "break_minutes", "is_special", "notes", "region"]
USER_COLS = ["id", "username", "role", "region"]
TPL_COLS = ["id", "name", "start_time", "end_time", "break_minutes"]
FIXED_SPECIAL_DAYS = {
    "01-01": "Yilbasi",
    "04-23": "23 Nisan Ulusal Egemenlik ve Cocuk Bayrami",
    "05-01": "Emek ve Dayanisma Gunu",
    "05-19": "19 Mayis Ataturk'u Anma Genclik ve Spor Bayrami",
    "07-15": "15 Temmuz Demokrasi ve Milli Birlik Gunu",
    "08-30": "30 Agustos Zafer Bayrami",
    "10-29": "29 Ekim Cumhuriyet Bayrami",
}
YEAR_SPECIAL_DAYS = {
    "2026-05-26": "Kurban Bayrami Arefesi",
    "2026-05-27": "Kurban Bayrami 1. Gun",
    "2026-05-28": "Kurban Bayrami 2. Gun",
    "2026-05-29": "Kurban Bayrami 3. Gun",
    "2026-05-30": "Kurban Bayrami 4. Gun",
}


def _emit(ctx, data, headers=None):
    if ctx.obj.get("json"):
        click.echo(json.dumps(data, ensure_ascii=False, default=str, indent=2))
        return
    if isinstance(data, list) and data and isinstance(data[0], dict):
        cols = headers or list(data[0].keys())
        click.echo(" | ".join(cols))
        click.echo("-" * 60)
        for r in data:
            click.echo(" | ".join(str(r.get(c, "")) for c in cols))
    elif isinstance(data, list):
        for r in data:
            click.echo(r)
    elif isinstance(data, dict):
        for k, v in data.items():
            click.echo(f"{k}: {v}")
    else:
        click.echo(str(data))


def _rows(rows, cols):
    return [dict(zip(cols, r)) for r in rows]


def _norm_text(value):
    text = str(value or "").strip().casefold().replace("ı", "i").replace("þ", "s")
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _find_shift_template(template_ref, required=True):
    if not template_ref:
        return None
    wanted = str(template_ref).strip()
    wanted_norm = _norm_text(wanted)
    templates = db.list_shift_templates()
    for tpl in templates:
        if str(tpl[0]) == wanted or _norm_text(tpl[1]) == wanted_norm:
            return tpl
    if not required:
        return None
    available = ", ".join(str(t[1]) for t in templates) or "sablon yok"
    raise click.ClickException(f"Sablon bulunamadi: {template_ref}. Mevcut: {available}")


def _find_first_shift_template(*template_refs):
    for ref in template_refs:
        tpl = _find_shift_template(ref, required=False)
        if tpl:
            return tpl
    return None


def _get_employee(employee_id):
    for row in db.list_employees():
        if int(row[0]) == int(employee_id):
            return dict(zip(EMP_COLS, row))
    raise click.ClickException(f"Calisan bulunamadi: {employee_id}")


def _special_day_name(work_date):
    try:
        parsed = calc_mod.parse_date(work_date)
    except Exception:
        return None
    return YEAR_SPECIAL_DAYS.get(parsed.isoformat()) or FIXED_SPECIAL_DAYS.get(parsed.strftime("%m-%d"))


def _select_auto_template(employee, work_date, start_time, end_time, notes=""):
    department = _norm_text(employee.get("department"))
    notes_key = _norm_text(notes)
    try:
        weekday = calc_mod.parse_date(work_date).weekday()
    except Exception as exc:
        raise click.ClickException(str(exc))

    if department == "stant":
        start = str(start_time or "").strip()
        end = str(end_time or "").strip()
        if (start, end) == ("10:00", "22:00") or "full" in notes_key:
            return _find_first_shift_template("stant full")
        if (start, end) == ("10:00", "18:00"):
            return _find_first_shift_template("stant 10-18")
        if (start, end) == ("14:00", "22:00"):
            return _find_first_shift_template("stant 14-22")
        return None

    if department in ("lojistik", "teknik") and weekday == 5:
        return _find_first_shift_template("Cumartesi 09-14", "cumartesi")
    if department == "lojistik" and weekday <= 4:
        return _find_first_shift_template("şöför", "şoför", "sofor")
    if department == "teknik" and weekday <= 4:
        return _find_first_shift_template("Hafta Ici 8.30-17.30")
    return None


def _append_note(notes, extra):
    notes = (notes or "").strip()
    if not extra:
        return notes
    if not notes:
        return extra
    if extra in notes:
        return notes
    return f"{notes} | {extra}"


def _apply_shift_template(template_ref, start_time, end_time, break_minutes, notes, template=None):
    tpl = template or _find_shift_template(template_ref)
    if not tpl:
        return start_time, end_time, 0 if break_minutes is None else break_minutes, notes
    _tpl_id, tpl_name, tpl_start, tpl_end, tpl_break = tpl
    return (
        start_time or tpl_start,
        end_time or tpl_end,
        tpl_break if break_minutes is None else break_minutes,
        _append_note(notes, f"Sablon: {tpl_name}"),
    )


@click.group(invoke_without_command=True)
@click.option("--json", "json_out", is_flag=True, help="Makine okunur JSON ciktisi.")
@click.pass_context
def cli(ctx, json_out):
    """Rainstaff puantaj CLI - personel, puantaj, izin, arac, rapor."""
    ctx.ensure_object(dict)
    ctx.obj["json"] = json_out
    db.init_db()
    if ctx.invoked_subcommand is None:
        _repl(ctx)


# ---------------------------------------------------------------- db / info
@cli.command("info")
@click.pass_context
def info(ctx):
    """Veritabani yolu ve ozet sayilar."""
    data = {
        "db_path": db.DB_PATH,
        "employees": len(db.list_employees()),
        "timesheets": len(db.list_timesheets()),
        "users": len(db.list_users()),
    }
    _emit(ctx, data)


# ---------------------------------------------------------------- employees
@cli.group()
def employee():
    """Personel yonetimi."""


@employee.command("list")
@click.option("--region", default=None)
@click.pass_context
def employee_list(ctx, region):
    _emit(ctx, _rows(db.list_employees(region=region), EMP_COLS), EMP_COLS)


@employee.command("add")
@click.option("--full-name", required=True)
@click.option("--identity-no", default="")
@click.option("--department", default="")
@click.option("--title", default="")
@click.option("--region", required=True)
@click.pass_context
def employee_add(ctx, full_name, identity_no, department, title, region):
    db.add_employee(full_name, identity_no, department, title, region)
    _emit(ctx, {"status": "ok", "added": full_name, "region": region})


@employee.command("update")
@click.argument("employee_id", type=int)
@click.option("--full-name", required=True)
@click.option("--identity-no", default="")
@click.option("--department", default="")
@click.option("--title", default="")
@click.option("--region", required=True)
@click.pass_context
def employee_update(ctx, employee_id, full_name, identity_no, department, title, region):
    db.update_employee(employee_id, full_name, identity_no, department, title, region)
    _emit(ctx, {"status": "ok", "updated": employee_id})


@employee.command("delete")
@click.argument("employee_id", type=int)
@click.pass_context
def employee_delete(ctx, employee_id):
    db.delete_employee(employee_id)
    _emit(ctx, {"status": "ok", "deleted": employee_id})


# ---------------------------------------------------------------- timesheets
@cli.group()
def timesheet():
    """Puantaj (mesai) kayitlari."""


@timesheet.command("list")
@click.option("--employee-id", type=int, default=None)
@click.option("--start-date", default=None)
@click.option("--end-date", default=None)
@click.option("--region", default=None)
@click.pass_context
def timesheet_list(ctx, employee_id, start_date, end_date, region):
    rows = db.list_timesheets(employee_id=employee_id, start_date=start_date,
                              end_date=end_date, region=region)
    _emit(ctx, _rows(rows, TS_COLS), TS_COLS)


@timesheet.command("add")
@click.option("--employee-id", type=int, required=True)
@click.option("--work-date", required=True, help="YYYY-MM-DD")
@click.option("--start", "start_time", default=None, help="HH:MM. --template ile verilmezse sablondan gelir.")
@click.option("--end", "end_time", default=None, help="HH:MM. --template ile verilmezse sablondan gelir.")
@click.option("--break-minutes", type=int, default=None, help="Verilmezse --template molasi, sablon yoksa 0.")
@click.option("--template", "template_ref", default=None, help="Vardiya sablon adi veya ID.")
@click.option("--auto-template/--no-auto-template", default=True, help="Sablon verilmezse departman/gun/saatten sec.")
@click.option("--special", "is_special", is_flag=True)
@click.option("--auto-special/--no-auto-special", default=True, help="Bilinen resmi tatillerde ozel gunu otomatik isaretle.")
@click.option("--notes", default="")
@click.option("--region", required=True)
@click.pass_context
def timesheet_add(ctx, employee_id, work_date, start_time, end_time,
                  break_minutes, template_ref, auto_template, is_special,
                  auto_special, notes, region):
    employee = _get_employee(employee_id)
    template = None
    if template_ref:
        template = _find_shift_template(template_ref)
    elif auto_template:
        template = _select_auto_template(employee, work_date, start_time, end_time, notes)
    start_time, end_time, break_minutes, notes = _apply_shift_template(
        template_ref, start_time, end_time, break_minutes, notes, template=template
    )
    if not start_time or not end_time:
        raise click.ClickException("--start ve --end zorunlu; veya --template ile sablondan doldurun.")

    special_name = _special_day_name(work_date) if auto_special else None
    if special_name:
        is_special = True
        notes = _append_note(notes, f"Ozel gun calismasi: {special_name}")
    elif is_special:
        notes = _append_note(notes, "Ozel gun calismasi")

    db.add_timesheet(employee_id, work_date, start_time, end_time,
                     break_minutes, 1 if is_special else 0, notes, region)
    _emit(ctx, {"status": "ok", "employee_id": employee_id, "work_date": work_date})


@timesheet.command("delete")
@click.argument("timesheet_id", type=int)
@click.pass_context
def timesheet_delete(ctx, timesheet_id):
    db.delete_timesheet(timesheet_id)
    _emit(ctx, {"status": "ok", "deleted": timesheet_id})


@timesheet.command("calc")
@click.option("--work-date", required=True, help="YYYY-MM-DD")
@click.option("--start", "start_time", default=None, help="HH:MM. --template ile verilmezse sablondan gelir.")
@click.option("--end", "end_time", default=None, help="HH:MM. --template ile verilmezse sablondan gelir.")
@click.option("--break-minutes", type=int, default=None, help="Verilmezse --template molasi, sablon yoksa 0.")
@click.option("--template", "template_ref", default=None, help="Vardiya sablon adi veya ID.")
@click.option("--auto-template/--no-auto-template", default=True, help="Sablon verilmezse departman/gun/saatten sec.")
@click.option("--special", "is_special", is_flag=True)
@click.option("--auto-special/--no-auto-special", default=True, help="Bilinen resmi tatillerde ozel gunu otomatik isaretle.")
@click.option("--department", default=None)
@click.pass_context
def timesheet_calc(ctx, work_date, start_time, end_time, break_minutes,
                   template_ref, auto_template, is_special, auto_special, department):
    """Bir gun icin saatleri hesaplar (DB'ye yazmaz)."""
    template = None
    if template_ref:
        template = _find_shift_template(template_ref)
    elif auto_template and department:
        template = _select_auto_template({"department": department}, work_date, start_time, end_time)
    start_time, end_time, break_minutes, _notes = _apply_shift_template(
        template_ref, start_time, end_time, break_minutes, "", template=template
    )
    if not start_time or not end_time:
        raise click.ClickException("--start ve --end zorunlu; veya --template ile sablondan doldurun.")
    if auto_special and _special_day_name(work_date):
        is_special = True
    settings = db.get_all_settings()
    res = calc_mod.calc_day_hours(work_date, start_time, end_time, break_minutes,
                                  settings, 1 if is_special else 0, department)
    keys = ["worked", "scheduled", "overtime", "night", "overnight",
            "special_normal", "special_overtime", "special_night"]
    _emit(ctx, dict(zip(keys, res)))


# ---------------------------------------------------------------- attendance
@cli.group()
def attendance():
    """Devamsizlik / gunluk durum kayitlari."""


@attendance.command("set")
@click.option("--employee-id", type=int, required=True)
@click.option("--work-date", required=True)
@click.option("--status", required=True)
@click.option("--reason", default="")
@click.option("--region", required=True)
@click.option("--source", default="cli")
@click.pass_context
def attendance_set(ctx, employee_id, work_date, status, reason, region, source):
    db.upsert_attendance_record(employee_id, work_date, status, reason, region, source)
    _emit(ctx, {"status": "ok", "employee_id": employee_id, "work_date": work_date})


@attendance.command("list")
@click.option("--employee-id", type=int, default=None)
@click.option("--start-date", default=None)
@click.option("--end-date", default=None)
@click.option("--status", default=None)
@click.option("--region", default=None)
@click.pass_context
def attendance_list(ctx, employee_id, start_date, end_date, status, region):
    rows = db.list_attendance_records(employee_id=employee_id, start_date=start_date,
                                      end_date=end_date, status=status, region=region)
    _emit(ctx, [list(r) for r in rows])


# ---------------------------------------------------------------- leave
@cli.group()
def leave():
    """Izin kayitlari."""


@leave.command("list")
@click.option("--employee-id", type=int, default=None)
@click.option("--start-date", default=None)
@click.option("--end-date", default=None)
@click.option("--status", default=None)
@click.option("--region", default=None)
@click.pass_context
def leave_list(ctx, employee_id, start_date, end_date, status, region):
    rows = db.list_leave_records(employee_id=employee_id, start_date=start_date,
                                 end_date=end_date, status=status, region=region)
    _emit(ctx, [list(r) for r in rows])


@leave.command("add")
@click.option("--employee-id", type=int, required=True)
@click.option("--start-date", required=True)
@click.option("--end-date", required=True)
@click.option("--leave-type", required=True)
@click.option("--reason", default="")
@click.option("--document-no", default="")
@click.option("--status", default="onayli")
@click.option("--region", required=True)
@click.pass_context
def leave_add(ctx, employee_id, start_date, end_date, leave_type, reason,
              document_no, status, region):
    db.add_leave_record(employee_id, start_date, end_date, leave_type, reason,
                        document_no, status, region)
    _emit(ctx, {"status": "ok", "employee_id": employee_id})


# ---------------------------------------------------------------- vehicles
@cli.group()
def vehicle():
    """Arac yonetimi."""


@vehicle.command("list")
@click.option("--region", default=None)
@click.pass_context
def vehicle_list(ctx, region):
    _emit(ctx, [list(r) for r in db.list_vehicles(region=region)])


# ---------------------------------------------------------------- drivers
@cli.group()
def driver():
    """Surucu yonetimi."""


@driver.command("list")
@click.option("--region", default=None)
@click.pass_context
def driver_list(ctx, region):
    _emit(ctx, [list(r) for r in db.list_drivers(region=region)])


# ---------------------------------------------------------------- shift templates
@cli.group("template")
def template():
    """Vardiya sablonlari."""


@template.command("list")
@click.pass_context
def template_list(ctx):
    _emit(ctx, _rows(db.list_shift_templates(), TPL_COLS), TPL_COLS)


@template.command("set")
@click.option("--name", required=True)
@click.option("--start", "start_time", required=True, help="HH:MM")
@click.option("--end", "end_time", required=True, help="HH:MM")
@click.option("--break-minutes", type=int, required=True)
@click.pass_context
def template_set(ctx, name, start_time, end_time, break_minutes):
    db.upsert_shift_template(name, start_time, end_time, break_minutes)
    _emit(ctx, {"status": "ok", "template": name})


@template.command("delete")
@click.argument("template_id", type=int)
@click.pass_context
def template_delete(ctx, template_id):
    db.delete_shift_template(template_id)
    _emit(ctx, {"status": "ok", "deleted": template_id})


# ---------------------------------------------------------------- settings/users
@cli.group()
def settings():
    """Ayarlar."""


@settings.command("list")
@click.pass_context
def settings_list(ctx):
    _emit(ctx, db.get_all_settings())


@settings.command("set")
@click.argument("key")
@click.argument("value")
@click.pass_context
def settings_set(ctx, key, value):
    db.set_setting(key, value)
    _emit(ctx, {"status": "ok", key: value})


@cli.command("users")
@click.pass_context
def users(ctx):
    """Kullanicilari listeler."""
    _emit(ctx, _rows(db.list_users(), USER_COLS), USER_COLS)


# ---------------------------------------------------------------- report
def _month_range(month_str: str) -> tuple[str, str]:
    """'2026-01' -> ('2026-01-01', '2026-01-31')."""
    import calendar as _cal
    from datetime import date as _date
    try:
        year, mon = month_str.split("-")
        y, m = int(year), int(mon)
    except Exception:
        raise click.ClickException("--month formati hatali. Ornek: 2026-01")
    last = _cal.monthrange(y, m)[1]
    return _date(y, m, 1).isoformat(), _date(y, m, last).isoformat()


@cli.command("report")
@click.option("--output", "-o", required=True, help="Hedef .xlsx yolu")
@click.option("--employee-id", type=int, default=None)
@click.option("--start-date", default=None, help="YYYY-MM-DD (veya --month verin).")
@click.option("--end-date", default=None, help="YYYY-MM-DD (veya --month verin).")
@click.option("--month", "month_str", default=None,
              help="'YYYY-MM' verirseniz baslangic/bitis otomatik secilir.")
@click.option("--region", default=None)
@click.pass_context
def report_cmd(ctx, output, employee_id, start_date, end_date, month_str, region):
    """Excel puantaj raporu uretir (openpyxl gerekir)."""
    try:
        import report as report_mod
    except ImportError as e:
        raise click.ClickException(
            f"Rapor modulu yuklenemedi ({e}). Gerekli: pip install openpyxl Pillow reportlab"
        )
    if month_str:
        start_date, end_date = _month_range(month_str)
    if not start_date or not end_date:
        raise click.ClickException("--start-date ve --end-date veya --month gerekli.")
    ts = db.list_timesheets(employee_id=employee_id, start_date=start_date,
                            end_date=end_date, region=region)
    att = db.list_attendance_records(employee_id=employee_id, start_date=start_date,
                                     end_date=end_date, status=None, region=region)
    lv = db.list_leave_records(employee_id=employee_id, start_date=start_date,
                               end_date=end_date, status=None, region=region)
    date_text = f"Tarih Araligi: {start_date} - {end_date}"
    report_mod.export_report(output, ts, db.get_all_settings(), date_text,
                             attendance_records=att, leave_records=lv,
                             start_date=start_date, end_date=end_date)
    _emit(ctx, {"status": "ok", "output": os.path.abspath(output),
                "rows": len(ts), "start_date": start_date, "end_date": end_date})


# ---------------------------------------------------------------- whatsapp
@cli.group("whatsapp")
def whatsapp():
    """WhatsApp grup mesajindan toplu puantaj (parse -> preview -> apply)."""


def _read_input_text(input_file, text):
    if input_file and input_file != "-":
        with open(input_file, "r", encoding="utf-8") as f:
            return f.read()
    if input_file == "-":
        return sys.stdin.read()
    if text:
        return text
    raise click.ClickException("--input veya --text vermelisiniz (--input - ile stdin).")


def _employees_by_id():
    return {
        int(row[0]): {
            "full_name": row[1],
            "identity_no": row[2],
            "department": row[3],
            "title": row[4],
            "region": row[5] if len(row) > 5 else None,
        }
        for row in db.list_employees()
    }


@whatsapp.command("parse")
@click.option("--input", "input_file", default=None,
              help="UTF-8 metin dosyasi. '-' verilirse stdin.")
@click.option("--text", default=None, help="Komut satirinda metin.")
@click.option("--default-date", default=None, help="Tarih bulunamadigi gun icin varsayilan (YYYY-MM-DD).")
@click.option("--region", default=None, help="Tum kayitlar icin bolge.")
@click.option("--out-json", default=None, help="Parse ciktisini bu JSON dosyaya yaz.")
@click.pass_context
def whatsapp_parse(ctx, input_file, text, default_date, region, out_json):
    """Mesaji ayristirir ve JSON dondurur (DB'ye yazmaz)."""
    raw = _read_input_text(input_file, text)
    entries = wa_mod.parse_text(raw, default_date=default_date, region=region)
    wa_mod.match_employees(entries, db.list_employees(), region=region)
    payload = wa_mod.entries_to_dicts(entries)
    if out_json:
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(payload, f, ensure_ascii=False, indent=2, default=str)
    _emit(ctx, payload if ctx.obj.get("json") else {
        "kayit_sayisi": len(payload),
        "eslesen": sum(1 for e in payload if e.get("employee_id")),
        "eslemeyen": sum(1 for e in payload if not e.get("employee_id")),
        "uyarili": sum(1 for e in payload if e.get("warnings")),
        "out_json": os.path.abspath(out_json) if out_json else None,
    })


@whatsapp.command("preview")
@click.option("--input", "input_file", default=None)
@click.option("--text", default=None)
@click.option("--records-json", default=None, help="Onceden parse edilmis JSON dosyasi.")
@click.option("--default-date", default=None)
@click.option("--region", default=None)
@click.option("--output", "-o", required=True, help="Hedef .xlsx onizleme dosyasi.")
@click.pass_context
def whatsapp_preview(ctx, input_file, text, records_json, default_date, region, output):
    """WhatsApp mesajini parse edip renkli Excel onizlemesi olusturur."""
    if records_json:
        with open(records_json, "r", encoding="utf-8") as f:
            payload = json.load(f)
    else:
        raw = _read_input_text(input_file, text)
        entries = wa_mod.parse_text(raw, default_date=default_date, region=region)
        wa_mod.match_employees(entries, db.list_employees(), region=region)
        payload = wa_mod.entries_to_dicts(entries)
    out_path = _load_preview_mod().build_preview(
        output, payload,
        employees_by_id=_employees_by_id(),
        settings=db.get_all_settings(),
    )
    _emit(ctx, {
        "status": "ok",
        "output": out_path,
        "kayit_sayisi": len(payload),
        "eslemeyen": sum(1 for e in payload if not e.get("employee_id")),
    })


@whatsapp.command("apply")
@click.option("--records", "records_json", default=None, help="Parse JSON dosyasi.")
@click.option("--xlsx", default=None, help="Onaylanmis preview Excel dosyasi (Detay sekmesi okunur).")
@click.option("--region", default=None, help="Kayit yapilirken kullanilacak varsayilan bolge.")
@click.option("--overwrite/--no-overwrite", default=False,
              help="Ayni gun + calisan icin onceki timesheet'leri sil.")
@click.option("--yes", "assume_yes", is_flag=True, help="Onay sormadan uygula.")
@click.pass_context
def whatsapp_apply(ctx, records_json, xlsx, region, overwrite, assume_yes):
    """Onaylanmis kayitlari DB'ye toplu yazar."""
    if not records_json and not xlsx:
        raise click.ClickException("--records veya --xlsx vermelisiniz.")
    if xlsx:
        entries = _load_preview_mod().read_entries_from_xlsx(xlsx)
    else:
        with open(records_json, "r", encoding="utf-8") as f:
            entries = json.load(f)
    if not entries:
        raise click.ClickException("Uygulanacak kayit yok.")
    if not assume_yes and sys.stdin.isatty() and not ctx.obj.get("json"):
        click.echo(f"{len(entries)} kayit DB'ye yazilacak. Devam edilsin mi? [y/N] ", nl=False)
        ans = input().strip().lower()
        if ans not in ("y", "yes", "e", "evet"):
            raise click.ClickException("Iptal edildi.")
    result = bulk_mod.apply_entries(
        db, entries, default_region=region, source="whatsapp", overwrite=overwrite,
    )
    _emit(ctx, result.to_dict())


@whatsapp.command("ingest")
@click.option("--input", "input_file", default=None)
@click.option("--text", default=None)
@click.option("--default-date", default=None)
@click.option("--region", default=None)
@click.option("--preview", "preview_out", default=None,
              help="Onizleme xlsx yolu (default: scratch dizininde).")
@click.option("--yes", "assume_yes", is_flag=True,
              help="On izleme olusturulduktan sonra dogrudan apply et.")
@click.option("--overwrite/--no-overwrite", default=False)
@click.pass_context
def whatsapp_ingest(ctx, input_file, text, default_date, region, preview_out, assume_yes, overwrite):
    """Tek komutla: parse + preview + (onay sonrasi) apply."""
    raw = _read_input_text(input_file, text)
    entries = wa_mod.parse_text(raw, default_date=default_date, region=region)
    wa_mod.match_employees(entries, db.list_employees(), region=region)
    payload = wa_mod.entries_to_dicts(entries)

    if not preview_out:
        from datetime import datetime as _dt
        preview_out = os.path.join(
            os.path.dirname(db.DB_PATH),
            f"whatsapp_puantaj_onizleme_{_dt.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        )
    out_path = _load_preview_mod().build_preview(
        preview_out, payload,
        employees_by_id=_employees_by_id(),
        settings=db.get_all_settings(),
    )

    summary = {
        "preview": out_path,
        "kayit_sayisi": len(payload),
        "eslemeyen": sum(1 for e in payload if not e.get("employee_id")),
        "uyarili": sum(1 for e in payload if e.get("warnings")),
    }

    if not assume_yes:
        if not sys.stdin.isatty() or ctx.obj.get("json"):
            _emit(ctx, {**summary, "status": "preview_only",
                        "info": "Onaylamak icin: cli-anything-puantaj whatsapp apply --xlsx <preview> --yes"})
            return
        click.echo(f"On izleme: {out_path}")
        click.echo(f"Toplam {summary['kayit_sayisi']} kayit, "
                   f"{summary['eslemeyen']} eslemeyen, {summary['uyarili']} uyarili.")
        click.echo("Excel'i incele. DB'ye yazilsin mi? [y/N] ", nl=False)
        ans = input().strip().lower()
        if ans not in ("y", "yes", "e", "evet"):
            _emit(ctx, {**summary, "status": "cancelled"})
            return

    result = bulk_mod.apply_entries(
        db, payload, default_region=region, source="whatsapp", overwrite=overwrite,
    )
    _emit(ctx, {**summary, **result.to_dict(), "status": "applied"})


# ---------------------------------------------------------------- department
@cli.group("department")
def department():
    """Departman yonetimi (employees.department TEXT alani uzerinden)."""


@department.command("list")
@click.option("--region", default=None)
@click.pass_context
def department_list(ctx, region):
    """Mevcut departmanlari ve calisan sayilarini listeler."""
    rows = db.list_employees(region=region)
    counter = {}
    for r in rows:
        dep = (r[3] or "(yok)").strip() or "(yok)"
        counter[dep] = counter.get(dep, 0) + 1
    data = [{"department": k, "calisan_sayisi": v}
            for k, v in sorted(counter.items(), key=lambda x: -x[1])]
    _emit(ctx, data, ["department", "calisan_sayisi"])


@department.command("employees")
@click.argument("name")
@click.option("--region", default=None)
@click.pass_context
def department_employees(ctx, name, region):
    """Departmandaki calisanlari listeler (buyuk/kucuk harf duyarsiz)."""
    target = _norm_text(name)
    rows = [r for r in db.list_employees(region=region) if _norm_text(r[3]) == target]
    _emit(ctx, _rows(rows, EMP_COLS), EMP_COLS)


@department.command("rename")
@click.option("--from", "old_name", required=True, help="Mevcut departman adi.")
@click.option("--to", "new_name", required=True, help="Yeni departman adi.")
@click.option("--region", default=None)
@click.pass_context
def department_rename(ctx, old_name, new_name, region):
    """Bir departmanin tum calisanlarini yeni isme atar."""
    target = _norm_text(old_name)
    rows = [r for r in db.list_employees(region=region) if _norm_text(r[3]) == target]
    count = 0
    for r in rows:
        eid, full_name, identity_no, _dep, title, emp_region = r[:6]
        db.update_employee(eid, full_name, identity_no, new_name, title, emp_region)
        count += 1
    _emit(ctx, {"status": "ok", "renamed": count, "from": old_name, "to": new_name})


@department.command("assign")
@click.argument("employee_id", type=int)
@click.option("--department", required=True)
@click.pass_context
def department_assign(ctx, employee_id, department):
    """Bir calisanin departmanini degistirir."""
    emp = _get_employee(employee_id)
    db.update_employee(employee_id, emp["full_name"], emp["identity_no"],
                       department, emp["title"], emp["region"])
    _emit(ctx, {"status": "ok", "employee_id": employee_id, "department": department})


# ---------------------------------------------------------------- timesheet bulk
@timesheet.command("bulk")
@click.option("--records", "records_json", required=True, help="JSON dosyasi.")
@click.option("--region", default=None, help="Eksik bolge icin varsayilan.")
@click.option("--overwrite/--no-overwrite", default=False)
@click.pass_context
def timesheet_bulk(ctx, records_json, region, overwrite):
    """JSON kayit listesinden toplu puantaj girer (whatsapp apply ile ayni motor)."""
    with open(records_json, "r", encoding="utf-8") as f:
        entries = json.load(f)
    result = bulk_mod.apply_entries(
        db, entries, default_region=region, source="cli_bulk", overwrite=overwrite,
    )
    _emit(ctx, result.to_dict())


# ---------------------------------------------------------------- REPL
def _repl(ctx):
    click.echo("cli-anything-puantaj REPL. 'help' yardim, 'exit' cikis.")
    click.echo(f"DB: {db.DB_PATH}")
    while True:
        try:
            line = input("puantaj> ").strip()
        except (EOFError, KeyboardInterrupt):
            click.echo()
            break
        if not line:
            continue
        if line in ("exit", "quit"):
            break
        if line == "help":
            line = "--help"
        try:
            cli.main(args=shlex.split(line), prog_name="puantaj",
                     standalone_mode=False, obj=ctx.obj)
        except SystemExit:
            pass
        except click.ClickException as e:
            e.show()
        except Exception as e:  # REPL surekli calismali
            click.echo(f"hata: {e}")


def main():
    cli(obj={})


if __name__ == "__main__":
    main()
