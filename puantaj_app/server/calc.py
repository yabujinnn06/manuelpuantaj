from datetime import datetime, date


def parse_time(value):
    return datetime.strptime(value, "%H:%M").time()

def parse_date(value):
    if value is None or value == "":
        raise ValueError("Tarih bos olamaz.")
    value = str(value).strip()
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Invalid date format: {value}")


def hours_between(start_time, end_time):
    start_dt = datetime.combine(date.today(), start_time)
    end_dt = datetime.combine(date.today(), end_time)
    if end_dt < start_dt:
        end_dt = end_dt.replace(day=end_dt.day + 1)
    delta = end_dt - start_dt
    return delta.total_seconds() / 3600.0


def _to_minutes(value):
    return value.hour * 60 + value.minute


def _overlap(start_a, end_a, start_b, end_b):
    return max(0, min(end_a, end_b) - max(start_a, start_b))


def night_hours_between(start_time, end_time):
    start_min = _to_minutes(start_time)
    end_min = _to_minutes(end_time)
    if end_min <= start_min:
        end_min += 24 * 60
    total = 0
    windows = [
        (22 * 60, 24 * 60),
        (0, 6 * 60),
        (22 * 60 + 24 * 60, 24 * 60 + 24 * 60),
        (0 + 24 * 60, 6 * 60 + 24 * 60),
    ]
    for w_start, w_end in windows:
        total += _overlap(start_min, end_min, w_start, w_end)
    return total / 60.0


def overnight_hours_between(start_time, end_time):
    start_min = _to_minutes(start_time)
    end_min = _to_minutes(end_time)
    if end_min <= start_min:
        end_min += 24 * 60
    if end_min <= 24 * 60:
        return 0.0
    return (end_min - 24 * 60) / 60.0


def is_sunday_non_stand(work_date, department=None):
    work_dt = parse_date(work_date)
    department_key = str(department or "").strip().upper()
    return work_dt.weekday() == 6 and department_key != "STANT"


def calc_day_hours(work_date, start_time, end_time, break_minutes, settings, is_special=0, department=None):
    work_dt = parse_date(work_date)
    weekday = work_dt.weekday()  # 0=Mon, 5=Sat, 6=Sun

    st = parse_time(start_time)
    et = parse_time(end_time)
    gross_hours = hours_between(st, et)
    
    # Break minutes validation
    break_minutes = max(0, int(break_minutes))
    if break_minutes > gross_hours * 60:
        break_minutes = int(gross_hours * 60)  # Max break = gross hours
    
    worked_hours = gross_hours - (break_minutes / 60.0)
    if worked_hours < 0:
        worked_hours = 0.0
    night_hours = night_hours_between(st, et)
    overnight_hours = overnight_hours_between(st, et)

    weekday_hours = float(settings.get("weekday_hours", "8") or 8)
    saturday_start = settings.get("saturday_start", "09:00")
    saturday_end = settings.get("saturday_end", "14:00")

    department_key = str(department or "").strip().upper()

    if weekday <= 4:
        scheduled_hours = weekday_hours
    elif weekday == 5:
        if department_key == "STANT":
            scheduled_hours = 8.0
        else:
            sat_hours = hours_between(parse_time(saturday_start), parse_time(saturday_end))
            scheduled_hours = sat_hours
    else:
        if department_key == "STANT":
            scheduled_hours = 8.0
        else:
            scheduled_hours = 0.0

    if is_special:
        scheduled_hours = 0.0
        overtime_hours = 0.0
        special_normal = worked_hours
        special_overtime = 0.0
        special_night = night_hours
    else:
        if is_sunday_non_stand(work_date, department):
            overtime_hours = 0.0
        elif scheduled_hours == 0.0:
            overtime_hours = max(0.0, gross_hours)
        else:
            overtime_hours = max(0.0, gross_hours - scheduled_hours)
        special_normal = 0.0
        special_overtime = 0.0
        special_night = 0.0

    return (
        round(worked_hours, 2),
        round(scheduled_hours, 2),
        round(overtime_hours, 2),
        round(night_hours, 2),
        round(overnight_hours, 2),
        round(special_normal, 2),
        round(special_overtime, 2),
        round(special_night, 2),
    )
