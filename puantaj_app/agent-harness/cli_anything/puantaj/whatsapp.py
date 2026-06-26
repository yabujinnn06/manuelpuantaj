"""WhatsApp grup sohbeti puantaj mesajlarini ayristirir.

Desteklenen formatlar (esnek):

- Tarih basligi + satir satir kayit:
    01.01.2026 Pazartesi
    Ahmet Yilmaz 08:00-17:00 60
    Mehmet Demir izinli
    Ayse Kaya 09:00-18:00 mola 60
    Hasan Ozturk raporlu

- WhatsApp disa aktarim formati:
    [01.01.2026 09:15] Murat Sef: Bugun gelenler:
    [01.01.2026 09:15] Murat Sef: - Ahmet Yilmaz 08-17
    [01.01.2026 09:16] Murat Sef: - Mehmet izinli

- Tek satirda tarihli kayit:
    01.01.2026 Ahmet Yilmaz 08:00-17:00

Cikti: parse edilmis kayit listesi (dict). Bu liste preview Excel
ve bulk apply tarafindan tuketilir.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field, asdict
from datetime import datetime, date
from typing import Iterable


STATUS_KEYWORDS = {
    "Izinli": ["izinli", "izin", "yillik izin", "yillikizin"],
    "Raporlu": ["raporlu", "rapor", "saglik raporu"],
    "Gelmedi": ["gelmedi", "gelmiyor", "yok", "devamsiz", "kayip"],
    "Mazeret": ["mazeret", "mazeretli", "mazaret"],
    "Tatil": ["tatil", "resmi tatil", "bayram"],
    "Diger": ["diger"],
}

WORK_KEYWORDS = ["calisti", "geldi", "var", "mevcut"]

# Resmi olmayan kisaltmalar (08-17 -> 08:00-17:00, 8.30-17.30 -> 08:30-17:30)
TIME_RANGE_PATTERNS = [
    re.compile(r"(?P<h1>\d{1,2})[:.](?P<m1>\d{2})\s*[-–—]\s*(?P<h2>\d{1,2})[:.](?P<m2>\d{2})"),
    re.compile(r"(?P<h1>\d{1,2})\s*[-–—]\s*(?P<h2>\d{1,2})(?!\d)"),
]

BREAK_PATTERN = re.compile(
    r"(?:mola|break|ara)\s*[:=]?\s*(?P<m>\d{1,3})|\b(?P<n>\d{1,3})\s*(?:dk|dakika|min)\b",
    re.IGNORECASE,
)

DATE_PATTERNS = [
    # 2026-01-15 / 2026/01/15
    re.compile(r"\b(?P<y>20\d{2})[-/](?P<m>\d{1,2})[-/](?P<d>\d{1,2})\b"),
    # 15.01.2026 / 15/01/2026 / 15-01-2026
    re.compile(r"\b(?P<d>\d{1,2})[./-](?P<m>\d{1,2})[./-](?P<y>20\d{2})\b"),
    # 15.01 (yil yok; varsayilan yili kullaniriz)
    re.compile(r"\b(?P<d>\d{1,2})[./-](?P<m>\d{1,2})\b(?!\s*[\d:])"),
]

# WhatsApp export line head:  [12.01.2026 09:15:30] Sender Name: rest
WHATSAPP_HEAD = re.compile(
    r"^\[?(?P<date>\d{1,2}[./-]\d{1,2}[./-]\d{2,4})[,\s]+\d{1,2}[:.]\d{2}(?:[:.]\d{2})?\]?\s*[-–]?\s*"
    r"(?P<sender>[^:]+):\s*(?P<body>.*)$"
)

# Tarih basliklari icin gunluk listede ay/gun cevirme
TR_MONTHS = {
    "ocak": 1, "subat": 2, "şubat": 2, "mart": 3, "nisan": 4, "mayis": 5, "mayıs": 5,
    "haziran": 6, "temmuz": 7, "agustos": 8, "ağustos": 8, "eylul": 9, "eylül": 9,
    "ekim": 10, "kasim": 11, "kasım": 11, "aralik": 12, "aralık": 12,
}
TR_MONTH_PATTERN = re.compile(
    r"\b(?P<d>\d{1,2})\s+(?P<mon>" + "|".join(TR_MONTHS.keys()) + r")(?:\s+(?P<y>20\d{2}))?\b",
    re.IGNORECASE,
)

# Notlarda yakalanacak fazla mesai/gece/ozel gun isaretleri
SPECIAL_FLAGS = re.compile(r"\b(?:ozel\s*gun|özel\s*gün|tatil\s*calisma|tatil\s+çalışma|bayram\s+mesa)", re.IGNORECASE)


@dataclass
class ParsedEntry:
    employee_name_raw: str
    work_date: str  # YYYY-MM-DD
    status: str  # ATTENDANCE_STATUSES'ten biri
    start_time: str | None = None
    end_time: str | None = None
    break_minutes: int = 0
    is_special: bool = False
    notes: str = ""
    region: str | None = None
    employee_id: int | None = None  # matcher dolduracak
    matched_name: str | None = None
    confidence: float = 0.0  # 0..1
    warnings: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return asdict(self)


# ---------------------------------------------------------------- normalize
def _strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def normalize(text: str) -> str:
    text = (text or "").strip().casefold()
    text = text.replace("ı", "i").replace("İ", "i")
    text = _strip_accents(text)
    return re.sub(r"\s+", " ", text)


def _pad_hour(h: str) -> str:
    h = h.zfill(2)
    return h


def _parse_date(token: str, default_year: int) -> str | None:
    token = token.strip()
    for pat in DATE_PATTERNS:
        m = pat.search(token)
        if not m:
            continue
        gd = m.groupdict()
        try:
            year = int(gd.get("y") or default_year)
            month = int(gd["m"])
            day = int(gd["d"])
            return date(year, month, day).isoformat()
        except (ValueError, KeyError):
            continue
    m = TR_MONTH_PATTERN.search(token)
    if m:
        try:
            day = int(m.group("d"))
            month = TR_MONTHS[m.group("mon").lower().replace("ı", "i")]
            year = int(m.group("y") or default_year)
            return date(year, month, day).isoformat()
        except (ValueError, KeyError):
            return None
    return None


def _match_status(lower_text: str) -> str | None:
    for status, keys in STATUS_KEYWORDS.items():
        for k in keys:
            if re.search(r"\b" + re.escape(k) + r"\b", lower_text):
                return status
    if any(re.search(r"\b" + re.escape(k) + r"\b", lower_text) for k in WORK_KEYWORDS):
        return "Calisti"
    return None


def _match_time_range(text: str) -> tuple[str, str] | None:
    for pat in TIME_RANGE_PATTERNS:
        m = pat.search(text)
        if not m:
            continue
        gd = m.groupdict()
        h1 = _pad_hour(gd["h1"])
        h2 = _pad_hour(gd["h2"])
        m1 = gd.get("m1") or "00"
        m2 = gd.get("m2") or "00"
        try:
            datetime.strptime(f"{h1}:{m1}", "%H:%M")
            datetime.strptime(f"{h2}:{m2}", "%H:%M")
        except ValueError:
            continue
        return f"{h1}:{m1}", f"{h2}:{m2}"
    return None


_TRAILING_BREAK = re.compile(r"\s+(\d{1,3})\s*$")


def _extract_break(text: str, had_time_range: bool = False) -> int:
    m = BREAK_PATTERN.search(text)
    if m:
        val = m.group("m") or m.group("n")
        try:
            n = int(val)
        except (TypeError, ValueError):
            n = -1
        if 0 <= n <= 360:
            return n
    if had_time_range:
        # "Ahmet 08-17 60" gibi: zaman araliginin ardindaki kuru sayi mola.
        m2 = _TRAILING_BREAK.search(text)
        if m2:
            try:
                n = int(m2.group(1))
                if 0 <= n <= 360:
                    return n
            except ValueError:
                pass
    return 0


def _strip_status_tokens(text: str, status: str | None) -> str:
    if not status:
        return text
    out = text
    for keys in STATUS_KEYWORDS.values():
        for k in keys:
            out = re.sub(r"\b" + re.escape(k) + r"\b", "", out, flags=re.IGNORECASE)
    for k in WORK_KEYWORDS:
        out = re.sub(r"\b" + re.escape(k) + r"\b", "", out, flags=re.IGNORECASE)
    return out


_RESIDUAL_NOISE = re.compile(
    r"\b(?:dk|dakika|min|saat|mola|break|ara)\b|\d+",
    re.IGNORECASE,
)


def _extract_name(text: str) -> str:
    text = re.sub(r"^[\s\-\*•·~>•]+", "", text)
    text = re.sub(r"\s*\(.*?\)", "", text)
    text = _RESIDUAL_NOISE.sub("", text)
    text = re.sub(r"\s+", " ", text).strip(" -:,;")
    return text.strip()


# ---------------------------------------------------------------- main parser
def parse_text(
    text: str,
    default_date: str | None = None,
    default_year: int | None = None,
    region: str | None = None,
) -> list[ParsedEntry]:
    """Serbest metni ayristirip ParsedEntry listesi dondurur."""
    if not text or not text.strip():
        return []
    if default_year is None:
        default_year = (
            int(default_date[:4]) if default_date and len(default_date) >= 4 and default_date[:4].isdigit()
            else datetime.now().year
        )
    current_date = default_date
    entries: list[ParsedEntry] = []
    lines = text.splitlines()
    for raw_line in lines:
        line = raw_line.strip()
        if not line:
            continue
        # WhatsApp head: timestamp + sender + body
        head = WHATSAPP_HEAD.match(line)
        body = line
        if head:
            head_date = _parse_date(head.group("date"), default_year)
            if head_date:
                current_date = head_date
            body = head.group("body").strip()
            if not body:
                continue
        # Tarih basligi tek basina mi?
        if not WHATSAPP_HEAD.match(line):
            stripped = re.sub(r"^[\s\-\*•·]+", "", body)
            d = _parse_date(stripped, default_year)
            if d:
                # Tarih ile birlikte ayni satirda kayit varsa devam etmek istiyoruz
                rest = stripped
                for pat in DATE_PATTERNS:
                    rest = pat.sub("", rest)
                rest = TR_MONTH_PATTERN.sub("", rest)
                rest = re.sub(r"\s+", " ", rest).strip(" -:,;")
                current_date = d
                if not rest:
                    continue
                body = rest

        # Devamsizlik bulundu mu?
        body_norm = normalize(body)
        status = _match_status(body_norm)
        time_range = _match_time_range(body)
        break_min = _extract_break(body, had_time_range=time_range is not None)
        is_special = bool(SPECIAL_FLAGS.search(body))

        if time_range and not status:
            status = "Calisti"
        if not status and not time_range:
            # Calisma satiri degil; muhtemelen aciklama. Atla.
            continue

        name_part = body
        if time_range:
            for pat in TIME_RANGE_PATTERNS:
                name_part = pat.sub("", name_part)
        name_part = BREAK_PATTERN.sub("", name_part)
        name_part = SPECIAL_FLAGS.sub("", name_part)
        name_part = _strip_status_tokens(name_part, status)
        name_part = _extract_name(name_part)
        if not name_part:
            continue

        warnings: list[str] = []
        if not current_date:
            warnings.append("Tarih belirlenemedi; --default-date veya satirda tarih girin.")
        if status == "Calisti" and not time_range:
            warnings.append("Calisma satirinda saat araligi bulunamadi.")
        if time_range:
            try:
                from calc import hours_between, parse_time  # type: ignore
                hb = hours_between(parse_time(time_range[0]), parse_time(time_range[1]))
                if hb <= 0:
                    warnings.append("Calisma suresi 0 veya negatif gozukuyor.")
                if break_min > hb * 60:
                    warnings.append("Mola dakikasi calisma suresinden buyuk.")
            except Exception:
                pass

        entries.append(
            ParsedEntry(
                employee_name_raw=name_part,
                work_date=current_date or "",
                status=status,
                start_time=time_range[0] if time_range else None,
                end_time=time_range[1] if time_range else None,
                break_minutes=break_min,
                is_special=is_special,
                notes="WhatsApp puantaj",
                region=region,
                warnings=warnings,
            )
        )

    return entries


# ---------------------------------------------------------------- name matching
def _name_tokens(name: str) -> set[str]:
    return {t for t in normalize(name).split() if t}


def match_employees(
    entries: Iterable[ParsedEntry],
    employees: list[tuple],
    region: str | None = None,
) -> None:
    """Her ParsedEntry icin employees uzerinden en iyi eslesmeyi yazar.

    employees: db.list_employees() ciktisi (id, full_name, identity_no, department, title, region).
    """
    candidates: list[tuple[int, str, str | None]] = []
    for row in employees:
        eid = row[0]
        full_name = row[1]
        emp_region = row[5] if len(row) > 5 else None
        if region and emp_region and emp_region != region:
            continue
        candidates.append((eid, full_name, emp_region))

    for entry in entries:
        raw_tokens = _name_tokens(entry.employee_name_raw)
        if not raw_tokens:
            entry.warnings.append("Isim bos.")
            continue
        best_id = None
        best_name = None
        best_score = 0.0
        for eid, full_name, _ in candidates:
            cand_tokens = _name_tokens(full_name)
            if not cand_tokens:
                continue
            inter = raw_tokens & cand_tokens
            union = raw_tokens | cand_tokens
            score = len(inter) / len(union) if union else 0.0
            # Tam tek-kelime baslangici da bonus
            if any(t in cand_tokens for t in raw_tokens):
                score = max(score, 0.5)
            if score > best_score:
                best_score = score
                best_id = eid
                best_name = full_name
        if best_score >= 0.66:
            entry.employee_id = best_id
            entry.matched_name = best_name
            entry.confidence = round(best_score, 2)
        elif best_score >= 0.34:
            entry.employee_id = best_id
            entry.matched_name = best_name
            entry.confidence = round(best_score, 2)
            entry.warnings.append(f"Dusuk guvenle eslesti: {best_name}")
        else:
            entry.warnings.append(
                f"Calisan eslesmedi. Aday yok veya skor dusuk (en iyi: {best_name or '-'})."
            )
            entry.matched_name = best_name
            entry.confidence = round(best_score, 2)


def entries_to_dicts(entries: list[ParsedEntry]) -> list[dict]:
    return [e.as_dict() for e in entries]
