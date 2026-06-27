"""WhatsApp grup sohbeti puantaj mesajlarini ayristirir.

Gercek "Rainwater puantaj" grubu formatina gore tasarlandi. Desteklenen
ana ozellikler:

- WhatsApp Turkce disa aktarim basligi:
    6.01.2026 ogleden sonra 4:50 - Hasan Teknik: <mesaj>
  Turkce saat ifadeleri (ogleden once/sonra, sabah, aksam, gece, ...) taninir.

- Cok satirli tek mesaj: bir baslik satirindan sonraki (baslik olmayan)
  satirlar ayni mesaja aittir ve birlikte islenir.

- Gonderen != calisan: gercek isim cogu zaman mesajin icindedir
  (telefon numarasindan gelenlerde ozellikle). Mesaj ici isim onceliklidir.

- Cogu kisi sadece cikis saati yazar. Giris yoksa departman varsayilani
  (apply_shift_defaults) uygulanir.

- Ozel ifadeler: "Full"->10:00-22:00, "Haftalik/Yillik/Tam gun izin"->Izinli,
  "Pazar mesaisi"->pazar calismasi, "Yilbasi"->Tatil.

- Bir mesajda birden fazla gun (tarih blogu) olabilir; her blok ayri islenir.

Akis: parse_text -> match_employees -> apply_shift_defaults.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field, asdict
from datetime import datetime, date
from typing import Iterable


# ---------------------------------------------------------------- sabitler
TR_TIME_OF_DAY = (
    "ogleden once", "ogleden sonra", "gece yarisi",
    "sabah", "aksamustu", "aksam", "ogle", "gece", "oglen",
)

# Mesaj basligi: "6.01.2026 <turkce-saat> 4:50 - <kalan>"
# Turkce kelimeler aksansiz normalize edilmis metin uzerinde aranir.
_HEAD_RE = re.compile(
    r"^(?P<date>\d{1,2}\.\d{1,2}\.\d{2,4})\s+"
    r"(?:ogleden once|ogleden sonra|gece yarisi|sabah|aksamustu|aksam|oglen|ogle|gece)\s+"
    r"\d{1,2}:\d{2}\s+-\s+(?P<rest>.*)$"
)

# Sistem mesaji kaliplari (gonderen yok veya puantaj disi)
_SYSTEM_MARKERS = (
    "kisisini ekledi", "kisisini cikardi", "kisisini eklediniz",
    "grubunu olusturdu", "sizi ekledi", "bir grup baglantisiyla katildi",
    "bu mesaj silindi", "bu mesaji sildiniz", "medya dahil edilmedi",
    "telefon numarasini", "artik yoneticisiniz", "bir mesaji sabitlediniz",
    "mesajlar ve aramalar uctan uca", "guvenlik kodu degisti",
)

# Calisma durumu anahtar kelimeleri (aksansiz, casefold)
_STATUS_PATTERNS = [
    ("Izinli", re.compile(r"\b(?:haftalik izin|yillik izin|tam gun izin|ucretsiz izin|"
                          r"izinli|izin|haftalik|haftal[i1]k|yillik)\b")),
    ("Raporlu", re.compile(r"\b(?:raporlu|rapor|saglik raporu)\b")),
    ("Mazeret", re.compile(r"\b(?:yarim gun izin|mazeret izni|mazeret|mazaret)\b")),
    ("Tatil", re.compile(r"\b(?:yilbasi|resmi tatil|bayram tatili)\b")),
]

# Pazar mesaisi isareti
_SUNDAY_WORK_RE = re.compile(r"pazar\s*mesa")
_FULL_RE = re.compile(r"\bfull\b")
_PAZAR_ONLY_RE = re.compile(r"^\s*pazar\s*$")

# Saat tokeni: 09:30 / 19.30 / 08;30 / 07/30 / 0800 / 5.30
_CLOCK = r"\d{1,2}\s*[:.;/]\s*\d{2}|\b\d{3,4}\b"

# "giris 09:30" (etiket once, saat sonra) - newline gecebilir (cok satirli)
_GIRIS_LABEL = re.compile(
    r"(?:giris|gir[i1]s)\s*(?:saati|saat)?\s*[:.\-]?[ \t]*(" + _CLOCK + r")"
)
_CIKIS_LABEL = re.compile(
    r"(?:cikis|c[i1]k[i1]s)\s*(?:saati|saat)?\s*[:.\-]?[ \t]*(" + _CLOCK + r")"
)
# "08:00 GIRIS" / "23:00 cikis" / "08:00. GIRIS" (saat once, etiket sonra) - ayni satir
# (aksi halde "Giris 10:00\nCikis 18:00"da 10:00'i cikisa baglar)
_GIRIS_LABEL_AFTER = re.compile(r"(" + _CLOCK + r")[ \t.]*(?:giris|gir[i1]s)")
_CIKIS_LABEL_AFTER = re.compile(r"(" + _CLOCK + r")[ \t.]*(?:cikis|c[i1]k[i1]s)")

# Tarih: 05.01.2026 / 05.01.26 / 5.1.2026
_DATE_RE = re.compile(r"\b(\d{1,2})\.(\d{1,2})\.(\d{2,4})\b")

# Departman bazli varsayilan giris saatleri
DEPARTMENT_DEFAULT_START = {
    "stant": None,      # vardiyaya gore hesaplanir
    "teknik": "08:30",
    "lojistik": "08:30",
    "sofor": "08:30",
    "ofis": "09:00",
    "ik": "09:00",
}
GENERIC_DEFAULT_START = "08:30"
FULL_SHIFT = ("10:00", "22:00")


@dataclass
class ParsedEntry:
    employee_name_raw: str
    work_date: str  # YYYY-MM-DD
    status: str
    start_time: str | None = None
    end_time: str | None = None
    break_minutes: int = 0
    is_special: bool = False
    is_sunday_work: bool = False
    start_assumed: bool = False  # giris departman varsayilanindan dolduruldu mu
    notes: str = ""
    region: str | None = None
    sender: str | None = None
    employee_id: int | None = None
    matched_name: str | None = None
    department: str | None = None
    confidence: float = 0.0
    warnings: list[str] = field(default_factory=list)

    def as_dict(self) -> dict:
        return asdict(self)


# ---------------------------------------------------------------- normalize
def _strip_accents(text: str) -> str:
    text = unicodedata.normalize("NFKD", text)
    return "".join(ch for ch in text if not unicodedata.combining(ch))


def _fold(text: str) -> str:
    """Aksansiz, kucuk harf, Turkce i/I duzeltmeli normalize."""
    text = (text or "").replace("İ", "i").replace("I", "i").replace("ı", "i")
    text = text.casefold()
    text = _strip_accents(text)
    return re.sub(r"\s+", " ", text).strip()


def _strip_math_bold(text: str) -> str:
    """Unicode matematiksel bold/italik harfleri ASCII'ye indirger (𝐁𝐚𝐬̧𝐚𝐤 -> Basak)."""
    out = []
    for ch in text:
        decomp = unicodedata.normalize("NFKC", ch)
        out.append(decomp)
    return unicodedata.normalize("NFKC", text)


def normalize_clock(token: str) -> str | None:
    """'19.30'->'19:30', '0800'->'08:00', '5.30'->'05:30', '08;30'->'08:30'."""
    if token is None:
        return None
    t = str(token).strip()
    t = re.sub(r"\s+", "", t)
    m = re.match(r"^(\d{1,2})[:.;/](\d{2})$", t)
    if m:
        h, mi = int(m.group(1)), int(m.group(2))
    else:
        m = re.match(r"^(\d{3,4})$", t)
        if not m:
            return None
        digits = m.group(1)
        if len(digits) == 3:
            h, mi = int(digits[0]), int(digits[1:])
        else:
            h, mi = int(digits[:2]), int(digits[2:])
    if h == 24:
        h = 0
    if not (0 <= h <= 23 and 0 <= mi <= 59):
        return None
    return f"{h:02d}:{mi:02d}"


def _parse_date(day: str, month: str, year: str, default_year: int) -> str | None:
    try:
        d, mo = int(day), int(month)
        y = int(year)
        if y < 100:
            y += 2000
        return date(y, mo, d).isoformat()
    except (ValueError, TypeError):
        return None


# ---------------------------------------------------------------- mesaj gruplama
def _split_messages(text: str, default_year: int):
    """(header_date, sender, body_lines) ureten jenerator. Multi-line birlestirir."""
    current = None  # dict(date, sender, lines)
    for raw in text.splitlines():
        line = raw.rstrip()
        folded = _fold(line)
        head = _HEAD_RE.match(folded)
        if head:
            # Yeni mesaj basligi. Onceki mesaji yayinla.
            if current is not None:
                yield current
            # rest'i ORIJINAL satirdan al (folded degil) ki isim/saat korunur
            rest_orig = _extract_rest_original(line)
            sender, body = _split_sender(rest_orig)
            parts = head.group("date").split(".")
            head_date = _parse_date(parts[0], parts[1], parts[2], default_year) \
                if len(parts) >= 3 else None
            current = {
                "date": head_date,
                "sender": sender,
                "lines": [body] if body else [],
            }
        else:
            if current is not None:
                current["lines"].append(line)
            # baslik yoksa ve hic mesaj baslamadiysa: atla (export ust bilgisi)
    if current is not None:
        yield current


def _extract_rest_original(line: str) -> str:
    """Orijinal satirdan baslik kismindan sonraki gercek metni alir."""
    # "6.01.2026 ogleden sonra 4:50 - REST"  -> REST
    m = re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}\s+\S.*?\s+\d{1,2}:\d{2}\s+-\s+(.*)$", line)
    if m:
        return m.group(1)
    # Turkce kelime iki sozcuklu olabilir (ogleden sonra); genis yakala
    m = re.match(r"^\d{1,2}\.\d{1,2}\.\d{2,4}\s+.+?\s+\d{1,2}:\d{2}\s+-\s+(.*)$", line)
    return m.group(1) if m else line


def _split_sender(rest: str) -> tuple[str | None, str]:
    """'Gonderen: body' -> (gonderen, body). ':' yoksa (None, rest)."""
    # Telefon numarasi gonderen: "+90 505 074 24 61: body"
    m = re.match(r"^(\+?\d[\d\s]{6,}):\s*(.*)$", rest)
    if m:
        return None, m.group(2)  # telefon -> gonderen yok say
    m = re.match(r"^([^:]{1,40}?):\s*(.*)$", rest)
    if m:
        sender = m.group(1).strip()
        return sender, m.group(2)
    return None, rest


def _is_system_message(sender: str | None, body_lines: list[str]) -> bool:
    joined = _fold(" ".join(body_lines))
    if any(mark in joined for mark in _SYSTEM_MARKERS):
        return True
    if sender:
        sfold = _fold(sender)
        if any(mark in sfold for mark in _SYSTEM_MARKERS):
            return True
    return False


# ---------------------------------------------------------------- isim cikarma
_NAME_NOISE = re.compile(
    r"\b(?:giris|gir[i1]s|cikis|c[i1]k[i1]s|saati|saat|full|izin|izinli|haftalik|"
    r"yillik|mazeret|rapor|raporlu|pazar|mesai|mesaisi|yilbasi|cumartesi|tatil|"
    r"sabah|aksam|gece|ogleden|tam|gun|yarim|abi|stant|teknik|rainwater|avm|"
    r"anatolium|ofis|demo|trafik|montaj|qr|kod|okut)\b"
)


def _looks_like_name(line: str) -> bool:
    folded = _fold(line)
    if not folded:
        return False
    if any(ch.isdigit() for ch in line):
        return False
    if ":" in line:
        return False
    # Saf isim satiri noise kelime (izinli, cumartesi, giris, full, ...) icermez.
    if _NAME_NOISE.search(folded):
        return False
    words = [w for w in re.split(r"\s+", line.strip()) if w]
    # 1-3 kelime ad-soyad
    if not (1 <= len(words) <= 3):
        return False
    letters = sum(ch.isalpha() for ch in line)
    return letters >= 3


def _extract_name(body_lines: list[str], sender: str | None) -> str:
    """Mesaj govdesinden en olasi calisan ismini cikarir; yoksa gonderen adini."""
    candidates = []
    for ln in body_lines:
        clean = _strip_math_bold(ln).strip()
        # Satir icinde gomulu isim de olabilir: "28.03.2026 Ugur Erturk 09:28 giris..."
        if _looks_like_name(clean):
            candidates.append(clean)
        elif _DATE_RE.search(clean) or re.search(_CLOCK, clean):
            # SADECE tarih/saat iceren satirlarda gomulu isim ara
            # ("28.03.2026 Ugur Erturk 09:28 giris ..."). Boylece duz prose
            # satirlari ("Stan kuruldu ertesi sabaha kadar") isim sayilmaz.
            words = re.findall(r"[^\W\d_]{2,}", clean, flags=re.UNICODE)
            words = [w for w in words if not _NAME_NOISE.search(_fold(w))]
            if len(words) >= 1:
                candidates.append(" ".join(words[:3]))
    if candidates:
        # En cok kelimeli / en uzun adayi sec (ad soyad)
        candidates.sort(key=lambda c: (len(c.split()), len(c)), reverse=True)
        return _clean_name(candidates[0])
    if sender and not _fold(sender).replace(" ", "").isdigit():
        return _clean_name(_strip_math_bold(sender))
    return ""


def _clean_name(name: str) -> str:
    name = re.sub(r"\s+", " ", name).strip(" -:,;.")
    return name


# ---------------------------------------------------------------- saat/durum cikarma
def _extract_times(text_block: str) -> tuple[str | None, str | None]:
    """Giris/cikis saatlerini cikarir. SATIR SATIR calisir ki cok-satirli
    'Giris 10:00 / Cikis 18:00' bloklarinda giris cikisa karismasin."""
    start = end = None
    for raw_line in text_block.splitlines():
        folded = _fold(raw_line)
        # Tarihleri maskele ki yil (2026) saat (20:26) olarak yakalanmasin.
        folded = _DATE_RE.sub(" ", folded)
        if not folded.strip():
            continue
        if start is None:
            # Once "saat + etiket" (09:28 giris), sonra "etiket + saat" (giris 09:28).
            m = _GIRIS_LABEL_AFTER.search(folded) or _GIRIS_LABEL.search(folded)
            if m:
                start = normalize_clock(m.group(1))
        if end is None:
            m = _CIKIS_LABEL_AFTER.search(folded) or _CIKIS_LABEL.search(folded)
            if m:
                end = normalize_clock(m.group(1))
    return start, end


def _match_status(text_block: str) -> str | None:
    folded = _fold(text_block)
    for status, pat in _STATUS_PATTERNS:
        if pat.search(folded):
            return status
    return None


# ---------------------------------------------------------------- blok bolme
def _split_date_blocks(body_lines: list[str], header_date: str | None, default_year: int):
    """Bir mesaj govdesini tarih bloklarina ayirir.

    Cogu mesajda tek tarih vardir; ama Hasan gibi bir mesajda birden fazla gun
    yazanlar icin her tarih yeni bir blok baslatir.
    Yield: (work_date, block_lines)
    """
    blocks = []
    current_lines = []
    current_date = None
    found_any_date = False
    for ln in body_lines:
        m = _DATE_RE.search(ln)
        if m:
            iso = _parse_date(m.group(1), m.group(2), m.group(3), default_year)
            if iso:
                # Yeni tarih: onceki blogu kapat
                if found_any_date and (current_lines or current_date):
                    blocks.append((current_date, current_lines))
                    current_lines = []
                current_date = iso
                found_any_date = True
        current_lines.append(ln)
    if current_lines or current_date:
        blocks.append((current_date if found_any_date else header_date, current_lines))
    if not blocks:
        blocks = [(header_date, body_lines)]
    return blocks


# ---------------------------------------------------------------- ana parser
def parse_text(text: str, default_date: str | None = None,
               default_year: int | None = None, region: str | None = None) -> list[ParsedEntry]:
    if not text or not text.strip():
        return []
    if default_year is None:
        default_year = datetime.now().year
        if default_date and len(default_date) >= 4 and default_date[:4].isdigit():
            default_year = int(default_date[:4])

    entries: list[ParsedEntry] = []
    for msg in _split_messages(text, default_year):
        sender = msg["sender"]
        body_lines = [ln for ln in msg["lines"] if ln.strip()]
        if not body_lines:
            continue
        if _is_system_message(sender, body_lines):
            continue

        # Mesaj genelinde tek kisi varsayimi: ismi mesaj seviyesinde belirle.
        # Cok-gunlu mesajlarda (Hasan gibi) isim hep en ustte/altta bir kez yazilir;
        # blok bazli isim aramak not satirlarindan yanlis isim cikarir.
        message_name = _extract_name(body_lines, sender)

        for work_date, block in _split_date_blocks(body_lines, msg["date"], default_year):
            block_text = "\n".join(block)
            folded = _fold(block_text)

            start, end = _extract_times(block_text)
            status = _match_status(block_text)
            is_full = bool(_FULL_RE.search(folded))
            is_sunday = bool(_SUNDAY_WORK_RE.search(folded))
            is_pazar_only = any(_PAZAR_ONLY_RE.match(_fold(l)) for l in block)

            # Bir calisma kaydi mi? (saat / full / durum / pazar mesaisi)
            has_signal = bool(start or end or status or is_full or is_sunday)
            if not has_signal and not is_pazar_only:
                # Sadece sohbet/aciklama -> atla
                continue

            # Blokta belirgin/farkli bir isim varsa onu, yoksa mesaj-seviyesi ismi kullan.
            name = message_name or _extract_name(block, sender)
            if not name:
                continue

            warnings: list[str] = []
            is_special = False
            note_bits = ["WhatsApp puantaj"]

            if is_full:
                start = start or FULL_SHIFT[0]
                end = end or FULL_SHIFT[1]
                note_bits.append("Full vardiya")

            if is_sunday:
                is_special = False
                note_bits.append("Pazar mesaisi")
            if is_pazar_only and not has_signal:
                status = status or "Tatil"
                note_bits.append("Pazar")

            # Durum belirleme
            if status in ("Izinli", "Raporlu", "Mazeret", "Tatil"):
                final_status = status
                start = end = None
            elif start or end or is_full or is_sunday:
                final_status = "Calisti"
            else:
                final_status = status or "Diger"

            # Tarih dogrulama / yil typo duzeltme
            if not work_date:
                warnings.append("Tarih belirlenemedi.")
            elif msg["date"]:
                try:
                    wd = date.fromisoformat(work_date)
                    hd = date.fromisoformat(msg["date"])
                    if abs((wd - hd).days) > 200 and wd.year != hd.year:
                        # Yil typo'su (orn 06.01.2016 -> 06.01.2026): mesaj gun/ay
                        # ayniysa mesaj yilini kullan.
                        try:
                            corrected = wd.replace(year=hd.year)
                            warnings.append(
                                f"Yil {wd.year}->{hd.year} duzeltildi (mesaj gunune gore).")
                            work_date = corrected.isoformat()
                        except ValueError:
                            warnings.append(
                                f"Tarih mesaj gununden uzak ({work_date}); kontrol et.")
                    elif abs((wd - hd).days) > 45:
                        warnings.append(
                            f"Tarih mesaj gununden uzak ({work_date} vs mesaj {msg['date']}); kontrol et.")
                except ValueError:
                    pass

            if final_status == "Calisti" and not end:
                warnings.append("Cikis saati bulunamadi.")

            entries.append(ParsedEntry(
                employee_name_raw=name,
                work_date=work_date or "",
                status=final_status,
                start_time=start,
                end_time=end,
                break_minutes=0,
                is_special=is_special,
                is_sunday_work=is_sunday,
                notes=" | ".join(note_bits),
                region=region,
                sender=sender,
                warnings=warnings,
            ))

    return entries


# ---------------------------------------------------------------- isim eslestirme
def _name_tokens(name: str) -> set[str]:
    folded = _fold(_strip_math_bold(name))
    return {t for t in folded.split() if len(t) >= 2}


def match_employees(entries: Iterable[ParsedEntry], employees: list[tuple],
                    region: str | None = None) -> None:
    candidates = []
    for row in employees:
        eid = row[0]
        full_name = row[1]
        department = row[3] if len(row) > 3 else None
        emp_region = row[5] if len(row) > 5 else None
        if region and emp_region and emp_region != region:
            continue
        candidates.append((eid, full_name, department, _name_tokens(full_name)))

    for entry in entries:
        raw_tokens = _name_tokens(entry.employee_name_raw)
        # Gonderen adini da aday tokenlara kat (mesaj ici isim zayifsa)
        sender_tokens = _name_tokens(entry.sender or "")
        search_tokens = raw_tokens | sender_tokens
        if not search_tokens:
            entry.warnings.append("Isim bulunamadi.")
            continue
        best = (None, None, None, 0.0)
        for eid, full_name, dept, cand_tokens in candidates:
            if not cand_tokens:
                continue
            inter = search_tokens & cand_tokens
            if not inter:
                continue
            union = search_tokens | cand_tokens
            score = len(inter) / len(union)
            # Soyad/ad tam eslesmesi bonus
            if cand_tokens <= search_tokens or search_tokens <= cand_tokens:
                score = max(score, 0.8)
            if len(inter) >= 2:
                score = max(score, 0.85)
            if score > best[3]:
                best = (eid, full_name, dept, score)
        eid, full_name, dept, score = best
        if score >= 0.6:
            entry.employee_id = eid
            entry.matched_name = full_name
            entry.department = dept
            entry.confidence = round(score, 2)
            if score < 0.75:
                entry.warnings.append(f"Orta guvenle eslesti: {full_name}")
        else:
            entry.matched_name = full_name
            entry.confidence = round(score, 2)
            entry.warnings.append(
                f"Calisan eslesmedi (en yakin: {full_name or '-'}). Excel'de duzeltin.")


# ---------------------------------------------------------------- vardiya varsayilanlari
def _dept_key(department: str | None) -> str:
    return _fold(department or "").replace(" ", "")


def department_default_start(department: str | None, end_time: str | None) -> tuple[str | None, bool]:
    """Giris saati yoksa departman/vardiya varsayilanini dondurur.

    Return: (start_time, assumed)  assumed=True ise tahmin edilmistir.
    """
    key = _dept_key(department)
    if "stant" in key:
        # Stant: cikisa gore vardiya tahmini
        if end_time:
            try:
                eh = int(end_time.split(":")[0])
            except (ValueError, IndexError):
                eh = None
            if eh is not None:
                if eh <= 18:
                    return "10:00", True
                return "14:00", True
        return "10:00", True
    for token, default in DEPARTMENT_DEFAULT_START.items():
        if token != "stant" and token in key and default:
            return default, True
    return GENERIC_DEFAULT_START, True


def apply_shift_defaults(entries: Iterable[ParsedEntry]) -> None:
    """Calisti kayitlarinda giris yoksa departman varsayilanini doldurur."""
    for e in entries:
        if e.status != "Calisti":
            continue
        if not e.start_time and e.end_time:
            start, assumed = department_default_start(e.department, e.end_time)
            if start:
                e.start_time = start
                e.start_assumed = assumed
                if assumed:
                    e.notes = (e.notes + " | Giris varsayilan").strip(" |")
                    e.warnings.append(
                        f"Giris saati mesajda yok; departman varsayilani ({start}) kullanildi.")


def entries_to_dicts(entries: list[ParsedEntry]) -> list[dict]:
    return [e.as_dict() for e in entries]
