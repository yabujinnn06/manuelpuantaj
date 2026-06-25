def parse_month(value):
    """Validate and normalize a month string in 'YYYY-MM' format."""
    if not value or not isinstance(value, str):
        raise ValueError("Ay bos olamaz. Ornek: 2026-01")
    value = value.strip()
    try:
        dt = datetime.strptime(value, "%Y-%m")
        return dt.strftime("%Y-%m")
    except Exception:
        raise ValueError("Ay formati gecersiz. Ornek: 2026-01")
import os
import csv
import zipfile
import logging
import traceback
import sys
import queue
import shutil
import calendar
from datetime import datetime, date, time, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import time as time_module

import calc
from openpyxl import load_workbook
from tkcalendar import DateEntry

import puantaj_db as db
import report

try:
    import winsound
except ImportError:
    winsound = None
try:
    import requests
except ImportError:
    requests = None

DATE_FMT = "YYYY-MM-DD"
TIME_FMT = "HH:MM"
KEEPALIVE_SECONDS = 300
REGIONS = ["Ankara", "Izmir", "Bursa", "Istanbul"]
VIEW_REGIONS = ["Tum Bolgeler"] + REGIONS
DEFAULT_OIL_INTERVAL_KM = 14000
DEFAULT_OIL_SOON_KM = 2000
LOG_DIR = os.path.join(os.path.dirname(db.DB_DIR), "logs")
LOG_PATH = os.path.join(LOG_DIR, "rainstaff.log")

ATTENDANCE_STATUSES = ["Calisti", "Izinli", "Gelmedi", "Raporlu", "Mazeret", "Tatil", "Diger"]
LEAVE_TYPES = ["Yillik Izin", "Raporlu", "Ucretsiz Izin", "Mazeret", "Diger"]
LEAVE_STATUSES = ["Onayli", "Beklemede", "Reddedildi"]

VEHICLE_CHECKLIST = [
    ("body_dent", "Govde ezik/cizik"),
    ("paint_damage", "Boya hasari"),
    ("interior_clean", "Ic temizligi"),
    ("smoke_smell", "Sigara kokusu"),
    ("tire_condition", "Lastik durumu"),
    ("lights", "Far/stop/sinyal"),
    ("glass", "Camlar"),
    ("warning_lamps", "Ikaz lambalari"),
    ("water_level", "Su seviyesi"),
]

EMP_HEADER_ALIASES = {
    "full_name": ["ad soyad", "adsoyad", "calisan", "calisan adi", "name", "full_name"],
    "identity_no": ["tckn", "tc", "tc kimlik", "identity", "identity_no"],
    "department": ["departman", "department"],
    "title": ["unvan", "title"],
}

TS_HEADER_ALIASES = {
    "employee": ["calisan", "ad soyad", "employee", "full_name", "name"],
    "work_date": ["tarih", "date", "work_date"],
    "start_time": ["giris", "start", "start_time"],
    "end_time": ["cikis", "end", "end_time"],
    "break_minutes": ["mola", "mola dk", "mola dakika", "break", "break_minutes"],
    "is_special": ["ozel gun", "ozel", "resmi tatil", "special", "is_special"],
    "notes": ["not", "notes", "aciklama"],
}


def normalize_header(value):
    """Lowercase and strip spaces/underscores for loose header matching."""
    return str(value or "").strip().lower().replace(" ", "").replace("_", "")


def build_header_aliases(alias_config):
    """Expand alias lists into normalized lookup sets per target column."""
    result = {}
    for target, aliases in alias_config.items():
        normalized = {normalize_header(target)}
        for alias in aliases:
            normalized.add(normalize_header(alias))
        result[target] = normalized
    return result


def days_until(date_str):
    """Return days from today to the given ISO date (positive = future, negative = past)."""
    if not date_str:
        return None
    try:
        target = datetime.strptime(date_str, "%Y-%m-%d").date()
    except ValueError:
        return None
    return (target - datetime.now().date()).days


def normalize_date_value(value):
    """Import icin flexible tarih normalizasyonu; Excel float veya string kabul eder."""
    if value is None or value == "":
        raise ValueError("Tarih bos olamaz.")
    if isinstance(value, str):
        return normalize_date(value)
    if isinstance(value, (int, float)):
        try:
            dt = datetime.fromordinal(int(value) + 693594)
            return dt.strftime("%Y-%m-%d")
        except (ValueError, OverflowError):
            raise ValueError(f"Tarih formati gecersiz: {value}")
    raise ValueError(f"Tarih formati gecersiz: {value}")


def normalize_time_value(value):
    """Import icin flexible saat normalizasyonu."""
    if value is None or value == "":
        raise ValueError("Saat bos olamaz.")
    if isinstance(value, str):
        return normalize_time(value)
    if isinstance(value, (int, float)):
        try:
            total_minutes = int(round(value * 24 * 60))
            hours = (total_minutes // 60) % 24
            minutes = total_minutes % 60
            return f"{hours:02d}:{minutes:02d}"
        except (ValueError, OverflowError):
            raise ValueError(f"Saat formati gecersiz: {value}")
    raise ValueError(f"Saat formati gecersiz: {value}")


def parse_bool(value):
    """String veya int boolean'a cevir."""
    if isinstance(value, bool):
        return value
    if isinstance(value, int):
        return value != 0
    if isinstance(value, str):
        return value.lower().strip() in {"1", "true", "yes", "evet", "e"}
    return bool(value)


def split_display_name(display, regions):
    """`Adi (Region)` formatindan baz adi ve bolgeyi ayir."""
    text = str(display or "").strip()
    if not text:
        return "", None
    if text.endswith(")") and "(" in text:
        base, suffix = text.rsplit("(", 1)
        base = base.strip()
        region = suffix[:-1].strip()  # sondaki ) kaldir
        if region in regions:
            return base, region
        if region == "-":  # belirsiz/boş
            return base, None
    return text, None


def week_start_from_date(value):
    """Verilen tarihi ait haftanin pazartesi baslangicina cek."""
    iso = normalize_date(str(value))
    d = datetime.strptime(iso, "%Y-%m-%d").date()
    start = d - timedelta(days=d.weekday())  # Pazartesi
    return start.strftime("%Y-%m-%d")


def week_end_from_start(week_start):
    """Hafta baslangicindan pazar gununu uret."""
    d = datetime.strptime(week_start, "%Y-%m-%d").date()
    return (d + timedelta(days=6)).strftime("%Y-%m-%d")


def calc_sunday_separate_hours(work_date, department, is_special, worked_hours):
    """Pazar (STANT haric) calisma saatini normal fazla mesaiden ayri takip et."""
    if is_special:
        return 0.0
    try:
        if calc.is_sunday_non_stand(work_date, department):
            return max(0.0, float(worked_hours))
    except Exception:
        return 0.0
    return 0.0


def normalize_time_in_var(var):
    """StringVar icindeki saati normalize eder; hatada eski degeri korur."""
    try:
        normalized = normalize_time(var.get())
        var.set(normalized)
    except Exception:
        pass


def ensure_app_dirs():
    data_dir = os.path.join(os.path.dirname(__file__), "data")
    if not os.path.isdir(data_dir):
        os.makedirs(data_dir, exist_ok=True)


def setup_logging():
    os.makedirs(LOG_DIR, exist_ok=True)
    logger = logging.getLogger("rainstaff")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    file_handler = logging.FileHandler(LOG_PATH, encoding="utf-8")
    file_handler.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    # Konsola da yaz (komut penceresinde anlık görmek için)
    try:
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        logger.addHandler(console_handler)
    except Exception:
        # Konsol eklenemese de dosyaya yazmaya devam etsin
        pass

    def handle_uncaught(exc_type, exc_value, exc_traceback):
        if issubclass(exc_type, KeyboardInterrupt):
            return
        logger.error("Unhandled exception", exc_info=(exc_type, exc_value, exc_traceback))

    sys.excepthook = handle_uncaught
    return logger


class LogQueueHandler(logging.Handler):
    def __init__(self, log_queue):
        super().__init__()
        self.log_queue = log_queue
        self.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))

    def emit(self, record):
        try:
            msg = self.format(record)
            self.log_queue.put(msg)
        except Exception:
            pass


def parse_int(value, default=0):
    try:
        return int(value)
    except (ValueError, TypeError):
        return default


def parse_float(value, default=0.0):
    try:
        return float(value)
    except (ValueError, TypeError):
        return default


def normalize_date(value):
    if value is None or value == "":
        raise ValueError("Tarih bos olamaz.")
    value = str(value).strip()
    if not value:
        raise ValueError("Tarih bos olamaz.")
    for fmt in ("%Y-%m-%d", "%d.%m.%Y"):
        try:
            return datetime.strptime(value, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    raise ValueError("Tarih formati gecersiz. Ornek: 2026-01-05 veya 05.01.2026")


def normalize_time(value):
    if value is None:
        raise ValueError("Saat bos olamaz.")
    if isinstance(value, datetime):
        return value.strftime("%H:%M")
    if isinstance(value, time):
        return value.strftime("%H:%M")
    if isinstance(value, (int, float)) and 0 <= value < 1:
        total_minutes = int(round(value * 24 * 60))
        hours = (total_minutes // 60) % 24
        minutes = total_minutes % 60
        return f"{hours:02d}:{minutes:02d}"
    text = str(value).strip()
    if not text:
        raise ValueError("Saat bos olamaz.")
    text = text.replace(".", ":")
    if text.isdigit() and len(text) in (3, 4):
        if len(text) == 3:
            text = "0" + text
        return f"{text[:2]}:{text[2:]}"
    if ":" in text:
        parts = text.split(":")
        if len(parts) >= 2:
            return f"{parts[0].zfill(2)}:{parts[1].zfill(2)}"
    raise ValueError("Saat formati gecersiz. Ornek: 09:30")


def normalize_vehicle_status(status):
    """Araç muayene durumunu normalize et (Olumsuz/Olumlu/Belirsiz)"""
    if status is None:
        return "Belirsiz"
    text = str(status).strip().lower()
    if "olumsuz" in text or "bad" in text or "0" in text or "no" in text:
        return "Olumsuz"
    if "olumlu" in text or "good" in text or "ok" in text or "1" in text or "yes" in text:
        return "Olumlu"
    return "Belirsiz"


EMP_HEADER_MAP = build_header_aliases(EMP_HEADER_ALIASES)
TS_HEADER_MAP = build_header_aliases(TS_HEADER_ALIASES)


def map_headers(header_row, header_map):
    mapping = {}
    for idx, value in enumerate(header_row):
        key = normalize_header(value)
        for target, aliases in header_map.items():
            if key in aliases:
                mapping[target] = idx
    return mapping


def load_tabular_file(path):
    ext = os.path.splitext(path)[1].lower()
    rows = []
    if ext == ".csv":
        with open(path, newline="", encoding="utf-8-sig") as handle:
            sample = handle.read(4096)
            handle.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample)
            except csv.Error:
                dialect = csv.get_dialect("excel")
            reader = csv.reader(handle, dialect)
            for row in reader:
                rows.append(row)
    else:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            rows.append(list(row))
    return [row for row in rows if any(cell is not None and str(cell).strip() for cell in row)]


def create_labeled_entry(parent, label, textvariable, width=24):
    frame = ttk.Frame(parent)
    ttk.Label(frame, text=label).pack(side=tk.LEFT, padx=(0, 8))
    ttk.Entry(frame, textvariable=textvariable, width=width).pack(side=tk.LEFT)
    return frame


def create_labeled_date(parent, label, textvariable, width=12):
    frame = ttk.Frame(parent)
    ttk.Label(frame, text=label).pack(side=tk.LEFT, padx=(0, 8))
    date_entry = DateEntry(frame, textvariable=textvariable, width=width, date_pattern="yyyy-mm-dd")
    date_entry.pack(side=tk.LEFT)
    return frame, date_entry


def create_time_entry(parent, label, textvariable, width=8):
    frame = ttk.Frame(parent)
    ttk.Label(frame, text=label).pack(side=tk.LEFT, padx=(0, 8))
    entry = ttk.Entry(frame, textvariable=textvariable, width=width)
    entry.pack(side=tk.LEFT)
    return frame


def set_time_vars(time_value, textvariable):
    try:
        normalized = normalize_time(time_value)
    except ValueError:
        return
    textvariable.set(normalized)


def clear_date_entry(entry):
    try:
        entry.delete(0, tk.END)
    except Exception:
        pass


def attach_tooltip(widget, text):
    if not text:
        return
    tip = {"window": None}

    def show_tip(_event=None):
        if tip["window"] or not widget.winfo_exists():
            return
        x = widget.winfo_pointerx() + 10
        y = widget.winfo_pointery() + 12
        win = tk.Toplevel(widget)
        win.wm_overrideredirect(True)
        win.wm_geometry(f"+{x}+{y}")
        label = tk.Label(
            win,
            text=text,
            bg="#202020",
            fg="#E0E0E0",
            font=("Segoe UI", 9),
            padx=8,
            pady=4,
            borderwidth=1,
            relief="solid",
        )
        label.pack()
        tip["window"] = win

    def hide_tip(_event=None):
        if tip["window"]:
            try:
                tip["window"].destroy()
            except Exception:
                pass
            tip["window"] = None

    widget.bind("<Enter>", show_tip)
    widget.bind("<Leave>", hide_tip)
    widget.bind("<ButtonPress>", hide_tip)


def insert_empty_row(tree, columns, message):
    values = [message] + [""] * (len(columns) - 1)
    tree.insert("", tk.END, values=values, tags=("empty",))


def create_kpi_card(parent, title, value_var, subtitle=None, theme=None, accent=None):
    if theme is None:
        theme = {
            "bg_content": "#1F1F1F",
            "bg_hover": "#2A2A2A",
            "text_primary": "#E0E0E0",
            "text_secondary": "#8C8C8C",
            "primary": "#5B9BD5",
        }
    accent = accent or theme.get("primary", "#5B9BD5")

    card = tk.Frame(
        parent,
        bg=theme["bg_content"],
        highlightthickness=1,
        highlightbackground=theme["bg_hover"],
    )
    top = tk.Frame(card, bg=accent, height=3)
    top.pack(fill=tk.X)
    body = tk.Frame(card, bg=theme["bg_content"])
    body.pack(fill=tk.BOTH, expand=True, padx=14, pady=10)
    title_lbl = tk.Label(body, text=title, bg=theme["bg_content"], fg=theme["text_secondary"], font=("Segoe UI", 9))
    title_lbl.pack(anchor="w")
    value_lbl = tk.Label(
        body,
        textvariable=value_var,
        bg=theme["bg_content"],
        fg=theme["text_primary"],
        font=("Segoe UI", 16, "bold"),
    )
    value_lbl.pack(anchor="w", pady=(4, 0))
    if subtitle:
        sub_lbl = tk.Label(
            body, text=subtitle, bg=theme["bg_content"], fg=theme["text_secondary"], font=("Segoe UI", 9)
        )
        sub_lbl.pack(anchor="w", pady=(4, 0))

    def _set_bg(bg):
        card.configure(bg=bg)
        body.configure(bg=bg)
        title_lbl.configure(bg=bg)
        value_lbl.configure(bg=bg)
        if subtitle:
            sub_lbl.configure(bg=bg)

    def _on_enter(_event=None):
        _set_bg(theme["bg_hover"])
        card.configure(highlightbackground=accent)

    def _on_leave(_event=None):
        _set_bg(theme["bg_content"])
        card.configure(highlightbackground=theme["bg_hover"])

    card.bind("<Enter>", _on_enter)
    card.bind("<Leave>", _on_leave)
    body.bind("<Enter>", _on_enter)
    body.bind("<Leave>", _on_leave)
    return card


def ensure_logo_asset(path):
    if os.path.isfile(path):
        return
    try:
        from PIL import Image, ImageDraw, ImageFont
    except Exception:
        return
    width, height = 320, 80
    img = Image.new("RGB", (width, height), "#e7eef6")
    draw = ImageDraw.Draw(img)
    draw.rounded_rectangle([8, 8, width - 8, height - 8], radius=18, fill="#cfe1f2")
    # Minimalist rain drop icon
    drop_x, drop_y = 28, 18
    draw.polygon(
        [(drop_x + 16, drop_y), (drop_x, drop_y + 24), (drop_x + 16, drop_y + 46), (drop_x + 32, drop_y + 24)],
        fill="#2f6fed",
        outline="#2a5fd1",
    )
    draw.ellipse([drop_x + 6, drop_y + 24, drop_x + 26, drop_y + 44], fill="#2f6fed", outline="#2a5fd1")
    draw.ellipse([drop_x + 14, drop_y + 10, drop_x + 20, drop_y + 16], fill="#bfe0ff")
    text = "RAINSTAFF"
    try:
        font = ImageFont.truetype("Consola.ttf", 22)
    except Exception:
        font = ImageFont.load_default()
    draw.text((84, 22), text, fill="#0f2438", font=font)
    try:
        sub_font = ImageFont.truetype("Consola.ttf", 14)
    except Exception:
        sub_font = ImageFont.load_default()
    draw.text((86, 48), "PUANTAJ", fill="#425466", font=sub_font)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    img.save(path)


# Rehber kaldırıldı - Modern ERP tasarımıyla değiştirildi


def load_logo_image(path, target_height=48):
    try:
        from PIL import Image, ImageTk
    except Exception:
        Image = None
        ImageTk = None

    if Image and ImageTk:
        try:
            img = Image.open(path)
            ratio = target_height / float(img.height)
            target_width = int(img.width * ratio)
            img = img.resize((target_width, target_height), Image.LANCZOS)
            return ImageTk.PhotoImage(img)
        except Exception:
            return None
    try:
        return tk.PhotoImage(file=path)
    except Exception:
        return None


class PuantajApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Rainstaff ERP - Puantaj Yönetimi")
        self.geometry("1280x800")
        self.minsize(1024, 720)

        self.logger = setup_logging()
        self.report_callback_exception = self._handle_tk_exception
        self.log_queue = queue.Queue()
        self.log_handler = LogQueueHandler(self.log_queue)
        if self.logger:
            self.logger.addHandler(self.log_handler)

        self.current_user = None
        self.current_region = None
        self.is_admin = False

        self.settings = db.get_all_settings()
        self.themes = {
            "Gece": {
                "bg_app": "#1E1E1E",
                "bg_content": "#2A2A2A",
                "bg_elevated": "#323232",
                "bg_input": "#363636",
                "bg_hover": "#3A3A3A",
                "text_primary": "#E0E0E0",
                "text_secondary": "#B0B0B0",
                "text_disabled": "#707070",
                "primary": "#5B9BD5",
                "primary_hover": "#7BB3E0",
                "accent_gold": "#C9A961"
            },
            "Sabah": {
                "bg_app": "#F1F3F5",
                "bg_content": "#FAFAFB",
                "bg_elevated": "#E6E8EB",
                "bg_input": "#F3F4F6",
                "bg_hover": "#E0E4EA",
                "text_primary": "#1F2933",
                "text_secondary": "#5B6470",
                "text_disabled": "#9AA3AD",
                "primary": "#3A7BD5",
                "primary_hover": "#2F6BBE",
                "accent_gold": "#DFA84A"
            },
            "Matrix": {
                "bg_app": "#0E1621",
                "bg_content": "#16212E",
                "bg_elevated": "#1F2B3A",
                "bg_input": "#203040",
                "bg_hover": "#2A3B4D",
                "text_primary": "#E6EEF5",
                "text_secondary": "#AAB8C6",
                "text_disabled": "#6E7B88",
                "primary": "#4F8CC9",
                "primary_hover": "#66A3E0",
                "accent_gold": "#7FB3E6"
            }
        }
        self.current_theme = self.settings.get("theme", "Gece")
        entry_region = self.settings.get("admin_entry_region", "Ankara")
        view_region = self.settings.get("admin_view_region", "Tum Bolgeler")
        if view_region == "ALL":
            view_region = "Tum Bolgeler"
        self.admin_entry_region_var = tk.StringVar(value=entry_region)
        self.admin_view_region_var = tk.StringVar(value=view_region)
        self.employee_map = {}
        self.employee_display_names = []
        self.employee_details = {}
        self.vehicle_map = {}
        self.driver_map = {}
        self.driver_display_names = []
        self.fault_map = {}
        self.service_visit_map = {}
        self.shift_template_map = {}
        self.status_var = tk.StringVar()
        self.ts_original = None
        self.ts_editing_id = None
        self.vehicle_original_plate = None
        self._tab_loaded = {}

        if not self._login_prompt():
            self.destroy()
            # Bu dosya, BACKUP_2026_01_18 içeriğiyle tamamen değiştirilecek...

        self._show_loading("Yukleniyor...")
        self.after(10, self._finish_startup)

    def _handle_tk_exception(self, exc, val, tb):
        if self.logger:
            self.logger.error("Tkinter callback error", exc_info=(exc, val, tb))
        messagebox.showerror("Hata", "Beklenmeyen bir hata olustu. Log kaydi alindi.")

    def _log_action(self, action, detail=""):
        if not self.logger:
            return
        user = self.current_user or "unknown"
        region = self.current_region or "ALL"
        suffix = f" | {detail}" if detail else ""
        self.logger.info("ACTION: %s | user=%s | region=%s%s", action, user, region, suffix)

    def _finish_startup(self):
        self.after(10, self._startup_step_prepare)
        if self.logger:
            self.logger.info("Uygulama basladi")

    def _startup_step_prepare(self):
        self.after(10, self._startup_step_style)

    def _startup_step_style(self):
        self._configure_style()
        self.after(10, self._startup_step_ui)

    def _startup_step_ui(self):
        self._build_ui()
        self.after(10, self._startup_step_data)

    def _startup_step_data(self):
        self._load_tab_data(self.tab_employees)
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._hide_loading()

    def _show_loading(self, text):
        overlay = tk.Toplevel(self)
        overlay.overrideredirect(True)
        overlay.attributes("-topmost", True)
        overlay.configure(bg="#0f1115")

        width, height = 420, 240
        screen_w = overlay.winfo_screenwidth()
        screen_h = overlay.winfo_screenheight()
        x = int((screen_w - width) / 2)
        y = int((screen_h - height) / 2)
        overlay.geometry(f"{width}x{height}+{x}+{y}")
        overlay.resizable(False, False)
        overlay.transient(self)
        overlay.grab_set()
        overlay.protocol("WM_DELETE_WINDOW", lambda: None)

        card = tk.Frame(overlay, bg="#151821", highlightbackground="#2b2f3a", highlightthickness=1)
        card.place(x=10, y=10, width=width - 20, height=height - 20)

        glow = tk.Canvas(card, bg="#151821", highlightthickness=0, height=60)
        glow.pack(fill=tk.X)
        glow.create_oval(20, -40, 200, 80, fill="#1f3550", outline="")
        glow.create_oval(120, -50, 320, 70, fill="#203c5a", outline="")

        logo_path = os.path.join(os.path.dirname(__file__), "assets", "rainstaff_logo_1.png")
        logo_img = load_logo_image(logo_path, target_height=56)
        if logo_img:
            self._loading_logo = logo_img
            tk.Label(card, image=logo_img, bg="#151821").pack(pady=(6, 4))
        else:
            tk.Label(card, text="RAINSTAFF", bg="#151821", fg="#7BB3E0",
                     font=("Segoe UI", 16, "bold")).pack(pady=(10, 4))

        tk.Label(card, text=text, bg="#151821", fg="#C8D3E0",
                 font=("Segoe UI", 11, "bold")).pack(pady=(0, 4))
        tk.Label(card, text="Sistem hazirlaniyor...", bg="#151821", fg="#6E7A8C",
                 font=("Segoe UI", 9)).pack(pady=(0, 8))

        spinner = tk.Canvas(card, width=140, height=30, bg="#151821", highlightthickness=0)
        spinner.pack(pady=(2, 6))
        dots = []
        for i in range(3):
            dot = spinner.create_oval(12 + i * 40, 10, 28 + i * 40, 26, fill="#7BB3E0", outline="")
            dots.append(dot)

        progress = tk.Canvas(card, width=260, height=10, bg="#151821", highlightthickness=0)
        progress.pack(pady=(6, 6))
        progress.create_rectangle(0, 0, 260, 10, fill="#1f242e", outline="")
        shimmer = progress.create_rectangle(-60, 0, 0, 10, fill="#7BB3E0", outline="")

        self._loading_overlay = overlay
        self._loading_spinner = spinner
        self._loading_started_at = time_module.time()

        def step(idx=0, pos=-60):
            if getattr(self, "_loading_overlay", None) is None:
                return
            for i, dot in enumerate(dots):
                color = "#7BB3E0" if i == idx else "#3d4a5f"
                spinner.itemconfigure(dot, fill=color)
            pos = pos + 8
            if pos > 260:
                pos = -60
            progress.coords(shimmer, pos, 0, pos + 60, 10)
            overlay.after(120, step, (idx + 1) % len(dots), pos)

        step()
        overlay.update_idletasks()
        overlay.lift()

    def _hide_loading(self):
        min_show = 0.6
        if getattr(self, "_loading_started_at", None):
            elapsed = time_module.time() - self._loading_started_at
            if elapsed < min_show:
                self.after(int((min_show - elapsed) * 1000), self._hide_loading)
                return
        if hasattr(self, "_loading_spinner"):
            self._loading_spinner = None
        if hasattr(self, "_loading_overlay"):
            try:
                self._loading_overlay.grab_release()
                self._loading_overlay.destroy()
            except Exception:
                pass
            self._loading_overlay = None
        self._loading_started_at = None

    def _start_keepalive(self):
        self._keepalive_stop = threading.Event()
        thread = threading.Thread(target=self._keepalive_worker, daemon=True)
        thread.start()

    def _keepalive_worker(self):
        while not self._keepalive_stop.wait(KEEPALIVE_SECONDS):
            if requests is None:
                continue
            settings = db.get_all_settings()
            enabled = settings.get("sync_enabled") == "1"
            sync_url = settings.get("sync_url", "").strip()
            token = settings.get("sync_token", "").strip()
            if not enabled or not sync_url:
                continue
            try:
                headers = {"X-API-KEY": token} if token else {}
                url = sync_url.rstrip("/") + "/health"
                requests.get(url, headers=headers, timeout=6)
            except Exception:
                pass

    def _on_close(self):
        if hasattr(self, "_keepalive_stop"):
            self._keepalive_stop.set()
        self.destroy()

    def _login_prompt(self):
        dialog = tk.Toplevel(self)
        dialog.title("Giris")
        dialog.resizable(False, False)
        dialog.configure(bg="#1E1E1E")  # Koyu arka plan
        dialog.transient(self)
        dialog.grab_set()

        # Card görünümü için sabit boyut ve merkezleme
        dialog.geometry("420x460")
        dialog.update_idletasks()
        if self.winfo_ismapped():
            px = self.winfo_rootx()
            py = self.winfo_rooty()
            pw = self.winfo_width()
            ph = self.winfo_height()
            dx = px + (pw - 420) // 2
            dy = py + (ph - 460) // 2
            dialog.geometry(f"420x460+{dx}+{dy}")

        card = tk.Frame(dialog, bg="#2A2A2A", highlightbackground="#3A3A3A", highlightthickness=1)
        card.place(x=16, y=16, relwidth=1, relheight=1, width=-32, height=-32)

        logo_path = os.path.join(os.path.dirname(__file__), "assets", "rainstaff_logo_1.png")
        logo_img = load_logo_image(logo_path, target_height=48)
        if logo_img:
            self._login_logo = logo_img
            tk.Label(card, image=logo_img, bg="#2A2A2A").pack(pady=(48, 12))
        else:
            tk.Label(card, text="RAINSTAFF", bg="#2A2A2A", fg="#C9A961",
                     font=("Segoe UI", 16, "bold")).pack(pady=(48, 12))

        tk.Label(card, text="Giris Yapin", bg="#2A2A2A", fg="#E0E0E0",
                 font=("Segoe UI", 16, "bold")).pack(pady=(0, 24))

        form_frame = tk.Frame(card, bg="#2A2A2A")
        form_frame.pack(padx=56, fill=tk.X)

        username_var = tk.StringVar()
        password_var = tk.StringVar()

        tk.Label(form_frame, text="Kullanici Adi", bg="#2A2A2A", fg="#5B9BD5",
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(0, 6))
        username_entry = tk.Entry(form_frame, textvariable=username_var,
                                  font=("Segoe UI", 11), relief="solid", bd=1,
                                  bg="#363636", fg="#E0E0E0",
                                  insertbackground="#5B9BD5",
                                  highlightthickness=1, highlightbackground="#454545")
        username_entry.pack(fill=tk.X, ipady=10)

        tk.Label(form_frame, text="Sifre", bg="#2A2A2A", fg="#5B9BD5",
                 font=("Segoe UI", 10, "bold"), anchor="w").pack(fill=tk.X, pady=(24, 6))
        password_entry = tk.Entry(form_frame, textvariable=password_var, show="●",
                                  font=("Segoe UI", 11), relief="solid", bd=1,
                                  bg="#363636", fg="#E0E0E0",
                                  insertbackground="#5B9BD5",
                                  highlightthickness=1, highlightbackground="#454545")
        password_entry.pack(fill=tk.X, ipady=10)

        status_var = tk.StringVar()
        error_label = tk.Label(card, textvariable=status_var, bg="#2A2A2A", fg="#E57373",
                               font=("Segoe UI", 9))
        error_label.pack(pady=(20, 0))

        success = {"ok": False}

        def attempt_login():
            user = db.verify_user(username_var.get().strip(), password_var.get().strip())
            if not user:
                status_var.set("❌ Kullanici adi veya sifre hatali")
                if self.logger:
                    self.logger.warning("Giris basarisiz: %s", username_var.get().strip())
                return
            self.current_user = user["username"]
            self.is_admin = user["role"] == "admin"
            region = user.get("region") or "Ankara"
            self.current_region = region
            if self.is_admin:
                self.admin_entry_region_var.set(region)
                self.admin_view_region_var.set("Tum Bolgeler")
            else:
                self.admin_entry_region_var.set(region)
                self.admin_view_region_var.set(region)
            success["ok"] = True
            dialog.destroy()

        btn_frame = tk.Frame(card, bg="#2A2A2A")
        btn_frame.pack(pady=(28, 40), padx=56, fill=tk.X)

        login_btn = tk.Button(btn_frame, text="Giris Yap", command=attempt_login,
                              bg="#5B9BD5", fg="#1E1E1E", font=("Segoe UI", 11, "bold"),
                              relief="flat", cursor="hand2", bd=0)
        login_btn.pack(fill=tk.X, ipady=12)

        def on_enter(_event):
            login_btn.config(bg="#7BB3E0")

        def on_leave(_event):
            login_btn.config(bg="#5B9BD5")

        login_btn.bind("<Enter>", on_enter)
        login_btn.bind("<Leave>", on_leave)

        dialog.bind("<Return>", lambda _e: attempt_login())
        username_entry.focus_set()

        self.wait_window(dialog)
        return success["ok"]

    def _toggle_theme(self):
        """Temalar arasında geçiş yap: Gece -> Sabah -> Matrix -> Gece"""
        theme_order = ["Gece", "Sabah", "Matrix"]
        current_idx = theme_order.index(self.current_theme) if self.current_theme in theme_order else 0
        next_idx = (current_idx + 1) % len(theme_order)
        self.current_theme = theme_order[next_idx]
        
        # Ayarı kaydet
        db.set_setting("theme", self.current_theme)
        
        # Tema ikonunu güncelle
        icons = {"Gece": "🌙", "Sabah": "☀️", "Matrix": "💚"}
        if hasattr(self, 'theme_btn'):
            self.theme_btn.config(text=icons.get(self.current_theme, "💡"))
        
        # Stili yeniden uygula
        self._configure_style()
        
        # Bildirim göster
        self.status_var.set(f"Tema: {self.current_theme}")

    def _view_region(self):
        if self.is_admin:
            value = self.admin_view_region_var.get().strip()
            if value in {"Tum Bolgeler", "ALL"}:
                return None
            return value or None
        return self.current_region

    def _refresh_region_views(self):
        self.refresh_employees()
        self.refresh_timesheets()
        self.refresh_dashboard()
        self.refresh_attendance()
        self.refresh_leave_records()
        self.refresh_admin_summary()

    def _entry_region(self):
        if self.is_admin:
            value = self.admin_entry_region_var.get().strip()
            return value or "Ankara"
        return self.current_region

    def _configure_style(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        # Dinamik tema renkleri
        theme = self.themes.get(self.current_theme, self.themes["Gece"])
        
        primary = theme["primary"]
        primary_hover = theme["primary_hover"]
        accent_gold = theme["accent_gold"]
        bg_app = theme["bg_app"]
        bg_content = theme["bg_content"]
        bg_elevated = theme["bg_elevated"]
        bg_input = theme["bg_input"]
        bg_hover = theme["bg_hover"]
        text_primary = theme["text_primary"]
        text_secondary = theme["text_secondary"]
        text_disabled = theme["text_disabled"]

        self.configure(bg=bg_app)
        self._ui_theme = theme

        style.configure("Header.TLabel",
            font=("Segoe UI", 18, "bold"),
            foreground=accent_gold,
            background=bg_app,
            padding=(0, 8, 0, 12))
        style.configure("SubHeader.TLabel",
            font=("Segoe UI", 12),
            foreground=text_secondary,
            background=bg_app,
            padding=(0, 4, 0, 8))

        style.configure("Section.TLabelframe",
            padding=(24, 20),
            relief="flat",
            borderwidth=1,
            background=bg_content)
        style.configure("Section.TLabelframe.Label",
            font=("Segoe UI", 11, "bold"),
            foreground=primary,
            background=bg_content,
            padding=(0, 0, 0, 8))

        style.configure("Accent.TButton",
            padding=(20, 10),
            background=primary,
            foreground=bg_app,
            borderwidth=0,
            relief="flat",
            font=("Segoe UI", 10, "bold"))
        style.map("Accent.TButton",
            background=[("active", primary_hover), ("pressed", primary_hover)],
            foreground=[("active", bg_app), ("pressed", bg_app)])

        style.configure("TButton",
            padding=(16, 8),
            background=bg_elevated,
            foreground=text_primary,
            borderwidth=1,
            relief="solid",
            font=("Segoe UI", 10))
        style.map("TButton",
            background=[("active", bg_hover), ("pressed", bg_hover)],
            foreground=[("active", primary), ("pressed", primary)])

        style.configure("Treeview",
            rowheight=36,
            fieldbackground=bg_content,
            background=bg_content,
            foreground=text_primary,
            borderwidth=1,
            font=("Segoe UI", 10))
        style.configure("Treeview.Heading",
            font=("Segoe UI", 10, "bold"),
            background=bg_elevated,
            foreground=primary,
            borderwidth=1,
            relief="flat",
            padding=(8, 8))
        style.map("Treeview.Heading",
            background=[("active", bg_hover)],
            foreground=[("active", accent_gold)])
        style.map("Treeview",
            background=[("selected", "#3A4A5A")],
            foreground=[("selected", text_primary)])

        style.configure("TFrame", background=bg_app)
        style.configure("Card.TFrame", background=bg_content, relief="flat")
        style.configure("TLabel",
            background=bg_app,
            font=("Segoe UI", 10),
            foreground=text_primary)
        style.configure("CardLabel.TLabel",
            background=bg_content,
            font=("Segoe UI", 10),
            foreground=text_primary)
        style.configure("Status.TLabel",
            background=bg_elevated,
            foreground=text_secondary,
            font=("Segoe UI", 9),
            padding=(12, 6))
        self._tree_colors = {
            "odd": bg_content,
            "even": bg_hover,
            "empty": bg_content,
            "text": text_primary,
            "muted": text_secondary,
        }

        style.configure("TNotebook",
            background=bg_app,
            borderwidth=0)
        style.configure("TNotebook.Tab",
            padding=(16, 8),
            background=bg_content,
            foreground=text_secondary,
            font=("Segoe UI", 10, "bold"))
        style.map(
            "TNotebook.Tab",
            background=[("selected", bg_elevated), ("active", bg_hover)],
            foreground=[("selected", text_primary), ("active", text_primary)],
        )

        style.configure("TEntry",
            padding=(12, 10),
            borderwidth=1,
            relief="solid",
            fieldbackground=bg_input,
            foreground=text_primary,
            selectbackground=bg_hover,
            selectforeground=text_primary,
            insertwidth=2,
            insertcolor=primary,
            font=("Segoe UI", 10))
        style.map("TEntry",
            fieldbackground=[("focus", bg_elevated)],
            bordercolor=[("focus", primary)])

        style.configure("TCombobox",
            padding=(12, 10),
            borderwidth=1,
            fieldbackground=bg_input,
            readonlybackground=bg_input,
            foreground=text_primary,
            selectbackground=bg_hover,
            selectforeground=text_primary,
            arrowsize=14,
            font=("Segoe UI", 10))
        style.map("TCombobox",
            fieldbackground=[("readonly", bg_input), ("!active", bg_input)],
            foreground=[("readonly", text_primary)],
            selectbackground=[("readonly", bg_hover)],
            selectforeground=[("readonly", text_primary)])

        # Dropdown list ve genel selection renkleri
        self.option_add("*TCombobox*Listbox.background", bg_content)
        self.option_add("*TCombobox*Listbox.foreground", text_primary)
        self.option_add("*TCombobox*Listbox.selectBackground", bg_hover)
        self.option_add("*TCombobox*Listbox.selectForeground", text_primary)
        self.option_add("*Entry.selectBackground", bg_hover)
        self.option_add("*Entry.selectForeground", text_primary)

        style.configure("TScrollbar",
            troughcolor=bg_content,
            background=bg_hover,
            bordercolor=bg_hover,
            arrowcolor=text_secondary)

        style.configure("TLabelframe", background=bg_content, borderwidth=0)
        style.configure("TLabelframe.Label", background=bg_content, foreground=text_secondary)
        style.configure("TMenubutton", background=bg_content, foreground=text_primary, borderwidth=0)
        style.configure("Toolbutton", background=bg_content, foreground=text_primary)
        style.configure("TCheckbutton", background=bg_app, foreground=text_primary, indicatorcolor=bg_input)

        # Hide notebook tabs (we use custom nav bar instead)
        style.layout("Hidden.TNotebook.Tab", [])
        style.configure("Hidden.TNotebook", tabmargins=0)

    def _apply_tree_zebra(self, tree):
        colors = getattr(self, "_tree_colors", None)
        if not colors:
            return
        tree.tag_configure("odd", background=colors["odd"], foreground=colors["text"])
        tree.tag_configure("even", background=colors["even"], foreground=colors["text"])
        tree.tag_configure("empty", background=colors["empty"], foreground=colors["muted"])

    def _auto_select_tree_first(self, tree, callback=None):
        for item in tree.get_children():
            tags = tree.item(item, "tags") or ()
            if "empty" in tags:
                continue
            tree.selection_set(item)
            tree.focus(item)
            tree.see(item)
            if callback:
                try:
                    callback()
                except TypeError:
                    callback(None)
            return True
        return False

    def _set_pane_sash(self, pane, ratio=0.7):
        try:
            total = max(1, pane.winfo_width())
            pane.sashpos(0, int(total * ratio))
        except Exception:
            pass

    def _build_ui(self):
        self.title("Rainstaff Puantaj")
        self.geometry("1280x800")
        self.minsize(1100, 700)
        theme = self.themes.get(self.current_theme, self.themes["Gece"])
        self.configure(bg=theme["bg_app"])
        header_bg = theme["bg_elevated"]
        header_fg = theme["text_primary"]
        sub_fg = theme["text_secondary"]

        header = tk.Frame(self, bg=header_bg, height=80)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        header.grid_columnconfigure(0, weight=1)
        header.grid_columnconfigure(1, weight=1)
        header.grid_columnconfigure(2, weight=0)

        left_block = tk.Frame(header, bg=header_bg)
        left_block.grid(row=0, column=0, sticky="w", padx=20, pady=12)

        ascii_path = os.path.join(os.path.dirname(__file__), "assets", "ascii.png")
        self._ascii_image = load_logo_image(ascii_path, target_height=28)
        if self._ascii_image:
            tk.Label(left_block, image=self._ascii_image, bg=header_bg).pack(side=tk.LEFT, padx=(0, 12))

        title_block = tk.Frame(left_block, bg=header_bg)
        title_block.pack(side=tk.LEFT)
        tk.Label(title_block, text="RAINSTAFF", bg=header_bg,
            fg=header_fg, font=("Segoe UI", 13, "bold")).pack(anchor="w")
        tk.Label(title_block, text="Puantaj Yönetimi", bg=header_bg,
            fg=sub_fg, font=("Segoe UI", 9)).pack(anchor="w")

        search_block = tk.Frame(header, bg=header_bg)
        search_block.grid(row=0, column=1, sticky="ew", padx=12, pady=16)
        search_block.grid_columnconfigure(1, weight=1)
        tk.Label(search_block, text="Hizli Arama", bg=header_bg, fg=sub_fg,
                 font=("Segoe UI", 9, "bold")).grid(row=0, column=0, sticky="w", padx=(0, 8))
        self.global_search_var = tk.StringVar()
        self.global_search_entry = ttk.Entry(search_block, textvariable=self.global_search_var)
        self.global_search_entry.grid(row=0, column=1, sticky="ew")
        self.global_search_entry.bind("<Return>", lambda _e: self._on_global_search())
        btn_search = ttk.Button(search_block, text="Ara", style="Accent.TButton", command=self._on_global_search)
        btn_search.grid(row=0, column=2, padx=8)
        attach_tooltip(self.global_search_entry, "Calisan / not / tarih anahtar kelime")
        attach_tooltip(btn_search, "Bulundugun sekmede ara")

        user_info = tk.Frame(header, bg=header_bg)
        user_info.grid(row=0, column=2, sticky="e", padx=20, pady=16)

        icons = {"Gece": "🌙", "Sabah": "☀️", "Matrix": "💚"}
        self.theme_btn = tk.Button(
            user_info,
            text=icons.get(self.current_theme, "💡"),
            command=self._toggle_theme,
            bg=header_bg,
            fg=theme["accent_gold"],
            font=("Segoe UI", 11, "bold"),
            relief="flat",
            cursor="hand2",
            bd=0,
            activebackground=header_bg,
            activeforeground=theme["accent_gold"],
        )
        self.theme_btn.pack(side=tk.RIGHT, padx=(0, 12))
        attach_tooltip(self.theme_btn, "Tema degistir (Gece / Sabah / Matrix)")

        user_text = f"{self.current_user}"
        if self.is_admin:
            user_text += " (Admin)"
        region_text = self._view_region() or "Tum Bolgeler"
        region_chip = tk.Label(
            user_info,
            text=region_text,
            bg=theme["bg_hover"],
            fg=header_fg,
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=3,
        )
        region_chip.pack(side=tk.RIGHT, padx=(0, 8))
        tk.Label(user_info, text=user_text, bg=header_bg,
            fg=sub_fg, font=("Segoe UI", 10)).pack(side=tk.RIGHT, padx=(0, 8))

        divider = tk.Frame(self, bg=theme["bg_hover"], height=1)
        divider.pack(fill=tk.X)

        self.notebook = ttk.Notebook(self, style="Hidden.TNotebook")

        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_employees = ttk.Frame(self.notebook)
        self.tab_timesheets = ttk.Frame(self.notebook)
        self.tab_attendance = ttk.Frame(self.notebook)
        self.tab_reports = ttk.Frame(self.notebook)
        self.tab_settings = ttk.Frame(self.notebook)
        self.tab_admin = ttk.Frame(self.notebook)
        self.tab_logs = ttk.Frame(self.notebook)

        self.notebook.add(self.tab_dashboard, text="Dashboard")
        self.notebook.add(self.tab_timesheets, text="Puantaj")
        self.notebook.add(self.tab_attendance, text="İzin & Yoklama")
        self.notebook.add(self.tab_employees, text="Çalışanlar")
        self.notebook.add(self.tab_reports, text="Raporlar")
        self.notebook.add(self.tab_admin, text="Yönetim")
        self.notebook.add(self.tab_settings, text="Ayarlar")
        self.notebook.add(self.tab_logs, text="Loglar")

        nav_frame = tk.Frame(self, bg=theme["bg_content"])
        nav_frame.pack(fill=tk.X, padx=10, pady=(6, 0))

        def nav_button(text, tab, icon=""):
            btn = tk.Button(
                nav_frame,
                text=f"{icon} {text}".strip(),
                command=lambda: self._switch_tab(tab),
                bg=theme["bg_content"],
                fg=header_fg,
                font=("Segoe UI", 9, "bold"),
                relief="flat",
                cursor="hand2",
                bd=0,
                padx=10,
                pady=6,
                activebackground=theme["bg_hover"],
                activeforeground=header_fg,
            )
            btn.pack(side=tk.LEFT, padx=4)
            attach_tooltip(btn, f"{text} sekmesine git")
            return btn

        self.nav_buttons = {
            "dashboard": nav_button("Dashboard", self.tab_dashboard, "📊"),
            "timesheets": nav_button("Puantaj", self.tab_timesheets, "🧾"),
            "attendance": nav_button("İzin/Yoklama", self.tab_attendance, "🗓️"),
            "employees": nav_button("Çalışanlar", self.tab_employees, "👥"),
            "reports": nav_button("Raporlar", self.tab_reports, "📄"),
            "admin": nav_button("Yönetim", self.tab_admin, "🧠"),
            "settings": nav_button("Ayarlar", self.tab_settings, "⚙️"),
            "logs": nav_button("Loglar", self.tab_logs, "🧾"),
        }

        self.notebook.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)

        self.tab_dashboard_body = self._make_tab_scrollable(self.tab_dashboard)
        self.tab_employees_body = self._make_tab_scrollable(self.tab_employees)
        self.tab_timesheets_body = self._make_tab_scrollable(self.tab_timesheets)
        self.tab_attendance_body = self._make_tab_scrollable(self.tab_attendance)
        self.tab_reports_body = self._make_tab_scrollable(self.tab_reports)
        self.tab_settings_body = self._make_tab_scrollable(self.tab_settings)
        self.tab_admin_body = self._make_tab_scrollable(self.tab_admin)
        self.tab_logs_body = self._make_tab_scrollable(self.tab_logs)

        self._build_dashboard_tab()
        self._build_employees_tab()
        self._build_timesheets_tab()
        self._build_attendance_tab()
        self._build_reports_tab()
        self._build_settings_tab()
        self._build_admin_tab()
        self._build_logs_tab()

        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)
        self._update_nav_highlight(self.tab_dashboard)

        status_bar = ttk.Label(self, textvariable=self.status_var, anchor=tk.W, style="Status.TLabel")
        status_bar.pack(fill=tk.X, padx=10, pady=(0, 8))
        self.status_var.set("Hazir")
        self.bind_all("<Control-f>", self._focus_global_search)
        self.bind_all("<Control-n>", self._quick_new_record)
        self.bind_all("<Control-s>", self._quick_save_record)

    def _on_tab_changed(self, _event):
        current = self.notebook.nametowidget(self.notebook.select())
        self._load_tab_data(current)
        if current is self.tab_timesheets:
            self.refresh_shift_templates()
        self._update_nav_highlight(current)

    def _switch_tab(self, tab):
        try:
            self.notebook.select(tab)
        except Exception:
            pass

    def _update_nav_highlight(self, current_tab):
        if not hasattr(self, "nav_buttons"):
            return
        theme = self.themes.get(self.current_theme, self.themes["Gece"])
        active_bg = theme["bg_hover"]
        idle_bg = theme["bg_content"]
        for key, btn in self.nav_buttons.items():
            tab = {
                "dashboard": self.tab_dashboard,
                "timesheets": self.tab_timesheets,
                "attendance": self.tab_attendance,
                "employees": self.tab_employees,
                "reports": self.tab_reports,
                "admin": self.tab_admin,
                "settings": self.tab_settings,
                "logs": self.tab_logs,
            }.get(key)
            if tab is current_tab:
                btn.configure(bg=active_bg)
            else:
                btn.configure(bg=idle_bg)

    def _on_global_search(self):
        text = self.global_search_var.get().strip()
        current = self.notebook.nametowidget(self.notebook.select())
        if current is self.tab_employees and hasattr(self, "emp_search_var"):
            self.emp_search_var.set(text)
            self.refresh_employees()
        elif current is self.tab_timesheets and hasattr(self, "ts_filter_search_var"):
            self.ts_filter_search_var.set(text)
            self.refresh_timesheets()
        elif current is self.tab_attendance:
            if hasattr(self, "att_filter_search_var"):
                self.att_filter_search_var.set(text)
                self.refresh_attendance()
            if hasattr(self, "leave_filter_search_var"):
                self.leave_filter_search_var.set(text)
                self.refresh_leave_records()
        elif current is self.tab_admin:
            self.admin_search_var.set(text)
            self.refresh_admin_summary()

    def _animate_stat(self, var, target, decimals=0, duration_ms=300):
        try:
            current = float(str(var.get()).replace(",", "."))
        except Exception:
            current = 0.0
        try:
            target_val = float(str(target).replace(",", "."))
        except Exception:
            target_val = 0.0
        steps = 10
        if duration_ms <= 0:
            steps = 1
        delta = (target_val - current) / steps if steps else 0
        delay = max(20, duration_ms // steps) if steps else duration_ms

        def step(i=0, value=current):
            if i >= steps:
                value = target_val
            else:
                value = value + delta
            if decimals == 0:
                var.set(str(int(round(value))))
            else:
                var.set(f"{value:.{decimals}f}")
            if i < steps:
                self.after(delay, step, i + 1, value)

        step()

    def _focus_global_search(self, _event=None):
        if hasattr(self, "global_search_entry"):
            self.global_search_entry.focus_set()
            self.global_search_entry.select_range(0, tk.END)

    def _quick_new_record(self, _event=None):
        current = self.notebook.nametowidget(self.notebook.select())
        if current is self.tab_timesheets:
            self.clear_timesheet_form()
        elif current is self.tab_employees:
            self.clear_employee_form()
        elif current is self.tab_attendance:
            self.clear_attendance_form()
            self.clear_leave_form()

    def _quick_save_record(self, _event=None):
        current = self.notebook.nametowidget(self.notebook.select())
        if current is self.tab_timesheets:
            self.add_or_update_timesheet()
        elif current is self.tab_employees:
            self.add_or_update_employee()
        elif current is self.tab_attendance:
            self.add_or_update_attendance()
            if (
                getattr(self, "leave_employee_var", None)
                and (self.leave_employee_var.get().strip() or self.leave_start_var.get().strip())
            ):
                self.add_or_update_leave()

    def _load_tab_data(self, tab):
        if self._tab_loaded.get(tab):
            return
        if tab is self.tab_dashboard:
            self.refresh_dashboard()
        elif tab is self.tab_employees:
            self.refresh_employees()
        elif tab is self.tab_timesheets:
            self.refresh_employees()
            self.refresh_timesheets()
            self.refresh_shift_templates()
        elif tab is self.tab_attendance:
            self.refresh_employees()
            self.refresh_attendance()
            self.refresh_leave_records()
        elif tab is self.tab_reports:
            self.refresh_report_archive()
        elif tab is self.tab_settings:
            self.refresh_shift_templates()
        elif tab is self.tab_admin:
            self.refresh_admin_summary()
        self._tab_loaded[tab] = True

    def _make_tab_scrollable(self, tab):
        canvas = tk.Canvas(tab, highlightthickness=0)
        vscroll = ttk.Scrollbar(tab, orient=tk.VERTICAL, command=canvas.yview)
        hscroll = ttk.Scrollbar(tab, orient=tk.HORIZONTAL, command=canvas.xview)
        canvas.configure(yscrollcommand=vscroll.set, xscrollcommand=hscroll.set)

        tab.rowconfigure(0, weight=1)
        tab.columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")
        hscroll.grid(row=1, column=0, sticky="ew")

        content = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind(
            "<Configure>",
            lambda e: canvas.itemconfigure(window_id, width=e.width),
        )
        content.bind("<Enter>", lambda _e: self._bind_canvas_mousewheel(canvas))
        content.bind("<Leave>", lambda _e: self._unbind_canvas_mousewheel(canvas))
        return content

    def _make_window_scrollable(self, window):
        canvas = tk.Canvas(window, highlightthickness=0)
        vscroll = ttk.Scrollbar(window, orient=tk.VERTICAL, command=canvas.yview)
        canvas.configure(yscrollcommand=vscroll.set)

        window.rowconfigure(0, weight=1)
        window.columnconfigure(0, weight=1)
        canvas.grid(row=0, column=0, sticky="nsew")
        vscroll.grid(row=0, column=1, sticky="ns")

        content = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=content, anchor="nw")
        content.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(window_id, width=e.width))
        content.bind("<Enter>", lambda _e: self._bind_canvas_mousewheel(canvas))
        content.bind("<Leave>", lambda _e: self._unbind_canvas_mousewheel(canvas))
        return content

    def _bind_canvas_mousewheel(self, canvas):
        canvas.bind_all("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

    def _unbind_canvas_mousewheel(self, canvas):
        canvas.unbind_all("<MouseWheel>")

    def _drain_log_queue(self):
        if not hasattr(self, "log_text"):
            return
        drained = False
        while True:
            try:
                msg = self.log_queue.get_nowait()
            except queue.Empty:
                break
            drained = True
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.insert(tk.END, msg + "\n")
            self.log_text.configure(state=tk.DISABLED)
        if drained:
            total_lines = int(self.log_text.index("end-1c").split(".")[0])
            if total_lines > 2000:
                self.log_text.configure(state=tk.NORMAL)
                self.log_text.delete("1.0", f"{total_lines - 2000}.0")
                self.log_text.configure(state=tk.DISABLED)
            self.log_text.see(tk.END)
        self.after(500, self._drain_log_queue)

    def clear_log_view(self):
        if hasattr(self, "log_text"):
            self.log_text.configure(state=tk.NORMAL)
            self.log_text.delete("1.0", tk.END)
            self.log_text.insert(tk.END, "Log temizlendi.\n")
            self.log_text.configure(state=tk.DISABLED)

    def trigger_sync(self, reason="manual", force=False):
        # Bulut senkron kaldirildi - artik islem yapmiyoruz.
        return

    def _sync_worker(self, sync_url, token, reason):
        """Senkronizasyon worker; upload + download + merge logic (19 Ocak)."""
        msg = None
        try:
            # Step 1: Upload local DB to server
            # Get region from settings or use default
            settings = db.get_all_settings()
            user_region = settings.get("user_region", "Ankara")
            current_region = user_region or "ALL"
            
            with open(db.DB_PATH, "rb") as handle:
                files = {"db": ("puantaj.db", handle, "application/octet-stream")}
                headers = {
                    "X-API-KEY": token,
                    "X-Region": current_region,
                    "X-Reason": reason
                }
                url = sync_url.rstrip("/") + "/sync"
                resp = requests.post(url, headers=headers, files=files, timeout=10)

            # Basarili upload doğrulama
            if resp.status_code != 200:
                msg = f"Senkron hatasi: Upload HTTP {resp.status_code}"
                if self.logger:
                    self.logger.warning("Cloud sync upload error: %s", msg)
                self.after(0, lambda: self._notify_sync_result(msg, reason))
                return
            
            # Log DEBUG info from server if available
            try:
                data = resp.json()
                if "debug_logs" in data and self.logger:
                    self.logger.info("--- SERVER SYNC DEBUG LOGS ---")
                    for log_line in data["debug_logs"]:
                        self.logger.info("SERVER: %s", log_line)
                    self.logger.info("------------------------------")
            except Exception as e:
                if self.logger:
                    self.logger.warning("Could not parse server debug logs: %s", str(e))

            # Step 2: Download merged DB from server
            headers = {"X-API-KEY": token}
            download_url = sync_url.rstrip("/") + "/sync/download"
            
            resp = requests.get(download_url, headers=headers, timeout=10)
            
            if resp.status_code != 200:
                msg = f"Senkron hatasi: Download HTTP {resp.status_code}"
                if self.logger:
                    self.logger.warning("Cloud sync download error: %s", msg)
                self.after(0, lambda: self._notify_sync_result(msg, reason))
                return

            # Step 3: Backup current local database
            import shutil
            backup_path = db.DB_PATH + ".sync_backup"
            if os.path.isfile(db.DB_PATH):
                shutil.copy2(db.DB_PATH, backup_path)

            # Step 4: Write downloaded database as new local DB
            with open(db.DB_PATH, "wb") as f:
                f.write(resp.content)

            # Step 4.5: Ensure DB schema is up to date (creates deleted_records table if missing)
            db.init_db()

            # Step 5: Refresh UI to reflect merged data
            self.after(0, self._refresh_all_after_sync)

            msg = "Senkron basarili"
            if self.logger:
                self.logger.info("Cloud sync completed (upload+download+merge): %s", reason)
            
        except requests.Timeout:
            msg = "Senkron hatasi: Baglanti timeout"
            if self.logger:
                self.logger.warning("Cloud sync timeout")
        except requests.RequestException as e:
            msg = f"Senkron hatasi: {str(e)[:80]}"
            if self.logger:
                self.logger.warning("Cloud sync request error: %s", str(e))
        except Exception as e:
            msg = f"Senkron hatasi: {str(e)[:80]}"
            if self.logger:
                self.logger.error("Cloud sync unexpected error: %s", str(e))

        if msg:
            self.after(0, lambda: self._notify_sync_result(msg, reason))

    def manual_sync(self):
        if not self.sync_enabled_var.get():
            messagebox.showwarning("Uyari", "Senkron kapali. Ayarlardan acin.")
            return
        self.trigger_sync("manual", force=True)

    def _notify_sync_result(self, message, reason):
        self.status_var.set(message)
        if reason == "manual":
            messagebox.showinfo("Senkron", message)

    def _refresh_all_after_sync(self):
        """Refresh all UI views after sync download to reflect merged data."""
        try:
            self.settings = db.get_all_settings()
            self.refresh_employees()
            self.refresh_timesheets()
            if hasattr(self, 'refresh_report_archive'):
                self.refresh_report_archive()
        except Exception as e:
            if self.logger:
                self.logger.warning("UI refresh after sync failed: %s", str(e))

    # Employees tab
    def _build_employees_tab(self):
        form = ttk.LabelFrame(self.tab_employees_body, text="Calisan Bilgisi", style="Section.TLabelframe")
        form.pack(fill=tk.X, padx=6, pady=6)

        self.emp_id_var = tk.StringVar()
        self.emp_name_var = tk.StringVar()
        self.emp_identity_var = tk.StringVar()
        self.emp_department_var = tk.StringVar()
        self.emp_title_var = tk.StringVar()

        row1 = ttk.Frame(form)
        row1.pack(fill=tk.X, pady=4)
        create_labeled_entry(row1, "Ad Soyad", self.emp_name_var, 30).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row1, "TCKN", self.emp_identity_var, 18).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row1, "Departman", self.emp_department_var, 18).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row1, "Unvan", self.emp_title_var, 18).pack(side=tk.LEFT, padx=6)

        btn_row = ttk.Frame(form)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="Kaydet", style="Accent.TButton", command=self.add_or_update_employee).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btn_row, text="Sil", command=self.delete_employee).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="Temizle", command=self.clear_employee_form).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Excel/CSV Iceri Aktar", command=self.import_employees).pack(side=tk.LEFT, padx=6)

        filter_frame = ttk.LabelFrame(self.tab_employees_body, text="Filtre", style="Section.TLabelframe")
        filter_frame.pack(fill=tk.X, padx=6, pady=6)
        self.emp_search_var = tk.StringVar()
        ttk.Label(filter_frame, text="Ara").pack(side=tk.LEFT, padx=(0, 6))
        emp_search_entry = ttk.Entry(filter_frame, textvariable=self.emp_search_var, width=26)
        emp_search_entry.pack(side=tk.LEFT)
        btn_emp_filter = ttk.Button(filter_frame, text="Guncelle", style="Accent.TButton", command=self.refresh_employees)
        btn_emp_filter.pack(side=tk.LEFT, padx=6)
        btn_emp_clear = ttk.Button(filter_frame, text="Temizle", command=self._clear_employee_search)
        btn_emp_clear.pack(side=tk.LEFT)
        attach_tooltip(emp_search_entry, "Ad, TCKN, departman, unvan veya bolge")
        attach_tooltip(btn_emp_filter, "Listeyi yenile")
        attach_tooltip(btn_emp_clear, "Aramayi temizle")

        emp_stats_row = ttk.Frame(self.tab_employees_body)
        emp_stats_row.pack(fill=tk.X, padx=6, pady=6)
        self.emp_stats = {
            "total": tk.StringVar(value="0"),
            "departments": tk.StringVar(value="0"),
            "regions": tk.StringVar(value="0"),
        }
        create_kpi_card(emp_stats_row, "Toplam Calisan", self.emp_stats["total"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)
        create_kpi_card(emp_stats_row, "Departman", self.emp_stats["departments"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)
        create_kpi_card(emp_stats_row, "Bolge", self.emp_stats["regions"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)

        pane = ttk.PanedWindow(self.tab_employees_body, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        list_frame = ttk.Frame(pane)
        detail_frame = ttk.LabelFrame(pane, text="Detay", style="Section.TLabelframe")
        pane.add(list_frame, weight=3)
        pane.add(detail_frame, weight=2)

        columns = ("id", "name", "identity", "department", "title", "region")
        self.employee_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.employee_tree.heading("id", text="ID")
        self.employee_tree.heading("name", text="Ad Soyad")
        self.employee_tree.heading("identity", text="TCKN")
        self.employee_tree.heading("department", text="Departman")
        self.employee_tree.heading("title", text="Unvan")
        self.employee_tree.heading("region", text="Bolge")
        self.employee_tree.column("id", width=60, anchor=tk.CENTER)
        self.employee_tree.column("name", width=220)
        self.employee_tree.column("identity", width=140)
        self.employee_tree.column("department", width=160)
        self.employee_tree.column("title", width=160)
        self.employee_tree.column("region", width=110)
        # Koyu tema zebra satırları
        self.employee_tree.tag_configure("odd", background="#252525", foreground="#E0E0E0")
        self.employee_tree.tag_configure("even", background="#1F1F1F", foreground="#E0E0E0")
        self.employee_tree.tag_configure("empty", background="#1F1F1F", foreground="#808080")
        emp_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.employee_tree.xview)
        emp_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.employee_tree.yview)
        self.employee_tree.configure(xscrollcommand=emp_xscroll.set, yscrollcommand=emp_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.employee_tree.grid(row=0, column=0, sticky="nsew")
        emp_yscroll.grid(row=0, column=1, sticky="ns")
        emp_xscroll.grid(row=1, column=0, sticky="ew")
        self.employee_tree.bind("<<TreeviewSelect>>", self.on_employee_select)
        self._apply_tree_zebra(self.employee_tree)

        # Detail panel
        self.emp_detail_name = tk.StringVar(value="-")
        self.emp_detail_dept = tk.StringVar(value="-")
        self.emp_detail_title = tk.StringVar(value="-")
        self.emp_detail_region = tk.StringVar(value="-")
        self.emp_detail_identity = tk.StringVar(value="-")

        drow1 = ttk.Frame(detail_frame)
        drow1.pack(fill=tk.X, pady=4)
        ttk.Label(drow1, text="Ad Soyad").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(drow1, textvariable=self.emp_detail_name).pack(side=tk.LEFT)

        drow2 = ttk.Frame(detail_frame)
        drow2.pack(fill=tk.X, pady=4)
        ttk.Label(drow2, text="Departman").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(drow2, textvariable=self.emp_detail_dept).pack(side=tk.LEFT)

        drow3 = ttk.Frame(detail_frame)
        drow3.pack(fill=tk.X, pady=4)
        ttk.Label(drow3, text="Unvan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(drow3, textvariable=self.emp_detail_title).pack(side=tk.LEFT)

        drow4 = ttk.Frame(detail_frame)
        drow4.pack(fill=tk.X, pady=4)
        ttk.Label(drow4, text="Bolge").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(drow4, textvariable=self.emp_detail_region).pack(side=tk.LEFT)

        drow5 = ttk.Frame(detail_frame)
        drow5.pack(fill=tk.X, pady=4)
        ttk.Label(drow5, text="TCKN").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(drow5, textvariable=self.emp_detail_identity).pack(side=tk.LEFT)

        recent_frame = ttk.LabelFrame(detail_frame, text="Bu Ay Mesai", style="Section.TLabelframe")
        recent_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.emp_recent_tree = ttk.Treeview(
            recent_frame,
            columns=("date", "worked", "overtime"),
            show="headings",
            height=8,
        )
        self.emp_recent_tree.heading("date", text="Tarih")
        self.emp_recent_tree.heading("worked", text="Calisilan")
        self.emp_recent_tree.heading("overtime", text="Fazla")
        self.emp_recent_tree.column("date", width=100)
        self.emp_recent_tree.column("worked", width=90)
        self.emp_recent_tree.column("overtime", width=90)
        self.emp_recent_tree.pack(fill=tk.BOTH, expand=True)
        self._apply_tree_zebra(self.emp_recent_tree)


    def refresh_employees(self):
        for item in self.employee_tree.get_children():
            self.employee_tree.delete(item)
        self.employee_map = {}
        self.employee_display_names = []
        self.employee_details = {}
        name_counts = {}
        search = ""
        if hasattr(self, "emp_search_var"):
            search = self.emp_search_var.get().strip().lower()
        employees = db.list_employees(region=self._view_region())
        if search:
            filtered = []
            for emp in employees:
                emp_id, name, identity_no, department, title, region = emp
                hay = " ".join(
                    [
                        str(name or ""),
                        str(identity_no or ""),
                        str(department or ""),
                        str(title or ""),
                        str(region or ""),
                    ]
                ).lower()
                if search in hay:
                    filtered.append(emp)
            employees = filtered
        for emp in employees:
            emp_id, name, identity_no, department, title, region = emp
            name_counts[name] = name_counts.get(name, 0) + 1
            self.employee_map[(name, region or "")] = emp_id
            tag = "odd" if len(self.employee_tree.get_children()) % 2 else "even"
            self.employee_tree.insert("", tk.END, values=emp, tags=(tag,))
            self.employee_details[(name, region or "")] = {
                "department": department or "",
                "title": title or "",
                "identity_no": identity_no or "",
                "region": region or "",
            }
        if not employees:
            insert_empty_row(self.employee_tree, ("id", "name", "identity", "department", "title", "region"), "Kayit yok")
        self.employee_display_names = []
        for emp in db.list_employees(region=self._view_region()):
            _emp_id, name, _identity_no, _department, _title, region = emp
            display = name
            if name_counts.get(name, 0) > 1:
                display = f"{name} ({region or '-'})"
            self.employee_display_names.append(display)
        self._refresh_employee_comboboxes()
        if hasattr(self, "emp_stats"):
            departments = {e[3] for e in employees if e[3]}
            regions = {e[5] for e in employees if e[5]}
            self._animate_stat(self.emp_stats["total"], len(employees), decimals=0)
            self._animate_stat(self.emp_stats["departments"], len(departments), decimals=0)
            self._animate_stat(self.emp_stats["regions"], len(regions), decimals=0)
        if not self._auto_select_tree_first(self.employee_tree, self.on_employee_select):
            if hasattr(self, "emp_detail_name"):
                self.emp_detail_name.set("-")
                self.emp_detail_identity.set("-")
                self.emp_detail_dept.set("-")
                self.emp_detail_title.set("-")
                self.emp_detail_region.set("-")
            if hasattr(self, "emp_recent_tree"):
                for item in self.emp_recent_tree.get_children():
                    self.emp_recent_tree.delete(item)

    def _clear_employee_search(self):
        if hasattr(self, "emp_search_var"):
            self.emp_search_var.set("")
        self.refresh_employees()

    def _refresh_employee_comboboxes(self):
        values = ["Tum Calisanlar"] + sorted(self.employee_display_names)
        if hasattr(self, "ts_employee_combo"):
            self.ts_employee_combo["values"] = values
        if hasattr(self, "ts_filter_combo"):
            self.ts_filter_combo["values"] = values
        if hasattr(self, "att_employee_combo"):
            self.att_employee_combo["values"] = sorted(self.employee_display_names)
        if hasattr(self, "att_filter_combo"):
            self.att_filter_combo["values"] = values
        if hasattr(self, "leave_employee_combo"):
            self.leave_employee_combo["values"] = sorted(self.employee_display_names)
        if hasattr(self, "leave_filter_combo"):
            self.leave_filter_combo["values"] = values
        if hasattr(self, "report_employee_combo"):
            self.report_employee_combo["values"] = values
        if hasattr(self, "admin_employee_combo"):
            self.admin_employee_combo["values"] = values
        if hasattr(self, "admin_department_combo"):
            departments = sorted({d["department"] for d in self.employee_details.values() if d["department"]})
            self.admin_department_combo["values"] = ["Tum Departmanlar"] + departments
        if hasattr(self, "admin_title_combo"):
            titles = sorted({d["title"] for d in self.employee_details.values() if d["title"]})
            self.admin_title_combo["values"] = ["Tum Unvanlar"] + titles

    def clear_employee_form(self):
        self.emp_id_var.set("")
        self.emp_name_var.set("")
        self.emp_identity_var.set("")
        self.emp_department_var.set("")
        self.emp_title_var.set("")

    def on_employee_select(self, _event=None):
        selected = self.employee_tree.selection()
        if not selected:
            return
        values = self.employee_tree.item(selected[0], "values")
        self.emp_id_var.set(values[0])
        self.emp_name_var.set(values[1])
        self.emp_identity_var.set(values[2])
        self.emp_department_var.set(values[3])
        self.emp_title_var.set(values[4])
        if hasattr(self, "emp_detail_name"):
            self.emp_detail_name.set(values[1])
            self.emp_detail_identity.set(values[2] or "-")
            self.emp_detail_dept.set(values[3] or "-")
            self.emp_detail_title.set(values[4] or "-")
            self.emp_detail_region.set(values[5] or "-")
            self._refresh_employee_recent_timesheets(parse_int(values[0]))

    def add_or_update_employee(self):
        name = self.emp_name_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Ad Soyad zorunlu.")
            return
        identity_no = self.emp_identity_var.get().strip()
        department = self.emp_department_var.get().strip()
        title = self.emp_title_var.get().strip()

        emp_id = self.emp_id_var.get().strip()
        if emp_id:
            db.update_employee(
                parse_int(emp_id),
                name,
                identity_no,
                department,
                title,
                self._entry_region(),
            )
            self._log_action("employee_update", f"id={emp_id} name={name}")
        else:
            db.add_employee(name, identity_no, department, title, self._entry_region())
            self._log_action("employee_add", f"name={name}")
        self.refresh_employees()
        self.clear_employee_form()
        self.trigger_sync("employee")

    def _refresh_employee_recent_timesheets(self, employee_id):
        if not hasattr(self, "emp_recent_tree"):
            return
        for item in self.emp_recent_tree.get_children():
            self.emp_recent_tree.delete(item)
        today = datetime.now().date()
        month_start = today.replace(day=1).strftime("%Y-%m-%d")
        month_end = today.strftime("%Y-%m-%d")
        records = db.list_timesheets(
            employee_id=employee_id,
            start_date=month_start,
            end_date=month_end,
            region=self._view_region(),
        )
        records = sorted(records, key=lambda r: r[4], reverse=True)
        for ts in records:
            (
                _ts_id,
                _emp_id,
                _name,
                department,
                work_date,
                start_time,
                end_time,
                break_minutes,
                is_special,
                _notes,
                _region,
            ) = ts
            try:
                worked, _scheduled, overtime, _night, _overnight, _s1, _s2, _s3 = calc.calc_day_hours(
                    work_date, start_time, end_time, break_minutes, self.settings, is_special, department
                )
            except Exception:
                worked = overtime = ""
            self.emp_recent_tree.insert("", tk.END, values=(work_date, worked, overtime))
        if not records:
            insert_empty_row(self.emp_recent_tree, ("date", "worked", "overtime"), "Kayit yok")

    def delete_employee(self):
        emp_id = self.emp_id_var.get().strip()
        if not emp_id:
            messagebox.showwarning("Uyari", "Silmek icin calisan secin.")
            return
        if messagebox.askyesno("Onay", "Calisani silmek istiyor musunuz?"):
            db.delete_employee(parse_int(emp_id))
            self._log_action("employee_delete", f"id={emp_id}")
            self.refresh_employees()
            self.clear_employee_form()
            self.trigger_sync("employee_delete")

    def notify(self, message, sound=True, duration_ms=2500):
        self.status_var.set(message)
        if sound and winsound:
            try:
                winsound.MessageBeep(winsound.MB_ICONASTERISK)
            except Exception:
                pass
        self.after(duration_ms, lambda: self.status_var.set(""))

    # Timesheets tab
    def _build_timesheets_tab(self):
        form = ttk.LabelFrame(self.tab_timesheets_body, text="Puantaj Girisi", style="Section.TLabelframe")
        form.pack(fill=tk.X, padx=6, pady=6)

        self.ts_id_var = tk.StringVar()
        self.ts_employee_var = tk.StringVar()
        self.ts_date_var = tk.StringVar()
        self.ts_start_var = tk.StringVar(value="09:00")
        self.ts_end_var = tk.StringVar(value="18:00")
        self.ts_break_var = tk.StringVar(value="60")
        self.ts_notes_var = tk.StringVar()
        self.ts_template_var = tk.StringVar()
        self.ts_special_var = tk.IntVar(value=0)

        row1 = ttk.Frame(form)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.ts_employee_combo = ttk.Combobox(row1, textvariable=self.ts_employee_var, width=28, state="readonly")
        self.ts_employee_combo.pack(side=tk.LEFT)
        date_frame, self.ts_date_entry = create_labeled_date(row1, "Tarih", self.ts_date_var, 12)
        date_frame.pack(side=tk.LEFT, padx=6)
        start_frame = create_time_entry(row1, "Giris", self.ts_start_var, 8)
        start_frame.pack(side=tk.LEFT, padx=6)
        end_frame = create_time_entry(row1, "Cikis", self.ts_end_var, 8)
        end_frame.pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row1, "Mola dk", self.ts_break_var, 8).pack(side=tk.LEFT, padx=6)
        for child in start_frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.bind("<FocusOut>", lambda _e: normalize_time_in_var(self.ts_start_var))
        for child in end_frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.bind("<FocusOut>", lambda _e: normalize_time_in_var(self.ts_end_var))

        row2 = ttk.Frame(form)
        row2.pack(fill=tk.X, pady=4)
        create_labeled_entry(row2, "Not", self.ts_notes_var, 60).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(row2, text="Ozel Gun", variable=self.ts_special_var).pack(side=tk.LEFT, padx=6)

        row3 = ttk.Frame(form)
        row3.pack(fill=tk.X, pady=4)
        ttk.Label(row3, text="Sablon").pack(side=tk.LEFT, padx=(0, 8))
        self.ts_template_combo = ttk.Combobox(row3, textvariable=self.ts_template_var, width=28, state="readonly")
        self.ts_template_combo.pack(side=tk.LEFT)
        ttk.Button(row3, text="Uygula", command=self.apply_shift_template).pack(side=tk.LEFT, padx=6)

        btn_row = ttk.Frame(form)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="Kaydet", style="Accent.TButton", command=self.add_or_update_timesheet).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(btn_row, text="Sil", command=self.delete_timesheet).pack(side=tk.LEFT)
        ttk.Button(btn_row, text="Temizle", command=self.clear_timesheet_form).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Excel/CSV Iceri Aktar", command=self.import_timesheets).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Panodan Yapistir", command=self.paste_timesheets).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Toplu Sablon Uygula", command=self.apply_template_to_selected).pack(
            side=tk.LEFT, padx=6
        )

        filter_frame = ttk.LabelFrame(self.tab_timesheets_body, text="Filtre", style="Section.TLabelframe")
        filter_frame.pack(fill=tk.X, padx=6, pady=6)
        self.ts_filter_employee = tk.StringVar(value="Tum Calisanlar")
        self.ts_filter_start = tk.StringVar()
        self.ts_filter_end = tk.StringVar()
        self.ts_filter_limit_var = tk.StringVar(value="500")

        ttk.Label(filter_frame, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.ts_filter_combo = ttk.Combobox(filter_frame, textvariable=self.ts_filter_employee, width=28, state="readonly")
        self.ts_filter_combo.pack(side=tk.LEFT)
        start_frame, self.ts_filter_start_entry = create_labeled_date(
            filter_frame, "Baslangic", self.ts_filter_start, 12
        )
        start_frame.pack(side=tk.LEFT, padx=6)
        end_frame, self.ts_filter_end_entry = create_labeled_date(filter_frame, "Bitis", self.ts_filter_end, 12)
        end_frame.pack(side=tk.LEFT, padx=6)
        self.ts_filter_search_var = tk.StringVar()
        ttk.Label(filter_frame, text="Ara").pack(side=tk.LEFT, padx=(12, 6))
        ts_search_entry = ttk.Entry(filter_frame, textvariable=self.ts_filter_search_var, width=18)
        ts_search_entry.pack(side=tk.LEFT)
        ttk.Label(filter_frame, text="Limit").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Entry(filter_frame, textvariable=self.ts_filter_limit_var, width=6).pack(side=tk.LEFT)
        btn_filter = ttk.Button(filter_frame, text="Filtrele", style="Accent.TButton", command=self.refresh_timesheets)
        btn_filter.pack(side=tk.LEFT, padx=6)
        attach_tooltip(ts_search_entry, "Calisan, not veya tarih")
        attach_tooltip(btn_filter, "Secilen tarih araligini uygula")
        btn_clear = ttk.Button(filter_frame, text="Temizle", command=self.clear_timesheet_filter)
        btn_clear.pack(side=tk.LEFT)
        attach_tooltip(btn_clear, "Filtreleri sifirla")
        clear_date_entry(self.ts_filter_start_entry)
        clear_date_entry(self.ts_filter_end_entry)

        stats_row = ttk.Frame(self.tab_timesheets_body)
        stats_row.pack(fill=tk.X, padx=6, pady=6)
        self.ts_stats = {
            "records": tk.StringVar(value="0"),
            "worked": tk.StringVar(value="0"),
            "overtime": tk.StringVar(value="0"),
            "night": tk.StringVar(value="0"),
            "sunday_days": tk.StringVar(value="0"),
            "sunday_hours": tk.StringVar(value="0"),
        }
        create_kpi_card(stats_row, "Kayit", self.ts_stats["records"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)
        create_kpi_card(stats_row, "Toplam Calisilan", self.ts_stats["worked"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(stats_row, "Toplam Fazla Mesai", self.ts_stats["overtime"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(stats_row, "Toplam Gece", self.ts_stats["night"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(stats_row, "Pazar Gun", self.ts_stats["sunday_days"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(stats_row, "Pazar Saat", self.ts_stats["sunday_hours"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )

        pane = ttk.PanedWindow(self.tab_timesheets_body, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.timesheet_pane = pane

        list_frame = ttk.Frame(pane)
        detail_frame = ttk.LabelFrame(pane, text="Detay", style="Section.TLabelframe")
        pane.add(list_frame, weight=4)
        pane.add(detail_frame, weight=2)

        columns = (
            "id",
            "employee",
            "date",
            "start",
            "end",
            "break",
            "worked",
            "scheduled",
            "overtime",
            "night",
            "overnight",
            "special",
            "special_normal",
            "special_overtime",
            "special_night",
            "notes",
            "region",
            "sunday_hours",
        )
        self.timesheet_tree = ttk.Treeview(list_frame, columns=columns, show="headings", selectmode="extended")
        self.timesheet_tree.heading("id", text="ID")
        self.timesheet_tree.heading("employee", text="Calisan")
        self.timesheet_tree.heading("date", text="Tarih")
        self.timesheet_tree.heading("start", text="Giris")
        self.timesheet_tree.heading("end", text="Cikis")
        self.timesheet_tree.heading("break", text="Mola")
        self.timesheet_tree.heading("worked", text="Calisilan")
        self.timesheet_tree.heading("scheduled", text="Plan")
        self.timesheet_tree.heading("overtime", text="Fazla Mesai")
        self.timesheet_tree.heading("night", text="Gece")
        self.timesheet_tree.heading("overnight", text="Geceye Tasan")
        self.timesheet_tree.heading("special", text="Ozel Gun")
        self.timesheet_tree.heading("special_normal", text="Ozel Gun Normal")
        self.timesheet_tree.heading("special_overtime", text="Ozel Gun Fazla")
        self.timesheet_tree.heading("special_night", text="Ozel Gun Gece")
        self.timesheet_tree.heading("notes", text="Not")
        self.timesheet_tree.heading("region", text="Bolge")
        self.timesheet_tree.heading("sunday_hours", text="Pazar Saat")
        self.timesheet_tree.column("id", width=60, anchor=tk.CENTER)
        self.timesheet_tree.column("employee", width=220)
        self.timesheet_tree.column("date", width=100)
        self.timesheet_tree.column("start", width=80)
        self.timesheet_tree.column("end", width=80)
        self.timesheet_tree.column("break", width=80)
        self.timesheet_tree.column("worked", width=90)
        self.timesheet_tree.column("scheduled", width=90)
        self.timesheet_tree.column("overtime", width=90)
        self.timesheet_tree.column("night", width=90)
        self.timesheet_tree.column("overnight", width=100)
        self.timesheet_tree.column("special", width=80)
        self.timesheet_tree.column("special_normal", width=110)
        self.timesheet_tree.column("special_overtime", width=110)
        self.timesheet_tree.column("special_night", width=110)
        self.timesheet_tree.column("notes", width=180)
        self.timesheet_tree.column("region", width=100)
        self.timesheet_tree.column("sunday_hours", width=100)
        # Koyu tema zebra satırları
        self.timesheet_tree.tag_configure("odd", background="#252525", foreground="#E0E0E0")
        self.timesheet_tree.tag_configure("even", background="#1F1F1F", foreground="#E0E0E0")
        self.timesheet_tree.tag_configure("empty", background="#1F1F1F", foreground="#808080")
        ts_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.timesheet_tree.xview)
        ts_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.timesheet_tree.yview)
        self.timesheet_tree.configure(xscrollcommand=ts_xscroll.set, yscrollcommand=ts_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.timesheet_tree.grid(row=0, column=0, sticky="nsew")
        ts_yscroll.grid(row=0, column=1, sticky="ns")
        ts_xscroll.grid(row=1, column=0, sticky="ew")
        self.timesheet_tree.bind("<<TreeviewSelect>>", self.on_timesheet_select)
        self._apply_tree_zebra(self.timesheet_tree)
        self.timesheet_tree.bind("<Button-3>", self.on_timesheet_right_click)

        # Detail panel
        self.ts_detail_employee = tk.StringVar(value="-")
        self.ts_detail_date = tk.StringVar(value="-")
        self.ts_detail_hours = tk.StringVar(value="-")
        self.ts_detail_overtime = tk.StringVar(value="-")
        self.ts_detail_night = tk.StringVar(value="-")
        self.ts_detail_notes = tk.StringVar(value="-")

        d1 = ttk.Frame(detail_frame)
        d1.pack(fill=tk.X, pady=4)
        ttk.Label(d1, text="Calisan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d1, textvariable=self.ts_detail_employee).pack(side=tk.LEFT)

        d2 = ttk.Frame(detail_frame)
        d2.pack(fill=tk.X, pady=4)
        ttk.Label(d2, text="Tarih").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d2, textvariable=self.ts_detail_date).pack(side=tk.LEFT)

        d3 = ttk.Frame(detail_frame)
        d3.pack(fill=tk.X, pady=4)
        ttk.Label(d3, text="Calisilan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d3, textvariable=self.ts_detail_hours).pack(side=tk.LEFT)

        d4 = ttk.Frame(detail_frame)
        d4.pack(fill=tk.X, pady=4)
        ttk.Label(d4, text="Fazla Mesai").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d4, textvariable=self.ts_detail_overtime).pack(side=tk.LEFT)

        d5 = ttk.Frame(detail_frame)
        d5.pack(fill=tk.X, pady=4)
        ttk.Label(d5, text="Gece").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d5, textvariable=self.ts_detail_night).pack(side=tk.LEFT)

        d6 = ttk.Frame(detail_frame)
        d6.pack(fill=tk.X, pady=4)
        ttk.Label(d6, text="Not").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d6, textvariable=self.ts_detail_notes, wraplength=220).pack(side=tk.LEFT)

        quick_frame = ttk.LabelFrame(detail_frame, text="Hizli Duzenle", style="Section.TLabelframe")
        quick_frame.pack(fill=tk.X, padx=6, pady=6)
        self.ts_quick_start_var = tk.StringVar()
        self.ts_quick_end_var = tk.StringVar()
        self.ts_quick_break_var = tk.StringVar()
        qrow = ttk.Frame(quick_frame)
        qrow.pack(fill=tk.X, pady=4)
        create_labeled_entry(qrow, "Giris", self.ts_quick_start_var, 8).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(qrow, "Cikis", self.ts_quick_end_var, 8).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(qrow, "Mola", self.ts_quick_break_var, 6).pack(side=tk.LEFT, padx=6)
        ttk.Button(qrow, text="Uygula", style="Accent.TButton", command=self.quick_update_timesheet).pack(
            side=tk.LEFT, padx=6
        )

        self.ts_menu = tk.Menu(self, tearoff=0)

        self.ts_menu.add_command(label="Duzenle", command=self.edit_selected_timesheet)
        self.ts_menu.add_command(label="Sil", command=self.delete_timesheet)
        self.after(120, lambda: self._set_pane_sash(self.timesheet_pane, 0.68))

    def clear_timesheet_form(self):
        self.ts_id_var.set("")
        self.ts_employee_var.set("")
        self.ts_date_var.set("")
        self.ts_start_var.set("09:00")
        self.ts_end_var.set("18:00")
        self.ts_break_var.set("60")
        self.ts_notes_var.set("")
        self.ts_template_var.set("")
        self.ts_special_var.set(0)
        self.ts_original = None
        self.ts_editing_id = None

    def clear_timesheet_filter(self):
        self.ts_filter_employee.set("Tum Calisanlar")
        self.ts_filter_start.set("")
        self.ts_filter_end.set("")
        if hasattr(self, "ts_filter_search_var"):
            self.ts_filter_search_var.set("")
        clear_date_entry(self.ts_filter_start_entry)
        clear_date_entry(self.ts_filter_end_entry)
        self.refresh_timesheets()

    def on_timesheet_select(self, _event=None):
        selected = self.timesheet_tree.selection()
        if not selected:
            return
        values = self.timesheet_tree.item(selected[0], "values")
        if not values:
            return
        if hasattr(self, "ts_detail_employee"):
            self.ts_detail_employee.set(values[1])
            self.ts_detail_date.set(values[2])
            self.ts_detail_hours.set(values[6])
            self.ts_detail_overtime.set(values[8])
            self.ts_detail_night.set(values[9])
            self.ts_detail_notes.set(values[15] or "-")
        if hasattr(self, "ts_quick_start_var"):
            self.ts_quick_start_var.set(values[3])
            self.ts_quick_end_var.set(values[4])
            self.ts_quick_break_var.set(values[5])

    def quick_update_timesheet(self):
        selected = self.timesheet_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Duzenlemek icin kayit secin.")
            return
        values = self.timesheet_tree.item(selected[0], "values")
        ts_id = parse_int(values[0])
        if not ts_id:
            return
        start_time = self.ts_quick_start_var.get().strip() or values[3]
        end_time = self.ts_quick_end_var.get().strip() or values[4]
        break_minutes = parse_int(self.ts_quick_break_var.get(), parse_int(values[5], 0))
        try:
            start_time = normalize_time(start_time)
            end_time = normalize_time(end_time)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        region = values[16] if len(values) > 16 else ""
        employee_id = self.employee_map.get((values[1], region)) or self.employee_map.get((values[1], ""))
        if not employee_id:
            messagebox.showwarning("Uyari", "Calisan bulunamadi.")
            return
        db.update_timesheet(
            ts_id,
            employee_id,
            values[2],
            start_time,
            end_time,
            break_minutes,
            1 if values[11] == "Evet" else 0,
            values[15],
            region or self._entry_region(),
        )
        self.refresh_timesheets()
        self.notify("Puantaj guncellendi.")

    def apply_template_to_selected(self):
        selected = self.timesheet_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Toplu islem icin kayit secin.")
            return
        tpl_name = self.ts_template_var.get().strip()
        if not tpl_name or tpl_name not in getattr(self, "shift_template_map", {}):
            messagebox.showwarning("Uyari", "Sablon secin.")
            return
        _tpl_id, _name, start_time, end_time, break_minutes = self.shift_template_map[tpl_name]
        updated = 0
        for item in selected:
            values = self.timesheet_tree.item(item, "values")
            ts_id = parse_int(values[0])
            if not ts_id:
                continue
            region = values[16] if len(values) > 16 else ""
            employee_id = self.employee_map.get((values[1], region)) or self.employee_map.get((values[1], ""))
            if not employee_id:
                continue
            db.update_timesheet(
                ts_id,
                employee_id,
                values[2],
                start_time,
                end_time,
                break_minutes,
                1 if values[11] == "Evet" else 0,
                values[15],
                region or self._entry_region(),
            )
            updated += 1
        self.refresh_timesheets()
        self.notify(f"{updated} kayit sablonla guncellendi.")

    def paste_timesheets(self):
        try:
            raw = self.clipboard_get()
        except Exception:
            messagebox.showwarning("Uyari", "Pano bos.")
            return
        if not raw:
            return
        lines = [line.strip() for line in raw.splitlines() if line.strip()]
        if not lines:
            return
        imported = 0
        for line in lines:
            parts = [p.strip() for p in line.replace(";", "\t").split("\t") if p.strip() != ""]
            if len(parts) < 4:
                parts = [p.strip() for p in line.split(",")]
            if len(parts) < 4:
                continue
            name = parts[0]
            work_date = parts[1]
            start_time = parts[2]
            end_time = parts[3]
            break_minutes = parse_int(parts[4], 60) if len(parts) > 4 else 60
            notes = parts[5] if len(parts) > 5 else ""
            base, region = split_display_name(name, REGIONS)
            employee_id = self.employee_map.get((base, region or "")) or self.employee_map.get(
                (base, self._entry_region())
            )
            if not employee_id:
                continue
            try:
                work_date = normalize_date(work_date)
                start_time = normalize_time(start_time)
                end_time = normalize_time(end_time)
            except ValueError:
                continue
            db.add_timesheet(
                employee_id,
                work_date,
                start_time,
                end_time,
                break_minutes,
                0,
                notes,
                self._entry_region(),
            )
            imported += 1
        self.refresh_timesheets()
        self.notify(f"Pano girisi tamamlandi: {imported} kayit.")

    def on_timesheet_right_click(self, event):
        row_id = self.timesheet_tree.identify_row(event.y)
        if row_id:
            self.timesheet_tree.selection_set(row_id)
            self.ts_menu.tk_popup(event.x_root, event.y_root)

    def edit_selected_timesheet(self):
        selected = self.timesheet_tree.selection()
        if not selected:
            return
        values = self.timesheet_tree.item(selected[0], "values")
        self.ts_editing_id = values[0]
        self.ts_id_var.set(values[0])
        region = values[16] if len(values) > 16 else ""
        display_name = values[1]
        if region:
            display_name = f"{values[1]} ({region})"
        self.ts_employee_var.set(display_name)
        self.ts_date_var.set(values[2])
        set_time_vars(values[3], self.ts_start_var)
        set_time_vars(values[4], self.ts_end_var)
        self.ts_break_var.set(values[5])
        self.ts_special_var.set(1 if values[11] == "Evet" else 0)
        self.ts_notes_var.set(values[15])
        self.ts_original = (values[1], values[2], values[3], values[4])

    def refresh_timesheets(self):
        for item in self.timesheet_tree.get_children():
            self.timesheet_tree.delete(item)

        employee_name = self.ts_filter_employee.get()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))
        start_date = self.ts_filter_start.get().strip() or None
        end_date = self.ts_filter_end.get().strip() or None
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records = db.list_timesheets(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            region=self._view_region(),
        )
        search = ""
        if hasattr(self, "ts_filter_search_var"):
            search = self.ts_filter_search_var.get().strip().lower()
        total_worked = total_overtime = total_night = 0.0
        total_sunday_hours = 0.0
        sunday_day_keys = set()
        shown = 0
        limit = parse_int(self.ts_filter_limit_var.get(), 0) if hasattr(self, "ts_filter_limit_var") else 0
        for ts in records:
            (
                ts_id,
                _emp_id,
                name,
                department,
                work_date,
                start_time,
                end_time,
                break_minutes,
                is_special,
                notes,
                region,
            ) = ts
            if search:
                hay = " ".join([str(name), str(work_date), str(notes or "")]).lower()
                if search not in hay:
                    continue
            try:
                (
                    worked,
                    scheduled,
                    overtime,
                    night_hours,
                    overnight_hours,
                    spec_norm,
                    spec_ot,
                    spec_night,
                ) = calc.calc_day_hours(
                    work_date,
                    start_time,
                    end_time,
                    break_minutes,
                    self.settings,
                    is_special,
                    department,
                )
                sunday_hours = calc_sunday_separate_hours(work_date, department, is_special, worked)
            except Exception:
                worked = scheduled = overtime = ""
                night_hours = overnight_hours = ""
                spec_norm = spec_ot = spec_night = ""
                sunday_hours = ""
            tag = "odd" if len(self.timesheet_tree.get_children()) % 2 else "even"
            self.timesheet_tree.insert(
                "",
                tk.END,
                values=(
                    ts_id,
                    name,
                    work_date,
                    start_time,
                    end_time,
                    break_minutes,
                    worked,
                    scheduled,
                    overtime,
                    night_hours,
                    overnight_hours,
                    "Evet" if is_special else "Hayir",
                    spec_norm,
                    spec_ot,
                    spec_night,
                    notes or "",
                    region or "",
                    sunday_hours,
                ),
                tags=(tag,),
            )
            if worked != "":
                total_worked += float(worked)
            if overtime != "":
                total_overtime += float(overtime)
            if night_hours != "":
                total_night += float(night_hours)
            if sunday_hours != "":
                sunday_float = float(sunday_hours)
                if sunday_float > 0:
                    total_sunday_hours += sunday_float
                    sunday_day_keys.add((name, work_date, region or ""))
            shown += 1
            if limit and shown >= limit:
                break
        if not self.timesheet_tree.get_children():
            insert_empty_row(
                self.timesheet_tree,
                (
                    "id",
                    "employee",
                    "date",
                    "start",
                    "end",
                    "break",
                    "worked",
                    "scheduled",
                    "overtime",
                    "night",
                    "overnight",
                    "special",
                    "special_normal",
                    "special_overtime",
                    "special_night",
                    "notes",
                    "region",
                    "sunday_hours",
                ),
                "Kayit yok",
            )
        if hasattr(self, "ts_stats"):
            self._animate_stat(self.ts_stats["records"], shown, decimals=0)
            self._animate_stat(self.ts_stats["worked"], total_worked, decimals=2)
            self._animate_stat(self.ts_stats["overtime"], total_overtime, decimals=2)
            self._animate_stat(self.ts_stats["night"], total_night, decimals=2)
            self._animate_stat(self.ts_stats["sunday_days"], len(sunday_day_keys), decimals=0)
            self._animate_stat(self.ts_stats["sunday_hours"], total_sunday_hours, decimals=2)

        if hasattr(self, "ts_filter_combo"):
            self.ts_filter_combo["values"] = ["Tum Calisanlar"] + sorted(self.employee_display_names)
        if not self._auto_select_tree_first(self.timesheet_tree, self.on_timesheet_select):
            if hasattr(self, "ts_detail_employee"):
                self.ts_detail_employee.set("-")
                self.ts_detail_date.set("-")
                self.ts_detail_hours.set("-")
                self.ts_detail_overtime.set("-")
                self.ts_detail_night.set("-")
                self.ts_detail_notes.set("-")

    def add_or_update_timesheet(self):
        name = self.ts_employee_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Calisan secin.")
            return
        employee_id = None
        base, region = split_display_name(name, REGIONS)
        if region is None:
            employee_id = self.employee_map.get((base, "")) or self.employee_map.get((base, self._entry_region()))
        else:
            employee_id = self.employee_map.get((base, region))
        if not employee_id:
            messagebox.showwarning("Uyari", "Calisan bulunamadi.")
            return
        work_date = self.ts_date_var.get().strip()
        if not work_date:
            messagebox.showwarning("Uyari", "Tarih zorunlu.")
            return
        try:
            work_date = normalize_date(work_date)
            start_time = normalize_time(self.ts_start_var.get())
            end_time = normalize_time(self.ts_end_var.get())
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        break_minutes = parse_int(self.ts_break_var.get(), 0)
        # Mola dakikasi validasyonu (0-480 dakika = 0-8 saat)
        if not (0 <= break_minutes <= 480):
            messagebox.showwarning("Uyari", "Mola dakikasi 0-480 arasinda olmalidir.")
            return
        notes = self.ts_notes_var.get().strip()
        is_special = 1 if self.ts_special_var.get() else 0
        # Bolge NULL kontrolu
        entry_region = self._entry_region()
        if not entry_region:
            messagebox.showwarning("Uyari", "Bolge tanimlanimis. Ayarlardan kontrol edin.")
            return

        ts_id = self.ts_editing_id
        if ts_id:
            db.update_timesheet(
                parse_int(ts_id),
                employee_id,
                work_date,
                start_time,
                end_time,
                break_minutes,
                is_special,
                notes,
                self._entry_region(),
            )
            self._log_action("timesheet_update", f"id={ts_id} date={work_date}")
        else:
            db.add_timesheet(
                employee_id,
                work_date,
                start_time,
                end_time,
                break_minutes,
                is_special,
                notes,
                self._entry_region(),
            )
            self._log_action("timesheet_add", f"employee_id={employee_id} date={work_date}")
        self.refresh_timesheets()
        self.clear_timesheet_form()
        self.trigger_sync("timesheet")
        self.notify("Puantaj kaydedildi.")

    def delete_timesheet(self):
        ts_id = self.ts_editing_id
        if not ts_id:
            selected = self.timesheet_tree.selection()
            if selected:
                values = self.timesheet_tree.item(selected[0], "values")
                ts_id = values[0]
        if not ts_id:
            messagebox.showwarning("Uyari", "Silmek icin puantaj secin.")
            return
        if messagebox.askyesno("Onay", "Puantaj kaydini silmek istiyor musunuz?"):
            db.delete_timesheet(parse_int(ts_id))
            self._log_action("timesheet_delete", f"id={ts_id}")
            self.refresh_timesheets()
            self.clear_timesheet_form()
            self.notify("Puantaj silindi.")
            self.trigger_sync("timesheet_delete")

    def refresh_shift_templates(self):
        templates = db.list_shift_templates()
        self.shift_template_map = {tpl[1]: tpl for tpl in templates}
        if hasattr(self, "ts_template_combo"):
            self.ts_template_combo["values"] = sorted(self.shift_template_map.keys())
            if not self.ts_template_var.get() and templates:
                self.ts_template_var.set(templates[0][1])
                self.apply_shift_template()
        if hasattr(self, "template_tree"):
            for item in self.template_tree.get_children():
                self.template_tree.delete(item)
            for tpl in templates:
                tag = "odd" if len(self.template_tree.get_children()) % 2 else "even"
                self.template_tree.insert("", tk.END, values=tpl, tags=(tag,))

    def apply_shift_template(self):
        name = self.ts_template_var.get().strip()
        if not name or name not in self.shift_template_map:
            messagebox.showwarning("Uyari", "Sablon secin.")
            return
        _tpl_id, _name, start_time, end_time, break_minutes = self.shift_template_map[name]
        set_time_vars(start_time, self.ts_start_var)
        set_time_vars(end_time, self.ts_end_var)
        self.ts_break_var.set(str(break_minutes))

    def save_shift_template(self):
        name = self.st_name_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Sablon adi zorunlu.")
            return
        try:
            start_time = normalize_time(self.st_start_var.get())
            end_time = normalize_time(self.st_end_var.get())
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        break_minutes = parse_int(self.st_break_var.get(), 0)
        db.upsert_shift_template(name, start_time, end_time, break_minutes)
        self._log_action("shift_template_save", f"name={name}")
        self.refresh_shift_templates()
        self.clear_shift_template_form()

    def delete_shift_template(self):
        tpl_id = self.st_id_var.get().strip()
        if not tpl_id:
            messagebox.showwarning("Uyari", "Silmek icin sablon secin.")
            return
        if messagebox.askyesno("Onay", "Sablonu silmek istiyor musunuz?"):
            db.delete_shift_template(parse_int(tpl_id))
            self._log_action("shift_template_delete", f"id={tpl_id}")
            self.refresh_shift_templates()
            self.clear_shift_template_form()

    def clear_shift_template_form(self):
        self.st_id_var.set("")
        self.st_name_var.set("")
        self.st_start_var.set("09:00")
        self.st_end_var.set("18:00")
        self.st_break_var.set("60")

    def on_template_select(self, _event=None):
        selected = self.template_tree.selection()
        if not selected:
            return
        values = self.template_tree.item(selected[0], "values")
        self.st_id_var.set(values[0])
        self.st_name_var.set(values[1])
        set_time_vars(values[2], self.st_start_var)
        set_time_vars(values[3], self.st_end_var)
        self.st_break_var.set(values[4])

    # Attendance & Leave tab
    def _build_attendance_tab(self):
        content = self.tab_attendance_body

        attendance_frame = ttk.LabelFrame(content, text="Yoklama Kaydi", style="Section.TLabelframe")
        attendance_frame.pack(fill=tk.X, padx=6, pady=6)

        self.att_id_var = tk.StringVar()
        self.att_employee_var = tk.StringVar()
        self.att_date_var = tk.StringVar()
        self.att_status_var = tk.StringVar(value=ATTENDANCE_STATUSES[0])
        self.att_reason_var = tk.StringVar()

        arow1 = ttk.Frame(attendance_frame)
        arow1.pack(fill=tk.X, pady=4)
        ttk.Label(arow1, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.att_employee_combo = ttk.Combobox(arow1, textvariable=self.att_employee_var, width=28, state="readonly")
        self.att_employee_combo.pack(side=tk.LEFT)
        date_frame, self.att_date_entry = create_labeled_date(arow1, "Tarih", self.att_date_var, 12)
        date_frame.pack(side=tk.LEFT, padx=6)
        ttk.Label(arow1, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        self.att_status_combo = ttk.Combobox(
            arow1, textvariable=self.att_status_var, values=ATTENDANCE_STATUSES, width=14, state="readonly"
        )
        self.att_status_combo.pack(side=tk.LEFT)

        arow2 = ttk.Frame(attendance_frame)
        arow2.pack(fill=tk.X, pady=4)
        create_labeled_entry(arow2, "Neden", self.att_reason_var, 60).pack(side=tk.LEFT, padx=6)

        abtn = ttk.Frame(attendance_frame)
        abtn.pack(fill=tk.X, pady=6)
        btn_att_save = ttk.Button(abtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_attendance)
        btn_att_save.pack(side=tk.LEFT, padx=6)
        attach_tooltip(btn_att_save, "Yoklama kaydini kaydet")
        btn_att_delete = ttk.Button(abtn, text="Sil", command=self.delete_attendance)
        btn_att_delete.pack(side=tk.LEFT)
        attach_tooltip(btn_att_delete, "Secili yoklama kaydini sil")
        btn_att_clear = ttk.Button(abtn, text="Temizle", command=self.clear_attendance_form)
        btn_att_clear.pack(side=tk.LEFT, padx=6)
        attach_tooltip(btn_att_clear, "Formu sifirla")

        att_filter = ttk.LabelFrame(content, text="Yoklama Filtre", style="Section.TLabelframe")
        att_filter.pack(fill=tk.X, padx=6, pady=6)
        self.att_filter_employee_var = tk.StringVar(value="Tum Calisanlar")
        self.att_filter_status_var = tk.StringVar(value="Tum Durumlar")
        self.att_filter_start_var = tk.StringVar()
        self.att_filter_end_var = tk.StringVar()

        ttk.Label(att_filter, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.att_filter_combo = ttk.Combobox(
            att_filter, textvariable=self.att_filter_employee_var, width=24, state="readonly"
        )
        self.att_filter_combo.pack(side=tk.LEFT)
        ttk.Label(att_filter, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        self.att_filter_status_combo = ttk.Combobox(
            att_filter,
            textvariable=self.att_filter_status_var,
            values=["Tum Durumlar"] + ATTENDANCE_STATUSES,
            width=14,
            state="readonly",
        )
        self.att_filter_status_combo.pack(side=tk.LEFT)
        start_frame, self.att_filter_start_entry = create_labeled_date(
            att_filter, "Baslangic", self.att_filter_start_var, 12
        )
        start_frame.pack(side=tk.LEFT, padx=6)
        end_frame, self.att_filter_end_entry = create_labeled_date(att_filter, "Bitis", self.att_filter_end_var, 12)
        end_frame.pack(side=tk.LEFT, padx=6)
        self.att_filter_search_var = tk.StringVar()
        ttk.Label(att_filter, text="Ara").pack(side=tk.LEFT, padx=(12, 6))
        att_search_entry = ttk.Entry(att_filter, textvariable=self.att_filter_search_var, width=16)
        att_search_entry.pack(side=tk.LEFT)
        btn_att_refresh = ttk.Button(att_filter, text="Guncelle", style="Accent.TButton", command=self.refresh_attendance)
        btn_att_refresh.pack(side=tk.LEFT, padx=6)
        attach_tooltip(att_search_entry, "Calisan, neden veya tarih")
        attach_tooltip(btn_att_refresh, "Yoklama listesini yenile")
        btn_att_clear = ttk.Button(att_filter, text="Temizle", command=self.clear_attendance_filter)
        btn_att_clear.pack(side=tk.LEFT)
        attach_tooltip(btn_att_clear, "Filtreleri sifirla")
        ttk.Button(att_filter, text="Toplu Durum", command=self.bulk_update_attendance_status).pack(
            side=tk.LEFT, padx=6
        )
        clear_date_entry(self.att_filter_start_entry)
        clear_date_entry(self.att_filter_end_entry)

        att_stats_row = ttk.Frame(content)
        att_stats_row.pack(fill=tk.X, padx=6, pady=6)
        self.att_stats = {
            "total": tk.StringVar(value="0"),
            "worked": tk.StringVar(value="0"),
            "leave": tk.StringVar(value="0"),
            "absent": tk.StringVar(value="0"),
        }
        create_kpi_card(att_stats_row, "Toplam", self.att_stats["total"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(att_stats_row, "Calisti", self.att_stats["worked"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(att_stats_row, "Izinli", self.att_stats["leave"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(att_stats_row, "Gelmedi", self.att_stats["absent"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )

        pane = ttk.PanedWindow(content, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        list_frame = ttk.Frame(pane)
        detail_frame = ttk.LabelFrame(pane, text="Detay", style="Section.TLabelframe")
        pane.add(list_frame, weight=4)
        pane.add(detail_frame, weight=2)
        columns = ("id", "date", "employee", "status", "reason", "source", "region")
        self.attendance_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.attendance_tree.heading("id", text="ID")
        self.attendance_tree.heading("date", text="Tarih")
        self.attendance_tree.heading("employee", text="Calisan")
        self.attendance_tree.heading("status", text="Durum")
        self.attendance_tree.heading("reason", text="Neden")
        self.attendance_tree.heading("source", text="Kaynak")
        self.attendance_tree.heading("region", text="Bolge")
        self.attendance_tree.column("id", width=60, anchor=tk.CENTER)
        self.attendance_tree.column("date", width=100)
        self.attendance_tree.column("employee", width=220)
        self.attendance_tree.column("status", width=120)
        self.attendance_tree.column("reason", width=240)
        self.attendance_tree.column("source", width=110)
        self.attendance_tree.column("region", width=90)
        self.attendance_tree.tag_configure("odd", background="#252525", foreground="#E0E0E0")
        self.attendance_tree.tag_configure("even", background="#1F1F1F", foreground="#E0E0E0")
        self.attendance_tree.tag_configure("empty", background="#1F1F1F", foreground="#808080")
        att_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.attendance_tree.xview)
        att_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.attendance_tree.yview)
        self.attendance_tree.configure(xscrollcommand=att_xscroll.set, yscrollcommand=att_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.attendance_tree.grid(row=0, column=0, sticky="nsew")
        att_yscroll.grid(row=0, column=1, sticky="ns")
        att_xscroll.grid(row=1, column=0, sticky="ew")
        self.attendance_tree.bind("<<TreeviewSelect>>", self.on_attendance_select)
        self._apply_tree_zebra(self.attendance_tree)

        self.att_detail_employee = tk.StringVar(value="-")
        self.att_detail_date = tk.StringVar(value="-")
        self.att_detail_status = tk.StringVar(value="-")
        self.att_detail_reason = tk.StringVar(value="-")
        self.att_detail_source = tk.StringVar(value="-")

        d1 = ttk.Frame(detail_frame); d1.pack(fill=tk.X, pady=4)
        ttk.Label(d1, text="Calisan").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(d1, textvariable=self.att_detail_employee).pack(side=tk.LEFT)
        d2 = ttk.Frame(detail_frame); d2.pack(fill=tk.X, pady=4)
        ttk.Label(d2, text="Tarih").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(d2, textvariable=self.att_detail_date).pack(side=tk.LEFT)
        d3 = ttk.Frame(detail_frame); d3.pack(fill=tk.X, pady=4)
        ttk.Label(d3, text="Durum").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(d3, textvariable=self.att_detail_status).pack(side=tk.LEFT)
        d4 = ttk.Frame(detail_frame); d4.pack(fill=tk.X, pady=4)
        ttk.Label(d4, text="Neden").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(d4, textvariable=self.att_detail_reason, wraplength=220).pack(side=tk.LEFT)
        d5 = ttk.Frame(detail_frame); d5.pack(fill=tk.X, pady=4)
        ttk.Label(d5, text="Kaynak").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(d5, textvariable=self.att_detail_source).pack(side=tk.LEFT)

        self._apply_tree_zebra(self.attendance_tree)

        leave_frame = ttk.LabelFrame(content, text="Izin Kaydi", style="Section.TLabelframe")
        leave_frame.pack(fill=tk.X, padx=6, pady=6)

        self.leave_id_var = tk.StringVar()
        self.leave_employee_var = tk.StringVar()
        self.leave_start_var = tk.StringVar()
        self.leave_end_var = tk.StringVar()
        self.leave_type_var = tk.StringVar(value=LEAVE_TYPES[0])
        self.leave_status_var = tk.StringVar(value="Onayli")
        self.leave_reason_var = tk.StringVar()
        self.leave_doc_no_var = tk.StringVar()
        self.leave_apply_attendance_var = tk.BooleanVar(value=True)
        self.leave_document_path = ""

        lrow1 = ttk.Frame(leave_frame)
        lrow1.pack(fill=tk.X, pady=4)
        ttk.Label(lrow1, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.leave_employee_combo = ttk.Combobox(lrow1, textvariable=self.leave_employee_var, width=28, state="readonly")
        self.leave_employee_combo.pack(side=tk.LEFT)
        lstart_frame, self.leave_start_entry = create_labeled_date(lrow1, "Baslangic", self.leave_start_var, 12)
        lstart_frame.pack(side=tk.LEFT, padx=6)
        lend_frame, self.leave_end_entry = create_labeled_date(lrow1, "Bitis", self.leave_end_var, 12)
        lend_frame.pack(side=tk.LEFT, padx=6)
        ttk.Label(lrow1, text="Tip").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Combobox(lrow1, textvariable=self.leave_type_var, values=LEAVE_TYPES, width=14, state="readonly").pack(
            side=tk.LEFT
        )
        ttk.Label(lrow1, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        status_combo = ttk.Combobox(
            lrow1, textvariable=self.leave_status_var, values=LEAVE_STATUSES, width=12, state="readonly"
        )
        status_combo.pack(side=tk.LEFT)
        status_combo.configure(state="disabled")

        lrow2 = ttk.Frame(leave_frame)
        lrow2.pack(fill=tk.X, pady=4)
        create_labeled_entry(lrow2, "Neden", self.leave_reason_var, 50).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(lrow2, "Belge No", self.leave_doc_no_var, 16).pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(lrow2, text="Yoklamaya isle", variable=self.leave_apply_attendance_var).pack(
            side=tk.LEFT, padx=8
        )

        lbtn = ttk.Frame(leave_frame)
        lbtn.pack(fill=tk.X, pady=6)
        btn_leave_save = ttk.Button(lbtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_leave)
        btn_leave_save.pack(
            side=tk.LEFT, padx=6
        )
        attach_tooltip(btn_leave_save, "Izin kaydini kaydet")
        btn_leave_delete = ttk.Button(lbtn, text="Sil", command=self.delete_leave)
        btn_leave_delete.pack(side=tk.LEFT)
        attach_tooltip(btn_leave_delete, "Secili izin kaydini sil")
        btn_leave_clear = ttk.Button(lbtn, text="Temizle", command=self.clear_leave_form)
        btn_leave_clear.pack(side=tk.LEFT, padx=6)
        attach_tooltip(btn_leave_clear, "Formu sifirla")
        btn_leave_form = ttk.Button(lbtn, text="Izin Formu Olustur", command=self.generate_leave_form)
        btn_leave_form.pack(side=tk.LEFT, padx=6)
        attach_tooltip(btn_leave_form, "Izin formunu Excel olarak olustur")
        btn_leave_open = ttk.Button(lbtn, text="Belgeyi Ac", command=self.open_leave_document)
        btn_leave_open.pack(side=tk.LEFT, padx=6)
        attach_tooltip(btn_leave_open, "Kayitli izin belgesini ac")

        leave_filter = ttk.LabelFrame(content, text="Izin Filtre", style="Section.TLabelframe")
        leave_filter.pack(fill=tk.X, padx=6, pady=6)
        self.leave_filter_employee_var = tk.StringVar(value="Tum Calisanlar")
        self.leave_filter_status_var = tk.StringVar(value="Tum Durumlar")
        self.leave_filter_start_var = tk.StringVar()
        self.leave_filter_end_var = tk.StringVar()

        ttk.Label(leave_filter, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.leave_filter_combo = ttk.Combobox(
            leave_filter, textvariable=self.leave_filter_employee_var, width=24, state="readonly"
        )
        self.leave_filter_combo.pack(side=tk.LEFT)
        ttk.Label(leave_filter, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Combobox(
            leave_filter,
            textvariable=self.leave_filter_status_var,
            values=["Tum Durumlar"] + LEAVE_STATUSES,
            width=14,
            state="readonly",
        ).pack(side=tk.LEFT)
        lstart_filter, self.leave_filter_start_entry = create_labeled_date(
            leave_filter, "Baslangic", self.leave_filter_start_var, 12
        )
        lstart_filter.pack(side=tk.LEFT, padx=6)
        lend_filter, self.leave_filter_end_entry = create_labeled_date(
            leave_filter, "Bitis", self.leave_filter_end_var, 12
        )
        lend_filter.pack(side=tk.LEFT, padx=6)
        self.leave_filter_search_var = tk.StringVar()
        ttk.Label(leave_filter, text="Ara").pack(side=tk.LEFT, padx=(12, 6))
        leave_search_entry = ttk.Entry(leave_filter, textvariable=self.leave_filter_search_var, width=16)
        leave_search_entry.pack(side=tk.LEFT)
        btn_leave_refresh = ttk.Button(
            leave_filter, text="Guncelle", style="Accent.TButton", command=self.refresh_leave_records
        )
        btn_leave_refresh.pack(side=tk.LEFT, padx=6)
        attach_tooltip(leave_search_entry, "Calisan, neden veya tarih")
        attach_tooltip(btn_leave_refresh, "Izin listesini yenile")
        btn_leave_clear = ttk.Button(leave_filter, text="Temizle", command=self.clear_leave_filter)
        btn_leave_clear.pack(side=tk.LEFT)
        attach_tooltip(btn_leave_clear, "Filtreleri sifirla")
        clear_date_entry(self.leave_filter_start_entry)
        clear_date_entry(self.leave_filter_end_entry)

        leave_stats_row = ttk.Frame(content)
        leave_stats_row.pack(fill=tk.X, padx=6, pady=6)
        self.leave_stats = {
            "total_days": tk.StringVar(value="0"),
            "approved": tk.StringVar(value="0"),
            "active": tk.StringVar(value="0"),
        }
        create_kpi_card(leave_stats_row, "Toplam Izin (Gun)", self.leave_stats["total_days"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(leave_stats_row, "Onayli Kayit", self.leave_stats["approved"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(leave_stats_row, "Aktif Izinli", self.leave_stats["active"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )

        leave_pane = ttk.PanedWindow(content, orient=tk.HORIZONTAL)
        leave_pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        leave_list_frame = ttk.Frame(leave_pane)
        leave_detail_frame = ttk.LabelFrame(leave_pane, text="Detay", style="Section.TLabelframe")
        leave_pane.add(leave_list_frame, weight=4)
        leave_pane.add(leave_detail_frame, weight=2)
        leave_columns = ("id", "employee", "start", "end", "days", "type", "status", "reason", "doc_no", "region")
        self.leave_tree = ttk.Treeview(leave_list_frame, columns=leave_columns, show="headings")
        self.leave_tree.heading("id", text="ID")
        self.leave_tree.heading("employee", text="Calisan")
        self.leave_tree.heading("start", text="Baslangic")
        self.leave_tree.heading("end", text="Bitis")
        self.leave_tree.heading("days", text="Gun")
        self.leave_tree.heading("type", text="Tip")
        self.leave_tree.heading("status", text="Durum")
        self.leave_tree.heading("reason", text="Neden")
        self.leave_tree.heading("doc_no", text="Belge No")
        self.leave_tree.heading("region", text="Bolge")
        self.leave_tree.column("id", width=60, anchor=tk.CENTER)
        self.leave_tree.column("employee", width=220)
        self.leave_tree.column("start", width=100)
        self.leave_tree.column("end", width=100)
        self.leave_tree.column("days", width=70, anchor=tk.CENTER)
        self.leave_tree.column("type", width=120)
        self.leave_tree.column("status", width=100)
        self.leave_tree.column("reason", width=220)
        self.leave_tree.column("doc_no", width=110)
        self.leave_tree.column("region", width=90)
        self.leave_tree.tag_configure("odd", background="#252525", foreground="#E0E0E0")
        self.leave_tree.tag_configure("even", background="#1F1F1F", foreground="#E0E0E0")
        self.leave_tree.tag_configure("empty", background="#1F1F1F", foreground="#808080")
        leave_xscroll = ttk.Scrollbar(leave_list_frame, orient=tk.HORIZONTAL, command=self.leave_tree.xview)
        leave_yscroll = ttk.Scrollbar(leave_list_frame, orient=tk.VERTICAL, command=self.leave_tree.yview)
        self.leave_tree.configure(xscrollcommand=leave_xscroll.set, yscrollcommand=leave_yscroll.set)
        leave_list_frame.columnconfigure(0, weight=1)
        leave_list_frame.rowconfigure(0, weight=1)
        self.leave_tree.grid(row=0, column=0, sticky="nsew")
        leave_yscroll.grid(row=0, column=1, sticky="ns")
        leave_xscroll.grid(row=1, column=0, sticky="ew")
        self.leave_tree.bind("<<TreeviewSelect>>", self.on_leave_select)
        self._apply_tree_zebra(self.leave_tree)

        self.leave_detail_employee = tk.StringVar(value="-")
        self.leave_detail_dates = tk.StringVar(value="-")
        self.leave_detail_type = tk.StringVar(value="-")
        self.leave_detail_status = tk.StringVar(value="-")
        self.leave_detail_reason = tk.StringVar(value="-")
        self.leave_detail_doc = tk.StringVar(value="-")

        l1 = ttk.Frame(leave_detail_frame); l1.pack(fill=tk.X, pady=4)
        ttk.Label(l1, text="Calisan").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l1, textvariable=self.leave_detail_employee).pack(side=tk.LEFT)
        l2 = ttk.Frame(leave_detail_frame); l2.pack(fill=tk.X, pady=4)
        ttk.Label(l2, text="Tarih").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l2, textvariable=self.leave_detail_dates).pack(side=tk.LEFT)
        l3 = ttk.Frame(leave_detail_frame); l3.pack(fill=tk.X, pady=4)
        ttk.Label(l3, text="Tip").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l3, textvariable=self.leave_detail_type).pack(side=tk.LEFT)
        l4 = ttk.Frame(leave_detail_frame); l4.pack(fill=tk.X, pady=4)
        ttk.Label(l4, text="Durum").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l4, textvariable=self.leave_detail_status).pack(side=tk.LEFT)
        l5 = ttk.Frame(leave_detail_frame); l5.pack(fill=tk.X, pady=4)
        ttk.Label(l5, text="Belge").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l5, textvariable=self.leave_detail_doc).pack(side=tk.LEFT)
        l6 = ttk.Frame(leave_detail_frame); l6.pack(fill=tk.X, pady=4)
        ttk.Label(l6, text="Neden").pack(side=tk.LEFT, padx=(0,6))
        ttk.Label(l6, textvariable=self.leave_detail_reason, wraplength=220).pack(side=tk.LEFT)

        history_frame = ttk.LabelFrame(leave_detail_frame, text="Durum Gecmisi", style="Section.TLabelframe")
        history_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.leave_history_tree = ttk.Treeview(
            history_frame,
            columns=("status", "note", "user", "date"),
            show="headings",
            height=6,
        )
        self.leave_history_tree.heading("status", text="Durum")
        self.leave_history_tree.heading("note", text="Not")
        self.leave_history_tree.heading("user", text="Kisi")
        self.leave_history_tree.heading("date", text="Tarih")
        self.leave_history_tree.column("status", width=90)
        self.leave_history_tree.column("note", width=160)
        self.leave_history_tree.column("user", width=100)
        self.leave_history_tree.column("date", width=130)
        self.leave_history_tree.pack(fill=tk.BOTH, expand=True)
        self._apply_tree_zebra(self.leave_history_tree)

        self._apply_tree_zebra(self.leave_tree)

        totals_frame = ttk.LabelFrame(content, text="Izin Toplamlari", style="Section.TLabelframe")
        totals_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.leave_totals_tree = ttk.Treeview(totals_frame, columns=("employee", "days"), show="headings", height=6)
        self.leave_totals_tree.heading("employee", text="Calisan")
        self.leave_totals_tree.heading("days", text="Toplam Gun")
        self.leave_totals_tree.column("employee", width=220)
        self.leave_totals_tree.column("days", width=120, anchor=tk.CENTER)
        totals_xscroll = ttk.Scrollbar(totals_frame, orient=tk.HORIZONTAL, command=self.leave_totals_tree.xview)
        totals_yscroll = ttk.Scrollbar(totals_frame, orient=tk.VERTICAL, command=self.leave_totals_tree.yview)
        self.leave_totals_tree.configure(xscrollcommand=totals_xscroll.set, yscrollcommand=totals_yscroll.set)
        totals_frame.columnconfigure(0, weight=1)
        totals_frame.rowconfigure(0, weight=1)
        self.leave_totals_tree.grid(row=0, column=0, sticky="nsew")
        self._apply_tree_zebra(self.leave_totals_tree)
        totals_yscroll.grid(row=0, column=1, sticky="ns")
        totals_xscroll.grid(row=1, column=0, sticky="ew")

    def clear_attendance_form(self):
        self.att_id_var.set("")
        self.att_employee_var.set("")
        self.att_date_var.set("")
        self.att_status_var.set(ATTENDANCE_STATUSES[0])
        self.att_reason_var.set("")

    def clear_attendance_filter(self):
        self.att_filter_employee_var.set("Tum Calisanlar")
        self.att_filter_status_var.set("Tum Durumlar")
        self.att_filter_start_var.set("")
        self.att_filter_end_var.set("")
        if hasattr(self, "att_filter_search_var"):
            self.att_filter_search_var.set("")
        clear_date_entry(self.att_filter_start_entry)
        clear_date_entry(self.att_filter_end_entry)
        self.refresh_attendance()

    def on_attendance_select(self, _event=None):
        selected = self.attendance_tree.selection()
        if not selected:
            return
        values = self.attendance_tree.item(selected[0], "values")
        att_id = values[0]
        if att_id:
            self.att_id_var.set(att_id)
        else:
            self.att_id_var.set("")
        self.att_employee_var.set(values[2])
        self.att_date_var.set(values[1])
        self.att_status_var.set(values[3])
        self.att_reason_var.set(values[4])
        if hasattr(self, "att_detail_employee"):
            self.att_detail_employee.set(values[2])
            self.att_detail_date.set(values[1])
            self.att_detail_status.set(values[3])
            self.att_detail_reason.set(values[4] or "-")
            self.att_detail_source.set(values[5] or "-")

    def add_or_update_attendance(self):
        name = self.att_employee_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Calisan secin.")
            return
        base, region = split_display_name(name, REGIONS)
        if region is None:
            employee_id = self.employee_map.get((base, "")) or self.employee_map.get((base, self._entry_region()))
        else:
            employee_id = self.employee_map.get((base, region))
        if not employee_id:
            messagebox.showwarning("Uyari", "Calisan bulunamadi.")
            return
        work_date = self.att_date_var.get().strip()
        if not work_date:
            messagebox.showwarning("Uyari", "Tarih zorunlu.")
            return
        try:
            work_date = normalize_date(work_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        status = self.att_status_var.get().strip() or "Calisti"
        reason = self.att_reason_var.get().strip()
        db.upsert_attendance_record(employee_id, work_date, status, reason, self._entry_region(), "Yoklama")
        self._log_action("attendance_save", f"employee_id={employee_id} date={work_date} status={status}")
        self.refresh_attendance()
        self.clear_attendance_form()
        self.notify("Yoklama kaydedildi.")

    def delete_attendance(self):
        att_id = self.att_id_var.get().strip()
        if not att_id:
            selected = self.attendance_tree.selection()
            if selected:
                values = self.attendance_tree.item(selected[0], "values")
                att_id = values[0]
        if not att_id:
            messagebox.showwarning("Uyari", "Silmek icin yoklama secin.")
            return
        if messagebox.askyesno("Onay", "Yoklama kaydini silmek istiyor musunuz?"):
            db.delete_attendance_record(parse_int(att_id))
            self._log_action("attendance_delete", f"id={att_id}")
            self.refresh_attendance()
            self.clear_attendance_form()

    def bulk_update_attendance_status(self):
        selected = self.attendance_tree.selection()
        if not selected:
            messagebox.showinfo("Bilgi", "Toplu islem icin satir secin.")
            return

        dialog = tk.Toplevel(self)
        dialog.title("Toplu Durum")
        dialog.geometry("360x200")
        dialog.transient(self)
        dialog.grab_set()

        status_var = tk.StringVar(value=ATTENDANCE_STATUSES[0])
        reason_var = tk.StringVar()

        row1 = ttk.Frame(dialog)
        row1.pack(fill=tk.X, pady=8, padx=10)
        ttk.Label(row1, text="Durum").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Combobox(row1, textvariable=status_var, values=ATTENDANCE_STATUSES, state="readonly").pack(
            side=tk.LEFT
        )

        row2 = ttk.Frame(dialog)
        row2.pack(fill=tk.X, pady=8, padx=10)
        ttk.Label(row2, text="Neden").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(row2, textvariable=reason_var, width=30).pack(side=tk.LEFT)

        def apply_bulk():
            status = status_var.get().strip()
            reason = reason_var.get().strip()
            for item in selected:
                values = self.attendance_tree.item(item, "values")
                name = values[2]
                work_date = values[1]
                region = values[6] if len(values) > 6 else ""
                base, reg = split_display_name(name, REGIONS)
                region_key = reg if reg is not None else (region or self._entry_region())
                employee_id = self.employee_map.get((base, region_key)) or self.employee_map.get((base, ""))
                if not employee_id:
                    continue
                db.upsert_attendance_record(employee_id, work_date, status, reason, region_key, "Toplu")
            dialog.destroy()
            self.refresh_attendance()
            self.notify("Toplu yoklama guncellendi.")

        btn_row = ttk.Frame(dialog)
        btn_row.pack(fill=tk.X, pady=12, padx=10)
        ttk.Button(btn_row, text="Uygula", style="Accent.TButton", command=apply_bulk).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Iptal", command=dialog.destroy).pack(side=tk.LEFT)

    def refresh_attendance(self):
        if not hasattr(self, "attendance_tree"):
            return
        for item in self.attendance_tree.get_children():
            self.attendance_tree.delete(item)

        employee_name = self.att_filter_employee_var.get().strip()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))

        status_filter = self.att_filter_status_var.get().strip()
        start_date = self.att_filter_start_var.get().strip() or None
        end_date = self.att_filter_end_var.get().strip() or None
        search = ""
        if hasattr(self, "att_filter_search_var"):
            search = self.att_filter_search_var.get().strip().lower()
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        merged = {}
        ts_records = db.list_timesheets(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            region=self._view_region(),
        )
        for (
            _ts_id,
            emp_id,
            name,
            _department,
            work_date,
            _start_time,
            _end_time,
            _break_minutes,
            _is_special,
            _notes,
            region,
        ) in ts_records:
            key = (emp_id, work_date)
            merged[key] = {
                "id": "",
                "date": work_date,
                "employee": name,
                "status": "Calisti",
                "reason": "Puantaj",
                "source": "Puantaj",
                "region": region or "",
            }

        attendance_records = db.list_attendance_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            status=None if status_filter == "Tum Durumlar" else status_filter,
            region=self._view_region(),
        )
        for att in attendance_records:
            att_id, emp_id, name, work_date, status, reason, source, region = att
            key = (emp_id, work_date)
            merged[key] = {
                "id": att_id,
                "date": work_date,
                "employee": name,
                "status": status,
                "reason": reason or "",
                "source": source or "Yoklama",
                "region": region or "",
            }

        rows = list(merged.values())
        if status_filter and status_filter != "Tum Durumlar":
            rows = [row for row in rows if row["status"] == status_filter]
        if search:
            rows = [
                row
                for row in rows
                if search in " ".join([row["employee"], row["date"], row["reason"]]).lower()
            ]

        rows.sort(key=lambda r: r["date"], reverse=True)
        total = worked = leave = absent = 0
        for row in rows:
            tag = "odd" if len(self.attendance_tree.get_children()) % 2 else "even"
            self.attendance_tree.insert(
                "",
                tk.END,
                values=(
                    row["id"],
                    row["date"],
                    row["employee"],
                    row["status"],
                    row["reason"],
                    row["source"],
                    row["region"],
                ),
                tags=(tag,),
            )
            total += 1
            if row["status"] == "Calisti":
                worked += 1
            elif row["status"] == "Izinli":
                leave += 1
            elif row["status"] == "Gelmedi":
                absent += 1
        if not self.attendance_tree.get_children():
            insert_empty_row(
                self.attendance_tree,
                ("id", "date", "employee", "status", "reason", "source", "region"),
                "Kayit yok",
            )
        if hasattr(self, "att_stats"):
            self._animate_stat(self.att_stats["total"], total, decimals=0)
            self._animate_stat(self.att_stats["worked"], worked, decimals=0)
            self._animate_stat(self.att_stats["leave"], leave, decimals=0)
            self._animate_stat(self.att_stats["absent"], absent, decimals=0)
        if not self._auto_select_tree_first(self.attendance_tree, self.on_attendance_select):
            if hasattr(self, "att_detail_employee"):
                self.att_detail_employee.set("-")
                self.att_detail_date.set("-")
                self.att_detail_status.set("-")
                self.att_detail_reason.set("-")
                self.att_detail_source.set("-")

    def clear_leave_form(self):
        self.leave_id_var.set("")
        self.leave_employee_var.set("")
        self.leave_start_var.set("")
        self.leave_end_var.set("")
        self.leave_type_var.set(LEAVE_TYPES[0])
        self.leave_status_var.set(LEAVE_STATUSES[0])
        self.leave_reason_var.set("")
        self.leave_doc_no_var.set("")
        self.leave_apply_attendance_var.set(True)
        self.leave_document_path = ""

    def clear_leave_filter(self):
        self.leave_filter_employee_var.set("Tum Calisanlar")
        self.leave_filter_status_var.set("Tum Durumlar")
        self.leave_filter_start_var.set("")
        self.leave_filter_end_var.set("")
        if hasattr(self, "leave_filter_search_var"):
            self.leave_filter_search_var.set("")
        clear_date_entry(self.leave_filter_start_entry)
        clear_date_entry(self.leave_filter_end_entry)
        self.refresh_leave_records()

    def on_leave_select(self, _event=None):
        selected = self.leave_tree.selection()
        if not selected:
            return
        values = self.leave_tree.item(selected[0], "values")
        if not values or not str(values[0]).isdigit():
            return
        self.leave_id_var.set(values[0])
        self.leave_employee_var.set(values[1])
        self.leave_start_var.set(values[2])
        self.leave_end_var.set(values[3])
        self.leave_type_var.set(values[5])
        self.leave_status_var.set(values[6])
        self.leave_reason_var.set(values[7])
        self.leave_doc_no_var.set(values[8])
        if hasattr(self, "leave_detail_employee"):
            self.leave_detail_employee.set(values[1])
            self.leave_detail_dates.set(f"{values[2]} - {values[3]}")
            self.leave_detail_type.set(values[5])
            self.leave_detail_status.set(values[6])
            self.leave_detail_reason.set(values[7] or "-")
            self.leave_detail_doc.set(values[8] or "-")
        if hasattr(self, "leave_history_tree"):
            self._refresh_leave_history(values[0])

    def _persist_leave_record(self):
        name = self.leave_employee_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Calisan secin.")
            return None
        base, region = split_display_name(name, REGIONS)
        if region is None:
            employee_id = self.employee_map.get((base, "")) or self.employee_map.get((base, self._entry_region()))
        else:
            employee_id = self.employee_map.get((base, region))
        if not employee_id:
            messagebox.showwarning("Uyari", "Calisan bulunamadi.")
            return None
        start_date = self.leave_start_var.get().strip()
        end_date = self.leave_end_var.get().strip()
        if not start_date or not end_date:
            messagebox.showwarning("Uyari", "Baslangic ve bitis tarihi zorunlu.")
            return None
        try:
            start_date = normalize_date(start_date)
            end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return None
        if end_date < start_date:
            messagebox.showwarning("Uyari", "Bitis tarihi baslangictan once olamaz.")
            return None
        leave_type = self.leave_type_var.get().strip()
        status = self.leave_status_var.get().strip() or "Onayli"
        reason = self.leave_reason_var.get().strip()
        doc_no = self.leave_doc_no_var.get().strip()
        leave_id = self.leave_id_var.get().strip()
        prev_status = None
        if leave_id:
            existing = db.get_leave_record(parse_int(leave_id))
            if existing:
                prev_status = existing[9]
        if leave_id:
            db.update_leave_record(
                parse_int(leave_id),
                employee_id,
                start_date,
                end_date,
                leave_type,
                reason,
                doc_no,
                status,
                self._entry_region(),
            )
        else:
            leave_id = db.add_leave_record(
                employee_id,
                start_date,
                end_date,
                leave_type,
                reason,
                doc_no,
                status,
                self._entry_region(),
            )
        if not prev_status or prev_status != status:
            try:
                db.add_leave_status_history(parse_int(leave_id), status, reason, self.current_user)
            except Exception:
                pass
        return leave_id

    def _refresh_leave_history(self, leave_id):
        if not hasattr(self, "leave_history_tree"):
            return
        for item in self.leave_history_tree.get_children():
            self.leave_history_tree.delete(item)
        try:
            history = db.list_leave_status_history(parse_int(leave_id))
        except Exception:
            history = []
        for status, note, user, changed_at in history:
            tag = "odd" if len(self.leave_history_tree.get_children()) % 2 else "even"
            self.leave_history_tree.insert(
                "",
                tk.END,
                values=(status, note or "", user or "", changed_at),
                tags=(tag,),
            )
        if not history:
            insert_empty_row(self.leave_history_tree, ("status", "note", "user", "date"), "Kayit yok")

    def set_leave_status(self, new_status):
        selected = self.leave_tree.selection()
        leave_id = self.leave_id_var.get().strip()
        if selected:
            values = self.leave_tree.item(selected[0], "values")
            leave_id = values[0]
        if not leave_id:
            messagebox.showwarning("Uyari", "Durum icin izin secin.")
            return
        leave = db.get_leave_record(parse_int(leave_id))
        if not leave:
            messagebox.showwarning("Uyari", "Kayit bulunamadi.")
            return
        (
            _lid,
            employee_id,
            _name,
            start_date,
            end_date,
            leave_type,
            reason,
            doc_no,
            _doc_path,
            _status,
            region,
        ) = leave
        db.update_leave_record(
            parse_int(leave_id),
            employee_id,
            start_date,
            end_date,
            leave_type,
            reason,
            doc_no,
            new_status,
            region,
        )
        try:
            db.add_leave_status_history(parse_int(leave_id), new_status, reason, self.current_user)
        except Exception:
            pass
        self.refresh_leave_records()
        self.notify(f"Izin durumu guncellendi: {new_status}")

    def add_or_update_leave(self):
        leave_id = self._persist_leave_record()
        if not leave_id:
            return
        if self.leave_apply_attendance_var.get():
            self._apply_leave_to_attendance(leave_id)
        self._log_action("leave_save", f"id={leave_id}")
        self.refresh_leave_records()
        self.clear_leave_form()
        self.notify("Izin kaydedildi.")

    def delete_leave(self):
        leave_id = self.leave_id_var.get().strip()
        if not leave_id:
            selected = self.leave_tree.selection()
            if selected:
                values = self.leave_tree.item(selected[0], "values")
                leave_id = values[0]
        if not leave_id:
            messagebox.showwarning("Uyari", "Silmek icin izin secin.")
            return
        if messagebox.askyesno("Onay", "Izin kaydini silmek istiyor musunuz?"):
            db.delete_leave_record(parse_int(leave_id))
            self._log_action("leave_delete", f"id={leave_id}")
            self.refresh_leave_records()
            self.clear_leave_form()

    def _apply_leave_to_attendance(self, leave_id):
        leave = db.get_leave_record(parse_int(leave_id))
        if not leave:
            return
        (
            _id,
            employee_id,
            _name,
            start_date,
            end_date,
            leave_type,
            reason,
            _doc_no,
            _doc_path,
            _status,
            region,
        ) = leave
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
            end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            return
        day = start_dt
        while day <= end_dt:
            status = "Izinli"
            note = leave_type
            if reason:
                note = f"{leave_type} - {reason}"
            db.upsert_attendance_record(employee_id, day.strftime("%Y-%m-%d"), status, note, region, "Izin")
            day += timedelta(days=1)

    def _calc_leave_days(self, start_date, end_date):
        try:
            start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
            end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
        except ValueError:
            return 0
        if end_dt < start_dt:
            return 0
        return (end_dt - start_dt).days + 1

    def refresh_leave_records(self):
        if not hasattr(self, "leave_tree"):
            return
        for item in self.leave_tree.get_children():
            self.leave_tree.delete(item)
        if hasattr(self, "leave_totals_tree"):
            for item in self.leave_totals_tree.get_children():
                self.leave_totals_tree.delete(item)

        employee_name = self.leave_filter_employee_var.get().strip()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))

        status_filter = self.leave_filter_status_var.get().strip()
        start_date = self.leave_filter_start_var.get().strip() or None
        end_date = self.leave_filter_end_var.get().strip() or None
        search = ""
        if hasattr(self, "leave_filter_search_var"):
            search = self.leave_filter_search_var.get().strip().lower()
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records = db.list_leave_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            status=None if status_filter == "Tum Durumlar" else status_filter,
            region=self._view_region(),
        )

        totals = {}
        total_days_approved = 0
        approved_count = 0
        active_count = 0
        today = datetime.now().date()
        for rec in records:
            (
                leave_id,
                _emp_id,
                name,
                start_date,
                end_date,
                leave_type,
                reason,
                doc_no,
                _doc_path,
                status,
                region,
            ) = rec
            if search:
                hay = " ".join([name, start_date, end_date, reason or "", doc_no or ""]).lower()
                if search not in hay:
                    continue
            days = self._calc_leave_days(start_date, end_date)
            tag = "odd" if len(self.leave_tree.get_children()) % 2 else "even"
            self.leave_tree.insert(
                "",
                tk.END,
                values=(
                    leave_id,
                    name,
                    start_date,
                    end_date,
                    days,
                    leave_type,
                    status,
                    reason or "",
                    doc_no or "",
                    region or "",
                ),
                tags=(tag,),
            )
            if status == "Onayli":
                totals[name] = totals.get(name, 0) + days
                total_days_approved += days
                approved_count += 1
                try:
                    start_dt = datetime.strptime(start_date, "%Y-%m-%d").date()
                    end_dt = datetime.strptime(end_date, "%Y-%m-%d").date()
                    if start_dt <= today <= end_dt:
                        active_count += 1
                except Exception:
                    pass
        if not self.leave_tree.get_children():
            insert_empty_row(
                self.leave_tree,
                ("id", "employee", "start", "end", "days", "type", "status", "reason", "doc_no", "region"),
                "Kayit yok",
            )

        for name, days in sorted(totals.items()):
            self.leave_totals_tree.insert("", tk.END, values=(name, days))
        if hasattr(self, "leave_stats"):
            self._animate_stat(self.leave_stats["total_days"], total_days_approved, decimals=0)
            self._animate_stat(self.leave_stats["approved"], approved_count, decimals=0)
            self._animate_stat(self.leave_stats["active"], active_count, decimals=0)
        if not self._auto_select_tree_first(self.leave_tree, self.on_leave_select):
            if hasattr(self, "leave_detail_employee"):
                self.leave_detail_employee.set("-")
                self.leave_detail_dates.set("-")
                self.leave_detail_type.set("-")
                self.leave_detail_status.set("-")
                self.leave_detail_reason.set("-")
                self.leave_detail_doc.set("-")
            if hasattr(self, "leave_history_tree"):
                for item in self.leave_history_tree.get_children():
                    self.leave_history_tree.delete(item)

    def generate_leave_form(self):
        leave_id = self._persist_leave_record()
        if not leave_id:
            return
        leave = db.get_leave_record(parse_int(leave_id))
        if not leave:
            return
        (
            _id,
            _emp_id,
            name,
            start_date,
            end_date,
            leave_type,
            reason,
            doc_no,
            _doc_path,
            _status,
            _region,
        ) = leave
        details = (
            self.employee_details.get((name, _region or ""))
            or self.employee_details.get((name, ""))
            or {}
        )
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"izin_formu_{name}_{start_date}.xlsx",
        )
        if not output_path:
            return
        report.export_leave_form(
            output_path,
            name,
            details.get("identity_no", ""),
            details.get("department", ""),
            details.get("title", ""),
            start_date,
            end_date,
            leave_type,
            reason,
            self.settings.get("company_name", ""),
            doc_no,
            self.settings.get("logo_path", ""),
        )
        db.update_leave_document(parse_int(leave_id), output_path, doc_no)
        self._log_action("leave_form_export", f"id={leave_id} file={os.path.basename(output_path)}")
        self.refresh_leave_records()
        messagebox.showinfo("Basarili", f"Izin formu kaydedildi: {output_path}")

    def open_leave_document(self):
        selected = self.leave_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Belge acmak icin izin secin.")
            return
        values = self.leave_tree.item(selected[0], "values")
        leave_id = parse_int(values[0])
        leave = db.get_leave_record(leave_id)
        if not leave:
            messagebox.showwarning("Uyari", "Belge bulunamadi.")
            return
        doc_path = leave[8]
        if not doc_path or not os.path.isfile(doc_path):
            messagebox.showwarning("Uyari", "Belge dosyasi bulunamadi.")
            return
        os.startfile(doc_path)

    def import_employees(self):
        path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV", "*.xlsx;*.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            rows = load_tabular_file(path)
        except Exception as exc:
            messagebox.showerror("Hata", f"Dosya okunamadi: {exc}")
            return
        if not rows:
            messagebox.showinfo("Bilgi", "Dosyada veri bulunamadi.")
            return

        header_map = map_headers(rows[0], EMP_HEADER_MAP)
        start_idx = 1 if "full_name" in header_map else 0
        existing_names = set(self.employee_map.keys())
        imported = 0
        skipped = 0
        for row in rows[start_idx:]:
            def cell(idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            if "full_name" in header_map:
                name = str(cell(header_map.get("full_name"))).strip()
                identity_no = str(cell(header_map.get("identity_no"))).strip()
                department = str(cell(header_map.get("department"))).strip()
                title = str(cell(header_map.get("title"))).strip()
            else:
                name = str(cell(0)).strip()
                identity_no = str(cell(1)).strip() if len(row) > 1 else ""
                department = str(cell(2)).strip() if len(row) > 2 else ""
                title = str(cell(3)).strip() if len(row) > 3 else ""

            key = (name, self._entry_region())
            if not name or key in existing_names:
                skipped += 1
                continue
            db.add_employee(name, identity_no, department, title, self._entry_region())
            existing_names.add(key)
            imported += 1

        self.refresh_employees()
        self._log_action("employee_import", f"file={os.path.basename(path)} added={imported} skipped={skipped}")
        messagebox.showinfo("Bilgi", f"Iceri aktarma tamamlandi. Eklenen: {imported}, Atlanan: {skipped}")

    def import_timesheets(self):
        if not self.employee_map:
            messagebox.showwarning("Uyari", "Once calisan ekleyin.")
            return
        path = filedialog.askopenfilename(
            filetypes=[("Excel/CSV", "*.xlsx;*.csv"), ("Excel", "*.xlsx"), ("CSV", "*.csv"), ("All", "*.*")]
        )
        if not path:
            return
        try:
            rows = load_tabular_file(path)
        except Exception as exc:
            messagebox.showerror("Hata", f"Dosya okunamadi: {exc}")
            return
        if not rows:
            messagebox.showinfo("Bilgi", "Dosyada veri bulunamadi.")
            return

        header_map = map_headers(rows[0], TS_HEADER_MAP)
        start_idx = 1 if "employee" in header_map else 0
        imported = 0
        skipped = 0
        missing_employee = 0

        for row in rows[start_idx:]:
            def cell(idx):
                return row[idx] if idx is not None and idx < len(row) else ""

            if "employee" in header_map:
                employee_name = str(cell(header_map.get("employee"))).strip()
                work_date = cell(header_map.get("work_date"))
                start_time = cell(header_map.get("start_time"))
                end_time = cell(header_map.get("end_time"))
                break_minutes = cell(header_map.get("break_minutes"))
                is_special = cell(header_map.get("is_special"))
                notes = str(cell(header_map.get("notes"))).strip()
            else:
                employee_name = str(cell(0)).strip()
                work_date = cell(1)
                start_time = cell(2)
                end_time = cell(3)
                break_minutes = cell(4) if len(row) > 4 else 0
                is_special = cell(5) if len(row) > 5 else 0
                notes = str(cell(6)).strip() if len(row) > 6 else ""

            if not employee_name:
                skipped += 1
                continue
            employee_id = None
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))
            if not employee_id:
                missing_employee += 1
                continue
            try:
                work_date = normalize_date_value(work_date)
                start_time = normalize_time_value(start_time)
                end_time = normalize_time_value(end_time)
            except ValueError:
                skipped += 1
                continue
            break_minutes = parse_int(break_minutes, 0)
            is_special = 1 if parse_bool(is_special) else 0

            db.add_timesheet(
                employee_id,
                work_date,
                start_time,
                end_time,
                break_minutes,
                is_special,
                notes,
                self._entry_region(),
            )
            imported += 1

        self.refresh_timesheets()
        self._log_action(
            "timesheet_import",
            f"file={os.path.basename(path)} added={imported} skipped={skipped} missing={missing_employee}",
        )
        messagebox.showinfo(
            "Bilgi",
            f"Iceri aktarma tamamlandi. Eklenen: {imported}, Atlanan: {skipped}, Calisan bulunamadi: {missing_employee}",
        )

    # Reports tab
    def _build_reports_tab(self):
        frame = ttk.LabelFrame(self.tab_reports_body, text="Excel Raporu", style="Section.TLabelframe")
        frame.pack(fill=tk.X, padx=6, pady=6)

        self.report_employee_var = tk.StringVar(value="Tum Calisanlar")
        self.report_start_var = tk.StringVar()
        self.report_end_var = tk.StringVar()
        self.report_use_dates = tk.BooleanVar(value=False)

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=6)
        ttk.Label(row1, text="Calisan").pack(side=tk.LEFT, padx=(0, 8))
        self.report_employee_combo = ttk.Combobox(row1, textvariable=self.report_employee_var, width=28, state="readonly")
        self.report_employee_combo.pack(side=tk.LEFT)
        report_start_frame, self.report_start_entry = create_labeled_date(
            row1, "Baslangic", self.report_start_var, 12
        )
        report_start_frame.pack(side=tk.LEFT, padx=6)
        report_end_frame, self.report_end_entry = create_labeled_date(
            row1, "Bitis", self.report_end_var, 12
        )
        report_end_frame.pack(side=tk.LEFT, padx=6)
        ttk.Button(row1, text="Rapor Olustur", style="Accent.TButton", command=self.export_report).pack(
            side=tk.LEFT, padx=6
        )
        clear_date_entry(self.report_start_entry)
        clear_date_entry(self.report_end_entry)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=(0, 6))
        ttk.Checkbutton(row2, text="Tarih filtresi kullan", variable=self.report_use_dates).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(row2, text="Son 7 Gun", command=lambda: self.apply_report_preset("last7")).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(row2, text="Bu Ay", command=lambda: self.apply_report_preset("this_month")).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(row2, text="Onceki Ay", command=lambda: self.apply_report_preset("prev_month")).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(row2, text="PDF + Excel", style="Accent.TButton", command=self.export_report_bundle).pack(
            side=tk.LEFT, padx=6
        )

        viewer = ttk.LabelFrame(self.tab_reports_body, text="Rapor Goruntule", style="Section.TLabelframe")
        viewer.pack(fill=tk.X, padx=6, pady=6)
        ttk.Button(viewer, text="XLSX Sec ve Goruntule", command=self.pick_and_preview_report).pack(
            side=tk.LEFT, padx=6, pady=6
        )

        archive = ttk.LabelFrame(self.tab_reports_body, text="Rapor Arsivi", style="Section.TLabelframe")
        archive.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        pane = ttk.PanedWindow(archive, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        list_frame = ttk.Frame(pane)
        detail_frame = ttk.LabelFrame(pane, text="Detay", style="Section.TLabelframe")
        pane.add(list_frame, weight=4)
        pane.add(detail_frame, weight=2)

        self.report_tree = ttk.Treeview(
            list_frame,
            columns=("id", "file", "created", "employee", "range"),
            show="headings",
            height=8,
        )
        self.report_tree.heading("id", text="ID")
        self.report_tree.heading("file", text="Dosya")
        self.report_tree.heading("created", text="Tarih")
        self.report_tree.heading("employee", text="Calisan")
        self.report_tree.heading("range", text="Aralik")
        self.report_tree.column("id", width=60, anchor=tk.CENTER)
        self.report_tree.column("file", width=360)
        self.report_tree.column("created", width=140)
        self.report_tree.column("employee", width=180)
        self.report_tree.column("range", width=180)
        report_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.report_tree.xview)
        report_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.report_tree.yview)
        self.report_tree.configure(xscrollcommand=report_xscroll.set, yscrollcommand=report_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.report_tree.grid(row=0, column=0, sticky="nsew")
        self.report_tree.bind("<<TreeviewSelect>>", self.on_report_select)
        self._apply_tree_zebra(self.report_tree)
        report_yscroll.grid(row=0, column=1, sticky="ns")
        report_xscroll.grid(row=1, column=0, sticky="ew")

        self.report_detail_file = tk.StringVar(value="-")
        self.report_detail_employee = tk.StringVar(value="-")
        self.report_detail_date = tk.StringVar(value="-")
        self.report_detail_range = tk.StringVar(value="-")
        self.report_detail_path = tk.StringVar(value="-")

        dr1 = ttk.Frame(detail_frame)
        dr1.pack(fill=tk.X, pady=4)
        ttk.Label(dr1, text="Dosya").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(dr1, textvariable=self.report_detail_file).pack(side=tk.LEFT)

        dr2 = ttk.Frame(detail_frame)
        dr2.pack(fill=tk.X, pady=4)
        ttk.Label(dr2, text="Calisan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(dr2, textvariable=self.report_detail_employee).pack(side=tk.LEFT)

        dr3 = ttk.Frame(detail_frame)
        dr3.pack(fill=tk.X, pady=4)
        ttk.Label(dr3, text="Tarih").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(dr3, textvariable=self.report_detail_date).pack(side=tk.LEFT)

        dr4 = ttk.Frame(detail_frame)
        dr4.pack(fill=tk.X, pady=4)
        ttk.Label(dr4, text="Aralik").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(dr4, textvariable=self.report_detail_range).pack(side=tk.LEFT)

        dr5 = ttk.Frame(detail_frame)
        dr5.pack(fill=tk.X, pady=4)
        ttk.Label(dr5, text="Yol").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(dr5, textvariable=self.report_detail_path, wraplength=220).pack(side=tk.LEFT)

        btn_row = ttk.Frame(archive)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="Arsivi Yenile", command=self.refresh_report_archive).pack(side=tk.LEFT, padx=6)
        ttk.Button(btn_row, text="Seciliyi Goruntule", command=self.preview_selected_report).pack(
            side=tk.LEFT, padx=6
        )

        note = ttk.Label(
            frame,
            text=f"Tarih formatlari: {DATE_FMT} / Saat formatlari: {TIME_FMT}",
            foreground="#444444",
        )
        note.pack(anchor=tk.W, padx=6, pady=(0, 6))

    def _collect_report_records(self, employee_id=None, start_date=None, end_date=None):
        region = self._view_region()
        timesheet_records = db.list_timesheets(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            region=region,
        )
        attendance_records = db.list_attendance_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            status=None,
            region=region,
        )
        leave_records = db.list_leave_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            status="Onayli",
            region=region,
        )
        return timesheet_records, attendance_records, leave_records

    def export_report(self):
        employee_name = self.report_employee_var.get().strip()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))
        start_date = self.report_start_var.get().strip() or None
        end_date = self.report_end_var.get().strip() or None
        try:
            if self.report_use_dates.get():
                if start_date:
                    start_date = normalize_date(start_date)
                if end_date:
                    end_date = normalize_date(end_date)
            else:
                start_date = None
                end_date = None
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records, attendance_records, leave_records = self._collect_report_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
        )
        if not records and not attendance_records and not leave_records:
            messagebox.showinfo("Bilgi", "Rapor icin veri bulunamadi.")
            return

        if employee_name and employee_name != "Tum Calisanlar":
            employee_slug = employee_name.replace(" ", "_")
        else:
            employee_slug = "tum_calisanlar"
        company_slug = (self.settings.get("company_name", "") or "rainstaff").strip().replace(" ", "_")
        filename = f"{company_slug}_rapor_{employee_slug}_{start_date or 'tum'}_{end_date or 'tum'}.xlsx"
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=filename,
        )
        if not output_path:
            return

        date_text = f"Tarih Araligi: {start_date or '-'} - {end_date or '-'}"
        try:
            report.export_report(
                output_path,
                records,
                db.get_all_settings(),
                date_text,
                attendance_records=attendance_records,
                leave_records=leave_records,
                start_date=start_date,
                end_date=end_date,
            )
        except ValueError as exc:
            messagebox.showerror("Hata", str(exc))
            return
        created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
        db.add_report_log(output_path, created_at, employee_name, start_date, end_date)
        self.refresh_report_archive()
        self._log_action(
            "report_export",
            f"file={os.path.basename(output_path)} employee={employee_name or 'Tum'} range={start_date or '-'}-{end_date or '-'}",
        )
        messagebox.showinfo("Basarili", f"Rapor kaydedildi: {output_path}")

    def refresh_report_archive(self):
        if not hasattr(self, "report_tree"):
            return
        for item in self.report_tree.get_children():
            self.report_tree.delete(item)
        for rep in db.list_report_logs():
            rep_id, file_path, created_at, employee, start_date, end_date = rep
            range_text = f"{start_date or '-'} - {end_date or '-'}"
            values = (rep_id, file_path, created_at, employee or "Tum Calisanlar", range_text)
            self.report_tree.insert("", tk.END, values=values)
        if not self._auto_select_tree_first(self.report_tree, self.on_report_select):
            if hasattr(self, "report_detail_file"):
                self.report_detail_file.set("-")
                self.report_detail_employee.set("-")
                self.report_detail_date.set("-")
                self.report_detail_range.set("-")
                self.report_detail_path.set("-")

    def on_report_select(self, _event=None):
        selected = self.report_tree.selection()
        if not selected:
            return
        values = self.report_tree.item(selected[0], "values")
        if hasattr(self, "report_detail_file"):
            self.report_detail_file.set(os.path.basename(values[1]) if values[1] else "-")
            self.report_detail_date.set(values[2] or "-")
            self.report_detail_range.set(values[4] or "-")
            self.report_detail_employee.set(values[3] or "-")
            self.report_detail_path.set(values[1] or "-")

    def pick_and_preview_report(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not path:
            return
        self.preview_report(path)

    def preview_selected_report(self):
        selected = self.report_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Once rapor secin.")
            return
        values = self.report_tree.item(selected[0], "values")
        path = values[1]
        self.preview_report(path)

    def preview_report(self, path):
        if not path or not os.path.isfile(path):
            messagebox.showwarning("Uyari", "Dosya bulunamadi.")
            return
        try:
            wb = load_workbook(path, data_only=True)
            ws = wb.active
        except Exception as exc:
            messagebox.showerror("Hata", f"Dosya acilamadi: {exc}")
            return

        preview = tk.Toplevel(self)
        preview.title(f"Rapor Goruntule - {os.path.basename(path)}")
        preview.geometry("1100x700")

        frame = ttk.Frame(preview)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        max_cols = min(ws.max_column, 25)
        max_rows = min(ws.max_row, 200)
        headers = [ws.cell(row=1, column=c).value or f"C{c}" for c in range(1, max_cols + 1)]
        tree = ttk.Treeview(frame, columns=list(range(max_cols)), show="headings")
        for idx, header in enumerate(headers):
            tree.heading(idx, text=str(header))
            tree.column(idx, width=120)
        tree.pack(fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        tree.configure(yscrollcommand=yscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)

        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(xscrollcommand=xscroll.set)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        for r in range(2, max_rows + 1):
            row_vals = []
            for c in range(1, max_cols + 1):
                val = ws.cell(row=r, column=c).value
                row_vals.append("" if val is None else val)
            tree.insert("", tk.END, values=row_vals)

    def apply_admin_month_filter(self):
        month_text = self.admin_month_filter_var.get().strip()
        if not month_text:
            return
        try:
            parse_month(month_text)
            month_dt = datetime.strptime(month_text, "%Y-%m")
            last_day = calendar.monthrange(month_dt.year, month_dt.month)[1]
            start_date = f"{month_text}-01"
            end_date = f"{month_text}-{last_day:02d}"
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        self.admin_start_var.set(start_date)
        self.admin_end_var.set(end_date)
        self.refresh_admin_summary()

    def refresh_admin_summary(self):
        if not hasattr(self, "admin_tree"):
            return
        for item in self.admin_tree.get_children():
            self.admin_tree.delete(item)
        if hasattr(self, "admin_alert_tree"):
            for item in self.admin_alert_tree.get_children():
                self.admin_alert_tree.delete(item)
        if hasattr(self, "admin_anomaly_tree"):
            for item in self.admin_anomaly_tree.get_children():
                self.admin_anomaly_tree.delete(item)

        employee_name = self.admin_employee_var.get().strip()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))
        department_filter = self.admin_department_var.get().strip()
        title_filter = self.admin_title_var.get().strip()
        search_text = self.admin_search_var.get().strip().lower()
        start_date = self.admin_start_var.get().strip() or None
        end_date = self.admin_end_var.get().strip() or None
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records = db.list_timesheets(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            region=self._view_region(),
        )
        totals = {}
        daily_overtime = {}
        dept_overtime = {}
        alerts = []
        work_days = {}
        for (
            _ts_id,
            _emp_id,
            name,
            record_department,
            work_date,
            start_time,
            end_time,
            break_minutes,
            is_special,
            _notes,
            _region,
        ) in records:
            details = self.employee_details.get((name, _region or ""), {})
            department = details.get("department", "") or record_department or ""
            title = details.get("title", "")
            if department_filter and department_filter != "Tum Departmanlar" and department != department_filter:
                continue
            if title_filter and title_filter != "Tum Unvanlar" and title != title_filter:
                continue
            if search_text:
                hay = " ".join([name, department, title, str(_notes or "")]).lower()
                if search_text not in hay:
                    continue

            (
                worked,
                _scheduled,
                overtime,
                night_hours,
                overnight_hours,
                spec_norm,
                spec_ot,
                spec_night,
            ) = calc.calc_day_hours(
                work_date,
                start_time,
                end_time,
                break_minutes,
                self.settings,
                is_special,
                department,
            )
            key = (name, _region or "")
            if key not in totals:
                totals[key] = {
                    "name": name,
                    "region": _region or "",
                    "department": department or "",
                    "title": title or "",
                    "worked": 0.0,
                    "overtime": 0.0,
                    "night": 0.0,
                    "overnight": 0.0,
                    "special": 0.0,
                }
            if not totals[key].get("department") and department:
                totals[key]["department"] = department
            if not totals[key].get("title") and title:
                totals[key]["title"] = title
            totals[key]["worked"] += worked
            totals[key]["overtime"] += overtime
            totals[key]["night"] += night_hours
            totals[key]["overnight"] += overnight_hours
            totals[key]["special"] += spec_norm + spec_ot + spec_night

            daily_overtime[work_date] = daily_overtime.get(work_date, 0.0) + overtime
            dept_key = department or "Bilinmeyen"
            dept_overtime[dept_key] = dept_overtime.get(dept_key, 0.0) + overtime

            try:
                gross_hours = calc.hours_between(calc.parse_time(start_time), calc.parse_time(end_time))
            except Exception:
                gross_hours = 0.0
            if gross_hours >= 12:
                alerts.append((work_date, name, "Uzun Mesai", f"{gross_hours:.1f}s"))
            if overnight_hours > 0:
                alerts.append((work_date, name, "Geceye Tasan", f"{overnight_hours:.1f}s"))
            if is_special:
                alerts.append((work_date, name, "Ozel Gun", f"{worked:.1f}s"))

            work_days.setdefault(name, set()).add(work_date)

        total_worked = sum(v["worked"] for v in totals.values())
        total_overtime = sum(v["overtime"] for v in totals.values())
        total_night = sum(v["night"] for v in totals.values())
        total_overnight = sum(v["overnight"] for v in totals.values())
        total_special = sum(v["special"] for v in totals.values())

        self.admin_stats["total_records"].set(str(len(records)))
        self.admin_stats["total_employees"].set(str(len(totals)))
        self.admin_stats["total_worked"].set(f"{total_worked:.2f}")
        self.admin_stats["total_overtime"].set(f"{total_overtime:.2f}")
        self.admin_stats["total_night"].set(f"{total_night:.2f}")
        self.admin_stats["total_overnight"].set(f"{total_overnight:.2f}")
        self.admin_stats["total_special"].set(f"{total_special:.2f}")
        avg_ot = (total_overtime / len(totals)) if totals else 0.0
        self.admin_stats["avg_overtime"].set(f"{avg_ot:.2f}")
        max_daily = max(daily_overtime.values()) if daily_overtime else 0.0
        self.admin_stats["max_daily"].set(f"{max_daily:.2f}")

        self.admin_detail_map = {}
        for _key, data in sorted(totals.items(), key=lambda x: x[1]["name"]):
            display_name = data["name"]
            if data.get("region"):
                display_name = f"{data['name']} ({data['region']})"
            self.admin_tree.insert(
                "",
                tk.END,
                values=(
                    display_name,
                    round(data["worked"], 2),
                    round(data["overtime"], 2),
                    round(data["night"], 2),
                    round(data["overnight"], 2),
                    round(data["special"], 2),
                ),
            )
            days = len(work_days.get(data["name"], []))
            avg_day = (data["worked"] / days) if days else 0.0
            self.admin_detail_map[display_name] = {
                "department": data.get("department") or "-",
                "title": data.get("title") or "-",
                "worked": round(data["worked"], 2),
                "overtime": round(data["overtime"], 2),
                "night": round(data["night"], 2),
                "overnight": round(data["overnight"], 2),
                "special": round(data["special"], 2),
                "avg_day": round(avg_day, 2),
            }

        if hasattr(self, "admin_alert_tree"):
            for work_date, name, issue, value in alerts[:200]:
                self.admin_alert_tree.insert("", tk.END, values=(work_date, name, issue, value))

        if hasattr(self, "admin_anomaly_tree"):
            anomalies = self._build_consecutive_day_anomalies(work_days)
            for name, period, issue in anomalies:
                self.admin_anomaly_tree.insert("", tk.END, values=(name, period, issue))
        if hasattr(self, "admin_tree"):
            if not self._auto_select_tree_first(self.admin_tree, self.on_admin_select):
                if hasattr(self, "admin_detail_name"):
                    self.admin_detail_name.set("-")
                    self.admin_detail_department.set("-")
                    self.admin_detail_title.set("-")
                    self.admin_detail_worked.set("-")
                    self.admin_detail_overtime.set("-")
                    self.admin_detail_night.set("-")
                    self.admin_detail_overnight.set("-")
                    self.admin_detail_special.set("-")
                    self.admin_detail_avg.set("-")

    def _build_consecutive_day_anomalies(self, work_days):
        anomalies = []
        for name, dates in work_days.items():
            try:
                sorted_dates = sorted(datetime.strptime(d, "%Y-%m-%d").date() for d in dates)
            except Exception:
                continue
            streak = 1
            start = sorted_dates[0] if sorted_dates else None
            for i in range(1, len(sorted_dates)):
                if (sorted_dates[i] - sorted_dates[i - 1]).days == 1:
                    streak += 1
                else:
                    if streak >= 6 and start:
                        anomalies.append((name, f"{start} - {sorted_dates[i-1]}", f"{streak} gun ust uste"))
                    streak = 1
                    start = sorted_dates[i]
            if streak >= 6 and start:
                anomalies.append((name, f"{start} - {sorted_dates[-1]}", f"{streak} gun ust uste"))
        return anomalies


    def package_monthly_reports(self):
        month_text = self.admin_month_var.get().strip()
        try:
            parse_month(month_text)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        logs = db.list_report_logs()
        files = [r[1] for r in logs if r[2].startswith(month_text)]
        files = [f for f in files if os.path.isfile(f)]
        if not files:
            messagebox.showinfo("Bilgi", "Secilen ay icin rapor bulunamadi.")
            return
        output_path = filedialog.asksaveasfilename(
            defaultextension=".zip",
            filetypes=[("ZIP", "*.zip")],
            initialfile=f"raporlar_{month_text}.zip",
        )
        if not output_path:
            return
        try:
            with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
                for file_path in files:
                    zf.write(file_path, arcname=os.path.basename(file_path))
        except Exception as exc:
            messagebox.showerror("Hata", f"Zip olusturulamadi: {exc}")
            return
        messagebox.showinfo("Basarili", f"Zip kaydedildi: {output_path}")

    def refresh_vehicles(self):
        if not hasattr(self, "vehicle_tree"):
            return
        for item in self.vehicle_tree.get_children():
            self.vehicle_tree.delete(item)
        self.vehicle_map = {}
        km = None
        oil_change_km = None
        oil_interval_km = None
        for vehicle in db.list_vehicles(region=self._view_region()):
            (
                vehicle_id,
                plate,
                brand,
                model,
                year,
                km,
                inspection_date,
                insurance_date,
                maintenance_date,
                oil_change_date,
                oil_change_km,
                oil_interval_km,
                _notes,
                region,
            ) = vehicle
            oil_status = "-"
            interval_km = oil_interval_km or DEFAULT_OIL_INTERVAL_KM
            if interval_km and oil_change_km is not None and km is not None:
                remaining = interval_km - (km - oil_change_km)
                oil_status = "Geldi" if remaining <= 0 else f"{remaining} km"
            self.vehicle_tree.insert(
                "",
                tk.END,
                values=(
                    vehicle_id,
                    plate,
                    brand,
                    model,
                    year,
                    km,
                    inspection_date,
                    insurance_date,
                    maintenance_date,
                    oil_status,
                    region or "",
                ),
            )
            self.vehicle_map[plate] = vehicle_id
        if hasattr(self, "inspect_vehicle_combo"):
            self.inspect_vehicle_combo["values"] = sorted(self.vehicle_map.keys())
        if hasattr(self, "fault_vehicle_combo"):
            self.fault_vehicle_combo["values"] = sorted(self.vehicle_map.keys())
        if hasattr(self, "service_vehicle_combo"):
            self.service_vehicle_combo["values"] = sorted(self.vehicle_map.keys())

    def refresh_drivers(self):
        if not hasattr(self, "driver_tree"):
            return
        for item in self.driver_tree.get_children():
            self.driver_tree.delete(item)
        self.driver_map = {}
        self.driver_display_names = []
        name_counts = {}
        drivers = db.list_drivers(region=self._view_region())
        for driver in drivers:
            driver_id, name, license_class, license_expiry, phone, _notes, region = driver
            name_counts[name] = name_counts.get(name, 0) + 1
            self.driver_tree.insert(
                "",
                tk.END,
                values=(driver_id, name, license_class, license_expiry, phone, region or ""),
            )
            self.driver_map[(name, region or "")] = driver_id
        for driver in drivers:
            _driver_id, name, _license_class, _license_expiry, _phone, _notes, region = driver
            display = name
            if name_counts.get(name, 0) > 1:
                display = f"{name} ({region or '-'})"
            self.driver_display_names.append(display)
        if hasattr(self, "inspect_driver_combo"):
            self.inspect_driver_combo["values"] = sorted(self.driver_display_names)

    def refresh_faults(self):
        if hasattr(self, "fault_tree"):
            for item in self.fault_tree.get_children():
                self.fault_tree.delete(item)
        self.fault_map = {}
        self.fault_display_by_id = {}
        for fault in db.list_vehicle_faults(region=self._view_region()):
            fault_id, vehicle_id, plate, title, desc, opened_date, closed_date, status, region = fault
            if hasattr(self, "fault_tree"):
                self.fault_tree.insert(
                    "",
                    tk.END,
                    values=(fault_id, plate, title, status, opened_date or "", closed_date or "", region or ""),
                )
            display = f"{plate} - {title} (#{fault_id})"
            self.fault_map[display] = fault_id
            self.fault_display_by_id[fault_id] = display
        if hasattr(self, "inspect_fault_combo"):
            values = [""] + list(self.fault_map.keys())
            self.inspect_fault_combo["values"] = values
        if hasattr(self, "service_fault_combo"):
            values = [""] + list(self.fault_map.keys())
            self.service_fault_combo["values"] = values

    def refresh_service_visits(self):
        if not hasattr(self, "service_tree"):
            return
        for item in self.service_tree.get_children():
            self.service_tree.delete(item)
        self.service_visit_map = {}
        for visit in db.list_vehicle_service_visits(region=self._view_region()):
            (
                visit_id,
                _vehicle_id,
                plate,
                _fault_id,
                fault_title,
                start_date,
                end_date,
                reason,
                cost,
                _notes,
                region,
            ) = visit
            self.service_tree.insert(
                "",
                tk.END,
                values=(
                    visit_id,
                    plate,
                    fault_title or "",
                    start_date,
                    end_date or ("Sanayide" if end_date is None or end_date == "" else ""),
                    f"{cost:.2f}" if cost is not None else "",
                    reason or "",
                    region or "",
                ),
            )
            self.service_visit_map[visit_id] = visit

    def clear_fault_form(self):
        self.fault_id_var.set("")
        self.fault_vehicle_var.set("")
        self.fault_title_var.set("")
        self.fault_desc_var.set("")
        self.fault_open_var.set("")
        self.fault_close_var.set("")
        self.fault_status_var.set("Acik")
        if hasattr(self, "fault_open_entry"):
            clear_date_entry(self.fault_open_entry)
        if hasattr(self, "fault_close_entry"):
            clear_date_entry(self.fault_close_entry)

    def add_or_update_fault(self):
        plate = self.fault_vehicle_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        title = self.fault_title_var.get().strip()
        if not title:
            messagebox.showwarning("Uyari", "Baslik zorunlu.")
            return
        opened_date = self.fault_open_var.get().strip()
        closed_date = self.fault_close_var.get().strip()
        if opened_date:
            try:
                opened_date = normalize_date(opened_date)
            except ValueError as exc:
                messagebox.showwarning("Uyari", str(exc))
                return
        if closed_date:
            try:
                closed_date = normalize_date(closed_date)
            except ValueError as exc:
                messagebox.showwarning("Uyari", str(exc))
                return
        status = self.fault_status_var.get().strip() or "Acik"
        desc = self.fault_desc_var.get().strip()
        fault_id = parse_int(self.fault_id_var.get())
        if fault_id:
            db.update_vehicle_fault(
                fault_id,
                vehicle_id,
                title,
                desc,
                opened_date,
                closed_date,
                status,
                self._entry_region(),
            )
            self._log_action("fault_update", f"id={fault_id} plate={plate} status={status}")
        else:
            db.add_vehicle_fault(
                vehicle_id,
                title,
                desc,
                opened_date,
                closed_date,
                status,
                self._entry_region(),
            )
            self._log_action("fault_add", f"plate={plate} title={title} status={status}")
        self.refresh_faults()
        self.refresh_vehicle_dashboard()
        self.clear_fault_form()
        messagebox.showinfo("Basarili", "Ariza kaydi kaydedildi.")
        self.trigger_sync("fault")

    def delete_fault(self):
        fault_id = parse_int(self.fault_id_var.get())
        if not fault_id:
            selected = self.fault_tree.selection() if hasattr(self, "fault_tree") else None
            if selected:
                values = self.fault_tree.item(selected[0], "values")
                fault_id = parse_int(values[0])
        if not fault_id:
            messagebox.showwarning("Uyari", "Silinecek ariza secin.")
            return
        if not messagebox.askyesno("Onay", "Ariza kaydi silinsin mi?"):
            return
        db.delete_vehicle_fault(fault_id)
        self._log_action("fault_delete", f"id={fault_id}")
        self.refresh_faults()
        self.refresh_vehicle_dashboard()
        self.clear_fault_form()
        self.trigger_sync("fault_delete")

    def on_fault_select(self, _event=None):
        selected = self.fault_tree.selection()
        if not selected:
            return
        values = self.fault_tree.item(selected[0], "values")
        fault_id = parse_int(values[0])
        fault = db.get_vehicle_fault(fault_id)
        if not fault:
            return
        (
            _fid,
            vehicle_id,
            title,
            desc,
            opened_date,
            closed_date,
            status,
        ) = fault
        plate = None
        for plate_name, vid in self.vehicle_map.items():
            if vid == vehicle_id:
                plate = plate_name
                break
        if plate:
            self.fault_vehicle_var.set(plate)
        self.fault_id_var.set(fault_id)
        self.fault_title_var.set(title or "")
        self.fault_desc_var.set(desc or "")
        self.fault_open_var.set(opened_date or "")
        self.fault_close_var.set(closed_date or "")
        self.fault_status_var.set(status or "Acik")

    def clear_service_visit_form(self):
        self.service_id_var.set("")
        self.service_vehicle_var.set("")
        self.service_fault_var.set("")
        self.service_start_var.set("")
        self.service_end_var.set("")
        self.service_reason_var.set("")
        self.service_cost_var.set("")
        self.service_notes_var.set("")
        self.service_in_shop_var.set(False)
        self._toggle_service_end_date()
        if hasattr(self, "service_start_entry"):
            clear_date_entry(self.service_start_entry)
        if hasattr(self, "service_end_entry"):
            clear_date_entry(self.service_end_entry)

    def add_or_update_service_visit(self):
        plate = self.service_vehicle_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        fault_display = self.service_fault_var.get().strip()
        fault_id = self.fault_map.get(fault_display) if fault_display else None
        start_date = self.service_start_var.get().strip()
        if not start_date:
            messagebox.showwarning("Uyari", "Gidis tarihi zorunlu.")
            return
        try:
            start_date = normalize_date(start_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        end_date = self.service_end_var.get().strip()
        if self.service_in_shop_var.get():
            end_date = ""
        elif end_date:
            try:
                end_date = normalize_date(end_date)
            except ValueError as exc:
                messagebox.showwarning("Uyari", str(exc))
                return
        reason = self.service_reason_var.get().strip()
        notes = self.service_notes_var.get().strip()
        cost = self.service_cost_var.get().strip()
        cost_val = None if cost == "" else parse_float(cost, 0.0)
        visit_id = parse_int(self.service_id_var.get())
        if visit_id:
            db.update_vehicle_service_visit(
                visit_id,
                vehicle_id,
                fault_id,
                start_date,
                end_date,
                reason,
                cost_val,
                notes,
                self._entry_region(),
            )
            self._log_action("service_visit_update", f"id={visit_id} plate={plate}")
        else:
            db.add_vehicle_service_visit(
                vehicle_id,
                fault_id,
                start_date,
                end_date,
                reason,
                cost_val,
                notes,
                self._entry_region(),
            )
            self._log_action("service_visit_add", f"plate={plate} start={start_date}")
        self.refresh_service_visits()
        self.refresh_vehicle_dashboard()
        self.clear_service_visit_form()
        messagebox.showinfo("Basarili", "Sanayi kaydi kaydedildi.")
        self.trigger_sync("service_visit")

    def delete_service_visit(self):
        visit_id = parse_int(self.service_id_var.get())
        if not visit_id:
            selected = self.service_tree.selection() if hasattr(self, "service_tree") else None
            if selected:
                values = self.service_tree.item(selected[0], "values")
                visit_id = parse_int(values[0])
        if not visit_id:
            messagebox.showwarning("Uyari", "Silinecek kaydi secin.")
            return
        if not messagebox.askyesno("Onay", "Sanayi kaydi silinsin mi?"):
            return
        db.delete_vehicle_service_visit(visit_id)
        self._log_action("service_visit_delete", f"id={visit_id}")
        self.refresh_service_visits()
        self.refresh_vehicle_dashboard()
        self.clear_service_visit_form()
        self.trigger_sync("service_visit_delete")

    def on_service_visit_select(self, _event=None):
        selected = self.service_tree.selection()
        if not selected:
            return
        values = self.service_tree.item(selected[0], "values")
        visit_id = parse_int(values[0])
        visit = db.get_vehicle_service_visit(visit_id)
        if not visit:
            return
        (
            _vid,
            vehicle_id,
            fault_id,
            start_date,
            end_date,
            reason,
            cost,
            notes,
        ) = visit
        plate = None
        for plate_name, vid in self.vehicle_map.items():
            if vid == vehicle_id:
                plate = plate_name
                break
        if plate:
            self.service_vehicle_var.set(plate)
        fault_display = ""
        if fault_id:
            for display, fid in self.fault_map.items():
                if fid == fault_id:
                    fault_display = display
                    break
        self.service_fault_var.set(fault_display)
        self.service_id_var.set(visit_id)
        self.service_start_var.set(start_date or "")
        self.service_end_var.set(end_date or "")
        self.service_in_shop_var.set(False if end_date else True)
        self._toggle_service_end_date()
        self.service_reason_var.set(reason or "")
        self.service_cost_var.set("" if cost is None else f"{cost:.2f}")
        self.service_notes_var.set(notes or "")

    def _toggle_service_end_date(self):
        state = "disabled" if self.service_in_shop_var.get() else "normal"
        if hasattr(self, "service_end_entry"):
            try:
                self.service_end_entry.configure(state=state)
            except tk.TclError:
                pass

    def refresh_vehicle_dashboard(self):
        if not hasattr(self, "vehicle_status_tree"):
            return
        if not hasattr(self, "dashboard_stats"):
            return
        for item in self.vehicle_status_tree.get_children():
            self.vehicle_status_tree.delete(item)
        for item in self.vehicle_alert_tree.get_children():
            self.vehicle_alert_tree.delete(item)
        for item in self.driver_alert_tree.get_children():
            self.driver_alert_tree.delete(item)

        vehicles = db.list_vehicles(region=self._view_region())
        drivers = db.list_drivers(region=self._view_region())
        driver_by_id = {row[0]: row for row in drivers}
        driver_latest = {}

        # Populate vehicle_map for alert clicks
        self.vehicle_map = {}

        oil_due = 0
        insp_due = 0
        ins_due = 0
        maint_due = 0
        lic_due = 0

        for vehicle in vehicles:
            (
                _vid,
                plate,
                brand,
                model,
                year,
                km,
                inspection_date,
                insurance_date,
                maintenance_date,
                oil_change_date,
                oil_change_km,
                oil_interval_km,
                _notes,
                region,
            ) = vehicle
            # Map plate to vehicle ID for alert clicks
            self.vehicle_map[plate] = _vid
            oil_status = "-"
            oil_flag = None
            interval_km = oil_interval_km or DEFAULT_OIL_INTERVAL_KM
            if interval_km and oil_change_km is not None and km is not None:
                remaining = interval_km - (km - oil_change_km)
                oil_status = "Geldi" if remaining <= 0 else f"{remaining} km"
                if remaining <= 0:
                    oil_due += 1
                    oil_flag = "oil_due"
                    self.vehicle_alert_tree.insert("", tk.END, values=(plate, "Yag Degisimi", "Geldi"))
                elif remaining <= DEFAULT_OIL_SOON_KM:
                    oil_flag = "oil_soon"

            insp_days = days_until(inspection_date)
            if insp_days is not None and insp_days <= 30:
                insp_due += 1
                detail = f"{inspection_date} ({insp_days} gun)"
                if insp_days < 0:
                    detail = f"{inspection_date} ({abs(insp_days)} gun gecikme)"
                self.vehicle_alert_tree.insert("", tk.END, values=(plate, "Muayene", detail))

            ins_days = days_until(insurance_date)
            if ins_days is not None and ins_days <= 30:
                ins_due += 1
                detail = f"{insurance_date} ({ins_days} gun)"
                if ins_days < 0:
                    detail = f"{insurance_date} ({abs(ins_days)} gun gecikme)"
                self.vehicle_alert_tree.insert("", tk.END, values=(plate, "Sigorta", detail))

            maint_days = days_until(maintenance_date)
            if maint_days is not None and maint_days <= 30:
                maint_due += 1
                detail = f"{maintenance_date} ({maint_days} gun)"
                if maint_days < 0:
                    detail = f"{maintenance_date} ({abs(maint_days)} gun gecikme)"
                self.vehicle_alert_tree.insert("", tk.END, values=(plate, "Bakim", detail))

            inspections = db.list_vehicle_inspections(vehicle_id=_vid, region=self._view_region())
            last_check = "-"
            last_driver = "-"
            if inspections:
                last_inspection = inspections[0]
                last_check = last_inspection[5]
                last_driver = last_inspection[4] or "-"
                driver_id = last_inspection[3]
                if driver_id:
                    current = driver_latest.get(driver_id)
                    if not current or last_inspection[5] > current[5]:
                        driver_latest[driver_id] = last_inspection

            if len(inspections) >= 2:
                current_inspection = inspections[0]
                previous_inspection = inspections[1]
                current_results = {
                    row[0]: normalize_vehicle_status(row[1])
                    for row in db.list_vehicle_inspection_results(current_inspection[0])
                }
                prev_results = {
                    row[0]: normalize_vehicle_status(row[1])
                    for row in db.list_vehicle_inspection_results(previous_inspection[0])
                }
                for item_key, label in VEHICLE_CHECKLIST:
                    if (
                        current_results.get(item_key) == "Olumsuz"
                        and prev_results.get(item_key) == "Olumsuz"
                    ):
                        self.vehicle_alert_tree.insert(
                            "",
                            tk.END,
                            values=(plate, "Tekrar Eden Sorun", f"{label} (2 hafta)"),
                        )

            self.vehicle_status_tree.insert(
                "",
                tk.END,
                values=(
                    plate,
                    km or "-",
                    oil_status,
                    inspection_date or "-",
                    insurance_date or "-",
                    maintenance_date or "-",
                    last_check,
                    last_driver,
                    region or "",
                ),
                tags=(oil_flag,) if oil_flag else (),
            )

        for driver in drivers:
            _did, name, _cls, license_expiry, _phone, _notes, _region = driver
            days = days_until(license_expiry)
            if days is not None and days <= 30:
                lic_due += 1
                detail = f"{license_expiry} ({days} gun)"
                if days < 0:
                    detail = f"{license_expiry} ({abs(days)} gun gecikme)"
                self.driver_alert_tree.insert("", tk.END, values=(name, "Ehliyet", detail))

        faults = db.list_vehicle_faults(region=self._view_region())
        now = datetime.now().date()
        faults_by_plate = {}
        for fault in faults:
            _fid, _vid, plate, title, _desc, opened_date, _closed_date, status, _region = fault
            faults_by_plate.setdefault(plate, []).append(fault)
            if status == "Acik":
                self.vehicle_alert_tree.insert("", tk.END, values=(plate, "Acik Ariza", title))

        for plate, items in faults_by_plate.items():
            title_counts = {}
            for fault in items:
                opened_date = fault[5]
                if not opened_date:
                    continue
                try:
                    opened_dt = datetime.strptime(opened_date, "%Y-%m-%d").date()
                except ValueError:
                    continue
                if (now - opened_dt).days <= 30:
                    title_counts[fault[3]] = title_counts.get(fault[3], 0) + 1
            for title, count in title_counts.items():
                if count >= 2:
                    self.vehicle_alert_tree.insert(
                        "",
                        tk.END,
                        values=(plate, "Tekrar Ariza (30 gun)", f"{title} x{count}"),
                    )

        fault_counts = {}
        for fault in faults:
            plate = fault[2]
            fault_counts[plate] = fault_counts.get(plate, 0) + 1
        top_faults = sorted(fault_counts.items(), key=lambda x: x[1], reverse=True)[:3]
        for plate, count in top_faults:
            self.vehicle_alert_tree.insert("", tk.END, values=(plate, "En Cok Ariza", f"{count} kayit"))

        for driver_id, inspection in driver_latest.items():
            driver = driver_by_id.get(driver_id)
            if not driver:
                continue
            name = driver[1]
            plate = inspection[2]
            results = db.list_vehicle_inspection_results(inspection[0])
            bad_items = []
            for item_key, label in VEHICLE_CHECKLIST:
                for res_key, status, _note in results:
                    if res_key == item_key and normalize_vehicle_status(status) == "Olumsuz":
                        bad_items.append(label)
                        break
            if bad_items:
                detail = f"{plate} - {len(bad_items)} olumsuz"
                self.driver_alert_tree.insert("", tk.END, values=(name, "Son Kontrol Sorun", detail))

        self.dashboard_stats["vehicles"].set(str(len(vehicles)))
        self.dashboard_stats["drivers"].set(str(len(drivers)))
        self.dashboard_stats["oil_due"].set(str(oil_due))
        self.dashboard_stats["inspection_due"].set(str(insp_due))
        self.dashboard_stats["insurance_due"].set(str(ins_due))
        self.dashboard_stats["maintenance_due"].set(str(maint_due))
        self.dashboard_stats["license_due"].set(str(lic_due))

    def clear_vehicle_form(self):
        self.vehicle_id_var.set("")
        self.vehicle_plate_var.set("")
        self.vehicle_brand_var.set("")
        self.vehicle_model_var.set("")
        self.vehicle_year_var.set("")
        self.vehicle_km_var.set("")
        self.vehicle_inspection_var.set("")
        self.vehicle_insurance_var.set("")
        self.vehicle_maintenance_var.set("")
        self.vehicle_oil_var.set("")
        self.vehicle_oil_km_var.set("")
        self.vehicle_oil_interval_var.set(str(DEFAULT_OIL_INTERVAL_KM))
        self.vehicle_notes_var.set("")
        self.vehicle_original_plate = None
        clear_date_entry(self.vehicle_insp_entry)
        clear_date_entry(self.vehicle_ins_entry)
        clear_date_entry(self.vehicle_maint_entry)
        clear_date_entry(self.vehicle_oil_entry)

    def on_vehicle_select(self, _event=None):
        selected = self.vehicle_tree.selection()
        if not selected:
            return
        values = self.vehicle_tree.item(selected[0], "values")
        vehicle_id = parse_int(values[0])
        row = db.get_vehicle(vehicle_id)
        if not row:
            return
        (
            _vid,
            plate,
            brand,
            model,
            year,
            km,
            inspection_date,
            insurance_date,
            maintenance_date,
            oil_change_date,
            oil_change_km,
            oil_interval_km,
            _notes,
            _region,
        ) = row
        self.vehicle_id_var.set(vehicle_id)
        self.vehicle_plate_var.set(plate)
        self.vehicle_original_plate = plate
        self.vehicle_brand_var.set(brand)
        self.vehicle_model_var.set(model)
        self.vehicle_year_var.set(year)
        self.vehicle_km_var.set(km or "")
        self.vehicle_inspection_var.set(inspection_date or "")
        self.vehicle_insurance_var.set(insurance_date or "")
        self.vehicle_maintenance_var.set(maintenance_date or "")
        self.vehicle_oil_var.set(oil_change_date or "")
        self.vehicle_oil_km_var.set(oil_change_km or "")
        self.vehicle_oil_interval_var.set(oil_interval_km or str(DEFAULT_OIL_INTERVAL_KM))

    def add_or_update_vehicle(self):
        plate = self.vehicle_plate_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Plaka zorunlu.")
            return
        brand = self.vehicle_brand_var.get().strip()
        model = self.vehicle_model_var.get().strip()
        year = self.vehicle_year_var.get().strip()
        km = parse_int(self.vehicle_km_var.get(), 0)
        inspection_date = self.vehicle_inspection_var.get().strip()
        insurance_date = self.vehicle_insurance_var.get().strip()
        maintenance_date = self.vehicle_maintenance_var.get().strip()
        oil_change_date = self.vehicle_oil_var.get().strip()
        oil_change_km = parse_int(self.vehicle_oil_km_var.get(), 0)
        oil_interval_km = parse_int(self.vehicle_oil_interval_var.get(), 0)
        notes = self.vehicle_notes_var.get().strip()
        try:
            inspection_date = normalize_date(inspection_date) if inspection_date else ""
            insurance_date = normalize_date(insurance_date) if insurance_date else ""
            maintenance_date = normalize_date(maintenance_date) if maintenance_date else ""
            oil_change_date = normalize_date(oil_change_date) if oil_change_date else ""
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        vehicle_id = self.vehicle_id_var.get().strip()
        if vehicle_id and self.vehicle_original_plate and plate != self.vehicle_original_plate:
            add_new = messagebox.askyesno(
                "Onay",
                "Plaka degisti. Yeni arac olarak eklensin mi?\n"
                "Evet: yeni kayit, Hayir: mevcut araci guncelle.",
            )
            if add_new:
                vehicle_id = ""
        try:
            if vehicle_id:
                db.update_vehicle(
                    parse_int(vehicle_id),
                    plate,
                    brand,
                    model,
                    year,
                    km,
                    inspection_date,
                    insurance_date,
                    maintenance_date,
                    oil_change_date,
                    oil_change_km,
                    oil_interval_km,
                    notes,
                    self._entry_region(),
                )
                self._log_action("vehicle_update", f"id={vehicle_id} plate={plate}")
            else:
                db.add_vehicle(
                    plate,
                    brand,
                    model,
                    year,
                    km,
                    inspection_date,
                    insurance_date,
                    maintenance_date,
                    oil_change_date,
                    oil_change_km,
                    oil_interval_km,
                    notes,
                    self._entry_region(),
                )
                self._log_action("vehicle_add", f"plate={plate}")
        except Exception:
            messagebox.showwarning("Uyari", "Arac kaydi eklenemedi. Plaka zaten var olabilir.")
            return
        self.refresh_vehicles()
        self.clear_vehicle_form()
        self.trigger_sync("vehicle")

    def delete_vehicle(self):
        vehicle_id = self.vehicle_id_var.get().strip()
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Silmek icin arac secin.")
            return
        if messagebox.askyesno("Onay", "Araci silmek istiyor musunuz?"):
            db.delete_vehicle(parse_int(vehicle_id))
            self._log_action("vehicle_delete", f"id={vehicle_id}")
            self.refresh_vehicles()
            self.clear_vehicle_form()
            self.trigger_sync("vehicle_delete")

    def clear_driver_form(self):
        self.driver_id_var.set("")
        self.driver_name_var.set("")
        self.driver_license_var.set("")
        self.driver_license_exp_var.set("")
        self.driver_phone_var.set("")
        self.driver_notes_var.set("")
        clear_date_entry(self.driver_exp_entry)

    def on_driver_select(self, _event=None):
        selected = self.driver_tree.selection()
        if not selected:
            return
        values = self.driver_tree.item(selected[0], "values")
        self.driver_id_var.set(values[0])
        self.driver_name_var.set(values[1])
        self.driver_license_var.set(values[2])
        self.driver_license_exp_var.set(values[3] or "")
        self.driver_phone_var.set(values[4])

    def add_or_update_driver(self):
        name = self.driver_name_var.get().strip()
        if not name:
            messagebox.showwarning("Uyari", "Ad Soyad zorunlu.")
            return
        license_class = self.driver_license_var.get().strip()
        license_expiry = self.driver_license_exp_var.get().strip()
        phone = self.driver_phone_var.get().strip()
        notes = self.driver_notes_var.get().strip()
        try:
            license_expiry = normalize_date(license_expiry) if license_expiry else ""
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        driver_id = self.driver_id_var.get().strip()
        if driver_id:
            db.update_driver(
                parse_int(driver_id),
                name,
                license_class,
                license_expiry,
                phone,
                notes,
                self._entry_region(),
            )
            self._log_action("driver_update", f"id={driver_id} name={name}")
        else:
            db.add_driver(name, license_class, license_expiry, phone, notes, self._entry_region())
            self._log_action("driver_add", f"name={name}")
        self.refresh_drivers()
        self.clear_driver_form()
        self.trigger_sync("driver")

    def delete_driver(self):
        driver_id = self.driver_id_var.get().strip()
        if not driver_id:
            messagebox.showwarning("Uyari", "Silmek icin surucu secin.")
            return
        if messagebox.askyesno("Onay", "Surucuyu silmek istiyor musunuz?"):
            db.delete_driver(parse_int(driver_id))
            self._log_action("driver_delete", f"id={driver_id}")
            self.refresh_drivers()
            self.clear_driver_form()
            self.trigger_sync("driver_delete")

    def save_vehicle_inspection(self):
        plate = self.inspect_vehicle_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        driver_name = self.inspect_driver_var.get().strip()
        driver_id = None
        if driver_name:
            base, region = split_display_name(driver_name, REGIONS)
            if region is None:
                driver_id = self.driver_map.get((base, "")) or self.driver_map.get((base, self._entry_region()))
            else:
                driver_id = self.driver_map.get((base, region))
        inspect_date = self.inspect_date_var.get().strip()
        if not inspect_date:
            messagebox.showwarning("Uyari", "Tarih zorunlu.")
            return
        inspect_km = parse_int(self.inspect_km_var.get(), 0)
        try:
            inspect_date = normalize_date(inspect_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        week_start = week_start_from_date(inspect_date)
        notes = self.inspect_notes_var.get().strip()
        fault_display = self.inspect_fault_var.get().strip()
        fault_id = self.fault_map.get(fault_display) if fault_display else None
        fault_status = self.inspect_fault_status_var.get().strip() if fault_id else None
        service_visit = 1 if self.inspect_service_var.get() else 0
        if service_visit and not fault_id:
            messagebox.showwarning("Uyari", "Sanayi icin ariza secin.")
            return
        inspection_id = db.add_vehicle_inspection(
            vehicle_id,
            driver_id,
            inspect_date,
            week_start,
            inspect_km,
            notes,
            fault_id=fault_id,
            fault_status=fault_status,
            service_visit=service_visit,
        )
        self._log_action(
            "vehicle_inspection_add",
            f"plate={plate} date={inspect_date} km={inspect_km}",
        )
        for item_key, _label in VEHICLE_CHECKLIST:
            status = self.inspect_item_vars[item_key].get()
            note = self.inspect_note_vars[item_key].get().strip()
            db.add_vehicle_inspection_result(inspection_id, item_key, status, note)
        if inspect_km:
            current = db.get_vehicle(vehicle_id)
            if current:
                (
                    _vid,
                    plate,
                    brand,
                    model,
                    year,
                    _km,
                    inspection_date,
                    insurance_date,
                    maintenance_date,
                    oil_change_date,
                    oil_change_km,
                    oil_interval_km,
                    notes,
                    _region,
                ) = current
                db.update_vehicle(
                    vehicle_id,
                    plate,
                    brand,
                    model,
                    year,
                    inspect_km,
                    inspection_date or "",
                    insurance_date or "",
                    maintenance_date or "",
                    oil_change_date or "",
                    oil_change_km or 0,
                    oil_interval_km or 0,
                    notes or "",
                    self._entry_region(),
                )
                self.refresh_vehicles()
        if fault_id and fault_status == "Kapandi":
            fault = db.get_vehicle_fault(fault_id)
            if fault:
                (
                    _fid,
                    f_vehicle_id,
                    title,
                    desc,
                    opened_date,
                    closed_date,
                    status,
                ) = fault
                if not closed_date:
                    closed_date = inspect_date
                db.update_vehicle_fault(
                    fault_id,
                    f_vehicle_id,
                    title,
                    desc,
                    opened_date,
                    closed_date,
                    "Kapandi",
                    self._entry_region(),
                )
                self.refresh_faults()
        messagebox.showinfo("Basarili", "Haftalik kontrol kaydedildi.")
        self.trigger_sync("vehicle_inspection")

    def on_inspect_vehicle_change(self, _event=None):
        plate = self.inspect_vehicle_var.get().strip()
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            return
        faults = db.list_open_vehicle_faults(vehicle_id=vehicle_id, region=self._view_region())
        if not faults:
            self.inspect_fault_var.set("")
            self.inspect_fault_status_var.set("Acik")
            return
        latest = faults[0]
        fault_id = latest[0]
        display = self.fault_display_by_id.get(fault_id)
        if display:
            self.inspect_fault_var.set(display)
            self.inspect_fault_status_var.set(latest[7] or "Acik")

    def _get_latest_inspection(self, vehicle_id, week_start):
        inspections = db.list_vehicle_inspections(
            vehicle_id=vehicle_id,
            week_start=week_start,
            region=self._view_region(),
        )
        if not inspections:
            return None
        inspections.sort(key=lambda x: x[5], reverse=True)
        return inspections[0]

    def compare_vehicle_week(self):
        for item in self.vehicle_compare_tree.get_children():
            self.vehicle_compare_tree.delete(item)
        plate = self.inspect_vehicle_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        inspect_date = self.inspect_date_var.get().strip()
        if not inspect_date:
            messagebox.showwarning("Uyari", "Tarih secin.")
            return
        try:
            inspect_date = normalize_date(inspect_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        week_start = week_start_from_date(inspect_date)
        prev_week = (datetime.strptime(week_start, "%Y-%m-%d").date() - timedelta(days=7)).strftime("%Y-%m-%d")

        current = self._get_latest_inspection(vehicle_id, week_start)
        previous = self._get_latest_inspection(vehicle_id, prev_week)
        if not current:
            messagebox.showwarning("Uyari", "Bu hafta icin kontrol bulunamadi.")
            return
        current_results = {row[0]: normalize_vehicle_status(row[1]) for row in db.list_vehicle_inspection_results(current[0])}
        prev_results = (
            {row[0]: normalize_vehicle_status(row[1]) for row in db.list_vehicle_inspection_results(previous[0])}
            if previous
            else {}
        )

        for item_key, label in VEHICLE_CHECKLIST:
            prev_status = prev_results.get(item_key, "-")
            curr_status = current_results.get(item_key, "-")
            if prev_status == curr_status:
                change = "Ayni"
            elif prev_status == "Olumsuz" and curr_status == "Olumsuz":
                change = "Tekrar eden sorun"
            elif curr_status == "Olumsuz" and prev_status != "Olumsuz":
                change = "Kotulesti"
            elif prev_status == "Olumsuz" and curr_status != "Olumsuz":
                change = "Iyilesti"
            else:
                change = "Degisti"
            self.vehicle_compare_tree.insert("", tk.END, values=(label, prev_status, curr_status, change))
        prev_fault_status = previous[10] if previous else None
        curr_fault_status = current[10]
        if prev_fault_status or curr_fault_status:
            prev_fault_status = prev_fault_status or "-"
            curr_fault_status = curr_fault_status or "-"
            if prev_fault_status == curr_fault_status:
                change = "Ayni"
            elif curr_fault_status == "Kapandi":
                change = "Iyilesti"
            elif prev_fault_status == "Kapandi" and curr_fault_status != "Kapandi":
                change = "Kotulesti"
            else:
                change = "Degisti"
            self.vehicle_compare_tree.insert(
                "",
                tk.END,
                values=("Ariza Durumu", prev_fault_status, curr_fault_status, change),
            )
        prev_service = "Evet" if previous and previous[11] else "Hayir"
        curr_service = "Evet" if current[11] else "Hayir"
        change = "Ayni" if prev_service == curr_service else "Degisti"
        self.vehicle_compare_tree.insert(
            "",
            tk.END,
            values=("Sanayiye Gitti", prev_service, curr_service, change),
        )

    def export_vehicle_weekly_report(self):
        plate = self.inspect_vehicle_var.get().strip()
        if not plate:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        inspect_date = self.inspect_date_var.get().strip()
        if not inspect_date:
            messagebox.showwarning("Uyari", "Tarih secin.")
            return
        try:
            inspect_date = normalize_date(inspect_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        week_start = week_start_from_date(inspect_date)
        prev_week = (datetime.strptime(week_start, "%Y-%m-%d").date() - timedelta(days=7)).strftime("%Y-%m-%d")
        current = self._get_latest_inspection(vehicle_id, week_start)
        previous = self._get_latest_inspection(vehicle_id, prev_week)
        if not current:
            messagebox.showwarning("Uyari", "Bu hafta icin kontrol bulunamadi.")
            return
        current_results = {row[0]: normalize_vehicle_status(row[1]) for row in db.list_vehicle_inspection_results(current[0])}
        prev_results = (
            {row[0]: normalize_vehicle_status(row[1]) for row in db.list_vehicle_inspection_results(previous[0])}
            if previous
            else {}
        )
        current_fault_id = current[9]
        current_fault_status = current[10] or ""
        current_service = bool(current[11])
        prev_fault_id = previous[9] if previous else None
        prev_fault_status = previous[10] if previous else ""
        prev_service = bool(previous[11]) if previous else False
        current_fault_title = ""
        prev_fault_title = ""
        if current_fault_id:
            fault = db.get_vehicle_fault(current_fault_id)
            if fault:
                current_fault_title = fault[2] or ""
        if prev_fault_id:
            fault = db.get_vehicle_fault(prev_fault_id)
            if fault:
                prev_fault_title = fault[2] or ""
        week_end = week_end_from_start(week_start)
        service_visits = db.list_vehicle_service_visits(
            vehicle_id=vehicle_id,
            start_date=week_start,
            end_date=week_end,
            region=self._view_region(),
        )

        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"arac_kontrol_{plate}_{week_start}.xlsx",
        )
        if not output_path:
            return
        vehicle = None
        for row in db.list_vehicles(region=self._view_region()):
            if row[1] == plate:
                vehicle = row
                break
        report.export_vehicle_weekly_report(
            output_path,
            plate,
            week_start,
            prev_week if previous else None,
            VEHICLE_CHECKLIST,
            prev_results,
            current_results,
            current[7],
            previous[7] if previous else None,
            vehicle,
            {
                "title": current_fault_title,
                "status": current_fault_status,
                "service": current_service,
            },
            {
                "title": prev_fault_title,
                "status": prev_fault_status,
                "service": prev_service,
            },
            service_visits,
        )
        self._log_action("vehicle_weekly_report", f"plate={plate} week={week_start} file={os.path.basename(output_path)}")

    def export_vehicle_card(self, plate):
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        vehicle = db.get_vehicle(vehicle_id)
        if not vehicle:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        inspections = db.list_vehicle_inspections(vehicle_id=vehicle_id, region=self._view_region())
        faults = db.list_vehicle_faults(vehicle_id=vehicle_id, region=self._view_region())
        services = db.list_vehicle_service_visits(vehicle_id=vehicle_id, region=self._view_region())
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"arac_karti_{plate}.xlsx",
        )
        if not output_path:
            return
        report.export_vehicle_card_report(
            output_path,
            plate,
            vehicle,
            inspections,
            faults,
            services,
        )
        self._log_action("vehicle_card_report", f"plate={plate} file={os.path.basename(output_path)}")
        messagebox.showinfo("Basarili", "Arac karti Excel olusturuldu.")
        messagebox.showinfo("Basarili", f"Rapor kaydedildi: {output_path}")

    def on_admin_right_click(self, event):
        row_id = self.admin_tree.identify_row(event.y)
        if row_id:
            self.admin_tree.selection_set(row_id)
            self.admin_menu.tk_popup(event.x_root, event.y_root)

    def on_admin_select(self, _event=None):
        selected = self.admin_tree.selection()
        if not selected:
            return
        values = self.admin_tree.item(selected[0], "values")
        name = values[0]
        detail = (self.admin_detail_map or {}).get(name, {})
        if hasattr(self, "admin_detail_name"):
            self.admin_detail_name.set(name)
            self.admin_detail_department.set(detail.get("department", "-"))
            self.admin_detail_title.set(detail.get("title", "-"))
            self.admin_detail_worked.set(detail.get("worked", "-"))
            self.admin_detail_overtime.set(detail.get("overtime", "-"))
            self.admin_detail_night.set(detail.get("night", "-"))
            self.admin_detail_overnight.set(detail.get("overnight", "-"))
            self.admin_detail_special.set(detail.get("special", "-"))
            self.admin_detail_avg.set(detail.get("avg_day", "-"))

    def on_admin_alert_select(self, _event=None):
        selected = self.admin_alert_tree.selection()
        if not selected:
            return
        values = self.admin_alert_tree.item(selected[0], "values")
        if hasattr(self, "admin_alert_detail"):
            self.admin_alert_detail.set(f"{values[0]} | {values[1]} | {values[2]}: {values[3]}")

    def on_admin_anomaly_select(self, _event=None):
        selected = self.admin_anomaly_tree.selection()
        if not selected:
            return
        values = self.admin_anomaly_tree.item(selected[0], "values")
        if hasattr(self, "admin_anomaly_detail"):
            self.admin_anomaly_detail.set(f"{values[0]} | {values[1]} | {values[2]}")

    def copy_selected_admin_rows(self):
        selected = self.admin_tree.selection()
        if not selected:
            messagebox.showinfo("Bilgi", "Kopyalamak icin satir secin.")
            return
        lines = []
        for item in selected:
            values = self.admin_tree.item(item, "values")
            lines.append("\t".join(str(v) for v in values))
        data = "\n".join(lines)
        self.clipboard_clear()
        self.clipboard_append(data)
        self.notify("Secili satirlar panoya kopyalandi.")

    def export_selected_admin_report(self):
        selected = self.admin_tree.selection()
        if not selected:
            messagebox.showinfo("Bilgi", "Rapor icin calisan secin.")
            return
        names = [self.admin_tree.item(item, "values")[0] for item in selected]
        clean_names = [split_display_name(n, REGIONS)[0] for n in names]
        start_date = self.admin_start_var.get().strip() or None
        end_date = self.admin_end_var.get().strip() or None
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return
        records, attendance_records, leave_records = self._collect_report_records(
            start_date=start_date,
            end_date=end_date,
        )
        name_set = set(clean_names)
        filtered = [r for r in records if r[2] in name_set]
        filtered_attendance = [r for r in attendance_records if len(r) > 2 and r[2] in name_set]
        filtered_leave = [r for r in leave_records if len(r) > 2 and r[2] in name_set]
        if not filtered and not filtered_attendance and not filtered_leave:
            messagebox.showinfo("Bilgi", "Secili calisanlar icin veri bulunamadi.")
            return
        filename = f"admin_secili_rapor_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=filename,
        )
        if not output_path:
            return
        date_text = f"Tarih Araligi: {start_date or '-'} - {end_date or '-'}"
        report.export_report(
            output_path,
            filtered,
            db.get_all_settings(),
            date_text,
            attendance_records=filtered_attendance,
            leave_records=filtered_leave,
            start_date=start_date,
            end_date=end_date,
        )
        messagebox.showinfo("Basarili", f"Rapor kaydedildi: {output_path}")

    def export_report_bundle(self):
        employee_name = self.report_employee_var.get().strip()
        employee_id = None
        if employee_name and employee_name != "Tum Calisanlar":
            base, region = split_display_name(employee_name, REGIONS)
            if region is None:
                employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                    (base, self._entry_region())
                )
            else:
                employee_id = self.employee_map.get((base, region))
        start_date = self.report_start_var.get().strip() or None
        end_date = self.report_end_var.get().strip() or None
        try:
            if self.report_use_dates.get():
                if start_date:
                    start_date = normalize_date(start_date)
                if end_date:
                    end_date = normalize_date(end_date)
            else:
                start_date = None
                end_date = None
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records, attendance_records, leave_records = self._collect_report_records(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
        )
        if not records and not attendance_records and not leave_records:
            messagebox.showinfo("Bilgi", "Rapor icin veri bulunamadi.")
            return

        if employee_name and employee_name != "Tum Calisanlar":
            employee_slug = employee_name.replace(" ", "_")
        else:
            employee_slug = "tum_calisanlar"
        company_slug = (self.settings.get("company_name", "") or "rainstaff").strip().replace(" ", "_")
        base_name = f"{company_slug}_rapor_{employee_slug}_{start_date or 'tum'}_{end_date or 'tum'}"
        output_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel", "*.xlsx")],
            initialfile=f"{base_name}.xlsx",
        )
        if not output_path:
            return
        date_text = f"Tarih Araligi: {start_date or '-'} - {end_date or '-'}"
        try:
            report.export_report(
                output_path,
                records,
                db.get_all_settings(),
                date_text,
                attendance_records=attendance_records,
                leave_records=leave_records,
                start_date=start_date,
                end_date=end_date,
            )
        except ValueError as exc:
            messagebox.showerror("Hata", str(exc))
            return

        pdf_path = os.path.splitext(output_path)[0] + ".pdf"
        pdf_ok = True
        try:
            report.export_report_pdf(pdf_path, records, db.get_all_settings(), date_text)
        except Exception as exc:
            pdf_ok = False
            messagebox.showwarning("Uyari", f"PDF olusturulamadi: {exc}")

        created_at = datetime.now().strftime("%Y-%m-%d %H:%M")
        db.add_report_log(output_path, created_at, employee_name, start_date, end_date)
        self.refresh_report_archive()
        if pdf_ok:
            messagebox.showinfo("Basarili", f"Excel + PDF kaydedildi:\n{output_path}\n{pdf_path}")
        else:
            messagebox.showinfo("Basarili", f"Excel kaydedildi: {output_path}")

    def apply_report_preset(self, preset):
        today = datetime.now().date()
        if preset == "last7":
            start = today - timedelta(days=6)
            end = today
        elif preset == "this_month":
            start = today.replace(day=1)
            end = today
        elif preset == "prev_month":
            first_day = today.replace(day=1)
            end = first_day - timedelta(days=1)
            start = end.replace(day=1)
        else:
            return
        self.report_use_dates.set(True)
        self.report_start_var.set(start.strftime("%Y-%m-%d"))
        self.report_end_var.set(end.strftime("%Y-%m-%d"))
        if hasattr(self, "report_start_entry"):
            self.report_start_entry.configure(state="normal")
        if hasattr(self, "report_end_entry"):
            self.report_end_entry.configure(state="normal")
        self.refresh_report_archive()

    def show_admin_employee_detail(self):
        selected = self.admin_tree.selection()
        if not selected:
            return
        values = self.admin_tree.item(selected[0], "values")
        employee_name = values[0]
        employee_id = None
        base, region = split_display_name(employee_name, REGIONS)
        if region is None:
            employee_id = self.employee_map.get((base, "")) or self.employee_map.get(
                (base, self._entry_region())
            )
        else:
            employee_id = self.employee_map.get((base, region))
        if not employee_id:
            messagebox.showwarning("Uyari", "Calisan bulunamadi.")
            return

        start_date = self.admin_start_var.get().strip() or None
        end_date = self.admin_end_var.get().strip() or None
        try:
            if start_date:
                start_date = normalize_date(start_date)
            if end_date:
                end_date = normalize_date(end_date)
        except ValueError as exc:
            messagebox.showwarning("Uyari", str(exc))
            return

        records = db.list_timesheets(
            employee_id=employee_id,
            start_date=start_date,
            end_date=end_date,
            region=self._view_region(),
        )
        detail = tk.Toplevel(self)
        detail.title(f"Mesai Detay - {employee_name}")
        detail.geometry("980x600")

        frame = ttk.Frame(detail)
        frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        cols = ("date", "start", "end", "worked", "overtime", "night", "overnight", "special", "note")
        tree = ttk.Treeview(frame, columns=cols, show="headings")
        tree.heading("date", text="Tarih")
        tree.heading("start", text="Giris")
        tree.heading("end", text="Cikis")
        tree.heading("worked", text="Calisilan")
        tree.heading("overtime", text="Fazla Mesai")
        tree.heading("night", text="Gece")
        tree.heading("overnight", text="Geceye Tasan")
        tree.heading("special", text="Ozel Gun")
        tree.heading("note", text="Not")
        tree.column("date", width=100)
        tree.column("start", width=70)
        tree.column("end", width=70)
        tree.column("worked", width=90)
        tree.column("overtime", width=90)
        tree.column("night", width=90)
        tree.column("overnight", width=110)
        tree.column("special", width=90)
        tree.column("note", width=240)
        tree.pack(fill=tk.BOTH, expand=True)

        yscroll = ttk.Scrollbar(frame, orient=tk.VERTICAL, command=tree.yview)
        xscroll = ttk.Scrollbar(frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        yscroll.pack(side=tk.RIGHT, fill=tk.Y)
        xscroll.pack(side=tk.BOTTOM, fill=tk.X)

        for (
            _ts_id,
            _emp_id,
            _name,
            department,
            work_date,
            start_time,
            end_time,
            break_minutes,
            is_special,
            notes,
            _region,
        ) in records:
            (
                worked,
                _scheduled,
                overtime,
                night_hours,
                overnight_hours,
                spec_norm,
                spec_ot,
                spec_night,
            ) = calc.calc_day_hours(
                work_date,
                start_time,
                end_time,
                break_minutes,
                self.settings,
                is_special,
                department,
            )
            special_total = round(spec_norm + spec_ot + spec_night, 2)
            tree.insert(
                "",
                tk.END,
                values=(
                    work_date,
                    start_time,
                    end_time,
                    worked,
                    overtime,
                    night_hours,
                    overnight_hours,
                    special_total,
                    notes or "",
                ),
            )

    # Settings tab
    def _build_settings_tab(self):
        frame = ttk.LabelFrame(self.tab_settings_body, text="Genel Ayarlar", style="Section.TLabelframe")
        frame.pack(fill=tk.X, padx=6, pady=6)

        kpi_row = ttk.Frame(self.tab_settings_body)
        kpi_row.pack(fill=tk.X, padx=6, pady=6)
        self.settings_stats = {
            "users": tk.StringVar(value="0"),
            "regions": tk.StringVar(value=str(len(REGIONS)))
        }
        create_kpi_card(kpi_row, "Kullanici", self.settings_stats["users"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)
        create_kpi_card(kpi_row, "Bolge", self.settings_stats["regions"], theme=self._ui_theme).pack(side=tk.LEFT, padx=6)
        try:
            self._animate_stat(self.settings_stats["users"], len(db.list_users()), decimals=0)
        except Exception:
            pass

        self.company_name_var = tk.StringVar(value=self.settings.get("company_name", ""))
        self.report_title_var = tk.StringVar(value=self.settings.get("report_title", "Puantaj ve Mesai Raporu"))
        self.weekday_hours_var = tk.StringVar(value=self.settings.get("weekday_hours", "8"))
        self.sat_start_var = tk.StringVar(value=self.settings.get("saturday_start", "09:00"))
        self.sat_end_var = tk.StringVar(value=self.settings.get("saturday_end", "14:00"))
        self.logo_path_var = tk.StringVar(value=self.settings.get("logo_path", ""))
        self.admin_entry_region_var.set(self.settings.get("admin_entry_region", "Ankara"))
        view_region = self.settings.get("admin_view_region", "Tum Bolgeler")
        if view_region == "ALL":
            view_region = "Tum Bolgeler"
        self.admin_view_region_var.set(view_region)

        row1 = ttk.Frame(frame)
        row1.pack(fill=tk.X, pady=4)
        create_labeled_entry(row1, "Kurum Adi", self.company_name_var, 40).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row1, "Rapor Basligi", self.report_title_var, 32).pack(side=tk.LEFT, padx=6)

        row2 = ttk.Frame(frame)
        row2.pack(fill=tk.X, pady=4)
        create_labeled_entry(row2, "Hafta Ici Saat", self.weekday_hours_var, 10).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row2, "Cumartesi Baslangic", self.sat_start_var, 10).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(row2, "Cumartesi Bitis", self.sat_end_var, 10).pack(side=tk.LEFT, padx=6)
        if self.is_admin:
            ttk.Label(row2, text="Kayit Bolge").pack(side=tk.LEFT, padx=(12, 6))
            entry_region_combo = ttk.Combobox(
                row2,
                textvariable=self.admin_entry_region_var,
                values=REGIONS,
                width=12,
                state="readonly",
            )
            entry_region_combo.pack(side=tk.LEFT, padx=6)
            ttk.Label(row2, text="Goruntuleme Bolge").pack(side=tk.LEFT, padx=(12, 6))
            view_region_combo = ttk.Combobox(
                row2,
                textvariable=self.admin_view_region_var,
                values=VIEW_REGIONS,
                width=14,
                state="readonly",
            )
            view_region_combo.pack(side=tk.LEFT, padx=6)
            view_region_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_region_views())

        row3 = ttk.Frame(frame)
        row3.pack(fill=tk.X, pady=4)
        ttk.Label(row3, text="Logo").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Entry(row3, textvariable=self.logo_path_var, width=60).pack(side=tk.LEFT)
        ttk.Button(row3, text="Sec", command=self.select_logo).pack(side=tk.LEFT, padx=6)

        btn_row = ttk.Frame(frame)
        btn_row.pack(fill=tk.X, pady=6)
        ttk.Button(btn_row, text="Kaydet", style="Accent.TButton", command=self.save_settings).pack(
            side=tk.LEFT, padx=6
        )

        data_frame = ttk.LabelFrame(self.tab_settings_body, text="Veri Yonetimi", style="Section.TLabelframe")
        data_frame.pack(fill=tk.X, padx=6, pady=6)
        drow1 = ttk.Frame(data_frame)
        drow1.pack(fill=tk.X, pady=4)
        ttk.Button(drow1, text="Log Klasoru", command=self.open_log_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(drow1, text="Veri Klasoru", command=self.open_data_folder).pack(side=tk.LEFT, padx=6)
        ttk.Button(drow1, text="Yedek Al", command=self.backup_database).pack(side=tk.LEFT, padx=6)
        ttk.Button(drow1, text="Yedek Geri Yukle", command=self.restore_database).pack(side=tk.LEFT, padx=6)

        drow2 = ttk.Frame(data_frame)
        drow2.pack(fill=tk.X, pady=4)
        ttk.Button(drow2, text="Veri Disari Aktar (ZIP)", command=self.export_data_zip).pack(side=tk.LEFT, padx=6)
        ttk.Button(drow2, text="Veri Iceri Aktar (ZIP)", command=self.import_data_zip).pack(side=tk.LEFT, padx=6)
        ttk.Label(
            data_frame,
            text="Yedek ve tasima islemleri veritabani dosyasini kopyalar. Islemden sonra uygulamayi yeniden baslatin.",
            foreground="#5f6a72",
        ).pack(anchor="w", padx=8, pady=(2, 6))

        template_frame = ttk.LabelFrame(self.tab_settings_body, text="Vardiya Sablonlari", style="Section.TLabelframe")
        template_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.st_id_var = tk.StringVar()
        self.st_name_var = tk.StringVar()
        self.st_start_var = tk.StringVar(value="09:00")
        self.st_end_var = tk.StringVar(value="18:00")
        self.st_break_var = tk.StringVar(value="60")

        trow1 = ttk.Frame(template_frame)
        trow1.pack(fill=tk.X, pady=4)
        create_labeled_entry(trow1, "Sablon Adi", self.st_name_var, 26).pack(side=tk.LEFT, padx=6)
        start_tpl_frame = create_time_entry(trow1, "Giris", self.st_start_var, 8)
        start_tpl_frame.pack(side=tk.LEFT, padx=6)
        end_tpl_frame = create_time_entry(trow1, "Cikis", self.st_end_var, 8)
        end_tpl_frame.pack(side=tk.LEFT, padx=6)
        create_labeled_entry(trow1, "Mola dk", self.st_break_var, 8).pack(side=tk.LEFT, padx=6)
        for child in start_tpl_frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.bind("<FocusOut>", lambda _e: normalize_time_in_var(self.st_start_var))
        for child in end_tpl_frame.winfo_children():
            if isinstance(child, ttk.Entry):
                child.bind("<FocusOut>", lambda _e: normalize_time_in_var(self.st_end_var))

        trow2 = ttk.Frame(template_frame)
        trow2.pack(fill=tk.X, pady=4)
        ttk.Button(trow2, text="Kaydet", style="Accent.TButton", command=self.save_shift_template).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(trow2, text="Sil", command=self.delete_shift_template).pack(side=tk.LEFT)
        ttk.Button(trow2, text="Temizle", command=self.clear_shift_template_form).pack(side=tk.LEFT, padx=6)

        list_frame = ttk.Frame(template_frame)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        columns = ("id", "name", "start", "end", "break")
        self.template_tree = ttk.Treeview(list_frame, columns=columns, show="headings")
        self.template_tree.heading("id", text="ID")
        self.template_tree.heading("name", text="Sablon")
        self.template_tree.heading("start", text="Giris")
        self.template_tree.heading("end", text="Cikis")
        self.template_tree.heading("break", text="Mola")
        self.template_tree.column("id", width=60, anchor=tk.CENTER)
        self.template_tree.column("name", width=220)
        self.template_tree.column("start", width=80)
        self.template_tree.column("end", width=80)
        self.template_tree.column("break", width=80)
        self.template_tree.tag_configure("odd", background="#252525", foreground="#E0E0E0")
        self.template_tree.tag_configure("even", background="#1F1F1F", foreground="#E0E0E0")
        tpl_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.template_tree.xview)
        tpl_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.template_tree.yview)
        self.template_tree.configure(xscrollcommand=tpl_xscroll.set, yscrollcommand=tpl_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.template_tree.grid(row=0, column=0, sticky="nsew")
        tpl_yscroll.grid(row=0, column=1, sticky="ns")
        tpl_xscroll.grid(row=1, column=0, sticky="ew")
        self.template_tree.bind("<<TreeviewSelect>>", self.on_template_select)
        self._apply_tree_zebra(self.template_tree)

    # Kullanım rehberi kaldırıldı - Modern ERP tasarımına geçildi



    def _build_logs_tab(self):
        frame = ttk.LabelFrame(self.tab_logs_body, text="Canli Log", style="Section.TLabelframe")
        frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        tools = ttk.Frame(frame)
        tools.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(tools, text="Temizle", command=self.clear_log_view).pack(side=tk.LEFT, padx=6)
        ttk.Button(tools, text="Log Dosyasi", command=self.open_log_file).pack(side=tk.LEFT, padx=6)
        ttk.Label(tools, text=LOG_PATH, foreground="#5f6a72").pack(side=tk.LEFT, padx=8)

        text_frame = tk.Frame(frame, bg="#0b1118")
        text_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = tk.Text(
            text_frame,
            height=26,
            wrap="none",
            bg="#0b1118",
            fg="#9FE870",
            insertbackground="#9FE870",
            font=("Consolas", 10),
        )
        self.log_text.tag_configure("banner", foreground="#7FD6FF")
        self.log_text.tag_configure("credit", foreground="#FFD46A")
        yscroll = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        xscroll = ttk.Scrollbar(text_frame, orient=tk.HORIZONTAL, command=self.log_text.xview)
        self.log_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        text_frame.columnconfigure(0, weight=1)
        text_frame.rowconfigure(0, weight=1)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        banner = (
            "     .  .  .     .  .  .     .  .  .     .  .  .\n"
            "   .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .\n"
            " .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .  .\n"
            "     .  .  .     .  .  .     .  .  .     .  .  .\n"
        )
        self.log_text.insert(tk.END, banner, "banner")
        self.log_text.insert(tk.END, "\nmade by @hsyncnorman\n\n", "credit")
        self.log_text.insert(tk.END, "Log ekranina hosgeldiniz.\n")
        self.log_text.configure(state=tk.DISABLED)
        footer = tk.Frame(frame, bg="#0b1118", height=10)
        footer.pack(fill=tk.X)
        self.after(300, self._drain_log_queue)

    def _build_vehicles_tab(self):
        vehicle_frame = ttk.LabelFrame(self.tab_vehicles_body, text="Arac Bilgisi", style="Section.TLabelframe")
        vehicle_frame.pack(fill=tk.X, padx=6, pady=6)

        self.vehicle_id_var = tk.StringVar()
        self.vehicle_plate_var = tk.StringVar()
        self.vehicle_brand_var = tk.StringVar()
        self.vehicle_model_var = tk.StringVar()
        self.vehicle_year_var = tk.StringVar()
        self.vehicle_km_var = tk.StringVar()
        self.vehicle_inspection_var = tk.StringVar()
        self.vehicle_insurance_var = tk.StringVar()
        self.vehicle_maintenance_var = tk.StringVar()
        self.vehicle_oil_var = tk.StringVar()
        self.vehicle_oil_km_var = tk.StringVar()
        self.vehicle_oil_interval_var = tk.StringVar(value=str(DEFAULT_OIL_INTERVAL_KM))
        self.vehicle_notes_var = tk.StringVar()

        vrow1 = ttk.Frame(vehicle_frame)
        vrow1.pack(fill=tk.X, pady=4)
        create_labeled_entry(vrow1, "Plaka", self.vehicle_plate_var, 12).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow1, "Marka", self.vehicle_brand_var, 14).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow1, "Model", self.vehicle_model_var, 14).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow1, "Yil", self.vehicle_year_var, 8).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow1, "KM", self.vehicle_km_var, 10).pack(side=tk.LEFT, padx=6)

        vrow2 = ttk.Frame(vehicle_frame)
        vrow2.pack(fill=tk.X, pady=4)
        insp_frame, self.vehicle_insp_entry = create_labeled_date(vrow2, "Muayene", self.vehicle_inspection_var, 12)
        insp_frame.pack(side=tk.LEFT, padx=6)
        ins_frame, self.vehicle_ins_entry = create_labeled_date(vrow2, "Sigorta", self.vehicle_insurance_var, 12)
        ins_frame.pack(side=tk.LEFT, padx=6)
        maint_frame, self.vehicle_maint_entry = create_labeled_date(
            vrow2, "Bakim", self.vehicle_maintenance_var, 12
        )
        maint_frame.pack(side=tk.LEFT, padx=6)
        oil_frame, self.vehicle_oil_entry = create_labeled_date(vrow2, "Yag", self.vehicle_oil_var, 12)
        oil_frame.pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow2, "Yag KM", self.vehicle_oil_km_var, 10).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow2, "Periyot KM", self.vehicle_oil_interval_var, 10).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(vrow2, "Not", self.vehicle_notes_var, 30).pack(side=tk.LEFT, padx=6)

        vbtn = ttk.Frame(vehicle_frame)
        vbtn.pack(fill=tk.X, pady=6)
        ttk.Button(vbtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_vehicle).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(vbtn, text="Sil", command=self.delete_vehicle).pack(side=tk.LEFT)
        ttk.Button(vbtn, text="Temizle", command=self.clear_vehicle_form).pack(side=tk.LEFT, padx=6)

        vlist_frame = ttk.Frame(self.tab_vehicles_body)
        vlist_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.vehicle_tree = ttk.Treeview(
            vlist_frame,
            columns=(
                "id",
                "plate",
                "brand",
                "model",
                "year",
                "km",
                "inspection",
                "insurance",
                "maintenance",
                "oil_due",
                "region",
            ),
            show="headings",
        )
        self.vehicle_tree.heading("id", text="ID")
        self.vehicle_tree.heading("plate", text="Plaka")
        self.vehicle_tree.heading("brand", text="Marka")
        self.vehicle_tree.heading("model", text="Model")
        self.vehicle_tree.heading("year", text="Yil")
        self.vehicle_tree.heading("km", text="KM")
        self.vehicle_tree.heading("inspection", text="Muayene")
        self.vehicle_tree.heading("insurance", text="Sigorta")
        self.vehicle_tree.heading("maintenance", text="Bakim")
        self.vehicle_tree.heading("oil_due", text="Yag Durum")
        self.vehicle_tree.heading("region", text="Bolge")
        self.vehicle_tree.column("id", width=60, anchor=tk.CENTER)
        self.vehicle_tree.column("plate", width=100)
        self.vehicle_tree.column("brand", width=120)
        self.vehicle_tree.column("model", width=120)
        self.vehicle_tree.column("year", width=70)
        self.vehicle_tree.column("km", width=80)
        self.vehicle_tree.column("inspection", width=100)
        self.vehicle_tree.column("insurance", width=100)
        self.vehicle_tree.column("maintenance", width=100)
        self.vehicle_tree.column("oil_due", width=120)
        self.vehicle_tree.column("region", width=100)
        v_xscroll = ttk.Scrollbar(vlist_frame, orient=tk.HORIZONTAL, command=self.vehicle_tree.xview)
        v_yscroll = ttk.Scrollbar(vlist_frame, orient=tk.VERTICAL, command=self.vehicle_tree.yview)
        self.vehicle_tree.configure(xscrollcommand=v_xscroll.set, yscrollcommand=v_yscroll.set)
        vlist_frame.columnconfigure(0, weight=1)
        vlist_frame.rowconfigure(0, weight=1)
        self.vehicle_tree.grid(row=0, column=0, sticky="nsew")
        v_yscroll.grid(row=0, column=1, sticky="ns")
        v_xscroll.grid(row=1, column=0, sticky="ew")
        self.vehicle_tree.bind("<<TreeviewSelect>>", self.on_vehicle_select)
        self.vehicle_tree.bind("<Double-1>", lambda _e: self.show_vehicle_card_from_list())

        vcard_row = ttk.Frame(self.tab_vehicles_body)
        vcard_row.pack(fill=tk.X, padx=6, pady=4)
        ttk.Button(vcard_row, text="Arac Karti", command=self.show_vehicle_card_from_list).pack(
            side=tk.LEFT, padx=6
        )

        driver_frame = ttk.LabelFrame(self.tab_vehicles_body, text="Surucu Bilgisi", style="Section.TLabelframe")
        driver_frame.pack(fill=tk.X, padx=6, pady=6)

        self.driver_id_var = tk.StringVar()
        self.driver_name_var = tk.StringVar()
        self.driver_license_var = tk.StringVar()
        self.driver_license_exp_var = tk.StringVar()
        self.driver_phone_var = tk.StringVar()
        self.driver_notes_var = tk.StringVar()

        drow1 = ttk.Frame(driver_frame)
        drow1.pack(fill=tk.X, pady=4)
        create_labeled_entry(drow1, "Ad Soyad", self.driver_name_var, 24).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(drow1, "Ehliyet Sinifi", self.driver_license_var, 12).pack(side=tk.LEFT, padx=6)
        d_exp_frame, self.driver_exp_entry = create_labeled_date(drow1, "Bitis", self.driver_license_exp_var, 12)
        d_exp_frame.pack(side=tk.LEFT, padx=6)
        create_labeled_entry(drow1, "Telefon", self.driver_phone_var, 14).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(drow1, "Not", self.driver_notes_var, 30).pack(side=tk.LEFT, padx=6)

        dbtn = ttk.Frame(driver_frame)
        dbtn.pack(fill=tk.X, pady=6)
        ttk.Button(dbtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_driver).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(dbtn, text="Sil", command=self.delete_driver).pack(side=tk.LEFT)
        ttk.Button(dbtn, text="Temizle", command=self.clear_driver_form).pack(side=tk.LEFT, padx=6)

        dlist_frame = ttk.Frame(self.tab_vehicles_body)
        dlist_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.driver_tree = ttk.Treeview(
            dlist_frame,
            columns=("id", "name", "license", "expiry", "phone", "region"),
            show="headings",
        )
        self.driver_tree.heading("id", text="ID")
        self.driver_tree.heading("name", text="Ad Soyad")
        self.driver_tree.heading("license", text="Ehliyet")
        self.driver_tree.heading("expiry", text="Bitis")
        self.driver_tree.heading("phone", text="Telefon")
        self.driver_tree.heading("region", text="Bolge")
        self.driver_tree.column("id", width=60, anchor=tk.CENTER)
        self.driver_tree.column("name", width=220)
        self.driver_tree.column("license", width=100)
        self.driver_tree.column("expiry", width=100)
        self.driver_tree.column("phone", width=120)
        self.driver_tree.column("region", width=100)
        d_xscroll = ttk.Scrollbar(dlist_frame, orient=tk.HORIZONTAL, command=self.driver_tree.xview)
        d_yscroll = ttk.Scrollbar(dlist_frame, orient=tk.VERTICAL, command=self.driver_tree.yview)
        self.driver_tree.configure(xscrollcommand=d_xscroll.set, yscrollcommand=d_yscroll.set)
        dlist_frame.columnconfigure(0, weight=1)
        dlist_frame.rowconfigure(0, weight=1)
        self.driver_tree.grid(row=0, column=0, sticky="nsew")
        d_yscroll.grid(row=0, column=1, sticky="ns")
        d_xscroll.grid(row=1, column=0, sticky="ew")
        self.driver_tree.bind("<<TreeviewSelect>>", self.on_driver_select)
        self.driver_tree.bind("<Double-1>", lambda _e: self.show_driver_detail_from_list())

        inspect_frame = ttk.LabelFrame(self.tab_vehicles_body, text="Haftalik Kontrol", style="Section.TLabelframe")
        inspect_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        self.inspect_vehicle_var = tk.StringVar()
        self.inspect_driver_var = tk.StringVar()
        self.inspect_date_var = tk.StringVar()
        self.inspect_km_var = tk.StringVar()
        self.inspect_notes_var = tk.StringVar()
        self.inspect_fault_var = tk.StringVar()
        self.inspect_fault_status_var = tk.StringVar(value="Acik")
        self.inspect_service_var = tk.BooleanVar(value=False)

        irow1 = ttk.Frame(inspect_frame)
        irow1.pack(fill=tk.X, pady=4)
        ttk.Label(irow1, text="Arac").pack(side=tk.LEFT, padx=(0, 6))
        self.inspect_vehicle_combo = ttk.Combobox(irow1, textvariable=self.inspect_vehicle_var, width=18, state="readonly")
        self.inspect_vehicle_combo.pack(side=tk.LEFT)
        self.inspect_vehicle_combo.bind("<<ComboboxSelected>>", self.on_inspect_vehicle_change)
        ttk.Label(irow1, text="Surucu").pack(side=tk.LEFT, padx=(12, 6))
        self.inspect_driver_combo = ttk.Combobox(irow1, textvariable=self.inspect_driver_var, width=18, state="readonly")
        self.inspect_driver_combo.pack(side=tk.LEFT)
        date_frame, self.inspect_date_entry = create_labeled_date(irow1, "Tarih", self.inspect_date_var, 12)
        date_frame.pack(side=tk.LEFT, padx=6)
        create_labeled_entry(irow1, "KM", self.inspect_km_var, 10).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(irow1, "Not", self.inspect_notes_var, 40).pack(side=tk.LEFT, padx=6)

        irow2 = ttk.Frame(inspect_frame)
        irow2.pack(fill=tk.X, pady=4)
        ttk.Label(irow2, text="Ariza").pack(side=tk.LEFT, padx=(0, 6))
        self.inspect_fault_combo = ttk.Combobox(
            irow2,
            textvariable=self.inspect_fault_var,
            width=40,
            state="readonly",
            values=[""],
        )
        self.inspect_fault_combo.pack(side=tk.LEFT)
        ttk.Label(irow2, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Combobox(
            irow2,
            textvariable=self.inspect_fault_status_var,
            values=["Acik", "Kapandi", "Takip"],
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT)
        ttk.Checkbutton(irow2, text="Sanayiye Gitti", variable=self.inspect_service_var).pack(
            side=tk.LEFT, padx=12
        )
        ttk.Label(
            inspect_frame,
            text="Ariza opsiyonel. Sanayiye gidis icin ariza secin; donus tarihi bilinmiyorsa bos kalabilir.",
            foreground="#5f6a72",
        ).pack(anchor="w", padx=8, pady=(2, 6))

        self.inspect_item_vars = {}
        self.inspect_note_vars = {}
        status_values = ["Olumlu", "Olumsuz", "Bilinmiyor"]
        for item_key, label in VEHICLE_CHECKLIST:
            row = ttk.Frame(inspect_frame)
            row.pack(fill=tk.X, pady=2)
            ttk.Label(row, text=label, width=22).pack(side=tk.LEFT, padx=6)
            status_var = tk.StringVar(value="Olumlu")
            note_var = tk.StringVar()
            ttk.Combobox(row, textvariable=status_var, values=status_values, width=10, state="readonly").pack(
                side=tk.LEFT
            )
            ttk.Entry(row, textvariable=note_var, width=60).pack(side=tk.LEFT, padx=6)
            self.inspect_item_vars[item_key] = status_var
            self.inspect_note_vars[item_key] = note_var

        ibtn = ttk.Frame(inspect_frame)
        ibtn.pack(fill=tk.X, pady=6)
        ttk.Button(ibtn, text="Kontrolu Kaydet", style="Accent.TButton", command=self.save_vehicle_inspection).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(ibtn, text="Haftalik Karsilastir", command=self.compare_vehicle_week).pack(side=tk.LEFT, padx=6)
        ttk.Button(ibtn, text="Excel Raporu", command=self.export_vehicle_weekly_report).pack(
            side=tk.LEFT, padx=6
        )

        compare_frame = ttk.LabelFrame(self.tab_vehicles_body, text="Haftalik Karsilastirma", style="Section.TLabelframe")
        compare_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.vehicle_compare_tree = ttk.Treeview(
            compare_frame,
            columns=("item", "prev", "current", "change"),
            show="headings",
            height=8,
        )
        self.vehicle_compare_tree.heading("item", text="Kontrol")
        self.vehicle_compare_tree.heading("prev", text="Onceki Hafta")
        self.vehicle_compare_tree.heading("current", text="Bu Hafta")
        self.vehicle_compare_tree.heading("change", text="Durum")
        self.vehicle_compare_tree.column("item", width=220)
        self.vehicle_compare_tree.column("prev", width=120)
        self.vehicle_compare_tree.column("current", width=120)
        self.vehicle_compare_tree.column("change", width=160)
        vc_xscroll = ttk.Scrollbar(compare_frame, orient=tk.HORIZONTAL, command=self.vehicle_compare_tree.xview)
        vc_yscroll = ttk.Scrollbar(compare_frame, orient=tk.VERTICAL, command=self.vehicle_compare_tree.yview)
        self.vehicle_compare_tree.configure(xscrollcommand=vc_xscroll.set, yscrollcommand=vc_yscroll.set)
        compare_frame.columnconfigure(0, weight=1)
        compare_frame.rowconfigure(0, weight=1)
        self.vehicle_compare_tree.grid(row=0, column=0, sticky="nsew")
        vc_yscroll.grid(row=0, column=1, sticky="ns")
        vc_xscroll.grid(row=1, column=0, sticky="ew")

    def _build_dashboard_tab(self):
        content = self.tab_dashboard_body

        kpi_frame = ttk.Frame(content)
        kpi_frame.pack(fill=tk.X, padx=6, pady=6)

        self.dash_stats = {
            "employees": tk.StringVar(value="0"),
            "timesheets": tk.StringVar(value="0"),
            "worked": tk.StringVar(value="0"),
            "overtime": tk.StringVar(value="0"),
            "scheduled": tk.StringVar(value="0"),
            "completion": tk.StringVar(value="0"),
        }
        create_kpi_card(kpi_frame, "Toplam Calisan", self.dash_stats["employees"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(kpi_frame, "Aralik Puantaj", self.dash_stats["timesheets"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(kpi_frame, "Aralik Calisilan", self.dash_stats["worked"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(kpi_frame, "Aralik Fazla Mesai", self.dash_stats["overtime"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(kpi_frame, "Planlanan", self.dash_stats["scheduled"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )
        create_kpi_card(kpi_frame, "Gerceklesen %", self.dash_stats["completion"], theme=self._ui_theme).pack(
            side=tk.LEFT, padx=6
        )

        dash_filter = ttk.LabelFrame(content, text="Trend Filtre", style="Section.TLabelframe")
        dash_filter.pack(fill=tk.X, padx=6, pady=6)
        self.dash_range_var = tk.StringVar(value="Son 7 Gun")
        self.dash_start_var = tk.StringVar()
        self.dash_end_var = tk.StringVar()
        ttk.Label(dash_filter, text="Aralik").pack(side=tk.LEFT, padx=(0, 6))
        dash_range_combo = ttk.Combobox(
            dash_filter,
            textvariable=self.dash_range_var,
            values=["Son 7 Gun", "Son 30 Gun", "Bu Ay", "Ozel"],
            state="readonly",
            width=12,
        )
        dash_range_combo.pack(side=tk.LEFT)
        dash_range_combo.bind("<<ComboboxSelected>>", lambda _e: self._toggle_dash_range())
        dstart_frame, self.dash_start_entry = create_labeled_date(dash_filter, "Baslangic", self.dash_start_var, 12)
        dstart_frame.pack(side=tk.LEFT, padx=6)
        dend_frame, self.dash_end_entry = create_labeled_date(dash_filter, "Bitis", self.dash_end_var, 12)
        dend_frame.pack(side=tk.LEFT, padx=6)
        ttk.Button(dash_filter, text="Uygula", style="Accent.TButton", command=self.refresh_dashboard).pack(
            side=tk.LEFT, padx=6
        )
        clear_date_entry(self.dash_start_entry)
        clear_date_entry(self.dash_end_entry)
        self._toggle_dash_range()

        chart_frame = ttk.LabelFrame(content, text="Haftalik Trend (Calisilan Saat)", style="Section.TLabelframe")
        chart_frame.pack(fill=tk.X, padx=6, pady=6)
        self.dash_chart = tk.Canvas(chart_frame, height=120, bg=self._ui_theme["bg_content"], highlightthickness=0)
        self.dash_chart.pack(fill=tk.X, padx=8, pady=8)

        self.dash_chart_hint = ttk.Label(
            chart_frame, text="Son 7 gunde toplam calisilan saat", foreground="#5f6a72"
        )
        self.dash_chart_hint.pack(anchor="w", padx=8, pady=(0, 4))
        summary_wrap = tk.Frame(chart_frame, bg=self._ui_theme["bg_hover"])
        summary_wrap.pack(fill=tk.X, padx=8, pady=(0, 6))
        self.dash_chart_summary = tk.Label(
            summary_wrap,
            text="-",
            bg=self._ui_theme["bg_hover"],
            fg=self._ui_theme["text_primary"],
            font=("Segoe UI", 9, "bold"),
            padx=8,
            pady=4,
        )
        self.dash_chart_summary.pack(side=tk.LEFT)

        self.dash_progress = tk.Canvas(chart_frame, height=10, bg=self._ui_theme["bg_hover"], highlightthickness=0)
        self.dash_progress.pack(fill=tk.X, padx=8, pady=(0, 8))

        activity_frame = ttk.LabelFrame(content, text="Son Aktiviteler", style="Section.TLabelframe")
        activity_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        term_frame = tk.Frame(activity_frame, bg="#0b1118")
        term_frame.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)
        self.dash_activity = tk.Text(
            term_frame,
            height=10,
            wrap="word",
            bg="#0b1118",
            fg="#9FE870",
            insertbackground="#9FE870",
            font=("Consolas", 9),
        )
        yscroll = ttk.Scrollbar(term_frame, orient=tk.VERTICAL, command=self.dash_activity.yview)
        self.dash_activity.configure(yscrollcommand=yscroll.set)
        term_frame.columnconfigure(0, weight=1)
        term_frame.rowconfigure(0, weight=1)
        self.dash_activity.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        self.dash_activity.configure(state=tk.DISABLED)

    def _toggle_dash_range(self):
        if not hasattr(self, "dash_range_var"):
            return
        state = "normal" if self.dash_range_var.get() == "Ozel" else "disabled"
        if hasattr(self, "dash_start_entry"):
            self.dash_start_entry.configure(state=state)
        if hasattr(self, "dash_end_entry"):
            self.dash_end_entry.configure(state=state)

    def _render_progress(self, ratio):
        if not hasattr(self, "dash_progress"):
            return
        canvas = self.dash_progress
        canvas.delete("all")
        width = canvas.winfo_width() or 600
        height = canvas.winfo_height() or 10
        ratio = max(0.0, min(1.0, ratio))
        fill_w = int(width * ratio)
        canvas.create_rectangle(0, 0, width, height, fill=self._ui_theme["bg_hover"], width=0)
        canvas.create_rectangle(0, 0, fill_w, height, fill=self._ui_theme["primary"], width=0)

    def _render_weekly_chart(self, values, labels=None):
        if not hasattr(self, "dash_chart"):
            return
        canvas = self.dash_chart
        canvas.delete("all")
        if not values:
            return
        width = canvas.winfo_width() or 600
        height = canvas.winfo_height() or 120
        max_val = max(values) if max(values) > 0 else 1
        bar_width = max(20, int(width / max(len(values), 1)) - 8)
        spacing = 8
        x = spacing
        for i, v in enumerate(values):
            bar_h = int((v / max_val) * (height - 20))
            y0 = height - 10 - bar_h
            canvas.create_rectangle(x, y0, x + bar_width, height - 10, fill=self._ui_theme["primary"], width=0)
            canvas.create_text(
                x + bar_width / 2,
                max(8, y0 - 8),
                text=f"{v:.1f}",
                fill=self._ui_theme["text_secondary"],
                font=("Segoe UI", 8),
            )
            if labels and i < len(labels):
                canvas.create_text(
                    x + bar_width / 2,
                    height - 2,
                    text=labels[i],
                    fill=self._ui_theme["text_secondary"],
                    font=("Segoe UI", 7),
                    anchor="s",
                )
            x += bar_width + spacing

    def refresh_dashboard(self):
        if not hasattr(self, "dash_stats"):
            return
        today = datetime.now().date()
        range_type = self.dash_range_var.get() if hasattr(self, "dash_range_var") else "Son 7 Gun"
        if range_type == "Bu Ay":
            start_dt = today.replace(day=1)
            end_dt = today
        elif range_type == "Son 30 Gun":
            start_dt = today - timedelta(days=29)
            end_dt = today
        elif range_type == "Ozel":
            try:
                start_dt = datetime.strptime(self.dash_start_var.get().strip(), "%Y-%m-%d").date()
                end_dt = datetime.strptime(self.dash_end_var.get().strip(), "%Y-%m-%d").date()
            except Exception:
                start_dt = today - timedelta(days=6)
                end_dt = today
        else:
            start_dt = today - timedelta(days=6)
            end_dt = today
        if end_dt < start_dt:
            start_dt, end_dt = end_dt, start_dt

        start_str = start_dt.strftime("%Y-%m-%d")
        end_str = end_dt.strftime("%Y-%m-%d")

        employees = db.list_employees(region=self._view_region())
        records = db.list_timesheets(start_date=start_str, end_date=end_str, region=self._view_region())

        total_worked = 0.0
        total_overtime = 0.0
        total_scheduled = 0.0
        for (
            _ts_id,
            _emp_id,
            _name,
            department,
            work_date,
            start_time,
            end_time,
            break_minutes,
            is_special,
            _notes,
            _region,
        ) in records:
            try:
                worked, _scheduled, overtime, _night, _overnight, _s1, _s2, _s3 = calc.calc_day_hours(
                    work_date, start_time, end_time, break_minutes, self.settings, is_special, department
                )
                total_worked += float(worked)
                total_overtime += float(overtime)
                total_scheduled += float(_scheduled)
            except Exception:
                pass

        self._animate_stat(self.dash_stats["employees"], len(employees), decimals=0)
        self._animate_stat(self.dash_stats["timesheets"], len(records), decimals=0)
        self._animate_stat(self.dash_stats["worked"], total_worked, decimals=2)
        self._animate_stat(self.dash_stats["overtime"], total_overtime, decimals=2)
        self._animate_stat(self.dash_stats["scheduled"], total_scheduled, decimals=2)
        completion = (total_worked / total_scheduled * 100) if total_scheduled else 0.0
        self._animate_stat(self.dash_stats["completion"], completion, decimals=1)
        self._render_progress(completion / 100 if completion else 0.0)

        span_days = (end_dt - start_dt).days + 1
        chart_days = min(14, max(7, span_days))
        chart_start = end_dt - timedelta(days=chart_days - 1)
        chart_start_str = chart_start.strftime("%Y-%m-%d")
        recent_records = db.list_timesheets(start_date=chart_start_str, end_date=end_str, region=self._view_region())
        values = []
        labels = []
        for i in range(chart_days - 1, -1, -1):
            day = end_dt - timedelta(days=i)
            day_str = day.strftime("%Y-%m-%d")
            day_records = [r for r in recent_records if r[4] == day_str]
            day_total = 0.0
            for r in day_records:
                try:
                    worked, _scheduled, _o, _n, _ov, _s1, _s2, _s3 = calc.calc_day_hours(
                        r[4], r[5], r[6], r[7], self.settings, r[8], r[3]
                    )
                    day_total += float(worked)
                except Exception:
                    pass
            values.append(day_total)
            labels.append(day.strftime("%d/%m"))
        self._render_weekly_chart(values, labels)
        if hasattr(self, "dash_chart_hint"):
            self.dash_chart_hint.configure(text=f"{start_str} - {end_str} calisilan saat trendi")
        if hasattr(self, "dash_chart_summary"):
            current_total = sum(values)
            avg_val = current_total / len(values) if values else 0.0
            prev_total = 0.0
            prev_start = start_dt - timedelta(days=span_days)
            prev_end = start_dt - timedelta(days=1)
            prev_records = db.list_timesheets(
                start_date=prev_start.strftime("%Y-%m-%d"),
                end_date=prev_end.strftime("%Y-%m-%d"),
                region=self._view_region(),
            )
            for r in prev_records:
                try:
                    worked, _scheduled, _o, _n, _ov, _s1, _s2, _s3 = calc.calc_day_hours(
                        r[4], r[5], r[6], r[7], self.settings, r[8], r[3]
                    )
                    prev_total += float(worked)
                except Exception:
                    pass
            change_text = "-"
            if prev_total > 0:
                diff = ((current_total - prev_total) / prev_total) * 100
                change_text = f"%{diff:+.1f}"
            self.dash_chart_summary.configure(
                text=f"Toplam: {current_total:.1f}s | Ortalama: {avg_val:.1f}s | Onceki doneme gore: {change_text}"
            )

        if hasattr(self, "dash_activity"):
            lines = []
            try:
                if os.path.isfile(LOG_PATH):
                    with open(LOG_PATH, "r", encoding="utf-8") as handle:
                        lines = handle.readlines()[-12:]
            except Exception:
                lines = []
            self.dash_activity.configure(state=tk.NORMAL)
            self.dash_activity.delete("1.0", tk.END)
            if lines:
                self.dash_activity.insert(tk.END, "rainstaff@dashboard:~$ tail -n 12 log\n")
                self.dash_activity.insert(tk.END, "".join(lines))
            else:
                self.dash_activity.insert(tk.END, "Aktivite bulunamadi.\n")
            self.dash_activity.configure(state=tk.DISABLED)

    def _build_service_tab(self):
        fault_frame = ttk.LabelFrame(self.tab_service_body, text="Ariza Kaydi", style="Section.TLabelframe")
        fault_frame.pack(fill=tk.X, padx=6, pady=6)

        self.fault_id_var = tk.StringVar()
        self.fault_vehicle_var = tk.StringVar()
        self.fault_title_var = tk.StringVar()
        self.fault_desc_var = tk.StringVar()
        self.fault_open_var = tk.StringVar()
        self.fault_close_var = tk.StringVar()
        self.fault_status_var = tk.StringVar(value="Acik")

        frow1 = ttk.Frame(fault_frame)
        frow1.pack(fill=tk.X, pady=4)
        ttk.Label(frow1, text="Arac").pack(side=tk.LEFT, padx=(0, 6))
        self.fault_vehicle_combo = ttk.Combobox(
            frow1, textvariable=self.fault_vehicle_var, width=18, state="readonly"
        )
        self.fault_vehicle_combo.pack(side=tk.LEFT)
        create_labeled_entry(frow1, "Baslik", self.fault_title_var, 24).pack(side=tk.LEFT, padx=6)

        frow2 = ttk.Frame(fault_frame)
        frow2.pack(fill=tk.X, pady=4)
        create_labeled_entry(frow2, "Aciklama", self.fault_desc_var, 60).pack(side=tk.LEFT, padx=6)

        frow3 = ttk.Frame(fault_frame)
        frow3.pack(fill=tk.X, pady=4)
        open_frame, self.fault_open_entry = create_labeled_date(frow3, "Acilis", self.fault_open_var, 12)
        open_frame.pack(side=tk.LEFT, padx=6)
        close_frame, self.fault_close_entry = create_labeled_date(frow3, "Kapanis", self.fault_close_var, 12)
        close_frame.pack(side=tk.LEFT, padx=6)
        ttk.Label(frow3, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Combobox(
            frow3,
            textvariable=self.fault_status_var,
            values=["Acik", "Kapandi", "Takip"],
            width=10,
            state="readonly",
        ).pack(side=tk.LEFT)

        fbtn = ttk.Frame(fault_frame)
        fbtn.pack(fill=tk.X, pady=6)
        ttk.Button(fbtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_fault).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(fbtn, text="Sil", command=self.delete_fault).pack(side=tk.LEFT)
        ttk.Button(fbtn, text="Temizle", command=self.clear_fault_form).pack(side=tk.LEFT, padx=6)

        fault_list = ttk.Frame(self.tab_service_body)
        fault_list.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.fault_tree = ttk.Treeview(
            fault_list,
            columns=("id", "plate", "title", "status", "opened", "closed", "region"),
            show="headings",
            height=8,
        )
        self.fault_tree.heading("id", text="ID")
        self.fault_tree.heading("plate", text="Plaka")
        self.fault_tree.heading("title", text="Baslik")
        self.fault_tree.heading("status", text="Durum")
        self.fault_tree.heading("opened", text="Acilis")
        self.fault_tree.heading("closed", text="Kapanis")
        self.fault_tree.heading("region", text="Bolge")
        self.fault_tree.column("id", width=60, anchor=tk.CENTER)
        self.fault_tree.column("plate", width=100)
        self.fault_tree.column("title", width=220)
        self.fault_tree.column("status", width=100)
        self.fault_tree.column("opened", width=120)
        self.fault_tree.column("closed", width=120)
        self.fault_tree.column("region", width=100)
        f_xscroll = ttk.Scrollbar(fault_list, orient=tk.HORIZONTAL, command=self.fault_tree.xview)
        f_yscroll = ttk.Scrollbar(fault_list, orient=tk.VERTICAL, command=self.fault_tree.yview)
        self.fault_tree.configure(xscrollcommand=f_xscroll.set, yscrollcommand=f_yscroll.set)
        fault_list.columnconfigure(0, weight=1)
        fault_list.rowconfigure(0, weight=1)
        self.fault_tree.grid(row=0, column=0, sticky="nsew")
        f_yscroll.grid(row=0, column=1, sticky="ns")
        f_xscroll.grid(row=1, column=0, sticky="ew")
        self.fault_tree.bind("<<TreeviewSelect>>", self.on_fault_select)

        service_frame = ttk.LabelFrame(self.tab_service_body, text="Sanayi Kaydi", style="Section.TLabelframe")
        service_frame.pack(fill=tk.X, padx=6, pady=6)

        self.service_id_var = tk.StringVar()
        self.service_vehicle_var = tk.StringVar()
        self.service_fault_var = tk.StringVar()
        self.service_start_var = tk.StringVar()
        self.service_end_var = tk.StringVar()
        self.service_reason_var = tk.StringVar()
        self.service_cost_var = tk.StringVar()
        self.service_notes_var = tk.StringVar()
        self.service_in_shop_var = tk.BooleanVar(value=False)

        srow1 = ttk.Frame(service_frame)
        srow1.pack(fill=tk.X, pady=4)
        ttk.Label(srow1, text="Arac").pack(side=tk.LEFT, padx=(0, 6))
        self.service_vehicle_combo = ttk.Combobox(
            srow1, textvariable=self.service_vehicle_var, width=18, state="readonly"
        )
        self.service_vehicle_combo.pack(side=tk.LEFT)
        ttk.Label(srow1, text="Ariza").pack(side=tk.LEFT, padx=(12, 6))
        self.service_fault_combo = ttk.Combobox(
            srow1, textvariable=self.service_fault_var, width=40, state="readonly", values=[""]
        )
        self.service_fault_combo.pack(side=tk.LEFT)

        srow2 = ttk.Frame(service_frame)
        srow2.pack(fill=tk.X, pady=4)
        start_frame, self.service_start_entry = create_labeled_date(srow2, "Gidis", self.service_start_var, 12)
        start_frame.pack(side=tk.LEFT, padx=6)
        end_frame, self.service_end_entry = create_labeled_date(srow2, "Donus", self.service_end_var, 12)
        end_frame.pack(side=tk.LEFT, padx=6)
        ttk.Checkbutton(
            srow2,
            text="Sanayide",
            variable=self.service_in_shop_var,
            command=self._toggle_service_end_date,
        ).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(srow2, "Masraf", self.service_cost_var, 10).pack(side=tk.LEFT, padx=6)

        srow3 = ttk.Frame(service_frame)
        srow3.pack(fill=tk.X, pady=4)
        create_labeled_entry(srow3, "Neden", self.service_reason_var, 40).pack(side=tk.LEFT, padx=6)
        create_labeled_entry(srow3, "Not", self.service_notes_var, 40).pack(side=tk.LEFT, padx=6)

        sbtn = ttk.Frame(service_frame)
        sbtn.pack(fill=tk.X, pady=6)
        ttk.Button(sbtn, text="Kaydet", style="Accent.TButton", command=self.add_or_update_service_visit).pack(
            side=tk.LEFT, padx=6
        )
        ttk.Button(sbtn, text="Sil", command=self.delete_service_visit).pack(side=tk.LEFT)
        ttk.Button(sbtn, text="Temizle", command=self.clear_service_visit_form).pack(side=tk.LEFT, padx=6)

        service_list = ttk.Frame(self.tab_service_body)
        service_list.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.service_tree = ttk.Treeview(
            service_list,
            columns=("id", "plate", "fault", "start", "end", "cost", "reason", "region"),
            show="headings",
            height=8,
        )
        self.service_tree.heading("id", text="ID")
        self.service_tree.heading("plate", text="Plaka")
        self.service_tree.heading("fault", text="Ariza")
        self.service_tree.heading("start", text="Gidis")
        self.service_tree.heading("end", text="Donus")
        self.service_tree.heading("cost", text="Masraf")
        self.service_tree.heading("reason", text="Neden")
        self.service_tree.heading("region", text="Bolge")
        self.service_tree.column("id", width=60, anchor=tk.CENTER)
        self.service_tree.column("plate", width=100)
        self.service_tree.column("fault", width=220)
        self.service_tree.column("start", width=120)
        self.service_tree.column("end", width=120)
        self.service_tree.column("cost", width=90)
        self.service_tree.column("reason", width=180)
        self.service_tree.column("region", width=100)
        s_xscroll = ttk.Scrollbar(service_list, orient=tk.HORIZONTAL, command=self.service_tree.xview)
        s_yscroll = ttk.Scrollbar(service_list, orient=tk.VERTICAL, command=self.service_tree.yview)
        self.service_tree.configure(xscrollcommand=s_xscroll.set, yscrollcommand=s_yscroll.set)
        service_list.columnconfigure(0, weight=1)
        service_list.rowconfigure(0, weight=1)
        self.service_tree.grid(row=0, column=0, sticky="nsew")
        s_yscroll.grid(row=0, column=1, sticky="ns")
        s_xscroll.grid(row=1, column=0, sticky="ew")
        self.service_tree.bind("<<TreeviewSelect>>", self.on_service_visit_select)

    def on_vehicle_status_right_click(self, event):
        row_id = self.vehicle_status_tree.identify_row(event.y)
        if not row_id:
            return
        self.vehicle_status_tree.selection_set(row_id)
        try:
            self.vehicle_status_menu.tk_popup(event.x_root, event.y_root)
        finally:
            self.vehicle_status_menu.grab_release()

    def _open_vehicle_card_from_alert(self):
        """Open vehicle card from alert tree row double-click"""
        selected = self.vehicle_alert_tree.selection()
        if not selected:
            return
        values = self.vehicle_alert_tree.item(selected[0], "values")
        if not values or len(values) < 1:
            return
        plate = values[0]  # First column is plate
        self._open_vehicle_card(plate)

    def show_vehicle_detail(self):
        selected = self.vehicle_status_tree.selection()
        if not selected:
            return
        values = self.vehicle_status_tree.item(selected[0], "values")
        if not values:
            return
        plate = values[0]
        self._open_vehicle_card(plate)

    def show_vehicle_card_from_list(self):
        selected = self.vehicle_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Arac secin.")
            return
        values = self.vehicle_tree.item(selected[0], "values")
        if not values:
            return
        plate = values[1]
        self._open_vehicle_card(plate)

    def show_driver_detail_from_list(self):
        selected = self.driver_tree.selection()
        if not selected:
            messagebox.showwarning("Uyari", "Surucu secin.")
            return
        values = self.driver_tree.item(selected[0], "values")
        if not values:
            return
        driver_id = parse_int(values[0])
        self._open_driver_card(driver_id)

    def _open_vehicle_card(self, plate):
        vehicle_id = self.vehicle_map.get(plate)
        if not vehicle_id:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return
        vehicle = db.get_vehicle(vehicle_id)
        if not vehicle:
            messagebox.showwarning("Uyari", "Arac bulunamadi.")
            return

        (
            _vid,
            _plate,
            brand,
            model,
            year,
            km,
            inspection_date,
            insurance_date,
            maintenance_date,
            oil_change_date,
            oil_change_km,
            oil_interval_km,
            notes,
            _region,
        ) = vehicle

        detail_win = tk.Toplevel(self)
        detail_win.title(f"Arac Karti - {plate}")
        detail_win.geometry("980x700")

        content = self._make_window_scrollable(detail_win)

        info = ttk.LabelFrame(content, text="Arac Bilgisi", style="Section.TLabelframe")
        info.pack(fill=tk.X, padx=10, pady=8)
        info_row1 = ttk.Frame(info)
        info_row1.pack(fill=tk.X, pady=4)
        ttk.Label(info_row1, text=f"Plaka: {plate}").pack(side=tk.LEFT, padx=6)
        ttk.Label(info_row1, text=f"Marka/Model: {brand} {model}").pack(side=tk.LEFT, padx=12)
        ttk.Label(info_row1, text=f"Yil: {year}").pack(side=tk.LEFT, padx=12)
        info_row2 = ttk.Frame(info)
        info_row2.pack(fill=tk.X, pady=4)
        ttk.Label(info_row2, text=f"KM: {km or '-'}").pack(side=tk.LEFT, padx=6)
        ttk.Label(info_row2, text=f"Yag Degisim: {oil_change_date or '-'}").pack(side=tk.LEFT, padx=12)
        ttk.Label(info_row2, text=f"Yag KM: {oil_change_km or '-'}").pack(side=tk.LEFT, padx=12)
        interval_km = oil_interval_km or DEFAULT_OIL_INTERVAL_KM
        oil_status = "-"
        if interval_km and oil_change_km is not None and km is not None:
            remaining = interval_km - (km - oil_change_km)
            oil_status = "Geldi" if remaining <= 0 else f"{remaining} km"
        ttk.Label(info_row2, text=f"Periyot: {interval_km or '-'} km").pack(side=tk.LEFT, padx=12)
        ttk.Label(info_row2, text=f"Yag Durum: {oil_status}").pack(side=tk.LEFT, padx=12)
        info_row3 = ttk.Frame(info)
        info_row3.pack(fill=tk.X, pady=4)
        ttk.Label(info_row3, text=f"Muayene: {inspection_date or '-'}").pack(side=tk.LEFT, padx=6)
        ttk.Label(info_row3, text=f"Sigorta: {insurance_date or '-'}").pack(side=tk.LEFT, padx=12)
        ttk.Label(info_row3, text=f"Bakim: {maintenance_date or '-'}").pack(side=tk.LEFT, padx=12)
        if notes:
            info_row4 = ttk.Frame(info)
            info_row4.pack(fill=tk.X, pady=4)
            ttk.Label(info_row4, text=f"Not: {notes}").pack(side=tk.LEFT, padx=6)

        btn_row = ttk.Frame(content)
        btn_row.pack(fill=tk.X, padx=10, pady=4)
        ttk.Button(btn_row, text="Excel Arac Karti", command=lambda: self.export_vehicle_card(plate)).pack(
            side=tk.LEFT, padx=6
        )

        inspections = db.list_vehicle_inspections(vehicle_id=vehicle_id, region=self._view_region())
        if inspections:
            last_driver = inspections[0][4] or "-"
            last_date = inspections[0][5] or "-"
            info_row5 = ttk.Frame(info)
            info_row5.pack(fill=tk.X, pady=4)
            ttk.Label(info_row5, text=f"Son Surucu: {last_driver}").pack(side=tk.LEFT, padx=6)
            ttk.Label(info_row5, text=f"Son Kontrol: {last_date}").pack(side=tk.LEFT, padx=12)
        faults = db.list_vehicle_faults(vehicle_id=vehicle_id, region=self._view_region())
        services = db.list_vehicle_service_visits(vehicle_id=vehicle_id, region=self._view_region())

        inspect_frame = ttk.LabelFrame(content, text="Kontroller", style="Section.TLabelframe")
        inspect_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        inspect_tree = ttk.Treeview(
            inspect_frame,
            columns=("date", "week", "driver", "km", "fault", "status", "service", "note"),
            show="headings",
            height=6,
        )
        inspect_tree.heading("date", text="Tarih")
        inspect_tree.heading("week", text="Hafta")
        inspect_tree.heading("driver", text="Surucu")
        inspect_tree.heading("km", text="KM")
        inspect_tree.heading("fault", text="Ariza")
        inspect_tree.heading("status", text="Durum")
        inspect_tree.heading("service", text="Sanayi")
        inspect_tree.heading("note", text="Not")
        inspect_tree.column("date", width=110)
        inspect_tree.column("week", width=110)
        inspect_tree.column("driver", width=160)
        inspect_tree.column("km", width=70)
        inspect_tree.column("fault", width=160)
        inspect_tree.column("status", width=90)
        inspect_tree.column("service", width=80)
        inspect_tree.column("note", width=180)
        i_xscroll = ttk.Scrollbar(inspect_frame, orient=tk.HORIZONTAL, command=inspect_tree.xview)
        i_yscroll = ttk.Scrollbar(inspect_frame, orient=tk.VERTICAL, command=inspect_tree.yview)
        inspect_tree.configure(xscrollcommand=i_xscroll.set, yscrollcommand=i_yscroll.set)
        inspect_frame.columnconfigure(0, weight=1)
        inspect_frame.rowconfigure(0, weight=1)
        inspect_tree.grid(row=0, column=0, sticky="nsew")
        i_yscroll.grid(row=0, column=1, sticky="ns")
        i_xscroll.grid(row=1, column=0, sticky="ew")

        for row in inspections[:20]:
            (
                _iid,
                _veh_id,
                _plate,
                _driver_id,
                driver_name,
                inspect_date,
                week_start,
                km_val,
                note_val,
                fault_id,
                fault_status,
                service_visit,
            ) = row
            fault_title = ""
            if fault_id:
                fault = db.get_vehicle_fault(fault_id)
                if fault:
                    fault_title = fault[2] or ""
            inspect_tree.insert(
                "",
                tk.END,
                values=(
                    inspect_date,
                    week_start,
                    driver_name or "-",
                    km_val or "-",
                    fault_title,
                    fault_status or "",
                    "Evet" if service_visit else "",
                    note_val or "",
                ),
            )

        fault_frame = ttk.LabelFrame(content, text="Ariza Kayitlari", style="Section.TLabelframe")
        fault_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        fault_tree = ttk.Treeview(
            fault_frame,
            columns=("title", "status", "opened", "closed"),
            show="headings",
            height=5,
        )
        fault_tree.heading("title", text="Baslik")
        fault_tree.heading("status", text="Durum")
        fault_tree.heading("opened", text="Acilis")
        fault_tree.heading("closed", text="Kapanis")
        fault_tree.column("title", width=240)
        fault_tree.column("status", width=100)
        fault_tree.column("opened", width=120)
        fault_tree.column("closed", width=120)
        f_xscroll = ttk.Scrollbar(fault_frame, orient=tk.HORIZONTAL, command=fault_tree.xview)
        f_yscroll = ttk.Scrollbar(fault_frame, orient=tk.VERTICAL, command=fault_tree.yview)
        fault_tree.configure(xscrollcommand=f_xscroll.set, yscrollcommand=f_yscroll.set)
        fault_frame.columnconfigure(0, weight=1)
        fault_frame.rowconfigure(0, weight=1)
        fault_tree.grid(row=0, column=0, sticky="nsew")
        f_yscroll.grid(row=0, column=1, sticky="ns")
        f_xscroll.grid(row=1, column=0, sticky="ew")
        for fault in faults:
            _fid, _vid, _plate, title, _desc, opened_date, closed_date, status, _region = fault
            fault_tree.insert("", tk.END, values=(title, status, opened_date or "", closed_date or ""))

        service_frame = ttk.LabelFrame(content, text="Sanayi Kayitlari", style="Section.TLabelframe")
        service_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        service_tree = ttk.Treeview(
            service_frame,
            columns=("fault", "start", "end", "cost", "reason"),
            show="headings",
            height=5,
        )
        service_tree.heading("fault", text="Ariza")
        service_tree.heading("start", text="Gidis")
        service_tree.heading("end", text="Donus")
        service_tree.heading("cost", text="Masraf")
        service_tree.heading("reason", text="Neden")
        service_tree.column("fault", width=220)
        service_tree.column("start", width=120)
        service_tree.column("end", width=120)
        service_tree.column("cost", width=90)
        service_tree.column("reason", width=200)
        s_xscroll = ttk.Scrollbar(service_frame, orient=tk.HORIZONTAL, command=service_tree.xview)
        s_yscroll = ttk.Scrollbar(service_frame, orient=tk.VERTICAL, command=service_tree.yview)
        service_tree.configure(xscrollcommand=s_xscroll.set, yscrollcommand=s_yscroll.set)
        service_frame.columnconfigure(0, weight=1)
        service_frame.rowconfigure(0, weight=1)
        service_tree.grid(row=0, column=0, sticky="nsew")
        s_yscroll.grid(row=0, column=1, sticky="ns")
        s_xscroll.grid(row=1, column=0, sticky="ew")
        for visit in services:
            _sid, _vid, _plate, _fid, title, start_date, end_date, reason, cost, _notes, _region = visit
            end_value = end_date or "Sanayide"
            cost_value = f"{cost:.2f}" if cost is not None else ""
            service_tree.insert("", tk.END, values=(title or "", start_date, end_value, cost_value, reason or ""))

        compare_frame = ttk.LabelFrame(content, text="Haftalik Karsilastirma", style="Section.TLabelframe")
        compare_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        compare_tree = ttk.Treeview(
            compare_frame,
            columns=("item", "prev", "current", "change"),
            show="headings",
            height=6,
        )
        compare_tree.heading("item", text="Kontrol")
        compare_tree.heading("prev", text="Onceki Hafta")
        compare_tree.heading("current", text="Bu Hafta")
        compare_tree.heading("change", text="Durum")
        compare_tree.column("item", width=220)
        compare_tree.column("prev", width=120)
        compare_tree.column("current", width=120)
        compare_tree.column("change", width=160)
        c_xscroll = ttk.Scrollbar(compare_frame, orient=tk.HORIZONTAL, command=compare_tree.xview)
        c_yscroll = ttk.Scrollbar(compare_frame, orient=tk.VERTICAL, command=compare_tree.yview)
        compare_tree.configure(xscrollcommand=c_xscroll.set, yscrollcommand=c_yscroll.set)
        compare_frame.columnconfigure(0, weight=1)
        compare_frame.rowconfigure(0, weight=1)
        compare_tree.grid(row=0, column=0, sticky="nsew")
        c_yscroll.grid(row=0, column=1, sticky="ns")
        c_xscroll.grid(row=1, column=0, sticky="ew")

        if len(inspections) >= 2:
            current_inspection = inspections[0]
            previous_inspection = inspections[1]
            current_results = {
                row[0]: normalize_vehicle_status(row[1])
                for row in db.list_vehicle_inspection_results(current_inspection[0])
            }
            prev_results = {
                row[0]: normalize_vehicle_status(row[1])
                for row in db.list_vehicle_inspection_results(previous_inspection[0])
            }
            for item_key, label in VEHICLE_CHECKLIST:
                prev_status = prev_results.get(item_key, "-")
                curr_status = current_results.get(item_key, "-")
                if prev_status == curr_status:
                    change = "Ayni"
                elif prev_status == "Olumsuz" and curr_status == "Olumsuz":
                    change = "Tekrar eden sorun"
                elif curr_status == "Olumsuz" and prev_status != "Olumsuz":
                    change = "Kotulesti"
                elif prev_status == "Olumsuz" and curr_status != "Olumsuz":
                    change = "Iyilesti"
                else:
                    change = "Degisti"
                compare_tree.insert("", tk.END, values=(label, prev_status, curr_status, change))
            prev_fault_status = previous_inspection[10] or "-"
            curr_fault_status = current_inspection[10] or "-"
            if prev_fault_status or curr_fault_status:
                if prev_fault_status == curr_fault_status:
                    change = "Ayni"
                elif curr_fault_status == "Kapandi":
                    change = "Iyilesti"
                elif prev_fault_status == "Kapandi" and curr_fault_status != "Kapandi":
                    change = "Kotulesti"
                else:
                    change = "Degisti"
                compare_tree.insert("", tk.END, values=("Ariza Durumu", prev_fault_status, curr_fault_status, change))
            prev_service = "Evet" if previous_inspection[11] else "Hayir"
            curr_service = "Evet" if current_inspection[11] else "Hayir"
            change = "Ayni" if prev_service == curr_service else "Degisti"
            compare_tree.insert("", tk.END, values=("Sanayiye Gitti", prev_service, curr_service, change))

    def _open_driver_card(self, driver_id):
        driver = db.get_driver(driver_id)
        if not driver:
            messagebox.showwarning("Uyari", "Surucu bulunamadi.")
            return
        _did, name, license_class, license_expiry, phone, notes = driver
        inspections = db.list_driver_inspections(driver_id, region=self._view_region())

        detail_win = tk.Toplevel(self)
        detail_win.title(f"Surucu Karti - {name}")
        detail_win.geometry("960x680")

        content = self._make_window_scrollable(detail_win)

        info = ttk.LabelFrame(content, text="Surucu Bilgisi", style="Section.TLabelframe")
        info.pack(fill=tk.X, padx=10, pady=8)
        info_row1 = ttk.Frame(info)
        info_row1.pack(fill=tk.X, pady=4)
        ttk.Label(info_row1, text=f"Ad Soyad: {name}").pack(side=tk.LEFT, padx=6)
        ttk.Label(info_row1, text=f"Ehliyet: {license_class or '-'}").pack(side=tk.LEFT, padx=12)
        ttk.Label(info_row1, text=f"Bitis: {license_expiry or '-'}").pack(side=tk.LEFT, padx=12)
        info_row2 = ttk.Frame(info)
        info_row2.pack(fill=tk.X, pady=4)
        ttk.Label(info_row2, text=f"Telefon: {phone or '-'}").pack(side=tk.LEFT, padx=6)
        if notes:
            ttk.Label(info_row2, text=f"Not: {notes}").pack(side=tk.LEFT, padx=12)

        vehicle_frame = ttk.LabelFrame(content, text="Surucunun Araclari", style="Section.TLabelframe")
        vehicle_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        vehicle_tree = ttk.Treeview(
            vehicle_frame,
            columns=("plate", "last_date", "count"),
            show="headings",
            height=5,
        )
        vehicle_tree.heading("plate", text="Plaka")
        vehicle_tree.heading("last_date", text="Son Kontrol")
        vehicle_tree.heading("count", text="Kontrol Sayisi")
        vehicle_tree.column("plate", width=140)
        vehicle_tree.column("last_date", width=140)
        vehicle_tree.column("count", width=120, anchor=tk.CENTER)
        v_xscroll = ttk.Scrollbar(vehicle_frame, orient=tk.HORIZONTAL, command=vehicle_tree.xview)
        v_yscroll = ttk.Scrollbar(vehicle_frame, orient=tk.VERTICAL, command=vehicle_tree.yview)
        vehicle_tree.configure(xscrollcommand=v_xscroll.set, yscrollcommand=v_yscroll.set)
        vehicle_frame.columnconfigure(0, weight=1)
        vehicle_frame.rowconfigure(0, weight=1)
        vehicle_tree.grid(row=0, column=0, sticky="nsew")
        v_yscroll.grid(row=0, column=1, sticky="ns")
        v_xscroll.grid(row=1, column=0, sticky="ew")

        vehicle_summary = {}
        for row in inspections:
            plate = row[2]
            inspect_date = row[5]
            entry = vehicle_summary.setdefault(plate, {"last": inspect_date, "count": 0})
            entry["count"] += 1
            if inspect_date and (not entry["last"] or inspect_date > entry["last"]):
                entry["last"] = inspect_date
        for plate, info_row in sorted(vehicle_summary.items()):
            vehicle_tree.insert("", tk.END, values=(plate, info_row["last"] or "-", info_row["count"]))

        history_frame = ttk.LabelFrame(content, text="Kontrol Gecmisi", style="Section.TLabelframe")
        history_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=6)
        history_tree = ttk.Treeview(
            history_frame,
            columns=("date", "week", "plate", "km", "fault", "status", "service", "note"),
            show="headings",
            height=8,
        )
        history_tree.heading("date", text="Tarih")
        history_tree.heading("week", text="Hafta")
        history_tree.heading("plate", text="Plaka")
        history_tree.heading("km", text="KM")
        history_tree.heading("fault", text="Ariza")
        history_tree.heading("status", text="Durum")
        history_tree.heading("service", text="Sanayi")
        history_tree.heading("note", text="Not")
        history_tree.column("date", width=110)
        history_tree.column("week", width=110)
        history_tree.column("plate", width=120)
        history_tree.column("km", width=70)
        history_tree.column("fault", width=160)
        history_tree.column("status", width=90)
        history_tree.column("service", width=80)
        history_tree.column("note", width=200)
        h_xscroll = ttk.Scrollbar(history_frame, orient=tk.HORIZONTAL, command=history_tree.xview)
        h_yscroll = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=history_tree.yview)
        history_tree.configure(xscrollcommand=h_xscroll.set, yscrollcommand=h_yscroll.set)
        history_frame.columnconfigure(0, weight=1)
        history_frame.rowconfigure(0, weight=1)
        history_tree.grid(row=0, column=0, sticky="nsew")
        h_yscroll.grid(row=0, column=1, sticky="ns")
        h_xscroll.grid(row=1, column=0, sticky="ew")

        for row in inspections[:50]:
            (
                _iid,
                _veh_id,
                plate,
                _driver_id,
                _driver_name,
                inspect_date,
                week_start,
                km_val,
                note_val,
                fault_id,
                fault_status,
                service_visit,
            ) = row
            fault_title = ""
            if fault_id:
                fault = db.get_vehicle_fault(fault_id)
                if fault:
                    fault_title = fault[2] or ""
            history_tree.insert(
                "",
                tk.END,
                values=(
                    inspect_date,
                    week_start,
                    plate,
                    km_val or "-",
                    fault_title,
                    fault_status or "",
                    "Evet" if service_visit else "",
                    note_val or "",
                ),
            )

    def _build_admin_tab(self):
        content = self.tab_admin_body

        filter_frame = ttk.LabelFrame(content, text="Filtre", style="Section.TLabelframe")
        filter_frame.pack(fill=tk.X, padx=6, pady=6)

        self.admin_employee_var = tk.StringVar(value="Tum Calisanlar")
        self.admin_department_var = tk.StringVar(value="Tum Departmanlar")
        self.admin_title_var = tk.StringVar(value="Tum Unvanlar")
        self.admin_start_var = tk.StringVar()
        self.admin_end_var = tk.StringVar()
        self.admin_search_var = tk.StringVar()
        self.admin_month_filter_var = tk.StringVar()

        ttk.Label(filter_frame, text="Calisan").pack(side=tk.LEFT, padx=(0, 6))
        self.admin_employee_combo = ttk.Combobox(
            filter_frame, textvariable=self.admin_employee_var, width=24, state="readonly"
        )
        self.admin_employee_combo.pack(side=tk.LEFT)
        ttk.Label(filter_frame, text="Departman").pack(side=tk.LEFT, padx=(12, 6))
        self.admin_department_combo = ttk.Combobox(
            filter_frame, textvariable=self.admin_department_var, width=18, state="readonly"
        )
        self.admin_department_combo.pack(side=tk.LEFT)
        ttk.Label(filter_frame, text="Unvan").pack(side=tk.LEFT, padx=(12, 6))
        self.admin_title_combo = ttk.Combobox(filter_frame, textvariable=self.admin_title_var, width=18, state="readonly")
        self.admin_title_combo.pack(side=tk.LEFT)

        row2 = ttk.Frame(filter_frame)
        row2.pack(fill=tk.X, pady=6)
        start_frame, self.admin_start_entry = create_labeled_date(row2, "Baslangic", self.admin_start_var, 12)
        start_frame.pack(side=tk.LEFT, padx=6)
        end_frame, self.admin_end_entry = create_labeled_date(row2, "Bitis", self.admin_end_var, 12)
        end_frame.pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="Ay").pack(side=tk.LEFT, padx=(12, 6))
        month_entry = ttk.Entry(row2, textvariable=self.admin_month_filter_var, width=10)
        month_entry.pack(side=tk.LEFT)
        btn_month = ttk.Button(row2, text="Ayı Uygula", style="Accent.TButton", command=self.apply_admin_month_filter)
        btn_month.pack(side=tk.LEFT, padx=6)
        attach_tooltip(month_entry, "Ornek: 2026-02")
        attach_tooltip(btn_month, "Secilen ayin baslangic ve bitisini uygular")
        clear_date_entry(self.admin_start_entry)
        clear_date_entry(self.admin_end_entry)
        ttk.Label(row2, text="Ara").pack(side=tk.LEFT, padx=(12, 6))
        ttk.Entry(row2, textvariable=self.admin_search_var, width=24).pack(side=tk.LEFT)
        btn_refresh = ttk.Button(row2, text="Guncelle", style="Accent.TButton", command=self.refresh_admin_summary)
        btn_refresh.pack(side=tk.LEFT, padx=12)
        attach_tooltip(btn_refresh, "Filtreleri uygula ve ozetleri yenile")

        summary = ttk.LabelFrame(content, text="Ozet", style="Section.TLabelframe")
        summary.pack(fill=tk.X, padx=6, pady=6)

        self.admin_stats = {
            "total_records": tk.StringVar(value="0"),
            "total_employees": tk.StringVar(value="0"),
            "total_worked": tk.StringVar(value="0"),
            "total_overtime": tk.StringVar(value="0"),
            "total_night": tk.StringVar(value="0"),
            "total_overnight": tk.StringVar(value="0"),
            "total_special": tk.StringVar(value="0"),
            "avg_overtime": tk.StringVar(value="0"),
            "max_daily": tk.StringVar(value="0"),
        }

        row1 = ttk.Frame(summary)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Kayit").pack(side=tk.LEFT, padx=6)
        ttk.Label(row1, textvariable=self.admin_stats["total_records"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row1, text="Calisan").pack(side=tk.LEFT, padx=18)
        ttk.Label(row1, textvariable=self.admin_stats["total_employees"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row1, text="Toplam Calisilan").pack(side=tk.LEFT, padx=18)
        ttk.Label(row1, textvariable=self.admin_stats["total_worked"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row1, text="Toplam Fazla Mesai").pack(side=tk.LEFT, padx=18)
        ttk.Label(row1, textvariable=self.admin_stats["total_overtime"]).pack(side=tk.LEFT, padx=6)

        row2 = ttk.Frame(summary)
        row2.pack(fill=tk.X, pady=4)
        ttk.Label(row2, text="Toplam Gece").pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, textvariable=self.admin_stats["total_night"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="Geceye Tasan").pack(side=tk.LEFT, padx=18)
        ttk.Label(row2, textvariable=self.admin_stats["total_overnight"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="Ozel Gun").pack(side=tk.LEFT, padx=18)
        ttk.Label(row2, textvariable=self.admin_stats["total_special"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="Ortalama Fazla Mesai").pack(side=tk.LEFT, padx=18)
        ttk.Label(row2, textvariable=self.admin_stats["avg_overtime"]).pack(side=tk.LEFT, padx=6)
        ttk.Label(row2, text="En Yuksek Gun").pack(side=tk.LEFT, padx=18)
        ttk.Label(row2, textvariable=self.admin_stats["max_daily"]).pack(side=tk.LEFT, padx=6)

        pane = ttk.PanedWindow(content, orient=tk.HORIZONTAL)
        pane.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        list_frame = ttk.LabelFrame(pane, text="Calisan Ozeti", style="Section.TLabelframe")
        detail_frame = ttk.LabelFrame(pane, text="Detay", style="Section.TLabelframe")
        pane.add(list_frame, weight=4)
        pane.add(detail_frame, weight=2)

        self.admin_tree = ttk.Treeview(
            list_frame,
            columns=("employee", "worked", "overtime", "night", "overnight", "special"),
            show="headings",
        )
        self.admin_tree.heading("employee", text="Calisan")
        self.admin_tree.heading("worked", text="Calisilan")
        self.admin_tree.heading("overtime", text="Fazla Mesai")
        self.admin_tree.heading("night", text="Gece")
        self.admin_tree.heading("overnight", text="Geceye Tasan")
        self.admin_tree.heading("special", text="Ozel Gun")
        self.admin_tree.column("employee", width=220)
        self.admin_tree.column("worked", width=90)
        self.admin_tree.column("overtime", width=90)
        self.admin_tree.column("night", width=90)
        self.admin_tree.column("overnight", width=110)
        self.admin_tree.column("special", width=90)
        admin_xscroll = ttk.Scrollbar(list_frame, orient=tk.HORIZONTAL, command=self.admin_tree.xview)
        admin_yscroll = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.admin_tree.yview)
        self.admin_tree.configure(xscrollcommand=admin_xscroll.set, yscrollcommand=admin_yscroll.set)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        self.admin_tree.grid(row=0, column=0, sticky="nsew")
        admin_yscroll.grid(row=0, column=1, sticky="ns")
        admin_xscroll.grid(row=1, column=0, sticky="ew")
        self.admin_tree.bind("<Button-3>", self.on_admin_right_click)
        self.admin_tree.bind("<<TreeviewSelect>>", self.on_admin_select)
        self._apply_tree_zebra(self.admin_tree)
        self.admin_menu = tk.Menu(self, tearoff=0)
        self.admin_menu.add_command(label="Detay", command=self.show_admin_employee_detail)

        admin_action_row = ttk.Frame(list_frame)
        admin_action_row.grid(row=2, column=0, sticky="ew", pady=6)
        ttk.Button(
            admin_action_row,
            text="Secili Calisanlar Raporu",
            style="Accent.TButton",
            command=self.export_selected_admin_report,
        ).pack(side=tk.LEFT, padx=6)
        ttk.Button(
            admin_action_row,
            text="Seciliyi Kopyala",
            command=self.copy_selected_admin_rows,
        ).pack(side=tk.LEFT, padx=6)

        self.admin_detail_name = tk.StringVar(value="-")
        self.admin_detail_department = tk.StringVar(value="-")
        self.admin_detail_title = tk.StringVar(value="-")
        self.admin_detail_worked = tk.StringVar(value="-")
        self.admin_detail_overtime = tk.StringVar(value="-")
        self.admin_detail_night = tk.StringVar(value="-")
        self.admin_detail_overnight = tk.StringVar(value="-")
        self.admin_detail_special = tk.StringVar(value="-")
        self.admin_detail_avg = tk.StringVar(value="-")

        d1 = ttk.Frame(detail_frame)
        d1.pack(fill=tk.X, pady=4)
        ttk.Label(d1, text="Calisan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d1, textvariable=self.admin_detail_name).pack(side=tk.LEFT)
        d2 = ttk.Frame(detail_frame)
        d2.pack(fill=tk.X, pady=4)
        ttk.Label(d2, text="Departman").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d2, textvariable=self.admin_detail_department).pack(side=tk.LEFT)
        d3 = ttk.Frame(detail_frame)
        d3.pack(fill=tk.X, pady=4)
        ttk.Label(d3, text="Unvan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d3, textvariable=self.admin_detail_title).pack(side=tk.LEFT)
        d4 = ttk.Frame(detail_frame)
        d4.pack(fill=tk.X, pady=4)
        ttk.Label(d4, text="Calisilan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d4, textvariable=self.admin_detail_worked).pack(side=tk.LEFT)
        d5 = ttk.Frame(detail_frame)
        d5.pack(fill=tk.X, pady=4)
        ttk.Label(d5, text="Fazla Mesai").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d5, textvariable=self.admin_detail_overtime).pack(side=tk.LEFT)
        d6 = ttk.Frame(detail_frame)
        d6.pack(fill=tk.X, pady=4)
        ttk.Label(d6, text="Gece").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d6, textvariable=self.admin_detail_night).pack(side=tk.LEFT)
        d7 = ttk.Frame(detail_frame)
        d7.pack(fill=tk.X, pady=4)
        ttk.Label(d7, text="Geceye Tasan").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d7, textvariable=self.admin_detail_overnight).pack(side=tk.LEFT)
        d8 = ttk.Frame(detail_frame)
        d8.pack(fill=tk.X, pady=4)
        ttk.Label(d8, text="Ozel Gun").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d8, textvariable=self.admin_detail_special).pack(side=tk.LEFT)
        d9 = ttk.Frame(detail_frame)
        d9.pack(fill=tk.X, pady=4)
        ttk.Label(d9, text="Ortalama/Gun").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Label(d9, textvariable=self.admin_detail_avg).pack(side=tk.LEFT)

        alert_frame = ttk.LabelFrame(content, text="Uyarilar", style="Section.TLabelframe")
        alert_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.admin_alert_tree = ttk.Treeview(
            alert_frame,
            columns=("date", "employee", "issue", "value"),
            show="headings",
            height=6,
        )
        self.admin_alert_tree.heading("date", text="Tarih")
        self.admin_alert_tree.heading("employee", text="Calisan")
        self.admin_alert_tree.heading("issue", text="Uyari")
        self.admin_alert_tree.heading("value", text="Deger")
        self.admin_alert_tree.column("date", width=100)
        self.admin_alert_tree.column("employee", width=220)
        self.admin_alert_tree.column("issue", width=180)
        self.admin_alert_tree.column("value", width=100)
        alert_xscroll = ttk.Scrollbar(alert_frame, orient=tk.HORIZONTAL, command=self.admin_alert_tree.xview)
        alert_yscroll = ttk.Scrollbar(alert_frame, orient=tk.VERTICAL, command=self.admin_alert_tree.yview)
        self.admin_alert_tree.configure(xscrollcommand=alert_xscroll.set, yscrollcommand=alert_yscroll.set)
        alert_frame.columnconfigure(0, weight=1)
        alert_frame.rowconfigure(0, weight=1)
        self.admin_alert_tree.grid(row=0, column=0, sticky="nsew")
        self._apply_tree_zebra(self.admin_alert_tree)
        alert_yscroll.grid(row=0, column=1, sticky="ns")
        alert_xscroll.grid(row=1, column=0, sticky="ew")
        self.admin_alert_tree.bind("<<TreeviewSelect>>", self.on_admin_alert_select)
        self.admin_alert_detail = tk.StringVar(value="-")
        ttk.Label(alert_frame, textvariable=self.admin_alert_detail, foreground="#5f6a72").grid(
            row=2, column=0, sticky="w", padx=6, pady=(4, 0)
        )

        anomaly_frame = ttk.LabelFrame(content, text="Anomali Listesi", style="Section.TLabelframe")
        anomaly_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)
        self.admin_anomaly_tree = ttk.Treeview(
            anomaly_frame,
            columns=("employee", "period", "issue"),
            show="headings",
            height=6,
        )
        self.admin_anomaly_tree.heading("employee", text="Calisan")
        self.admin_anomaly_tree.heading("period", text="Donem")
        self.admin_anomaly_tree.heading("issue", text="Durum")
        self.admin_anomaly_tree.column("employee", width=220)
        self.admin_anomaly_tree.column("period", width=160)
        self.admin_anomaly_tree.column("issue", width=220)
        anom_xscroll = ttk.Scrollbar(anomaly_frame, orient=tk.HORIZONTAL, command=self.admin_anomaly_tree.xview)
        anom_yscroll = ttk.Scrollbar(anomaly_frame, orient=tk.VERTICAL, command=self.admin_anomaly_tree.yview)
        self.admin_anomaly_tree.configure(xscrollcommand=anom_xscroll.set, yscrollcommand=anom_yscroll.set)
        anomaly_frame.columnconfigure(0, weight=1)
        anomaly_frame.rowconfigure(0, weight=1)
        self.admin_anomaly_tree.grid(row=0, column=0, sticky="nsew")
        self._apply_tree_zebra(self.admin_anomaly_tree)
        anom_yscroll.grid(row=0, column=1, sticky="ns")
        anom_xscroll.grid(row=1, column=0, sticky="ew")
        self.admin_anomaly_tree.bind("<<TreeviewSelect>>", self.on_admin_anomaly_select)
        self.admin_anomaly_detail = tk.StringVar(value="-")
        ttk.Label(anomaly_frame, textvariable=self.admin_anomaly_detail, foreground="#5f6a72").grid(
            row=2, column=0, sticky="w", padx=6, pady=(4, 0)
        )

        pack_frame = ttk.LabelFrame(content, text="Rapor Paketleme", style="Section.TLabelframe")
        pack_frame.pack(fill=tk.X, padx=6, pady=6)
        ttk.Label(pack_frame, text="Ay (YYYY-MM)").pack(side=tk.LEFT, padx=6)
        self.admin_month_var = tk.StringVar()
        ttk.Entry(pack_frame, textvariable=self.admin_month_var, width=10).pack(side=tk.LEFT)
        ttk.Button(pack_frame, text="Raporlari Zip Yap", command=self.package_monthly_reports).pack(
            side=tk.LEFT, padx=6
        )

        self.refresh_admin_summary()

    def select_logo(self):
        path = filedialog.askopenfilename(
            filetypes=[("Images", "*.png;*.jpg;*.jpeg;*.bmp"), ("All", "*.*")]
        )
        if path:
            self.logo_path_var.set(path)

    def save_settings(self):
        prev_entry_region = self.settings.get("admin_entry_region", "Ankara")
        prev_view_region = self.settings.get("admin_view_region", "Tum Bolgeler")
        db.set_setting("company_name", self.company_name_var.get().strip())
        db.set_setting("report_title", self.report_title_var.get().strip())
        db.set_setting("weekday_hours", self.weekday_hours_var.get().strip())
        db.set_setting("saturday_start", self.sat_start_var.get().strip())
        db.set_setting("saturday_end", self.sat_end_var.get().strip())
        db.set_setting("logo_path", self.logo_path_var.get().strip())
        if self.is_admin:
            new_entry_region = self.admin_entry_region_var.get().strip() or "Ankara"
            new_view_region = self.admin_view_region_var.get().strip() or "Tum Bolgeler"
            db.set_setting("admin_entry_region", new_entry_region)
            db.set_setting("admin_view_region", new_view_region)
        self.settings = db.get_all_settings()
        if hasattr(self, "settings_stats"):
            try:
                user_count = len(db.list_users())
            except Exception:
                user_count = 0
            self._animate_stat(self.settings_stats["users"], user_count, decimals=0)
        if self.is_admin and prev_view_region != (self.admin_view_region_var.get().strip() or "Tum Bolgeler"):
            self._refresh_region_views()
        self._log_action("settings_save")
        messagebox.showinfo("Basarili", "Ayarlar kaydedildi.")

    def open_log_folder(self):
        try:
            os.makedirs(LOG_DIR, exist_ok=True)
            os.startfile(LOG_DIR)
        except Exception:
            messagebox.showerror("Hata", "Log klasoru acilamadi.")

    def open_log_file(self):
        try:
            if not os.path.isfile(LOG_PATH):
                with open(LOG_PATH, "a", encoding="utf-8"):
                    pass
            os.startfile(LOG_PATH)
        except Exception:
            messagebox.showerror("Hata", "Log dosyasi acilamadi.")

    def open_data_folder(self):
        try:
            os.makedirs(db.DB_DIR, exist_ok=True)
            os.startfile(db.DB_DIR)
        except Exception:
            messagebox.showerror("Hata", "Veri klasoru acilamadi.")

    def backup_database(self):
        try:
            default_name = f"puantaj_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.db"
            path = filedialog.asksaveasfilename(
                defaultextension=".db",
                initialfile=default_name,
                filetypes=[("SQLite DB", "*.db"), ("All Files", "*.*")],
            )
            if not path:
                return
            backup_path = db.create_backup(path)
            if self.logger:
                self.logger.info("Manual backup created: %s", backup_path)
            self._log_action("backup_create", f"file={os.path.basename(backup_path)}")
            messagebox.showinfo("Basarili", f"Yedek olusturuldu: {backup_path}")
        except Exception as exc:
            if self.logger:
                self.logger.exception("Backup failed")
            messagebox.showerror("Hata", f"Yedek alinamadi: {exc}")

    def restore_database(self):
        if not messagebox.askyesno(
            "Onay",
            "Bu islem mevcut veritabaniyi degistirecek. Devam etmek istiyor musun?",
        ):
            return
        try:
            path = filedialog.askopenfilename(
                filetypes=[("SQLite DB", "*.db"), ("All Files", "*.*")]
            )
            if not path:
                return
            db.restore_backup(path)
            if self.logger:
                self.logger.info("Database restored from: %s", path)
            self._log_action("backup_restore", f"file={os.path.basename(path)}")
            messagebox.showinfo("Basarili", "Yedek geri yuklendi. Uygulamayi yeniden baslatin.")
        except Exception as exc:
            if self.logger:
                self.logger.exception("Restore failed")
            messagebox.showerror("Hata", f"Geri yukleme basarisiz: {exc}")

    def export_data_zip(self):
        try:
            default_name = f"rainstaff_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
            path = filedialog.asksaveasfilename(
                defaultextension=".zip",
                initialfile=default_name,
                filetypes=[("ZIP", "*.zip"), ("All Files", "*.*")],
            )
            if not path:
                return
            export_path = db.export_data_zip(path)
            if self.logger:
                self.logger.info("Data exported: %s", export_path)
            self._log_action("data_export", f"file={os.path.basename(export_path)}")
            messagebox.showinfo("Basarili", f"Disari aktarildi: {export_path}")
        except Exception as exc:
            if self.logger:
                self.logger.exception("Export failed")
            messagebox.showerror("Hata", f"Disari aktarma basarisiz: {exc}")

    def import_data_zip(self):
        if not messagebox.askyesno(
            "Onay",
            "Bu islem mevcut veritabaniyi degistirecek. Devam etmek istiyor musun?",
        ):
            return
        try:
            path = filedialog.askopenfilename(
                filetypes=[("ZIP", "*.zip"), ("All Files", "*.*")]
            )
            if not path:
                return
            db.import_data_zip(path)
            if self.logger:
                self.logger.info("Data imported from: %s", path)
            self._log_action("data_import", f"file={os.path.basename(path)}")
            messagebox.showinfo("Basarili", "Iceri aktarma tamamlandi. Uygulamayi yeniden baslatin.")
        except Exception as exc:
            if self.logger:
                self.logger.exception("Import failed")
            messagebox.showerror("Hata", f"Iceri aktarma basarisiz: {exc}")

    def _build_stock_tab(self):
        """Build stock inventory management tab"""
        upload_frame = ttk.LabelFrame(self.tab_stock_body, text="Excel Yukle", style="Section.TLabelframe")
        upload_frame.pack(fill=tk.X, padx=6, pady=6)

        self.stock_file_var = tk.StringVar(value="Dosya secilmedi")
        self.stock_region_var = tk.StringVar(value="Ankara")
        self.stock_file_path = None

        row1 = ttk.Frame(upload_frame)
        row1.pack(fill=tk.X, pady=4)
        ttk.Label(row1, text="Dosya").pack(side=tk.LEFT, padx=(0, 8))
        ttk.Label(row1, textvariable=self.stock_file_var, foreground="#5B9BD5").pack(side=tk.LEFT, padx=6)
        ttk.Button(row1, text="Excel Sec", command=self.select_stock_file).pack(side=tk.LEFT, padx=6)

        row2 = ttk.Frame(upload_frame)
        row2.pack(fill=tk.X, pady=4)
        ttk.Label(row2, text="Bolge").pack(side=tk.LEFT, padx=(0, 8))
        region_combo = ttk.Combobox(row2, textvariable=self.stock_region_var, values=REGIONS, width=14, state="readonly")
        region_combo.pack(side=tk.LEFT, padx=6)
        ttk.Button(row2, text="Yukle", style="Accent.TButton", command=self.upload_stock_file).pack(side=tk.LEFT, padx=6)

        self.stock_status_var = tk.StringVar(value="")
        status_label = ttk.Label(upload_frame, textvariable=self.stock_status_var, foreground="#B0B0B0")
        status_label.pack(anchor="w", padx=6, pady=4)

        # Stock list
        list_frame = ttk.LabelFrame(self.tab_stock_body, text="Stok Envanteri", style="Section.TLabelframe")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=6, pady=6)

        filter_row = ttk.Frame(list_frame)
        filter_row.pack(fill=tk.X, pady=4)
        ttk.Label(filter_row, text="Bolge").pack(side=tk.LEFT, padx=6)
        self.stock_filter_bolge = tk.StringVar(value="ALL")
        bolge_filter = ttk.Combobox(filter_row, textvariable=self.stock_filter_bolge, 
                                     values=["ALL"] + REGIONS, width=12, state="readonly")
        bolge_filter.pack(side=tk.LEFT, padx=6)
        bolge_filter.bind("<<ComboboxSelected>>", lambda _: self.refresh_stock_list())

        ttk.Label(filter_row, text="Durum").pack(side=tk.LEFT, padx=(12, 6))
        self.stock_filter_durum = tk.StringVar(value="ALL")
        durum_filter = ttk.Combobox(filter_row, textvariable=self.stock_filter_durum,
                                     values=["ALL", "VAR", "YOK", "FAZLA"], width=10, state="readonly")
        durum_filter.pack(side=tk.LEFT, padx=6)
        durum_filter.bind("<<ComboboxSelected>>", lambda _: self.refresh_stock_list())

        ttk.Label(filter_row, text="Ara").pack(side=tk.LEFT, padx=(12, 6))
        self.stock_search_var = tk.StringVar()
        search_entry = ttk.Entry(filter_row, textvariable=self.stock_search_var, width=20)
        search_entry.pack(side=tk.LEFT, padx=6)
        search_entry.bind("<Return>", lambda _: self.refresh_stock_list())
        ttk.Button(filter_row, text="Filtrele", command=self.refresh_stock_list).pack(side=tk.LEFT, padx=6)

        # Treeview container
        tree_container = ttk.Frame(list_frame)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        # Create header frame for column titles - Parent info
        header_frame = tk.Frame(tree_container, bg="#252525", height=25)
        header_frame.grid(row=0, column=0, sticky="ew", columnspan=2)
        header_frame.grid_propagate(False)
        
        headers = ["Stok Kodu", "Ürün Adı", "Seri Sayısı"]
        widths = [120, 280, 100]
        
        for i, (header, width) in enumerate(zip(headers, widths)):
            lbl = tk.Label(header_frame, text=header, bg="#252525", fg="#FFD700", 
                          font=("Segoe UI", 10, "bold"), anchor="w", padx=5)
            lbl.pack(side=tk.LEFT, padx=5, pady=5)
        
        # Treeview columns - 3 column layout
        columns = ("stok_adi", "seri_sayisi")
        self.stock_tree = ttk.Treeview(tree_container, columns=columns, show="tree headings", height=20)

        self.stock_tree.heading("#0", text="Stok Kodu")
        self.stock_tree.heading("stok_adi", text="Ürün Adı")
        self.stock_tree.heading("seri_sayisi", text="Seri Sayısı")
        
        self.stock_tree.column("#0", width=120, anchor="w")
        self.stock_tree.column("stok_adi", width=280, anchor="w")
        self.stock_tree.column("seri_sayisi", width=100, anchor="center")

        self.stock_tree.tag_configure("parent", background="#252525", foreground="#FFD700", font=("Segoe UI", 10, "bold"))
        self.stock_tree.tag_configure("child", background="#1f1f1f", foreground="#e0e0e0", font=("Segoe UI", 9))

        stock_xscroll = ttk.Scrollbar(tree_container, orient=tk.HORIZONTAL, command=self.stock_tree.xview)
        stock_yscroll = ttk.Scrollbar(tree_container, orient=tk.VERTICAL, command=self.stock_tree.yview)
        self.stock_tree.configure(xscrollcommand=stock_xscroll.set, yscrollcommand=stock_yscroll.set)

        tree_container.columnconfigure(0, weight=1)
        tree_container.rowconfigure(1, weight=1)
        self.stock_tree.grid(row=1, column=0, sticky="nsew")
        stock_yscroll.grid(row=1, column=1, sticky="ns")
        stock_xscroll.grid(row=2, column=0, sticky="ew")
        
        # Click handler for expand/collapse
        self.stock_tree.bind("<Button-1>", self._on_stock_tree_click)

    def select_stock_file(self):
        """Select Excel file for stock upload"""
        path = filedialog.askopenfilename(
            filetypes=[("Excel", "*.xlsx;*.xls"), ("XLSX", "*.xlsx"), ("XLS", "*.xls"), ("All", "*.*")]
        )
        if path:
            self.stock_file_path = path
            filename = os.path.basename(path)
            self.stock_file_var.set(filename)
            self.stock_status_var.set("")

    def upload_stock_file(self):
        """Upload stock Excel file to server"""
        if not self.stock_file_path:
            messagebox.showwarning("Uyari", "Once dosya secin.")
            return

        if not os.path.isfile(self.stock_file_path):
            messagebox.showwarning("Uyari", "Dosya bulunamadi.")
            return

        bolge = self.stock_region_var.get().strip()
        if not bolge:
            messagebox.showwarning("Uyari", "Bolge secin.")
            return

        self.stock_status_var.set("Dosya isleniyor...")
        self.after(100, self._stock_upload_worker, self.stock_file_path, bolge)

    def _stock_upload_worker(self, file_path, bolge):
        """Process stock file upload in background"""
        try:
            # Read Excel locally first
            rows = load_tabular_file(file_path)
            if not rows:
                self.stock_status_var.set("Dosyada veri bulunamadi!")
                messagebox.showwarning("Uyari", "Dosyada veri bulunamadi.")
                return

            # Check if first row is empty or has no valid headers
            first_row_empty = all(cell is None or str(cell).strip() == '' for cell in rows[0])
            
            # Also check if headers are valid (contain 'stok' or 'seri' keywords)
            headers = [str(h).strip().lower() if h else '' for h in rows[0]]
            has_valid_headers = any('stok' in h or 'seri' in h for h in headers)
            
            if first_row_empty or not has_valid_headers:
                # No headers - use default column indices and start from row 0
                if self.logger:
                    self.logger.info("Stock upload: No header row detected, using default indices, starting from row 0")
                stok_kod_idx = 0
                stok_adi_idx = 1
                seri_no_idx = 2
                seri_sayi_idx = 3
                start_row = 0  # NO HEADERS - start from row 0 (data starts immediately)
            else:
                # Parse headers (flexible)
                if self.logger:
                    self.logger.info("Stock upload: Valid headers found, parsing headers")
                
                stok_kod_idx = next((i for i, h in enumerate(headers) if 'stok' in h and 'kod' in h), 0)
                stok_adi_idx = next((i for i, h in enumerate(headers) if 'stok' in h and ('adi' in h or 'ad' in h)), 1)
                seri_no_idx = next((i for i, h in enumerate(headers) if 'seri' in h and 'no' in h), 2)
                seri_sayi_idx = next((i for i, h in enumerate(headers) if 'seri' in h and 'say' in h), 3)
                start_row = 1  # Skip header row

            # Parse nested Excel structure
            # Format: Stok header row followed by seri_no child rows (with empty stok_kod)
            imported = 0
            with db.get_conn() as conn:
                cursor = conn.cursor()
                cursor.execute("DELETE FROM stock_inventory WHERE bolge = ?", (bolge,))

                i = start_row
                while i < len(rows):
                    row = rows[i]

                    # Check if this is a product header (stok_kod not empty)
                    stok_kod = str(row[stok_kod_idx]).strip() if stok_kod_idx < len(row) and row[stok_kod_idx] else ''

                    if stok_kod and stok_kod not in ['', 'nan', 'None', 'None']:
                        # This is a product header
                        stok_adi = str(row[stok_adi_idx]).strip() if stok_adi_idx < len(row) and row[stok_adi_idx] else ''
                        seri_sayi = 0
                        try:
                            seri_sayi = int(row[seri_sayi_idx]) if seri_sayi_idx < len(row) and row[seri_sayi_idx] else 0
                        except (ValueError, TypeError):
                            pass

                        # Collect all following seri_no rows (child rows where stok_kod is empty)
                        i += 1
                        seri_count = 0
                        while i < len(rows):
                            child_row = rows[i]
                            child_stok_kod = str(child_row[stok_kod_idx]).strip() if stok_kod_idx < len(child_row) and child_row[stok_kod_idx] else ''

                            # If stok_kod is empty/None, this is a seri_no row
                            child_stok_kod_value = child_row[stok_kod_idx] if stok_kod_idx < len(child_row) else None
                            
                            if child_stok_kod_value is None or not child_stok_kod or child_stok_kod in ['', 'nan', 'None']:
                                try:
                                    # Seri no is in seri_no column (column 2) for child rows
                                    seri_no_value = child_row[seri_no_idx] if seri_no_idx < len(child_row) else None
                                    seri_no = str(seri_no_value).strip() if seri_no_value is not None else ''

                                    # Extract actual serial number (remove numbering like "1 ST87088" or "1. ST87088")
                                    if seri_no:
                                        parts = seri_no.split(maxsplit=1)
                                        # Fix: Handle both "1" and "1." prefixes
                                        if len(parts) == 2 and parts[0].replace('.', '', 1).isdigit():
                                            seri_no = parts[1]
                                    
                                    # Skip if seri_no is empty (but allow pure numbers as valid serials)
                                    if seri_no and seri_no not in ['', 'nan', 'None', 'None']:
                                        cursor.execute(
                                            """INSERT INTO stock_inventory
                                               (stok_kod, stok_adi, seri_no, durum, tarih, girdi_yapan, bolge, adet, updated_at)
                                               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)""",
                                            (stok_kod, stok_adi, seri_no, "OK", datetime.now().strftime("%Y-%m-%d"),
                                             "system", bolge, 1, datetime.now().isoformat())
                                        )
                                        imported += 1
                                        seri_count += 1

                                    i += 1
                                except Exception as e:
                                    if self.logger:
                                        self.logger.debug(f"Stock seri row error: {e}")
                                    i += 1
                                    break
                            else:
                                # Next product header found
                                break
                    else:
                        i += 1

            self.stock_status_var.set(f"✓ {imported} kayit yuklendi ({bolge})")
            self._log_action("stock_upload", f"file={os.path.basename(file_path)} region={bolge} count={imported}")
            
            # Refresh view
            self.refresh_stock_list()
            messagebox.showinfo("Basarili", f"{imported} stok kaydı yüklendi.")
            
            # Trigger sync
            self.trigger_sync("stock_upload")

        except Exception as e:
            self.stock_status_var.set(f"✗ Hata: {str(e)[:50]}")
            if self.logger:
                self.logger.error(f"Stock upload error: {e}")
            messagebox.showerror("Hata", f"Yukleme basarisiz: {str(e)[:100]}")

    def refresh_stock_list(self):
        """Refresh stock inventory list"""
        if not hasattr(self, "stock_tree"):
            return

        for item in self.stock_tree.get_children():
            self.stock_tree.delete(item)

        try:
            with db.get_conn() as conn:
                cursor = conn.cursor()

                query = "SELECT * FROM stock_inventory WHERE 1=1"
                params = []

                bolge_filter = self.stock_filter_bolge.get()
                if bolge_filter and bolge_filter != "ALL":
                    query += " AND bolge = ?"
                    params.append(bolge_filter)

                durum_filter = self.stock_filter_durum.get()
                if durum_filter and durum_filter != "ALL":
                    query += " AND durum = ?"
                    params.append(durum_filter)

                search = self.stock_search_var.get().strip()
                if search:
                    query += " AND (stok_kod LIKE ? OR stok_adi LIKE ? OR seri_no LIKE ?)"
                    search_term = f"%{search}%"
                    params.extend([search_term, search_term, search_term])

                query += " ORDER BY stok_kod, seri_no"
                cursor.execute(query, params)
                rows = cursor.fetchall()

            # Group by stok_kod
            grouped = {}
            for row in rows:
                stok_kod, stok_adi, seri_no, durum, tarih, girdi_yapan, bolge, adet, *_ = row
                if stok_kod not in grouped:
                    grouped[stok_kod] = {
                        'stok_adi': stok_adi,
                        'items': []
                    }
                grouped[stok_kod]['items'].append({
                    'seri_no': seri_no,
                    'durum': durum,
                    'tarih': tarih,
                    'girdi_yapan': girdi_yapan,
                    'bolge': bolge,
                    'adet': adet
                })

            # Insert hierarchical list - parent headers with children
            for stok_kod, data in grouped.items():
                # Parent row: stok_kod | stok_adi | seri_sayisi
                parent_id = self.stock_tree.insert("", tk.END, text=stok_kod,
                    values=(
                        data['stok_adi'] or "",
                        f"{len(data['items'])} seri"
                    ),
                    tags=("parent",),
                    open=False
                )
                
                # Child rows: ONLY seri_no (shown in #0 column via text)
                for item in data['items']:
                    self.stock_tree.insert(parent_id, tk.END, text=item['seri_no'] or "",
                        values=("", ""),
                        tags=("child",)
                    )

        except Exception as e:
            if self.logger:
                self.logger.error(f"Stock list refresh error: {e}")

    def _on_stock_tree_click(self, event):
        """Handle treeview click for expand/collapse"""
        item = self.stock_tree.identify('item', event.x, event.y)
        if not item:
            return
        
        # Check if this is a parent item (has children)
        children = self.stock_tree.get_children(item)
        if children:
            # Toggle open state
            current_open = self.stock_tree.item(item, 'open')
            self.stock_tree.item(item, open=not current_open)


if __name__ == "__main__":
    ensure_app_dirs()
    db.init_db()
    app = PuantajApp()
    app.mainloop()
