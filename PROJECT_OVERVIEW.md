# Rainstaff Puantaj Sistemi - Proje Genel Bakış

Bu dosya tüm sistemi tanımak için tek referans noktasıdır. Gelecekteki Claude oturumları
bu dökümana bakarak projenin mevcut durumunu hızlıca kavrayabilir.

Son taranma tarihi: 2026-05-21
Kod tabanı toplamı: ~12.700 satır Python + ~6.700 satır HTML/CSS

## 1. Sistem Mimarisi (Üst Seviye)

İki bileşenli hibrit mimari:

```
                        ┌─────────────────────┐
                        │  Render.com (Free)  │
                        │  Flask Dashboard    │
                        │  SQLite (/data)     │
                        └──────────▲──────────┘
                                   │
                  HTTP POST /sync  │  HTTP GET /sync/download
                  multipart .db    │  binary .db
                                   │
        ┌──────────────────────────┴──────────────────────────┐
        │                                                     │
   ┌────┴───────┐                                       ┌─────┴──────┐
   │ Desktop PC │                                       │ Desktop PC │
   │  (Tkinter) │                                       │  (Tkinter) │
   │  SQLite    │                                       │  SQLite    │
   │  %APPDATA% │                                       │  %APPDATA% │
   └────────────┘                                       └────────────┘
```

- **Desktop (Tkinter)** veri girişi yapar, lokal SQLite'a yazar, ardından sunucuya tüm DB
  dosyasını yükler.
- **Server (Flask)** read-only dashboard sunar; gelen DB dosyasını mevcut master DB ile
  `_merge_databases()` ile birleştirir (deleted_records tablosu ile silme yayılımı).
- Aynı DB dosyası (puantaj.db) hem desktop hem server tarafında kullanılır; iki tarafta
  da aynı `puantaj_db.py` modülü çalışır (server tarafı kopyası vardır, neredeyse birebir
  aynı).

Önemli: README.md "sync_enabled=1 ile aktif" derken, mevcut kodda
`server/app.py:20`'de `SYNC_ENABLED = False`. Yani şu an sunucu tarafında /sync endpoint'i
410 Gone dönüyor. Desktop tarafında `trigger_sync()` da no-op (app.py:1551). Bulut senkron
aktif değil.

## 2. Dizin Yapısı

```
puantaj/
├── puantaj_app/                  # ANA UYGULAMA KODU
│   ├── app.py                    # 8249 satır - Tkinter UI + tüm iş mantığı
│   ├── puantaj_db.py             # 1265 satır - SQLite şema + CRUD
│   ├── calc.py                   # 135 satır - Saat/mesai hesaplaması
│   ├── report.py                 # 985 satır - Excel/PDF rapor üretimi
│   ├── requirements.txt          # Desktop bağımlılıkları
│   ├── Rainstaff.spec            # PyInstaller spec (exe build)
│   ├── assets/                   # Logo, ascii.png
│   ├── data/                     # Lokal DB klasörü (dev fallback)
│   ├── build/, dist/             # PyInstaller çıktıları
│   ├── server/                   # FLASK WEB SUNUCUSU
│   │   ├── app.py                # 1006 satır - Flask routes + sync
│   │   ├── puantaj_db.py         # 1053 satır - desktop DB modülünün kopyası
│   │   ├── calc.py               # desktop calc'ın kopyası
│   │   ├── requirements.txt      # Sunucu bağımlılıkları
│   │   ├── render.yaml           # Render.com deploy config
│   │   ├── templates/            # 14 HTML template
│   │   │   ├── base.html         # Layout + 3 tema CSS (Matrix/Cyber/Sabah)
│   │   │   ├── modern_dashboard.html  # 2613 satır - ana panel
│   │   │   ├── login.html
│   │   │   ├── stock.html, vehicles.html, drivers.html,
│   │   │   ├── vehicle_faults.html, alerts.html, reports.html,
│   │   │   └── dashboard.html (eski), 404/500/error.html
│   │   └── static/
│   │       ├── img/, matrix-rain.js
│   └── (60+ utility / migration script: add_*, fix_*, test_*)
│
├── README.md                     # Proje notları (Ocak 2026 itibariyle güncel)
├── PROJECT_OVERVIEW.md           # BU DOSYA
└── (30+ ad-hoc analiz/log/MD dosyası: PROGRESS_LOG_*, STATUS_*, FIX_*, vs.)
```

## 3. Veritabanı Şeması

SQLite, dosya konumu:
- Windows: `%APPDATA%\Rainstaff\data\puantaj.db`
- Render: `/data/puantaj.db` (persistent disk) veya `/tmp/rainstaff_data/puantaj.db`
  (fallback - restart'ta silinir).

`init_db()` (puantaj_db.py:98) idempotent, açılışta her zaman çağrılır. ALTER TABLE ile
in-place migration yapan üç fonksiyon var:
`_ensure_timesheet_columns`, `_ensure_vehicle_columns`, `_ensure_region_columns`.

### Tablolar (14 tablo)

| Tablo                          | Amaç                                                  |
|--------------------------------|-------------------------------------------------------|
| `employees`                    | Çalışan kartları (TC, departman, ünvan, bölge)        |
| `timesheets`                   | Günlük puantaj girişleri (FK -> employees)            |
| `attendance_records`           | Günlük yoklama (Calisti/Izinli/Gelmedi/Raporlu/...)   |
| `leave_records`                | İzin kayıtları (tip, doküman, durum)                  |
| `leave_status_history`         | İzin durumu değişiklik geçmişi                        |
| `shift_templates`              | Vardiya şablonları (varsayılan 2 kayıt)               |
| `vehicles`                     | Araç kartları (muayene/sigorta/bakım/yağ tarihleri)   |
| `drivers`                      | Sürücüler (ehliyet sınıfı/bitiş)                      |
| `vehicle_faults`               | Araç arızaları (Acik/Kapali)                          |
| `vehicle_inspections`          | Haftalık kontroller (driver_id, fault_id linki var)   |
| `vehicle_inspection_results`   | Checklist sonuçları (9 madde)                         |
| `vehicle_service_visits`       | Servis ziyaretleri (sanayi)                           |
| `stock_inventory`              | Excel ile yüklenen stok (seri_no UNIQUE)              |
| `users`                        | Login (SHA-256 password_hash, role, region)           |
| `settings`                     | Key/value sistem ayarları                             |
| `reports`                      | Üretilen rapor arşivi (file_path)                     |
| `deleted_records`              | Silme yayılımı için (table_name, record_id, deleted_at) |

### Varsayılan Kullanıcılar (puantaj_db.py:61)

```python
DEFAULT_USERS = [
    ("ankara1",   "060106", "user",  "Ankara"),
    ("izmir1",    "350235", "user",  "Izmir"),
    ("bursa1",    "160316", "user",  "Bursa"),
    ("istanbul1", "340434", "user",  "Istanbul"),
    ("admin",     "748774", "admin", "ALL"),
]
```

Parolalar SHA-256 hash'leniyor (`hash_password()`, salt yok - güvenlik açığı, kırılabilir).

### Varsayılan Ayarlar (`DEFAULT_SETTINGS`)

```python
weekday_hours: "8"           # Standart hafta içi saat
saturday_start: "09:00"
saturday_end:   "14:00"
sync_enabled:   "0"          # Mevcut kodda zaten False, ayar etkisiz
admin_entry_region:  "Ankara"
admin_view_region:   "Tum Bolgeler"
```

## 4. Saat Hesaplama Kuralları (calc.py)

`calc_day_hours(work_date, start, end, break_min, settings, is_special, department)` →
`(worked, scheduled, overtime, night, overnight, special_normal, special_overtime, special_night)`

- **Çalışılan saat:** `(end - start) - break_minutes/60`. End < start ise +1 gün
  (gece vardiyası desteklenir).
- **Plan saat:** Pazartesi-Cuma `weekday_hours`, Cumartesi `saturday_end - saturday_start`,
  Pazar 0.
- **Fazla mesai:** `worked - scheduled` (negatifse 0).
- **Gece saati:** 22:00-06:00 aralığına denk gelen kısım (overlap hesaplaması).
- **Geceye taşan:** 24:00'ü aşan kısım.
- **Özel gün (is_special=1):** Tüm saat `special_normal`'a yazılır, plan ve fazla mesai 0.
- **Pazar (STANT departmanı hariç):** Fazla mesaiye eklenmez, ayrı takip edilir
  (`calc_sunday_separate_hours` - app.py:186, report.py:25).
- **STANT departmanı:** Cumartesi/Pazar planı 8 saat olarak sayılır (perakende stant
  operasyonu için).

## 5. Desktop Uygulama (app.py)

### Açılış Akışı
1. `PuantajApp.__init__` (app.py:565) Tk root oluşturur, logger ve log kuyruğu kurar.
2. `_login_prompt()` (app.py:803) modal login - başarısızsa uygulama kapanır.
3. `_show_loading("Yukleniyor...")` overlay.
4. `_finish_startup()` -> `_startup_step_style` -> `_startup_step_ui` ->
   `_startup_step_data` (asenkron, 10ms aralıklarla).
5. `_build_ui()` (app.py:1168) tüm sekmeleri kurar.

### UI Sekmeleri (Notebook + nav buttons)

| Sekme           | Build fonksiyonu               | Ne yapar                                          |
|-----------------|--------------------------------|---------------------------------------------------|
| Dashboard       | `_build_dashboard_tab` (6588)  | KPI kartları, haftalık grafik, son işlemler     |
| Puantaj         | `_build_timesheets_tab` (2014) | Günlük puantaj girişi/listesi, vardiya uygula    |
| İzin/Yoklama    | `_build_attendance_tab` (2786) | Yoklama statüleri + izin kayıtları + doküman    |
| Çalışanlar      | `_build_employees_tab` (1667)  | Personel CRUD, son puantajları gösterir          |
| Raporlar        | `_build_reports_tab` (4028)    | Excel rapor üretimi, arşiv listesi               |
| Yönetim (admin) | `_build_admin_tab` (7506)      | Yönetici özeti, anomali tespiti, paket export    |
| Araçlar         | `_build_vehicles_tab` (6318)   | Araç/sürücü/arıza/servis (filo)                  |
| Servis          | `_build_service_tab` (6883)    | Sanayi ziyaretleri, haftalık araç kontrol        |
| Stok            | `_build_stock_tab` (7905)      | Excel'den stok yükle, ağaç (parent/child) gösterim |
| Ayarlar         | `_build_settings_tab` (6119)   | Şirket bilgisi, çalışma saatleri, vardiya şablon, yedek/restore |
| Loglar          | `_build_logs_tab` (6273)       | Canlı log akışı                                  |

### Üç Tema (`self.themes`)
- **Gece** (varsayılan, koyu mavi/gri)
- **Sabah** (açık tema)
- **Matrix** (koyu mavi varyant)

Tema header'daki butondan döngüsel olarak değişir, settings tablosunda persist edilir
(`db.set_setting("theme", ...)`).

### Önemli Mantık
- **Bölge bazlı görünüm:** Admin "Tum Bolgeler" görebilir; admin dışı sadece kendi bölgesi.
  `_view_region()` (app.py:933) ve `_entry_region()` (app.py:949) bunu yönetir.
- **Klavye kısayolları:** Ctrl+F (arama), Ctrl+N (yeni kayıt), Ctrl+S (kaydet).
- **Keepalive:** `_keepalive_worker` (app.py:781) her 5 dk'da sunucuyu /health pingler
  (Render free tier uyumasın diye). Çalışsa da SYNC_ENABLED=False olduğu için anlamsız.
- **Tüm CRUD `_log_action()` ile loglanır** (LOG_PATH = `%APPDATA%/Rainstaff/logs/rainstaff.log`).
- **Silme yayılımı:** `delete_employee/timesheet/vehicle/...` her silmede `deleted_records`
  tablosuna kayıt atar. Sunucu merge sırasında bunları kullanır (kullanılırdı).

### Excel Import
- Çalışan import: `import_employees` (app.py:3889), header alias eşleme
  (`EMP_HEADER_ALIASES`).
- Puantaj import: `import_timesheets` (app.py:3936), `TS_HEADER_ALIASES`.
- Stok import: `_stock_upload_worker` (app.py:8032), başlıksız Excel'i de destekler
  (heuristik: ilk satır boşsa varsayılan kolon indexleri).

## 6. Sunucu (server/app.py)

### Endpoints

| Endpoint              | Method | Auth     | Ne yapar                                |
|-----------------------|--------|----------|-----------------------------------------|
| `/health`             | GET    | public   | `{"status":"healthy", ...}` JSON       |
| `/auto-sync`          | GET/POST| public  | SYNC_ENABLED=False, 410 dönüyor         |
| `/diagnostic-final`   | GET    | public   | DEBUG endpoint, parola hash dump (!)    |
| `/sync`               | POST   | public   | DB upload + merge (SYNC_ENABLED=False)  |
| `/sync/download`      | GET    | public   | Merged DB dosyası iner                  |
| `/sync/reset`         | POST   | reset key| DB'yi siler ('rainstaff2026reset' key)  |
| `/login`              | POST   | public   | Session login                           |
| `/logout`             | GET    | -        | Session clear                           |
| `/`                   | GET    | -        | login veya dashboard'a yönlendirme      |
| `/dashboard`          | GET    | session  | modern_dashboard.html (KPI cards)       |
| `/alerts`             | GET    | session  | alerts.html                             |
| `/reports`            | GET    | session  | reports.html                            |
| `/stock`              | GET    | session  | stock.html                              |
| `/vehicles`           | GET    | session  | VEHICLE_MODULE_ENABLED=False -> 404     |
| `/drivers`            | GET    | session  | aynı şekilde devre dışı                 |
| `/vehicle-faults`     | GET    | session  | aynı şekilde devre dışı                 |
| `/api/employee-timesheets/<emp_id>` | GET | session | Çalışan puantajı JSON, calc.py kullanır |
| `/api/timesheets`     | GET    | session  | Tüm puantajlar JSON                     |
| `/api/employee-overtime` | GET | session | Çalışan + toplam fazla mesai            |
| `/api/vehicles`       | GET    | session  | VEHICLE_MODULE_ENABLED=False -> 410     |
| `/api/drivers`        | GET    | session  | aynı                                    |
| `/api/vehicle-faults` | GET    | session  | aynı                                    |
| `/api/stock-data`     | GET    | session  | stok envanteri grouped by stok_kod      |

### Güvenlik Notları (kritik)
- `app.secret_key` çevre değişkeninden gelir, **fallback default** vardır:
  `'dev-secret-key-change-in-production'`. Production'da SECRET_KEY env tanımlanmazsa
  herkesin tahmin edebileceği session secret kullanılır.
- `/diagnostic-final` parola hash'lerini AÇIK ŞEKİLDE döndürüyor. Public endpoint.
  Kaldırılmalı.
- `/sync/reset` reset key kod içinde sabit (`'rainstaff2026reset'`). Public endpoint.
- `/sync` ve `/sync/download` hiçbir token doğrulaması yapmıyor (eski README'de X-API-KEY
  vardı ama kod onu kontrol etmiyor). SYNC_ENABLED=False olduğu için şu an sömürülemez
  ama açılırsa risk.

### Merge Mantığı (`_merge_databases`, app.py:200)
1. Incoming DB'den `deleted_records` çek, master'da uygula.
2. Master'ın kendi `deleted_records`'ını da topla -> `all_deleted` set.
3. Tablo tablo (timesheets, employees, vehicles, drivers) `INSERT OR REPLACE` ile
   incoming'i master'a kopyala. Silinmiş kayıtları atla.
4. `stock_inventory` özel - **full replace** (önce master tamamen silinir, sonra incoming
   yazılır). Stok için silme yayılımı yerine "son yükleme kazanır" stratejisi.

Eksik: `attendance_records`, `leave_records`, `shift_templates`, `vehicle_faults`,
`vehicle_inspections`, `vehicle_service_visits`, `reports`, `users`, `settings` merge
edilmiyor. Yalnız 4 tablo + stok merge oluyor.

## 7. Rapor Üretimi (report.py)

5 farklı rapor fonksiyonu:

1. `export_report()` (line 132) - Aylık/dönemsel puantaj Excel raporu. 17 kolon, başlık,
   logo, toplam satırı.
2. `export_report_pdf()` (line 616) - reportlab ile PDF (opsiyonel import, hata yutar).
3. `export_leave_form()` (line 477) - İzin formu (PDF).
4. `export_vehicle_weekly_report()` (line 698) - Haftalık araç kontrol raporu.
5. `export_vehicle_card_report()` (line 854) - Tek araç kartı (geçmiş ziyaretler/arızalar).

`unpack_timesheet_record` (line 36) iki farklı tuple format'ını destekler (11-uzunluk yeni,
10-uzunluk eski) - geriye dönük şema desteği var.

## 8. Deployment (Render)

`server/render.yaml`:
- Tip: `web`, runtime `python`, region `frankfurt`, plan `free`.
- Build: `pip install -r requirements.txt`
- Start: `gunicorn app:app`
- ENV vars: `PYTHON_VERSION=3.11.0`, `DATABASE_URL` (rainstaff-db connection string),
  `SECRET_KEY` (generated).
- Bir PostgreSQL DB tanımlı (`rainstaff-db`) ama kod **PostgreSQL kullanmıyor** - hala
  SQLite. `DATABASE_URL` env değişkeni okunmuyor. `psycopg2-binary==2.9.9`
  `requirements.txt`'te ama dead dependency.
- Persistent disk yok render.yaml içinde. `puantaj_db.py:26` `/data` var mı diye bakıyor;
  yoksa `/tmp/rainstaff_data` (her restart'ta sıfırlanır).

### Bağımlılıklar
**Desktop** (`puantaj_app/requirements.txt`):
```
openpyxl==3.1.5
Pillow>=11.0.0
tkcalendar==1.6.1
requests==2.32.3
reportlab>=4.0.0
```

**Sunucu** (`puantaj_app/server/requirements.txt`):
```
Flask==3.0.0
Werkzeug==3.0.1
gunicorn==21.2.0
openpyxl==3.1.5
requests==2.32.3
psycopg2-binary==2.9.9   # KULLANILMIYOR
python-dotenv==1.0.0
```

## 9. Bilinen Sorunlar / Borçlar

### Kritik (satışa engel)
1. **Bulut senkron kapalı** (`SYNC_ENABLED=False`, `trigger_sync()` no-op). Çok lokasyonlu
   pazarlama vaadi için aktive edilmeli.
2. **Render free tier persistent disk yok** -> `/tmp` kullanılıyor, restart'ta DB sıfırlanır.
3. **Tek-kiracılı tasarım.** Birden fazla müşteri = birden fazla deploy. Multi-tenant değil.
4. **Güvenlik açıkları:**
   - Hardcoded SECRET_KEY fallback
   - SHA-256 (salt yok) parola hash
   - Public `/diagnostic-final` parola hash dump
   - Hardcoded `/sync/reset` key
   - `/sync` endpoint'inde token doğrulaması yok (kod yorum satırı ile devre dışı)
5. **Mobil yok.** Dashboard responsive ama veri girişi sadece Tkinter desktop'tan.
6. **KVKK/audit log eksik.** `_log_action` dosyaya yazıyor ama "kim ne zaman neyi
   değiştirdi" formatında değil (yalnız "ACTION + user").
7. **Default kullanıcı parolaları kodda açık** (`DEFAULT_USERS`). İlk kurulumda mecburen
   değiştirilmiyor.

### Orta
8. **Merge eksik:** attendance, leave, faults, inspections, service_visits, settings sync
   edilmiyor. Tek bir lokasyonun verisi diğerine aktarılmaz.
9. **DB path heuristik:** Windows'ta %APPDATA%, Linux'ta /data veya /tmp - belirsizlik.
10. **app.py 8249 satır** tek dosya - test edilemez, refactor zor.
11. **Çok sayıda ad-hoc script** (`add_*.py`, `fix_*.py`, 30+ MD log dosyası) kod tabanını
    kirletiyor.
12. **`modern_dashboard.html`** 2613 satır, baştan `78906o` gibi bozuk karakter ile
    başlıyor (line 1 - muhtemelen düzeltme gerekecek).
13. **VEHICLE_MODULE_ENABLED=False** sunucuda - araç sayfaları desktop'ta var, web'de
    devre dışı.
14. **psycopg2-binary** requirements'ta ama kullanılmıyor - kaldırılmalı veya PostgreSQL'e
    geçilmeli.

### Küçük
15. Reportlab opsiyonel - PDF üretimi sessizce başarısız olabilir.
16. Birden fazla yedek dosya (`app.py.BACKUP_2026_01_18`, `puantaj_db.py.BROKEN_BACKUP`)
    repo içinde.
17. Token/sync_token settings'te tutuluyor ama kontrol edilmiyor.

## 10. Geliştirme Notları (Gelecek Oturumlar İçin)

### Çalıştırma
- **Desktop:** `cd puantaj_app && python app.py` (Windows). Tkinter + tkcalendar
  gerekir.
- **Sunucu (lokal):** `cd puantaj_app/server && python app.py` -> `http://localhost:5000`.
- **PyInstaller build:** `Rainstaff.spec` kullanılıyor (build/, dist/ klasörleri var).

### Test Dosyaları
puantaj_app içinde `test_*.py` dosyaları var ama bağımsız komut dosyaları, pytest fixture
yok. Bunlar ad-hoc smoke test'ler.

### Modül İlişkisi
- `puantaj_app/app.py` -> `puantaj_db` (lokal modül), `calc`, `report`, `tkcalendar`,
  `openpyxl`, `requests`.
- `puantaj_app/server/app.py` -> üst dizinin `puantaj_db.py`'ını `sys.path` hilesiyle
  import ediyor ama kendi içinde de bir kopyası var (server/puantaj_db.py). Hangisi
  yüklendiği `__file__` ile sürpriz olabilir.

### Düzenleme Yaparken Dikkat
- `puantaj_db.py` desktop ve server arasında **kopyalanmış** - bir tarafta yapılan şema
  değişikliği diğerine taşınmalı.
- `app.py` 8000+ satır, edit yaparken bağlamı (sınıf metodu mu, top-level mi) doğrula.
- Türkçe karakter desteği var ama dosya yolları ASCII tutuluyor (logging'de gözüküyor).
- Stok merge full-replace (atomic) - bunu değiştirirken silme yayılımı ekle.

### Çalışma Tarzı (kullanıcının tercihi)
- Read existing files before writing. Don't re-read unless changed.
- Thorough in reasoning, concise in output.
- No emojis or em-dashes (kod tabanında bazı emojiler var - tema ikonları için).
- Code first, explanation after.

## 11. Kullanıcı Hedefi

Sahibi (canorman0640@gmail.com) bu sistemi şirketlere SaaS veya one-time olarak
pazarlamak istiyor. Hedef segment: çok lokasyonlu temizlik/güvenlik/lojistik/taşeron
firmaları (50-300 personel, filolu).

Tahmini pazarlama fiyatı (mevcut hali):
- One-time: 15.000-40.000 TL (KOBİ)
- SaaS: 1.500-3.500 TL/ay
- Mobil + bordro entegrasyonu eklenirse: 5.000-9.000 TL/ay, 75.000-150.000 TL one-time.

Satışa hazırlık için yol haritası ayrı görüşmede tartışıldı (şirket kuruluşu, KVKK,
multi-tenant, mobil, landing page, ilk pilot müşteriler).
