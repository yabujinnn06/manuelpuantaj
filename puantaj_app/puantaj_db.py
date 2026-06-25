"""
Rainstaff Database Module
SQLite database operations for employee management, timesheets, vehicles, etc.
"""

import os
import sqlite3
import shutil
import zipfile
import hashlib
from datetime import datetime, timedelta
from contextlib import contextmanager

# ============================================================================
# CONFIGURATION
# ============================================================================

APP_NAME = "Rainstaff"
LOCAL_DB_DIR = os.path.join(os.path.dirname(__file__), "data")

if os.name == 'nt':  # Windows
    APPDATA_DIR = os.path.join(os.environ.get("APPDATA", LOCAL_DB_DIR), APP_NAME, "data")
    DB_DIR = APPDATA_DIR
else:  # Linux / Server (Render)
    # Check for Render persistent disk at /data
    if os.path.exists("/data"):
        DB_DIR = "/data"
    else:
        # Fallback to /tmp if /data (persistent disk) is not attached
        DB_DIR = "/tmp/rainstaff_data"

# Ensure DB directory exists
if not os.path.exists(DB_DIR):
    try:
        os.makedirs(DB_DIR, exist_ok=True)
    except OSError:
        # Last resort fallback
        import tempfile
        DB_DIR = os.path.join(tempfile.gettempdir(), "rainstaff_data")
        os.makedirs(DB_DIR, exist_ok=True)

DB_PATH = os.path.join(DB_DIR, "puantaj.db")
BACKUP_DIR = os.path.join(DB_DIR, "backups")
BACKUP_MARKER = os.path.join(BACKUP_DIR, "last_backup.txt")
EXPORT_DIR = os.path.join(DB_DIR, "exports")

DEFAULT_SETTINGS = {
    "company_name": "",
    "report_title": "Rainstaff Puantaj ve Mesai Raporu",
    "weekday_hours": "8",
    "saturday_start": "09:00",
    "saturday_end": "14:00",
    "logo_path": "",
    "sync_enabled": "0",
    "sync_url": "",
    "sync_token": "",
    "admin_entry_region": "Ankara",
    "admin_view_region": "Tum Bolgeler",
}

DEFAULT_USERS = [
    ("ankara1", "060106", "user", "Ankara"),
    ("izmir1", "350235", "user", "Izmir"),
    ("bursa1", "160316", "user", "Bursa"),
    ("istanbul1", "340434", "user", "Istanbul"),
    ("admin", "748774", "admin", "ALL"),
]

# ============================================================================
# DATABASE CONNECTION
# ============================================================================

def ensure_db_dir():
    """Ensure database directory exists"""
    if not os.path.isdir(DB_DIR):
        os.makedirs(DB_DIR, exist_ok=True)

@contextmanager
def get_conn():
    """SQLite connection with automatic commit and rollback"""
    ensure_db_dir()
    conn = sqlite3.connect(DB_PATH, timeout=30.0)
    try:
        conn.execute("PRAGMA foreign_keys = ON;")
        conn.execute("PRAGMA busy_timeout = 30000;")
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()

# ============================================================================
# DATABASE INITIALIZATION
# ============================================================================

def init_db():
    """Initialize database schema"""
    with get_conn() as conn:
        # Employees table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                identity_no TEXT,
                department TEXT,
                title TEXT,
                region TEXT
            );
        """)
        
        # Timesheets table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS timesheets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                work_date TEXT NOT NULL,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                break_minutes INTEGER NOT NULL DEFAULT 0,
                is_special INTEGER NOT NULL DEFAULT 0,
                notes TEXT,
                region TEXT,
                FOREIGN KEY (employee_id) REFERENCES employees (id) ON DELETE CASCADE
            );
        """)
        
        # Settings table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS settings (
                key TEXT PRIMARY KEY,
                value TEXT NOT NULL
            );
        """)
        
        # Reports table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS reports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                file_path TEXT NOT NULL,
                created_at TEXT NOT NULL,
                employee TEXT,
                start_date TEXT,
                end_date TEXT
            );
        """)
        
        # Shift templates table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS shift_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE,
                start_time TEXT NOT NULL,
                end_time TEXT NOT NULL,
                break_minutes INTEGER NOT NULL DEFAULT 0
            );
        """)

        # Attendance records table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS attendance_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                work_date TEXT NOT NULL,
                status TEXT NOT NULL,
                reason TEXT,
                source TEXT,
                region TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(employee_id, work_date),
                FOREIGN KEY (employee_id) REFERENCES employees (id) ON DELETE CASCADE
            );
        """)

        # Leave records table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS leave_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER NOT NULL,
                start_date TEXT NOT NULL,
                end_date TEXT NOT NULL,
                leave_type TEXT,
                reason TEXT,
                document_no TEXT,
                document_path TEXT,
                status TEXT DEFAULT 'Onayli',
                region TEXT,
                created_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (employee_id) REFERENCES employees (id) ON DELETE CASCADE
            );
        """)

        # Leave status history table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS leave_status_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                leave_id INTEGER NOT NULL,
                status TEXT NOT NULL,
                note TEXT,
                changed_by TEXT,
                changed_at TEXT DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (leave_id) REFERENCES leave_records (id) ON DELETE CASCADE
            );
        """)
        
        # Vehicles table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicles (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plate TEXT NOT NULL UNIQUE,
                brand TEXT,
                model TEXT,
                year TEXT,
                km INTEGER,
                inspection_date TEXT,
                insurance_date TEXT,
                maintenance_date TEXT,
                oil_change_date TEXT,
                oil_change_km INTEGER,
                oil_interval_km INTEGER,
                notes TEXT,
                region TEXT
            );
        """)
        
        # Drivers table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS drivers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                full_name TEXT NOT NULL,
                license_class TEXT,
                license_expiry TEXT,
                phone TEXT,
                notes TEXT,
                region TEXT
            );
        """)
        
        # Users table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT NOT NULL UNIQUE,
                password_hash TEXT NOT NULL,
                role TEXT NOT NULL,
                region TEXT NOT NULL
            );
        """)
        
        # Vehicle faults table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicle_faults (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vehicle_id INTEGER NOT NULL,
                title TEXT NOT NULL,
                description TEXT,
                opened_date TEXT,
                closed_date TEXT,
                status TEXT DEFAULT 'Acik',
                region TEXT,
                FOREIGN KEY (vehicle_id) REFERENCES vehicles (id) ON DELETE CASCADE
            );
        """)
        
        # Vehicle inspections table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicle_inspections (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vehicle_id INTEGER NOT NULL,
                driver_id INTEGER,
                inspection_date TEXT NOT NULL,
                week_start TEXT NOT NULL,
                km INTEGER,
                notes TEXT,
                fault_id INTEGER,
                fault_status TEXT,
                service_visit INTEGER DEFAULT 0,
                FOREIGN KEY (vehicle_id) REFERENCES vehicles (id) ON DELETE CASCADE,
                FOREIGN KEY (driver_id) REFERENCES drivers (id) ON DELETE SET NULL,
                FOREIGN KEY (fault_id) REFERENCES vehicle_faults (id) ON DELETE SET NULL
            );
        """)
        
        # Vehicle service visits table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicle_service_visits (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                vehicle_id INTEGER NOT NULL,
                fault_id INTEGER,
                start_date TEXT NOT NULL,
                end_date TEXT,
                reason TEXT,
                cost REAL,
                notes TEXT,
                region TEXT,
                FOREIGN KEY (vehicle_id) REFERENCES vehicles (id) ON DELETE CASCADE,
                FOREIGN KEY (fault_id) REFERENCES vehicle_faults (id) ON DELETE SET NULL
            );
        """)
        
        # Vehicle inspection results table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS vehicle_inspection_results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                inspection_id INTEGER NOT NULL,
                item_key TEXT NOT NULL,
                status TEXT NOT NULL,
                note TEXT,
                FOREIGN KEY (inspection_id) REFERENCES vehicle_inspections (id) ON DELETE CASCADE
            );
        """)
        
        # Stock inventory table
        conn.execute("""
            CREATE TABLE IF NOT EXISTS stock_inventory (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                stok_kod TEXT,
                stok_adi TEXT,
                seri_no TEXT NOT NULL UNIQUE,
                durum TEXT,
                tarih TEXT,
                girdi_yapan TEXT,
                bolge TEXT NOT NULL,
                adet INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
        """)
        
        # Deleted records tracking table (for multi-PC sync)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS deleted_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                table_name TEXT NOT NULL,
                record_id INTEGER NOT NULL,
                deleted_at TEXT NOT NULL,
                deleted_by TEXT
            );
        """)
        
        # Insert default settings
        for key, value in DEFAULT_SETTINGS.items():
            conn.execute(
                "INSERT OR IGNORE INTO settings (key, value) VALUES (?, ?);",
                (key, value)
            )
        
        # Insert default shift templates
        cur = conn.execute("SELECT COUNT(*) FROM shift_templates;")
        if cur.fetchone()[0] == 0:
            conn.executemany(
                "INSERT INTO shift_templates (name, start_time, end_time, break_minutes) VALUES (?, ?, ?, ?);",
                [
                    ("Hafta Ici 09-18", "09:00", "18:00", 60),
                    ("Cumartesi 09-14", "09:00", "14:00", 0),
                ]
            )
        
        # Ensure schema is up to date
        _ensure_timesheet_columns(conn)
        _ensure_vehicle_columns(conn)
        _ensure_region_columns(conn)
        _ensure_deleted_records_table(conn)
        _seed_default_users(conn)
        
        conn.commit()
    
    _backup_db_if_needed()

def _ensure_timesheet_columns(conn):
    """Ensure timesheets table has region column"""
    try:
        cursor = conn.execute("PRAGMA table_info(timesheets)")
        columns = [row[1] for row in cursor.fetchall()]
        if 'region' not in columns:
            conn.execute("ALTER TABLE timesheets ADD COLUMN region TEXT;")
    except Exception:
        pass

def _ensure_vehicle_columns(conn):
    """Ensure vehicles table has region column"""
    try:
        cursor = conn.execute("PRAGMA table_info(vehicles)")
        columns = [row[1] for row in cursor.fetchall()]
        if 'region' not in columns:
            conn.execute("ALTER TABLE vehicles ADD COLUMN region TEXT;")
    except Exception:
        pass

def _ensure_region_columns(conn):
    """Ensure all tables have region columns"""
    tables_needing_region = [
        'employees', 'drivers', 'vehicle_faults', 'vehicle_service_visits'
    ]
    for table in tables_needing_region:
        try:
            cursor = conn.execute(f"PRAGMA table_info({table})")
            columns = [row[1] for row in cursor.fetchall()]
            if 'region' not in columns:
                conn.execute(f"ALTER TABLE {table} ADD COLUMN region TEXT;")
        except Exception:
            pass

def _ensure_deleted_records_table(conn):
    """Ensure deleted_records table exists"""
    try:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS deleted_records (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                table_name TEXT NOT NULL,
                record_id INTEGER NOT NULL,
                deleted_at TEXT NOT NULL,
                deleted_by TEXT
            );
        """)
    except Exception:
        pass

def _seed_default_users(conn):
    """Seed default users if users table is empty"""
    try:
        cursor = conn.execute("SELECT COUNT(*) FROM users;")
        if cursor.fetchone()[0] == 0:
            for username, password, role, region in DEFAULT_USERS:
                password_hash = hash_password(password)
                conn.execute(
                    "INSERT INTO users (username, password_hash, role, region) VALUES (?, ?, ?, ?);",
                    (username, password_hash, role, region)
                )
    except Exception:
        pass

# ============================================================================
# USER AUTHENTICATION
# ============================================================================

def hash_password(password):
    """Hash password using SHA-256"""
    return hashlib.sha256(password.encode()).hexdigest()

def verify_password(password, password_hash):
    """Verify password against hash"""
    return hash_password(password) == password_hash

def verify_user(username, password):
    """Verify user credentials and return user dict if valid"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, username, password_hash, role, region FROM users WHERE username = ?;",
            (username,)
        )
        row = cursor.fetchone()
        if row and verify_password(password, row[2]):
            return {
                "id": row[0],
                "username": row[1],
                "password_hash": row[2],
                "role": row[3],
                "region": row[4]
            }
    return None

def get_user(username):
    """Get user by username"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, username, password_hash, role, region FROM users WHERE username = ?;",
            (username,)
        )
        row = cursor.fetchone()
        if row:
            return {
                "id": row[0],
                "username": row[1],
                "password_hash": row[2],
                "role": row[3],
                "region": row[4]
            }
    return None

def list_users():
    """List all users"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, username, role, region FROM users ORDER BY username;"
        )
        return cursor.fetchall()

# ============================================================================
# SETTINGS
# ============================================================================

def get_all_settings():
    """Get all settings as a dictionary"""
    with get_conn() as conn:
        cursor = conn.execute("SELECT key, value FROM settings;")
        return {row[0]: row[1] for row in cursor.fetchall()}

def set_setting(key, value):
    """Set a setting value"""
    with get_conn() as conn:
        conn.execute(
            "INSERT OR REPLACE INTO settings (key, value) VALUES (?, ?);",
            (key, value)
        )

# ============================================================================
# EMPLOYEES
# ============================================================================

def list_employees(region=None):
    """List all employees, optionally filtered by region"""
    with get_conn() as conn:
        if region:
            cursor = conn.execute(
                "SELECT id, full_name, identity_no, department, title, region FROM employees WHERE region = ? ORDER BY full_name;",
                (region,)
            )
        else:
            cursor = conn.execute(
                "SELECT id, full_name, identity_no, department, title, region FROM employees ORDER BY full_name;"
            )
        return cursor.fetchall()

def add_employee(full_name, identity_no, department, title, region):
    """Add a new employee"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO employees (full_name, identity_no, department, title, region) VALUES (?, ?, ?, ?, ?);",
            (full_name, identity_no, department, title, region)
        )

def update_employee(employee_id, full_name, identity_no, department, title, region):
    """Update an existing employee"""
    with get_conn() as conn:
        conn.execute(
            "UPDATE employees SET full_name = ?, identity_no = ?, department = ?, title = ?, region = ? WHERE id = ?;",
            (full_name, identity_no, department, title, region, employee_id)
        )

def delete_employee(employee_id):
    """Delete an employee"""
    with get_conn() as conn:
        conn.execute("DELETE FROM employees WHERE id = ?;", (employee_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("employees", employee_id, datetime.now().isoformat())
        )

def get_all_employees():
    """Get all employees (for server API)"""
    return list_employees()

# ============================================================================
# TIMESHEETS
# ============================================================================

def list_timesheets(employee_id=None, start_date=None, end_date=None, region=None):
    """List timesheets with optional filters"""
    with get_conn() as conn:
        query = """
            SELECT t.id, t.employee_id, e.full_name, e.department, t.work_date, t.start_time, t.end_time,
                   t.break_minutes, t.is_special, t.notes, t.region
            FROM timesheets t
            JOIN employees e ON t.employee_id = e.id
            WHERE 1=1
        """
        params = []
        
        if employee_id:
            query += " AND t.employee_id = ?"
            params.append(employee_id)
        if start_date:
            query += " AND t.work_date >= ?"
            params.append(start_date)
        if end_date:
            query += " AND t.work_date <= ?"
            params.append(end_date)
        if region:
            query += " AND t.region = ?"
            params.append(region)
        
        query += " ORDER BY t.work_date DESC, e.full_name;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def add_timesheet(employee_id, work_date, start_time, end_time, break_minutes, is_special, notes, region):
    """Add a new timesheet entry"""
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO timesheets (employee_id, work_date, start_time, end_time, break_minutes, is_special, notes, region)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?);""",
            (employee_id, work_date, start_time, end_time, break_minutes, is_special, notes, region)
        )

def update_timesheet(timesheet_id, employee_id, work_date, start_time, end_time, break_minutes, is_special, notes, region):
    """Update an existing timesheet entry"""
    with get_conn() as conn:
        conn.execute(
            """UPDATE timesheets SET employee_id = ?, work_date = ?, start_time = ?, end_time = ?,
               break_minutes = ?, is_special = ?, notes = ?, region = ? WHERE id = ?;""",
            (employee_id, work_date, start_time, end_time, break_minutes, is_special, notes, region, timesheet_id)
        )

def delete_timesheet(timesheet_id):
    """Delete a timesheet entry"""
    with get_conn() as conn:
        conn.execute("DELETE FROM timesheets WHERE id = ?;", (timesheet_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("timesheets", timesheet_id, datetime.now().isoformat())
        )

def get_all_timesheets():
    """Get all timesheets (for server API)"""
    return list_timesheets()

# ============================================================================
# SHIFT TEMPLATES
# ============================================================================

def list_shift_templates():
    """List all shift templates"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, name, start_time, end_time, break_minutes FROM shift_templates ORDER BY name;"
        )
        return cursor.fetchall()

def upsert_shift_template(name, start_time, end_time, break_minutes):
    """Insert or update a shift template"""
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO shift_templates (name, start_time, end_time, break_minutes)
               VALUES (?, ?, ?, ?)
               ON CONFLICT(name) DO UPDATE SET start_time = ?, end_time = ?, break_minutes = ?;""",
            (name, start_time, end_time, break_minutes, start_time, end_time, break_minutes)
        )

def delete_shift_template(template_id):
    """Delete a shift template"""
    with get_conn() as conn:
        conn.execute("DELETE FROM shift_templates WHERE id = ?;", (template_id,))

# ============================================================================
# ATTENDANCE & LEAVE
# ============================================================================

def upsert_attendance_record(employee_id, work_date, status, reason, region, source):
    """Insert or update an attendance record for a day"""
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO attendance_records (employee_id, work_date, status, reason, source, region)
               VALUES (?, ?, ?, ?, ?, ?)
               ON CONFLICT(employee_id, work_date)
               DO UPDATE SET status = excluded.status,
                             reason = excluded.reason,
                             source = excluded.source,
                             region = excluded.region;""",
            (employee_id, work_date, status, reason, source, region),
        )


def delete_attendance_record(record_id):
    """Delete an attendance record"""
    with get_conn() as conn:
        conn.execute("DELETE FROM attendance_records WHERE id = ?;", (record_id,))


def list_attendance_records(employee_id=None, start_date=None, end_date=None, status=None, region=None):
    """List attendance records with employee info"""
    query = """
        SELECT a.id, a.employee_id, e.full_name, a.work_date, a.status, a.reason, a.source, a.region
        FROM attendance_records a
        JOIN employees e ON a.employee_id = e.id
        WHERE 1=1
    """
    params = []
    if employee_id:
        query += " AND a.employee_id = ?"
        params.append(employee_id)
    if start_date:
        query += " AND a.work_date >= ?"
        params.append(start_date)
    if end_date:
        query += " AND a.work_date <= ?"
        params.append(end_date)
    if status:
        query += " AND a.status = ?"
        params.append(status)
    if region:
        query += " AND a.region = ?"
        params.append(region)
    query += " ORDER BY a.work_date DESC, e.full_name;"
    with get_conn() as conn:
        cursor = conn.execute(query, params)
        return cursor.fetchall()


def add_leave_record(employee_id, start_date, end_date, leave_type, reason, document_no, status, region):
    """Add a leave record and return its id"""
    with get_conn() as conn:
        cursor = conn.execute(
            """INSERT INTO leave_records (employee_id, start_date, end_date, leave_type, reason, document_no, status, region)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?);""",
            (employee_id, start_date, end_date, leave_type, reason, document_no, status, region),
        )
        return cursor.lastrowid


def update_leave_record(leave_id, employee_id, start_date, end_date, leave_type, reason, document_no, status, region):
    """Update a leave record"""
    with get_conn() as conn:
        conn.execute(
            """UPDATE leave_records
               SET employee_id = ?, start_date = ?, end_date = ?, leave_type = ?, reason = ?,
                   document_no = ?, status = ?, region = ?
               WHERE id = ?;""",
            (employee_id, start_date, end_date, leave_type, reason, document_no, status, region, leave_id),
        )


def update_leave_document(leave_id, document_path, document_no):
    """Update leave record document fields"""
    with get_conn() as conn:
        conn.execute(
            """UPDATE leave_records SET document_path = ?, document_no = ? WHERE id = ?;""",
            (document_path, document_no, leave_id),
        )


def delete_leave_record(leave_id):
    """Delete a leave record"""
    with get_conn() as conn:
        conn.execute("DELETE FROM leave_records WHERE id = ?;", (leave_id,))


def get_leave_record(leave_id):
    """Get a leave record with employee info"""
    query = """
        SELECT l.id, l.employee_id, e.full_name, l.start_date, l.end_date, l.leave_type,
               l.reason, l.document_no, l.document_path, l.status, l.region
        FROM leave_records l
        JOIN employees e ON l.employee_id = e.id
        WHERE l.id = ?;
    """
    with get_conn() as conn:
        cursor = conn.execute(query, (leave_id,))
        return cursor.fetchone()


def list_leave_records(employee_id=None, start_date=None, end_date=None, status=None, region=None):
    """List leave records with filters"""
    query = """
        SELECT l.id, l.employee_id, e.full_name, l.start_date, l.end_date, l.leave_type,
               l.reason, l.document_no, l.document_path, l.status, l.region
        FROM leave_records l
        JOIN employees e ON l.employee_id = e.id
        WHERE 1=1
    """
    params = []
    if employee_id:
        query += " AND l.employee_id = ?"
        params.append(employee_id)
    if start_date:
        query += " AND l.end_date >= ?"
        params.append(start_date)
    if end_date:
        query += " AND l.start_date <= ?"
        params.append(end_date)
    if status:
        query += " AND l.status = ?"
        params.append(status)
    if region:
        query += " AND l.region = ?"
        params.append(region)
    query += " ORDER BY l.start_date DESC, e.full_name;"
    with get_conn() as conn:
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def add_leave_status_history(leave_id, status, note=None, changed_by=None):
    """Add leave status history entry"""
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO leave_status_history (leave_id, status, note, changed_by)
               VALUES (?, ?, ?, ?);""",
            (leave_id, status, note, changed_by),
        )

def list_leave_status_history(leave_id):
    """List leave status history for a leave record"""
    with get_conn() as conn:
        cursor = conn.execute(
            """SELECT status, note, changed_by, changed_at
               FROM leave_status_history
               WHERE leave_id = ?
               ORDER BY changed_at DESC;""",
            (leave_id,),
        )
        return cursor.fetchall()

# REPORTS
# ============================================================================

def add_report_log(file_path, created_at, employee, start_date, end_date):
    """Add a report log entry"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO reports (file_path, created_at, employee, start_date, end_date) VALUES (?, ?, ?, ?, ?);",
            (file_path, created_at, employee, start_date, end_date)
        )

def list_report_logs():
    """List all report logs"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, file_path, created_at, employee, start_date, end_date FROM reports ORDER BY created_at DESC;"
        )
        return cursor.fetchall()

# ============================================================================
# VEHICLES
# ============================================================================

def list_vehicles(region=None):
    """List all vehicles, optionally filtered by region"""
    with get_conn() as conn:
        if region:
            cursor = conn.execute(
                """SELECT id, plate, brand, model, year, km, inspection_date, insurance_date,
                          maintenance_date, oil_change_date, oil_change_km, oil_interval_km, notes, region
                   FROM vehicles WHERE region = ? ORDER BY plate;""",
                (region,)
            )
        else:
            cursor = conn.execute(
                """SELECT id, plate, brand, model, year, km, inspection_date, insurance_date,
                          maintenance_date, oil_change_date, oil_change_km, oil_interval_km, notes, region
                   FROM vehicles ORDER BY plate;"""
            )
        return cursor.fetchall()

def get_vehicle(vehicle_id):
    """Get a vehicle by ID"""
    with get_conn() as conn:
        cursor = conn.execute(
            """SELECT id, plate, brand, model, year, km, inspection_date, insurance_date,
                      maintenance_date, oil_change_date, oil_change_km, oil_interval_km, notes, region
               FROM vehicles WHERE id = ?;""",
            (vehicle_id,)
        )
        return cursor.fetchone()

def add_vehicle(plate, brand, model, year, km, inspection_date, insurance_date, maintenance_date,
                oil_change_date, oil_change_km, oil_interval_km, notes, region):
    """Add a new vehicle"""
    with get_conn() as conn:
        conn.execute(
            """INSERT INTO vehicles (plate, brand, model, year, km, inspection_date, insurance_date,
                                     maintenance_date, oil_change_date, oil_change_km, oil_interval_km, notes, region)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);""",
            (plate, brand, model, year, km, inspection_date, insurance_date, maintenance_date,
             oil_change_date, oil_change_km, oil_interval_km, notes, region)
        )

def update_vehicle(vehicle_id, plate, brand, model, year, km, inspection_date, insurance_date,
                   maintenance_date, oil_change_date, oil_change_km, oil_interval_km, notes, region):
    """Update an existing vehicle"""
    with get_conn() as conn:
        conn.execute(
            """UPDATE vehicles SET plate = ?, brand = ?, model = ?, year = ?, km = ?,
                                   inspection_date = ?, insurance_date = ?, maintenance_date = ?,
                                   oil_change_date = ?, oil_change_km = ?, oil_interval_km = ?, notes = ?, region = ?
               WHERE id = ?;""",
            (plate, brand, model, year, km, inspection_date, insurance_date, maintenance_date,
             oil_change_date, oil_change_km, oil_interval_km, notes, region, vehicle_id)
        )

def delete_vehicle(vehicle_id):
    """Delete a vehicle"""
    with get_conn() as conn:
        conn.execute("DELETE FROM vehicles WHERE id = ?;", (vehicle_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("vehicles", vehicle_id, datetime.now().isoformat())
        )

# ============================================================================
# DRIVERS
# ============================================================================

def list_drivers(region=None):
    """List all drivers, optionally filtered by region"""
    with get_conn() as conn:
        if region:
            cursor = conn.execute(
                "SELECT id, full_name, license_class, license_expiry, phone, notes, region FROM drivers WHERE region = ? ORDER BY full_name;",
                (region,)
            )
        else:
            cursor = conn.execute(
                "SELECT id, full_name, license_class, license_expiry, phone, notes, region FROM drivers ORDER BY full_name;"
            )
        return cursor.fetchall()

def get_driver(driver_id):
    """Get a driver by ID"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, full_name, license_class, license_expiry, phone, notes FROM drivers WHERE id = ?;",
            (driver_id,)
        )
        return cursor.fetchone()

def add_driver(full_name, license_class, license_expiry, phone, notes, region):
    """Add a new driver"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO drivers (full_name, license_class, license_expiry, phone, notes, region) VALUES (?, ?, ?, ?, ?, ?);",
            (full_name, license_class, license_expiry, phone, notes, region)
        )

def update_driver(driver_id, full_name, license_class, license_expiry, phone, notes, region):
    """Update an existing driver"""
    with get_conn() as conn:
        conn.execute(
            "UPDATE drivers SET full_name = ?, license_class = ?, license_expiry = ?, phone = ?, notes = ?, region = ? WHERE id = ?;",
            (full_name, license_class, license_expiry, phone, notes, region, driver_id)
        )

def delete_driver(driver_id):
    """Delete a driver"""
    with get_conn() as conn:
        conn.execute("DELETE FROM drivers WHERE id = ?;", (driver_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("drivers", driver_id, datetime.now().isoformat())
        )

# ============================================================================
# VEHICLE FAULTS
# ============================================================================

def list_vehicle_faults(vehicle_id=None, region=None):
    """List vehicle faults"""
    with get_conn() as conn:
        query = """
            SELECT f.id, f.vehicle_id, v.plate, f.title, f.description, f.opened_date,
                   f.closed_date, f.status, f.region
            FROM vehicle_faults f
            JOIN vehicles v ON f.vehicle_id = v.id
            WHERE 1=1
        """
        params = []
        
        if vehicle_id:
            query += " AND f.vehicle_id = ?"
            params.append(vehicle_id)
        if region:
            query += " AND f.region = ?"
            params.append(region)
        
        query += " ORDER BY f.opened_date DESC;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def list_open_vehicle_faults(vehicle_id=None, region=None):
    """List open vehicle faults"""
    with get_conn() as conn:
        query = """
            SELECT f.id, f.vehicle_id, v.plate, f.title, f.description, f.opened_date,
                   f.closed_date, f.status, f.region
            FROM vehicle_faults f
            JOIN vehicles v ON f.vehicle_id = v.id
            WHERE f.status = 'Acik'
        """
        params = []
        
        if vehicle_id:
            query += " AND f.vehicle_id = ?"
            params.append(vehicle_id)
        if region:
            query += " AND f.region = ?"
            params.append(region)
        
        query += " ORDER BY f.opened_date DESC;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def get_vehicle_fault(fault_id):
    """Get a vehicle fault by ID"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, vehicle_id, title, description, opened_date, closed_date, status FROM vehicle_faults WHERE id = ?;",
            (fault_id,)
        )
        return cursor.fetchone()

def add_vehicle_fault(vehicle_id, title, description, opened_date, closed_date, status, region):
    """Add a new vehicle fault"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO vehicle_faults (vehicle_id, title, description, opened_date, closed_date, status, region) VALUES (?, ?, ?, ?, ?, ?, ?);",
            (vehicle_id, title, description, opened_date, closed_date, status, region)
        )

def update_vehicle_fault(fault_id, vehicle_id, title, description, opened_date, closed_date, status, region):
    """Update an existing vehicle fault"""
    with get_conn() as conn:
        conn.execute(
            "UPDATE vehicle_faults SET vehicle_id = ?, title = ?, description = ?, opened_date = ?, closed_date = ?, status = ?, region = ? WHERE id = ?;",
            (vehicle_id, title, description, opened_date, closed_date, status, region, fault_id)
        )

def delete_vehicle_fault(fault_id):
    """Delete a vehicle fault"""
    with get_conn() as conn:
        conn.execute("DELETE FROM vehicle_faults WHERE id = ?;", (fault_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("vehicle_faults", fault_id, datetime.now().isoformat())
        )

# ============================================================================
# VEHICLE INSPECTIONS
# ============================================================================

def list_vehicle_inspections(vehicle_id=None, week_start=None, region=None):
    """List vehicle inspections"""
    with get_conn() as conn:
        query = """
            SELECT i.id, i.vehicle_id, v.plate, i.driver_id, d.full_name, i.inspection_date,
                   i.week_start, i.km, i.notes, i.fault_id, i.fault_status, i.service_visit
            FROM vehicle_inspections i
            JOIN vehicles v ON i.vehicle_id = v.id
            LEFT JOIN drivers d ON i.driver_id = d.id
            WHERE 1=1
        """
        params = []
        
        if vehicle_id:
            query += " AND i.vehicle_id = ?"
            params.append(vehicle_id)
        if week_start:
            query += " AND i.week_start = ?"
            params.append(week_start)
        if region:
            query += " AND v.region = ?"
            params.append(region)
        
        query += " ORDER BY i.inspection_date DESC;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def list_driver_inspections(driver_id, region=None):
    """List inspections by driver"""
    with get_conn() as conn:
        query = """
            SELECT i.id, i.vehicle_id, v.plate, i.driver_id, d.full_name, i.inspection_date,
                   i.week_start, i.km, i.notes, i.fault_id, i.fault_status, i.service_visit
            FROM vehicle_inspections i
            JOIN vehicles v ON i.vehicle_id = v.id
            LEFT JOIN drivers d ON i.driver_id = d.id
            WHERE i.driver_id = ?
        """
        params = [driver_id]
        
        if region:
            query += " AND v.region = ?"
            params.append(region)
        
        query += " ORDER BY i.inspection_date DESC;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def add_vehicle_inspection(vehicle_id, driver_id, inspection_date, week_start, km, notes,
                           fault_id=None, fault_status=None, service_visit=0):
    """Add a new vehicle inspection"""
    with get_conn() as conn:
        cursor = conn.execute(
            """INSERT INTO vehicle_inspections (vehicle_id, driver_id, inspection_date, week_start, km, notes,
                                                 fault_id, fault_status, service_visit)
               VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?);""",
            (vehicle_id, driver_id, inspection_date, week_start, km, notes, fault_id, fault_status, service_visit)
        )
        return cursor.lastrowid

def list_vehicle_inspection_results(inspection_id):
    """List inspection results for a specific inspection"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT item_key, status, note FROM vehicle_inspection_results WHERE inspection_id = ?;",
            (inspection_id,)
        )
        return cursor.fetchall()

def add_vehicle_inspection_result(inspection_id, item_key, status, note):
    """Add an inspection result item"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO vehicle_inspection_results (inspection_id, item_key, status, note) VALUES (?, ?, ?, ?);",
            (inspection_id, item_key, status, note)
        )

# ============================================================================
# VEHICLE SERVICE VISITS
# ============================================================================

def list_vehicle_service_visits(vehicle_id=None, start_date=None, end_date=None, region=None):
    """List vehicle service visits"""
    with get_conn() as conn:
        query = """
            SELECT s.id, s.vehicle_id, v.plate, s.fault_id, f.title, s.start_date, s.end_date,
                   s.reason, s.cost, s.notes, s.region
            FROM vehicle_service_visits s
            JOIN vehicles v ON s.vehicle_id = v.id
            LEFT JOIN vehicle_faults f ON s.fault_id = f.id
            WHERE 1=1
        """
        params = []
        
        if vehicle_id:
            query += " AND s.vehicle_id = ?"
            params.append(vehicle_id)
        if start_date:
            query += " AND s.start_date >= ?"
            params.append(start_date)
        if end_date:
            query += " AND s.start_date <= ?"
            params.append(end_date)
        if region:
            query += " AND s.region = ?"
            params.append(region)
        
        query += " ORDER BY s.start_date DESC;"
        
        cursor = conn.execute(query, params)
        return cursor.fetchall()

def get_vehicle_service_visit(visit_id):
    """Get a service visit by ID"""
    with get_conn() as conn:
        cursor = conn.execute(
            "SELECT id, vehicle_id, fault_id, start_date, end_date, reason, cost, notes FROM vehicle_service_visits WHERE id = ?;",
            (visit_id,)
        )
        return cursor.fetchone()

def add_vehicle_service_visit(vehicle_id, fault_id, start_date, end_date, reason, cost, notes, region):
    """Add a new service visit"""
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO vehicle_service_visits (vehicle_id, fault_id, start_date, end_date, reason, cost, notes, region) VALUES (?, ?, ?, ?, ?, ?, ?, ?);",
            (vehicle_id, fault_id, start_date, end_date, reason, cost, notes, region)
        )

def update_vehicle_service_visit(visit_id, vehicle_id, fault_id, start_date, end_date, reason, cost, notes, region):
    """Update an existing service visit"""
    with get_conn() as conn:
        conn.execute(
            "UPDATE vehicle_service_visits SET vehicle_id = ?, fault_id = ?, start_date = ?, end_date = ?, reason = ?, cost = ?, notes = ?, region = ? WHERE id = ?;",
            (vehicle_id, fault_id, start_date, end_date, reason, cost, notes, region, visit_id)
        )

def delete_vehicle_service_visit(visit_id):
    """Delete a service visit"""
    with get_conn() as conn:
        conn.execute("DELETE FROM vehicle_service_visits WHERE id = ?;", (visit_id,))
        # Track deletion
        conn.execute(
            "INSERT INTO deleted_records (table_name, record_id, deleted_at) VALUES (?, ?, ?);",
            ("vehicle_service_visits", visit_id, datetime.now().isoformat())
        )

# ============================================================================
# BACKUP & RESTORE
# ============================================================================

def _backup_db_if_needed():
    """Create automatic backup if needed (once per day)"""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        
        # Check last backup time
        if os.path.exists(BACKUP_MARKER):
            with open(BACKUP_MARKER, 'r') as f:
                last_backup = f.read().strip()
                last_date = datetime.fromisoformat(last_backup).date()
                if last_date >= datetime.now().date():
                    return  # Already backed up today
        
        # Create backup
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = os.path.join(BACKUP_DIR, f"puantaj_auto_{timestamp}.db")
        shutil.copy2(DB_PATH, backup_path)
        
        # Update marker
        with open(BACKUP_MARKER, 'w') as f:
            f.write(datetime.now().isoformat())
        
        # Clean old backups (keep last 7 days)
        _cleanup_old_backups()
    except Exception:
        pass

def _cleanup_old_backups():
    """Remove backups older than 7 days"""
    try:
        cutoff = datetime.now() - timedelta(days=7)
        for filename in os.listdir(BACKUP_DIR):
            if filename.startswith("puantaj_auto_") and filename.endswith(".db"):
                filepath = os.path.join(BACKUP_DIR, filename)
                if os.path.getmtime(filepath) < cutoff.timestamp():
                    os.remove(filepath)
    except Exception:
        pass

def create_backup(output_path):
    """Create a manual backup"""
    shutil.copy2(DB_PATH, output_path)
    return output_path

def restore_backup(backup_path):
    """Restore database from backup"""
    shutil.copy2(backup_path, DB_PATH)

def export_data_zip(output_path):
    """Export database and backups as ZIP"""
    os.makedirs(EXPORT_DIR, exist_ok=True)
    
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
        # Add main database
        zf.write(DB_PATH, "puantaj.db")
        
        # Add backups if they exist
        if os.path.exists(BACKUP_DIR):
            for filename in os.listdir(BACKUP_DIR):
                if filename.endswith(".db"):
                    filepath = os.path.join(BACKUP_DIR, filename)
                    zf.write(filepath, f"backups/{filename}")
    
    return output_path

def import_data_zip(zip_path):
    """Import database from ZIP"""
    with zipfile.ZipFile(zip_path, 'r') as zf:
        # Extract main database
        zf.extract("puantaj.db", DB_DIR)
        
        # Extract backups if present
        for name in zf.namelist():
            if name.startswith("backups/"):
                zf.extract(name, DB_DIR)
