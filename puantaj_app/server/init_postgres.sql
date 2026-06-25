-- RainStaff PostgreSQL Schema
-- This will be automatically executed when database is created

-- Employees table
CREATE TABLE IF NOT EXISTS employees (
    id SERIAL PRIMARY KEY,
    full_name TEXT NOT NULL,
    identity_no TEXT,
    department TEXT,
    title TEXT,
    region TEXT
);

-- Timesheets table
CREATE TABLE IF NOT EXISTS timesheets (
    id SERIAL PRIMARY KEY,
    employee_id INTEGER NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    work_date TEXT NOT NULL,
    start_time TEXT NOT NULL,
    end_time TEXT NOT NULL,
    break_minutes INTEGER NOT NULL DEFAULT 0,
    is_special INTEGER NOT NULL DEFAULT 0,
    notes TEXT,
    region TEXT
);

-- Settings table
CREATE TABLE IF NOT EXISTS settings (
    key TEXT PRIMARY KEY,
    value TEXT NOT NULL
);

-- Reports table
CREATE TABLE IF NOT EXISTS reports (
    id SERIAL PRIMARY KEY,
    file_path TEXT NOT NULL,
    created_at TEXT NOT NULL,
    employee TEXT,
    start_date TEXT,
    end_date TEXT
);

-- Shift templates table
CREATE TABLE IF NOT EXISTS shift_templates (
    id SERIAL PRIMARY KEY,
    name TEXT NOT NULL UNIQUE,
    start_time TEXT NOT NULL,
    end_time TEXT NOT NULL,
    break_minutes INTEGER NOT NULL DEFAULT 0
);

-- Vehicles table
CREATE TABLE IF NOT EXISTS vehicles (
    id SERIAL PRIMARY KEY,
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

-- Drivers table
CREATE TABLE IF NOT EXISTS drivers (
    id SERIAL PRIMARY KEY,
    full_name TEXT NOT NULL,
    license_class TEXT,
    license_expiry TEXT,
    phone TEXT,
    notes TEXT,
    region TEXT
);

-- Users table
CREATE TABLE IF NOT EXISTS users (
    id SERIAL PRIMARY KEY,
    username TEXT NOT NULL UNIQUE,
    password_hash TEXT NOT NULL,
    role TEXT NOT NULL,
    region TEXT NOT NULL
);

-- Vehicle faults table
CREATE TABLE IF NOT EXISTS vehicle_faults (
    id SERIAL PRIMARY KEY,
    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
    title TEXT NOT NULL,
    description TEXT,
    opened_date TEXT,
    closed_date TEXT,
    status TEXT DEFAULT 'Acik',
    region TEXT
);

-- Vehicle inspections table
CREATE TABLE IF NOT EXISTS vehicle_inspections (
    id SERIAL PRIMARY KEY,
    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
    driver_id INTEGER REFERENCES drivers(id) ON DELETE SET NULL,
    inspection_date TEXT NOT NULL,
    week_start TEXT NOT NULL,
    km INTEGER,
    notes TEXT,
    fault_id INTEGER REFERENCES vehicle_faults(id) ON DELETE SET NULL,
    fault_status TEXT,
    service_visit INTEGER DEFAULT 0
);

-- Vehicle service visits table
CREATE TABLE IF NOT EXISTS vehicle_service_visits (
    id SERIAL PRIMARY KEY,
    vehicle_id INTEGER NOT NULL REFERENCES vehicles(id) ON DELETE CASCADE,
    fault_id INTEGER REFERENCES vehicle_faults(id) ON DELETE SET NULL,
    start_date TEXT NOT NULL,
    end_date TEXT,
    reason TEXT,
    cost REAL,
    notes TEXT,
    region TEXT
);

-- Vehicle inspection results table
CREATE TABLE IF NOT EXISTS vehicle_inspection_results (
    id SERIAL PRIMARY KEY,
    inspection_id INTEGER NOT NULL REFERENCES vehicle_inspections(id) ON DELETE CASCADE,
    item_key TEXT NOT NULL,
    status TEXT NOT NULL,
    note TEXT
);

-- Stock inventory table
CREATE TABLE IF NOT EXISTS stock_inventory (
    id SERIAL PRIMARY KEY,
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

-- Deleted records tracking table
CREATE TABLE IF NOT EXISTS deleted_records (
    id SERIAL PRIMARY KEY,
    table_name TEXT NOT NULL,
    record_id INTEGER NOT NULL,
    deleted_at TEXT NOT NULL,
    deleted_by TEXT
);

-- Insert default settings
INSERT INTO settings (key, value) VALUES 
    ('company_name', ''),
    ('report_title', 'Rainstaff Puantaj ve Mesai Raporu'),
    ('weekday_hours', '8'),
    ('saturday_start', '09:00'),
    ('saturday_end', '14:00'),
    ('logo_path', ''),
    ('sync_enabled', '0'),
    ('sync_url', ''),
    ('sync_token', ''),
    ('admin_entry_region', 'Ankara'),
    ('admin_view_region', 'Tum Bolgeler')
ON CONFLICT (key) DO NOTHING;

-- Insert default shift templates
INSERT INTO shift_templates (name, start_time, end_time, break_minutes) VALUES 
    ('Hafta Ici 09-18', '09:00', '18:00', 60),
    ('Cumartesi 09-14', '09:00', '14:00', 0)
ON CONFLICT (name) DO NOTHING;

-- Insert default users (passwords are hashed with SHA-256)
INSERT INTO users (username, password_hash, role, region) VALUES 
    ('ankara1', 'e10adc3949ba59abbe56e057f20f883e', 'user', 'Ankara'),
    ('izmir1', 'e10adc3949ba59abbe56e057f20f883e', 'user', 'Izmir'),
    ('bursa1', 'e10adc3949ba59abbe56e057f20f883e', 'user', 'Bursa'),
    ('istanbul1', 'e10adc3949ba59abbe56e057f20f883e', 'user', 'Istanbul'),
    ('admin', 'c33367701511b4f6020ec61ded352059', 'admin', 'ALL')
ON CONFLICT (username) DO NOTHING;

-- Create indexes for better performance
CREATE INDEX IF NOT EXISTS idx_timesheets_employee ON timesheets(employee_id);
CREATE INDEX IF NOT EXISTS idx_timesheets_date ON timesheets(work_date);
CREATE INDEX IF NOT EXISTS idx_timesheets_region ON timesheets(region);
CREATE INDEX IF NOT EXISTS idx_vehicles_plate ON vehicles(plate);
CREATE INDEX IF NOT EXISTS idx_vehicles_region ON vehicles(region);
CREATE INDEX IF NOT EXISTS idx_stock_seri ON stock_inventory(seri_no);
CREATE INDEX IF NOT EXISTS idx_stock_bolge ON stock_inventory(bolge);
