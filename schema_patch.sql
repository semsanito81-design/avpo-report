CREATE TABLE IF NOT EXISTS manual_edits (
    report_date  TEXT NOT NULL,
    host_ksc     TEXT NOT NULL,
    column_name  TEXT NOT NULL,
    value        TEXT,
    updated_at   TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (report_date, host_ksc, column_name)
);