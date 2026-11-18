# -*- coding: utf-8 -*-
"""
АВПО: загрузка CSV из KSC -> SQLite -> Excel-отчёт.

Один лист «Отчёт АВПО»:
  [Имя хоста в AD][Имя хоста в KSC][Версия KES][Версия агента]
  + по 3 колонки на каждый четверг (Статус / Отчет об устранении / Причина).
Четверги добавляются по мере поступления данных, месяцы идут подряд.

Запуск:
    python avpo.py
    python avpo.py --csv data --date 13.11.2026 --out out/report.xlsx
"""
from __future__ import annotations

import argparse
import glob
import os
import re
import sqlite3
import sys
from datetime import date, datetime, timedelta
from pathlib import Path

import pandas as pd

# ----------------------------- Пути ----------------------------------------

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
OUT_DIR  = APP_DIR / "out"
DB_PATH  = APP_DIR / "avpo.db"
DATA_DIR.mkdir(exist_ok=True)
OUT_DIR.mkdir(exist_ok=True)

RU_MONTHS = ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь",
             "Июль", "Август", "Сентябрь", "Октябрь", "Ноябрь", "Декабрь"]

# ----------------------------- Схема БД ------------------------------------

DDL = """
CREATE TABLE IF NOT EXISTS hosts (
    id              INTEGER PRIMARY KEY AUTOINCREMENT,
    host_ksc        TEXT NOT NULL,
    host_ad         TEXT,
    ip              TEXT,
    group_full      TEXT,
    site            TEXT,
    kes_version     TEXT,
    agent_version   TEXT,
    agent_running   TEXT,
    status          TEXT,
    status_desc     TEXT,
    reason          TEXT,
    last_connect    TEXT,
    last_update_av  TEXT,
    av_release      TEXT,
    need_reboot     TEXT,
    source_file     TEXT,
    loaded_at       TEXT DEFAULT (datetime('now')),
    UNIQUE(host_ksc, ip)
);
CREATE INDEX IF NOT EXISTS ix_hosts_site   ON hosts(site);
CREATE INDEX IF NOT EXISTS ix_hosts_status ON hosts(status);
CREATE INDEX IF NOT EXISTS ix_hosts_ip     ON hosts(ip);

CREATE TABLE IF NOT EXISTS snapshots (
    report_date    TEXT NOT NULL,
    host_ksc       TEXT NOT NULL,
    ip             TEXT NOT NULL DEFAULT '',
    host_ad        TEXT,
    kes_version    TEXT,
    agent_version  TEXT,
    status         TEXT,
    status_desc    TEXT,
    reason         TEXT,
    resolved_at    TEXT,
    source_file    TEXT,
    loaded_at      TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (report_date, host_ksc, ip)
);
CREATE INDEX IF NOT EXISTS ix_snap_date ON snapshots(report_date);
CREATE INDEX IF NOT EXISTS ix_snap_host ON snapshots(host_ksc);

CREATE TABLE IF NOT EXISTS load_log (
    id         INTEGER PRIMARY KEY AUTOINCREMENT,
    file_name  TEXT,
    rows_total INTEGER,
    rows_new   INTEGER,
    rows_upd   INTEGER,
    loaded_at  TEXT DEFAULT (datetime('now'))
);
CREATE TABLE IF NOT EXISTS manual_edits (
    report_date  TEXT NOT NULL,
    host_ksc     TEXT NOT NULL,
    column_name  TEXT NOT NULL,   -- 'resolved_at' | 'reason' | 'status'
    value        TEXT,
    updated_at   TEXT DEFAULT (datetime('now')),
    PRIMARY KEY (report_date, host_ksc, column_name)
);
"""

# ------------------------- Нормализация ------------------------------------

DOMAIN_RE = re.compile(r"\.gtm\.gazprom\.ru$", re.IGNORECASE)

REASON_MAP = {
    "защита выключена": "Защита выключена",
    "приложение безопасности не запущено": "Приложение безопасности не запущено",
    "обнаружены активные угрозы": "Обнаружены активные угрозы",
    "базы устарели": "Базы устарели",
    "базы приложения сильно устарели": "Базы приложения сильно устарели",
    "давно не подключался": "Давно не подключался",
    "давно не выполнялся поиск вредоносного по": "Давно не выполнялся поиск вредоносного ПО",
    "cryptossl": "Уст. обновления версии...",
    "перезапустите приложение": "Уст. обновления версии...",
}

COL_ALIASES = {
    "имя": "host_ksc",
    "последнее подключение к серверу администрирования": "last_connect",
    "агент администрирования запущен": "agent_running",
    "статус": "status",
    "описание статуса": "status_desc",
    "версия агента администрирования": "agent_version",
    "точка распространения": "dist_point",
    "последнее обновление антивирусных баз": "last_update_av",
    "версия приложения": "kes_version",
    "выпуск антивирусных баз": "av_release",
    "ip-адрес": "ip",
    "требуется перезагрузка": "need_reboot",
    "полное название группы": "group_full",
}


def _norm_header(h: str) -> str:
    return h.strip().lstrip("\ufeff").lower()


def _parse_dt(s: str):
    if not s or not str(s).strip():
        return None
    s = str(s).strip()
    for fmt in ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y"):
        try:
            return datetime.strptime(s, fmt).isoformat()
        except ValueError:
            continue
    return None


def _short_ad(host_ksc: str) -> str:
    return DOMAIN_RE.sub("", host_ksc.strip()).upper()


def _site_from_group(group: str) -> str:
    if not group:
        return ""
    tail = group.split("/")[-1].strip()
    m = re.match(r"^\d+(?:\.\d+){0,3}\s+(.*)$", tail)
    return (m.group(1) if m else tail).strip()


def _reason_from_desc(desc: str) -> str:
    if not desc:
        return ""
    d = desc.lower()
    hits = []
    for key, label in REASON_MAP.items():
        if key in d and label not in hits:
            hits.append(label)
    if hits:
        return "; ".join(hits)
    # fallback: не теряем неизвестную причину
    return desc.strip()


def _status_norm(status: str) -> str:
    s = (status or "").strip().lower()
    if s.startswith("ок"):
        return "ОК"
    if s.startswith("предупр"):
        return "Предупреждение"
    if s.startswith("крит"):
        return "Критический"
    return status.strip() if status else ""


# -------------------------- Чтение CSV -------------------------------------

def read_ksc_csv(path) -> pd.DataFrame:
    raw = pd.read_csv(path, sep=None, engine="python",
                      encoding="utf-8-sig", dtype=str, keep_default_na=False)
    rename = {}
    for col in raw.columns:
        key = _norm_header(col)
        if key in COL_ALIASES:
            rename[col] = COL_ALIASES[key]
    raw = raw.rename(columns=rename)
    for c in COL_ALIASES.values():
        if c not in raw.columns:
            raw[c] = ""
    return raw


def transform(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df["host_ksc"] = df["host_ksc"].str.strip()
    df = df[df["host_ksc"] != ""]
    df["host_ad"]    = df["host_ksc"].map(_short_ad)
    df["site"]       = df["group_full"].map(_site_from_group)
    df["status"]     = df["status"].map(_status_norm)
    df["reason"]     = df["status_desc"].map(_reason_from_desc)
    df["last_connect"]   = df["last_connect"].map(_parse_dt)
    df["last_update_av"] = df["last_update_av"].map(_parse_dt)
    return df[[
        "host_ksc", "host_ad", "ip", "group_full", "site",
        "kes_version", "agent_version", "agent_running",
        "status", "status_desc", "reason",
        "last_connect", "last_update_av", "av_release", "need_reboot",
    ]]


# --------------------------- Работа с датами --------------------------------

def thursday_of(d: date) -> date:
    return d - timedelta(days=(d.weekday() - 3) % 7)


def last_thursday(today: date | None = None) -> date:
    return thursday_of(today or date.today())


def parse_date_arg(s: str) -> date:
    s = s.strip()
    for fmt in ("%d.%m.%Y", "%Y-%m-%d", "%d.%m.%y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    raise ValueError(f"Неверная дата: {s} (ДД.ММ.ГГГГ или ГГГГ-ММ-ДД)")


def extract_date_from_filename(path) -> date | None:
    name = Path(path).name
    for pat in (r"(\d{4})[.\-](\d{2})[.\-](\d{2})",
                r"(\d{2})[.\-](\d{2})[.\-](\d{4})"):
        m = re.search(pat, name)
        if m:
            try:
                if len(m.group(1)) == 4:
                    return date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                return date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                pass
    return None


# -------------------------- Загрузка в БД ----------------------------------

def init_db(conn: sqlite3.Connection) -> None:
    conn.executescript(DDL)
    conn.commit()


def load_csv_to_db(path, conn, report_date: str | None = None,
                   resolved_map: dict[str, str] | None = None):
    if report_date is None:
        d = extract_date_from_filename(path) or last_thursday()
        report_date = d.isoformat()
    resolved_map = resolved_map or {}
    df = transform(read_ksc_csv(path))
    cur = conn.cursor()
    new = upd = 0
    for _, r in df.iterrows():
        cur.execute("SELECT id FROM hosts WHERE host_ksc=? AND ip=?", (r.host_ksc, r.ip))
        row = cur.fetchone()
        if row:
            cur.execute("""UPDATE hosts SET host_ad=?, group_full=?, site=?,
                    kes_version=?, agent_version=?, agent_running=?,
                    status=?, status_desc=?, reason=?,
                    last_connect=?, last_update_av=?, av_release=?, need_reboot=?,
                    source_file=?, loaded_at=datetime('now') WHERE id=?""",
                (r.host_ad, r.group_full, r.site, r.kes_version, r.agent_version,
                 r.agent_running, r.status, r.status_desc, r.reason,
                 r.last_connect, r.last_update_av, r.av_release, r.need_reboot,
                 os.path.basename(str(path)), row[0]))
            upd += 1
        else:
            cur.execute("""INSERT INTO hosts (host_ksc, host_ad, ip, group_full, site,
                    kes_version, agent_version, agent_running,
                    status, status_desc, reason,
                    last_connect, last_update_av, av_release, need_reboot, source_file)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                (r.host_ksc, r.host_ad, r.ip, r.group_full, r.site,
                 r.kes_version, r.agent_version, r.agent_running,
                 r.status, r.status_desc, r.reason,
                 r.last_connect, r.last_update_av, r.av_release, r.need_reboot,
                 os.path.basename(str(path))))
            new += 1

        resolved = resolved_map.get(r.host_ksc)
        cur.execute("""INSERT OR REPLACE INTO snapshots
            (report_date, host_ksc, ip, host_ad, kes_version, agent_version,
             status, status_desc, reason, resolved_at, source_file)
            VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (report_date, r.host_ksc, r.ip or "", r.host_ad, r.kes_version,
             r.agent_version, r.status, r.status_desc, r.reason, resolved,
             os.path.basename(str(path))))

    conn.execute("""INSERT INTO load_log(file_name, rows_total, rows_new, rows_upd)
                    VALUES (?,?,?,?)""",
                 (f"{os.path.basename(str(path))} @ {report_date}",
                  len(df), new, upd))
    conn.commit()
    return len(df), new, upd


# ------------------ Сборка единого листа отчёта ----------------------------

BASE_COLS = ["Имя хоста в AD", "Имя хоста в KSC",
             "Версия KES", "Версия агента Администрирования"]

SERVERS_FILE = APP_DIR / "servers.txt"

def _load_servers() -> set[str]:
    if not SERVERS_FILE.exists():
        return set()
    out = set()
    for line in SERVERS_FILE.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if line and not line.startswith("#"):
            out.add(line.upper())
    return out

def build_combined_frame(conn: sqlite3.Connection,
                         year: int | None = None,
                         month: int | None = None
                         ) -> tuple[pd.DataFrame, list[str]]:
    """
    Возвращает DataFrame с колонками для одного месяца (или для всех, если
    year/month = None) и список дат этого месяца в порядке возрастания.
    """
    # --- даты ---
    if year is not None and month is not None:
        start = date(year, month, 1).isoformat()
        # первое число следующего месяца
        if month == 12:
            end = date(year + 1, 1, 1).isoformat()
        else:
            end = date(year, month + 1, 1).isoformat()
        dates = [r[0] for r in conn.execute(
            "SELECT DISTINCT report_date FROM snapshots "
            "WHERE report_date >= ? AND report_date < ? ORDER BY report_date",
            (start, end))]
    else:
        dates = [r[0] for r in conn.execute(
            "SELECT DISTINCT report_date FROM snapshots ORDER BY report_date")]

    if not dates:
        return pd.DataFrame(columns=BASE_COLS), []

    # --- снимки ---
    placeholders = ",".join("?" * len(dates))
    snap = pd.read_sql_query(
        f"""SELECT report_date, host_ksc, host_ad, kes_version, agent_version,
                   status, reason, resolved_at
            FROM snapshots WHERE report_date IN ({placeholders})""",
        conn, params=dates)

    # --- базовый список хостов ---
    # берём из hosts (последнее состояние) + всех, кто встречался в snapshots за эти даты
    hosts_now  = pd.read_sql_query(
        "SELECT host_ksc, host_ad, kes_version, agent_version FROM hosts", conn)
    hosts_hist = snap[["host_ksc", "host_ad", "kes_version", "agent_version"]]
    base = (pd.concat([hosts_now, hosts_hist])
              .sort_values(["host_ksc", "kes_version"], ascending=[True, False])
              .drop_duplicates("host_ksc")
              .reset_index(drop=True))

    # сортировка: АРМы → серверы
    servers = _load_servers()
    base["_is_server"] = base["host_ad"].str.upper().isin(servers).astype(int)
    base = base.sort_values(["_is_server", "host_ad"], kind="stable") \
               .reset_index(drop=True)

    # --- блоки по датам ---
    data = {
        "Имя хоста в AD": base["host_ad"].values,
        "Имя хоста в KSC": base["host_ksc"].values,
        "Версия KES": base["kes_version"].values,
        "Версия агента Администрирования": base["agent_version"].values,
    }

    for iso in dates:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
        tag = f"{d:%d.%m.%Y}"
        sl = snap[snap.report_date == iso].set_index("host_ksc")
        data[f"Статус на {tag}"] = sl["status"].reindex(base.host_ksc).values
        data[f"Отчет об устранении (с указанием даты) {tag}"] = \
            sl["resolved_at"].reindex(base.host_ksc).values
        data[f"Причина возникновения события {tag}"] = \
            sl["reason"].reindex(base.host_ksc).values

    # --- ручные правки (накладываем только за этот месяц) ---
    if year is not None and month is not None:
        edits = pd.read_sql_query(
            "SELECT report_date, host_ksc, column_name, value FROM manual_edits "
            "WHERE report_date >= ? AND report_date < ?",
            conn, params=(start, end))
    else:
        edits = pd.read_sql_query(
            "SELECT report_date, host_ksc, column_name, value FROM manual_edits", conn)

    if not edits.empty:
        emap: dict[str, dict[str, dict[str, str]]] = {}
        for _, e in edits.iterrows():
            emap.setdefault(e.report_date, {}) \
                .setdefault(e.host_ksc, {})[e.column_name] = e.value

        for iso in dates:
            d = datetime.strptime(iso, "%Y-%m-%d").date()
            tag = f"{d:%d.%m.%Y}"
            col_status   = f"Статус на {tag}"
            col_resolved = f"Отчет об устранении (с указанием даты) {tag}"
            col_reason   = f"Причина возникновения события {tag}"
            for col, default in ((col_status, [None]*len(base)),
                                 (col_resolved, [None]*len(base)),
                                 (col_reason, [None]*len(base))):
                if col not in data:
                    data[col] = default
            for idx, host in enumerate(base.host_ksc):
                hv = emap.get(iso, {}).get(host, {})
                if hv.get("status"):
                    data[col_status][idx] = hv["status"]
                if hv.get("resolved_at"):
                    data[col_resolved][idx] = hv["resolved_at"]
                if hv.get("reason"):
                    data[col_reason][idx] = hv["reason"]

    return pd.DataFrame(data), dates

def _import_edits_from_sheet(xls: pd.ExcelFile, sheet_name: str, conn: sqlite3.Connection) -> int:
    """Вспомогательная функция: импорт правок с одного листа."""
    # Читаем лист с двумя строками заголовка
    try:
        raw = pd.read_excel(xls, sheet_name=sheet_name, header=[0, 1], dtype=str)
    except Exception:
        return 0

    if raw.empty:
        return 0

    # Склеиваем MultiIndex-заголовок в плоский список
    flat_cols = []
    for top, bot in raw.columns:
        top_s = "" if pd.isna(top) else str(top).strip()
        bot_s = "" if pd.isna(bot) else str(bot).strip()
        if top_s in ("Имя хоста в AD", "Имя хоста в KSC", "Версия KES",
                     "Версия агента Администрирования"):
            flat_cols.append(top_s)
        else:
            flat_cols.append(bot_s)  # например 'Статус на 13.11.2026'
    raw.columns = flat_cols

    if "Имя хоста в KSC" not in raw.columns:
        return 0

    cur = conn.cursor()
    saved = 0
    date_re = re.compile(r"(\d{2}\.\d{2}\.\d{4})$")
    col_map = {
        "Статус на ": "status",
        "Отчет об устранении (с указанием даты) ": "resolved_at",
        "Причина возникновения события ": "reason",
    }

    for col in raw.columns:
        for prefix, db_col in col_map.items():
            if not col.startswith(prefix):
                continue
            m = date_re.search(col)
            if not m:
                continue
            try:
                d = datetime.strptime(m.group(1), "%d.%m.%Y").date()
            except ValueError:
                continue
            iso = d.isoformat()

            for _, row in raw.iterrows():
                host = row.get("Имя хоста в KSC")
                if not host or pd.isna(host):
                    continue
                val = row.get(col)
                if val is None or pd.isna(val) or str(val).strip() == "":
                    continue
                
                cur.execute("""INSERT OR REPLACE INTO manual_edits
                    (report_date, host_ksc, column_name, value, updated_at)
                    VALUES (?,?,?,?,datetime('now'))""",
                    (iso, str(host).strip(), db_col, str(val).strip()))
                saved += 1
            break  # переходим к следующей колонке

    return saved


def import_manual_edits_from_excel(path, conn: sqlite3.Connection) -> int:
    """
    Читает ВСЕ листы отчёта, вытаскивает блоки 'Статус на ДД.ММ.ГГГГ',
    'Отчет об устранении...', 'Причина...' и складывает непустые значения
    в manual_edits. Возвращает число сохранённых ячеек.
    """
    xls = pd.ExcelFile(path, engine="openpyxl")
    total_saved = 0
    
    # Проходим по всем листам в книге
    for sheet in xls.sheet_names:
        # Пропускаем служебные листы (если они вдруг остались от старых версий)
        if sheet in ("Сводка", "Критические", "Предупреждения", "Метаданные"):
            continue
        
        saved = _import_edits_from_sheet(xls, sheet, conn)
        total_saved += saved

    conn.commit()
    return total_saved


# ---------------------------- Excel ----------------------------------------

EXCEL_COLS = [
    ("host_ad",       "Имя хоста в AD"),
    ("host_ksc",      "Имя хоста в KSC"),
    ("kes_version",   "Версия KES"),
    ("agent_version", "Версия агента Администрирования"),
    ("status",        "Статус"),
    ("reason",        "Причина возникновения события"),
    ("status_desc",   "Описание статуса (KSC)"),
    ("last_connect",  "Последнее подключение"),
    ("av_release",    "Выпуск антивирусных баз"),
    ("last_update_av","Последнее обновление баз"),
    ("ip",            "IP-адрес"),
    ("need_reboot",   "Требуется перезагрузка"),
    ("group_full",    "Полное название группы"),
]


def _df_from_sql(conn, sql, params=()) -> pd.DataFrame:
    return pd.read_sql_query(sql, conn, params=params)


def _write_flat(writer, df: pd.DataFrame, sheet: str):
    out = df[[c for c, _ in EXCEL_COLS]].rename(columns=dict(EXCEL_COLS))
    out.to_excel(writer, sheet_name=sheet[:31], index=False)


def _autosize_flat(writer):
    from openpyxl.utils import get_column_letter
    for ws in writer.book.worksheets:
        if ws.title == "Отчёт АВПО":
            continue
        for i, col in enumerate(ws.iter_cols(min_row=1, max_row=ws.max_row), start=1):
            maxlen = max((len(str(c.value)) if c.value is not None else 0) for c in col)
            ws.column_dimensions[get_column_letter(i)].width = min(maxlen + 2, 60)
        ws.freeze_panes = "A2"


def write_combined_sheet(writer, df: pd.DataFrame, dates: list[str],
                         sheet_name: str = "Отчёт АВПО"):
    from openpyxl.styles import PatternFill, Font, Alignment
    from openpyxl.utils import get_column_letter

    ws = writer.book.create_sheet(sheet_name[:31], 0)   # ← имя листа из параметра

    # 1) базовые колонки — объединяем по вертикали
    for i, name in enumerate(BASE_COLS, start=1):
        ws.cell(row=1, column=i, value=name)
        ws.merge_cells(start_row=1, start_column=i, end_row=2, end_column=i)
        ws.cell(row=2, column=i, value=None)

    # 2) блоки по датам
    col = len(BASE_COLS) + 1
    for iso in dates:
        d = datetime.strptime(iso, "%Y-%m-%d").date()
        ws.cell(row=1, column=col, value=f"Проверка на {d:%d.%m.%Y}")
        ws.merge_cells(start_row=1, start_column=col, end_row=1, end_column=col + 2)
        for j, sub in enumerate([
            f"Статус на {d:%d.%m.%Y}",
            "Отчет об устранении (с указанием даты)",
            "Причина возникновения события",
        ]):
            ws.cell(row=2, column=col + j, value=sub)
        col += 3

    # 3) данные
    for r_i, row in enumerate(df.itertuples(index=False), start=3):
        for c_i, val in enumerate(row, start=1):
            ws.cell(row=r_i, column=c_i, value=None if pd.isna(val) else val)

    # 4) оформление шапки
    head_fill = PatternFill("solid", fgColor="DDEBF7")
    for row in (1, 2):
        for c_i in range(1, col):
            cell = ws.cell(row=row, column=c_i)
            cell.fill = head_fill
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center",
                                       wrap_text=True)

    # 5) ширины и закрепление
    for i in range(1, col):
        ws.column_dimensions[get_column_letter(i)].width = 28 if i <= 4 else 24
    ws.freeze_panes = "E3"
    ws.auto_filter.ref = f"A2:{get_column_letter(col - 1)}{ws.max_row}"


def build_excel(conn, out_path, report_date: str | None = None) -> Path:
    out_path = Path(out_path)

    n = conn.execute("SELECT COUNT(*) FROM hosts").fetchone()[0]
    if n == 0:
        raise RuntimeError("В БД нет данных — сначала загрузите CSV.")

    # все месяцы, за которые есть снимки
    months = pd.read_sql_query(
        "SELECT DISTINCT substr(report_date, 1, 7) AS ym "
        "FROM snapshots ORDER BY ym", conn)["ym"].tolist()
    if not months:
        raise RuntimeError("В БД нет ни одного снимка (snapshots).")

    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        for ym in months:
            y, m = map(int, ym.split("-"))
            combined, dates = build_combined_frame(conn, y, m)
            if combined.empty or not dates:
                continue
            sheet_name = RU_MONTHS[m - 1]  # "Ноябрь", "Декабрь"
            write_combined_sheet(writer, combined, dates,
                                 sheet_name=sheet_name)

        # если по каким-то причинам ни один лист не создался — создаём пустой
        if not writer.book.sheetnames:
            ws = writer.book.create_sheet("Отчёт")
            ws.cell(row=1, column=1, value="Нет данных")

    return out_path


# ------------------------------- CLI ---------------------------------------

def _cmd_report(csv_glob: str, out: str, date_str: str | None):
    conn = sqlite3.connect(DB_PATH)
    init_db(conn)
    files = sorted(glob.glob(csv_glob))
    if not files:
        print(f"Не найдено CSV по маске: {csv_glob}")
        sys.exit(1)
    if date_str:
        try:
            d = parse_date_arg(date_str)
            report_date = d.isoformat()
        except ValueError as e:
            print(e); sys.exit(1)
        if d.weekday() != 3:
            print(f"Внимание: {d:%d.%m.%Y} — не четверг.")
    else:
        report_date = None
    for f in files:
        t, n, u = load_csv_to_db(f, conn, report_date=report_date)
        print(f"{os.path.basename(f)}: всего {t}, новых {n}, обновлено {u}")
    path = build_excel(conn, out)
    conn.close()
    print(f"Готово: {path}")


# ------------------------------- GUI ---------------------------------------

def run_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk

    root = tk.Tk()
    root.title("АВПО: формирование отчёта")
    root.geometry("760x500")

    conn = sqlite3.connect(DB_PATH)
    init_db(conn)

    top = ttk.Frame(root); top.pack(fill="x", padx=8, pady=6)
    date_var = tk.StringVar(value=last_thursday().strftime("%d.%m.%Y"))
    ttk.Label(top, text="Дата проверки (четверг):").pack(side="left")
    ttk.Entry(top, textvariable=date_var, width=12).pack(side="left", padx=4)

    def get_date():
        try:
            d = parse_date_arg(date_var.get())
            if d.weekday() != 3:
                if not messagebox.askyesno("Внимание",
                        f"{d:%d.%m.%Y} — не четверг. Продолжить?"):
                    return None
            return d.isoformat()
        except ValueError:
            messagebox.showerror("Ошибка", "Дата в формате ДД.ММ.ГГГГ")
            return None

    log = tk.Text(root, height=16, width=94, state="disabled")
    log.pack(padx=8, pady=6, fill="both", expand=True)

    def logln(msg):
        log.configure(state="normal")
        log.insert("end", f"[{datetime.now():%H:%M:%S}] {msg}\n")
        log.see("end")
        log.configure(state="disabled")
        root.update_idletasks()

    def add_csv():
        rd = get_date()
        if not rd: return
        paths = filedialog.askopenfilenames(
            title="Выберите CSV из KSC",
            filetypes=[("CSV", "*.csv"), ("Все файлы", "*.*")],
            initialdir=str(DATA_DIR))
        if not paths: return
        for p in paths:
            try:
                t, n, u = load_csv_to_db(p, conn, report_date=rd)
                logln(f"{os.path.basename(p)} @ {rd}: всего {t}, новых {n}, обновлено {u}")
            except Exception as e:
                logln(f"ОШИБКА {os.path.basename(p)}: {e}")

    def load_folder():
        rd = get_date()
        if not rd: return
        files = sorted(DATA_DIR.glob("*.csv"))
        if not files:
            logln("В папке data/ нет CSV"); return
        for p in files:
            try:
                t, n, u = load_csv_to_db(p, conn, report_date=rd)
                logln(f"{p.name} @ {rd}: всего {t}, новых {n}, обновлено {u}")
            except Exception as e:
                logln(f"ОШИБКА {p.name}: {e}")

    def make_report():
        try:
            out = OUT_DIR / f"Отчет_АВПО_{datetime.now():%Y%m%d_%H%M}.xlsx"
            path = build_excel(conn, out)
            logln(f"Отчёт сохранён: {path}")
            if messagebox.askyesno("Готово",
                    f"Отчёт сохранён:\n{path}\n\nОткрыть папку?"):
                if sys.platform.startswith("win"):
                    os.startfile(OUT_DIR)  # type: ignore
                elif sys.platform == "darwin":
                    os.system(f'open "{OUT_DIR}"')
                else:
                    os.system(f'xdg-open "{OUT_DIR}"')
        except Exception as e:
            logln(f"ОШИБКА формирования отчёта: {e}")

    def clear_db():
        if not messagebox.askyesno("Подтверждение",
                "Очистить БД (удалятся и хосты, и история снимков)?"):
            return
        for t in ("hosts", "snapshots", "load_log"):
            conn.execute(f"DELETE FROM {t}")
        conn.commit()
        logln("БД очищена.")

    def show_history():
        rows = conn.execute(
            "SELECT report_date, COUNT(*) FROM snapshots "
            "GROUP BY report_date ORDER BY report_date").fetchall()
        if not rows:
            logln("История пуста."); return
        logln("Снимки в БД (дата: кол-во хостов):")
        for d, c in rows:
            logln(f"  {d}: {c}")

    btns = ttk.Frame(root); btns.pack(fill="x", padx=8, pady=6)
    ttk.Button(btns, text="Добавить CSV…",         command=add_csv).pack(side="left")
    ttk.Button(btns, text="Загрузить папку data/",  command=load_folder).pack(side="left", padx=6)
    ttk.Button(btns, text="Сформировать отчёт",     command=make_report).pack(side="left")
    ttk.Button(btns, text="История снимков",        command=show_history).pack(side="left", padx=6)
    ttk.Button(btns, text="Очистить БД",            command=clear_db).pack(side="right")

    logln(f"БД:            {DB_PATH}")
    logln(f"Папка CSV:     {DATA_DIR}")
    logln(f"Папка отчётов: {OUT_DIR}")
    logln("Порядок: 1) дата четверга  2) загрузить CSV  3) сформировать отчёт")

    def import_edits():
        p = filedialog.askopenfilename(
            title="Выберите ранее сохранённый отчёт (xlsx) с вашими правками",
            filetypes=[("Excel", "*.xlsx")],
            initialdir=str(OUT_DIR))
        if not p:
            return
        try:
            n = import_manual_edits_from_excel(p, conn)
            logln(f"Импортировано ручных правок: {n} (из {os.path.basename(p)})")
        except Exception as e:
            logln(f"ОШИБКА импорта правок: {e}")

    ttk.Button(btns, text="Импортировать правки из Excel",
               command=import_edits).pack(side="left", padx=6)

    root.mainloop()
    conn.close()


# ------------------------------- Entry -------------------------------------

if __name__ == "__main__":
    ap = argparse.ArgumentParser(description="Отчёт АВПО из CSV KSC")
    ap.add_argument("--csv",  help="Маска CSV, напр. 'data/*.csv'")
    ap.add_argument("--out",  help="Путь к Excel-отчёту")
    ap.add_argument("--date", help="Дата проверки (ДД.ММ.ГГГГ или ГГГГ-ММ-ДД)")
    args = ap.parse_args()
    if args.csv and args.out:
        _cmd_report(args.csv, args.out, args.date)
    else:
        run_gui()