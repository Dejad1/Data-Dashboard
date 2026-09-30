"""The databank: a single SQLite file that holds every loaded row.

The workbook is always rebuilt from here, so charts, formatting and
formulas are never damaged by appending data.  Each fact table has a primary
key equal to the natural key in the spec, so a re-sent file replaces rows
instead of double counting.
"""
from __future__ import annotations

import re
import sqlite3
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

SCHEMA = """
CREATE TABLE IF NOT EXISTS wsf_cells (
    week_ending TEXT, report_date TEXT, area TEXT, zone TEXT, cell TEXT,
    male INTEGER, female INTEGER, children INTEGER,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, cell));
CREATE TABLE IF NOT EXISTS wsf_zone (
    week_ending TEXT, report_date TEXT, area TEXT, zone TEXT,
    cells_total INTEGER, cells_reported INTEGER,
    male INTEGER, female INTEGER, children INTEGER,
    origin TEXT, source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, zone));
CREATE TABLE IF NOT EXISTS midweek (
    week_ending TEXT, report_date TEXT, area TEXT, zone TEXT,
    male INTEGER, female INTEGER, children INTEGER,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, zone));
CREATE TABLE IF NOT EXISTS chop (
    week_ending TEXT, report_date TEXT, area TEXT, zone TEXT,
    male INTEGER, female INTEGER, children INTEGER,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, zone));
CREATE TABLE IF NOT EXISTS community (
    week_ending TEXT, report_date TEXT, church TEXT, service_type TEXT,
    male INTEGER, female INTEGER, children INTEGER,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, church, service_type));
CREATE TABLE IF NOT EXISTS transport_ops (
    week_ending TEXT, report_date TEXT, category TEXT, vehicle_id TEXT,
    area TEXT, zone TEXT, hired_by_type TEXT,
    trips INTEGER, ridership INTEGER, operational TEXT,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, vehicle_id));
CREATE TABLE IF NOT EXISTS transport_fin (
    week_ending TEXT, report_date TEXT, category TEXT, area TEXT, zone TEXT,
    cost_type TEXT, amount REAL, paid_by TEXT,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (week_ending, category, area, zone, cost_type));
CREATE TABLE IF NOT EXISTS dq_log (
    logged_on TEXT, source_file TEXT, sheet TEXT, source_row INTEGER,
    stream TEXT, severity TEXT, issue TEXT, detail TEXT, week_ending TEXT);
CREATE TABLE IF NOT EXISTS load_log (
    loaded_on TEXT, source_file TEXT, sheet TEXT, stream TEXT,
    rows_loaded INTEGER, rows_replaced INTEGER, rows_rejected INTEGER);
"""

# stream -> (table, natural key columns, value columns in insert order)
TABLES = {
    "WSF_CELLS": ("wsf_cells", ["week_ending", "cell"],
                  ["week_ending", "report_date", "area", "zone", "cell", "male", "female",
                   "children", "source_file", "loaded_on"]),
    "WSF_ZONE": ("wsf_zone", ["week_ending", "zone"],
                 ["week_ending", "report_date", "area", "zone", "cells_total", "cells_reported",
                  "male", "female", "children", "origin", "source_file", "loaded_on"]),
    "MIDWEEK": ("midweek", ["week_ending", "zone"],
                ["week_ending", "report_date", "area", "zone", "male", "female", "children",
                 "source_file", "loaded_on"]),
    "CHOP": ("chop", ["week_ending", "zone"],
             ["week_ending", "report_date", "area", "zone", "male", "female", "children",
              "source_file", "loaded_on"]),
    "COMMUNITY": ("community", ["week_ending", "church", "service_type"],
                  ["week_ending", "report_date", "church", "service_type", "male", "female",
                   "children", "source_file", "loaded_on"]),
    "TRANSPORT_OPS": ("transport_ops", ["week_ending", "vehicle_id"],
                      ["week_ending", "report_date", "category", "vehicle_id", "area", "zone",
                       "hired_by_type", "trips", "ridership", "operational", "source_file",
                       "loaded_on"]),
    "TRANSPORT_FINANCE": ("transport_fin", ["week_ending", "category", "area", "zone", "cost_type"],
                          ["week_ending", "report_date", "category", "area", "zone", "cost_type",
                           "amount", "paid_by", "source_file", "loaded_on"]),
}


class Databank:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.con = sqlite3.connect(self.path)
        self.con.executescript(SCHEMA)

    def close(self) -> None:
        self.con.commit()
        self.con.close()

    # -- facts ------------------------------------------------------------
    def upsert(self, key: str, df: pd.DataFrame) -> tuple[int, int]:
        """Insert rows, replacing any row with the same natural key.

        Returns (rows written, rows that replaced an existing row)."""
        if df.empty:
            return 0, 0
        table, nat_key, cols = TABLES[key]
        cur = self.con.cursor()
        cur.execute(f"CREATE TEMP TABLE IF NOT EXISTS _incoming AS SELECT * FROM {table} WHERE 0")
        cur.execute("DELETE FROM _incoming")
        rows = df[cols].astype(object).where(df[cols].notna(), None).values.tolist()
        placeholders = ",".join("?" * len(cols))
        cur.executemany(f"INSERT INTO _incoming ({','.join(cols)}) VALUES ({placeholders})", rows)
        join = " AND ".join(f"t.{k} = i.{k}" for k in nat_key)
        replaced = cur.execute(
            f"SELECT COUNT(*) FROM {table} t JOIN _incoming i ON {join}").fetchone()[0]
        cur.executemany(f"INSERT OR REPLACE INTO {table} ({','.join(cols)}) VALUES ({placeholders})",
                        rows)
        cur.execute("DROP TABLE _incoming")
        self.con.commit()
        return len(rows), replaced

    def read(self, table: str, where: str = "", params: tuple = ()) -> pd.DataFrame:
        sql = f"SELECT * FROM {table}" + (f" WHERE {where}" if where else "")
        return pd.read_sql_query(sql, self.con, params=params)

    def execute(self, sql: str, params: tuple = ()) -> list:
        return self.con.execute(sql, params).fetchall()

    # -- logs -------------------------------------------------------------
    def log_issues(self, issues: list[dict]) -> None:
        if not issues:
            return
        cols = ["logged_on", "source_file", "sheet", "source_row", "stream", "severity", "issue",
                "detail", "week_ending"]
        self.con.executemany(
            f"INSERT INTO dq_log ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
            [[i.get(c) for c in cols] for i in issues])
        self.con.commit()

    def log_load(self, **kw) -> None:
        cols = ["loaded_on", "source_file", "sheet", "stream", "rows_loaded", "rows_replaced",
                "rows_rejected"]
        self.con.execute(f"INSERT INTO load_log ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                         [kw.get(c) for c in cols])
        self.con.commit()


# --------------------------------------------------------------------------
# Master lists (CSV files in <root>/masters, edited by staff)
# --------------------------------------------------------------------------

MASTER_FILES = {
    # Area_Name must stay the 2nd column: workbook formulas read MASTER_AREAS!B:B
    "areas": ("MASTER_AREAS.csv", ["Area_ID", "Area_Name", "Active", "Area_No", "Aliases"]),
    "zones": ("MASTER_ZONES.csv", ["Zone_ID", "Zone_Name", "Area_Name", "Has_CHOP", "Active"]),
    "cells": ("MASTER_CELLS.csv", ["Cell", "Zone_Name", "Operational"]),
    "community": ("MASTER_COMMUNITY.csv", ["Church_ID", "Church_Name", "Location", "Active"]),
    "fleet": ("MASTER_FLEET.csv", ["Vehicle_ID", "Category", "Capacity", "Owner", "Status",
                                   "Area_Name", "Zone_Name", "Hired_By_Type"]),
}

# Columns that may be missing from older master files (filled with blanks)
OPTIONAL_COLUMNS = {"Area_No", "Aliases"}


def norm_key(value) -> str:
    """Matching key for names from incoming files: letters and digits only, lower case.

    "IYANA- ODO", "Iyana Odo" and "IYANA-ODO" all become "iyanaodo"."""
    return re.sub(r"[^a-z0-9]", "", str(value).replace("\u2011", "-").lower())


def yes(value) -> bool:
    return str(value).strip().lower() in {"y", "yes", "true", "1", "active", "operational"}


@dataclass
class Masters:
    areas: pd.DataFrame
    zones: pd.DataFrame
    cells: pd.DataFrame
    community: pd.DataFrame
    fleet: pd.DataFrame

    @classmethod
    def load(cls, folder: Path) -> "Masters":
        frames = {}
        for key, (fname, cols) in MASTER_FILES.items():
            path = folder / fname
            if path.exists():
                df = pd.read_csv(path, dtype=str, keep_default_na=False)
                for c in OPTIONAL_COLUMNS & set(cols):
                    if c not in df.columns:
                        df[c] = ""
                missing = [c for c in cols if c not in df.columns]
                if missing:
                    raise ValueError(f"{fname} is missing columns: {missing}")
                frames[key] = df[cols].copy()
            else:
                frames[key] = pd.DataFrame(columns=cols)
        m = cls(**frames)
        m._index()
        return m

    @staticmethod
    def write_blank(folder: Path) -> None:
        folder.mkdir(parents=True, exist_ok=True)
        for fname, cols in MASTER_FILES.values():
            path = folder / fname
            if not path.exists():
                pd.DataFrame(columns=cols).to_csv(path, index=False)

    def _index(self) -> None:
        z = self.zones
        self.area_by_key = {}
        for _, r in self.areas.iterrows():
            keys = [r["Area_Name"], r["Area_ID"], r["Area_No"]] + str(r["Aliases"]).split(";")
            for k in keys:
                if str(k).strip():
                    self.area_by_key[norm_key(k)] = r["Area_Name"]
        self.area_no = {r["Area_No"]: r["Area_Name"] for _, r in self.areas.iterrows() if str(r["Area_No"]).strip()}
        self.zone_by_key = {norm_key(n): n for n in z["Zone_Name"]}
        self.zone_by_key.update({norm_key(i): n for i, n in zip(z["Zone_ID"], z["Zone_Name"])})
        self.zone_area = dict(zip(z["Zone_Name"], z["Area_Name"]))
        self.zone_has_chop = {n: yes(f) for n, f in zip(z["Zone_Name"], z["Has_CHOP"])}
        self.active_zones = z.loc[z["Active"].map(yes), ["Zone_Name", "Area_Name"]]
        c = self.cells
        self.cell_by_key = {norm_key(n): n for n in c["Cell"]}
        self.cell_zone = dict(zip(c["Cell"], c["Zone_Name"]))
        self.operational_cells = c.loc[c["Operational"].map(yes)]
        cc = self.community
        self.church_by_key = {norm_key(n): n for n in cc["Church_Name"]}
        self.church_by_key.update({norm_key(i): n for i, n in zip(cc["Church_ID"], cc["Church_Name"])})
        self.church_by_key.update({norm_key(i): n for i, n in zip(cc["Location"], cc["Church_Name"])})
        f = self.fleet
        self.vehicle_by_key = {norm_key(v): v for v in f["Vehicle_ID"]}
        self.vehicle = f.set_index("Vehicle_ID").to_dict("index")
