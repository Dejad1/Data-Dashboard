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
CREATE TABLE IF NOT EXISTS transport_runs (
    week_ending TEXT, report_date TEXT, report TEXT, category TEXT, vehicle_type TEXT, area TEXT, zone_code TEXT,
    location TEXT, vehicle_id TEXT, buses INTEGER, capacity REAL, trips INTEGER,
    male INTEGER, female INTEGER, children INTEGER, cost REAL, paid_by TEXT, status TEXT, service TEXT,
    remarks_category TEXT, remarks TEXT, in_church TEXT, location_type TEXT, source_row INTEGER,
    source_file TEXT, loaded_on TEXT);
CREATE TABLE IF NOT EXISTS transport_costs (
    week_ending TEXT, report_date TEXT, report TEXT, category TEXT, area TEXT, vehicle_id TEXT, vehicle_reg TEXT,
    cost_type TEXT,
    hub_status TEXT, trips REAL, quantity REAL, unit_price REAL, amount REAL, payable TEXT, paid_by TEXT,
    source_row INTEGER, source_file TEXT, loaded_on TEXT);
CREATE TABLE IF NOT EXISTS transport_budget (
    week_ending TEXT, report_date TEXT, report TEXT, area_no INTEGER, area TEXT, buses_allocated INTEGER,
    cost_per_bus REAL, expected_spend REAL, payable_this_week REAL, source_row INTEGER,
    source_file TEXT, loaded_on TEXT);
CREATE TABLE IF NOT EXISTS service_areas (
    stream TEXT, week_ending TEXT, report_date TEXT, area TEXT, area_no INTEGER, total_zones INTEGER,
    zones_with_report INTEGER, chop_zones INTEGER, community_zones INTEGER, cells_total INTEGER,
    cells_reported INTEGER, male INTEGER, female INTEGER, children INTEGER,
    source_file TEXT, loaded_on TEXT,
    PRIMARY KEY (stream, week_ending, area));
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
                  "male", "female", "children", "first_timers", "new_converts", "testimonies", "origin",
                  "source_file", "loaded_on"]),
    "MIDWEEK": ("midweek", ["week_ending", "zone"],
                ["week_ending", "report_date", "area", "zone", "male", "female", "children", "first_timers",
                 "new_converts", "testimonies", "source_file", "loaded_on"]),
    "CHOP": ("chop", ["week_ending", "zone"],
             ["week_ending", "report_date", "area", "zone", "male", "female", "children", "first_timers",
              "new_converts", "testimonies", "source_file", "loaded_on"]),
    "COMMUNITY": ("community", ["week_ending", "church", "service_type"],
                  ["week_ending", "report_date", "church", "service_type", "male", "female",
                   "children", "first_timers", "new_converts", "testimonies", "pastors_present",
                   "source_file", "loaded_on"]),
}


class Databank:
    def __init__(self, path: Path):
        self.path = Path(path)
        self.con = sqlite3.connect(self.path)
        self.con.executescript(SCHEMA)
        self._migrate()

    # columns added after a databank was first created: (table, column)
    ADDED_COLUMNS = [("transport_runs", "location_type"), ("transport_costs", "vehicle_reg")] + [
        (t, c) for t in ("wsf_zone", "midweek", "chop", "community")
        for c in ("first_timers", "new_converts", "testimonies")] + [("community", "pastors_present")]

    def _migrate(self) -> None:
        for table, column in self.ADDED_COLUMNS:
            have = {r[1] for r in self.con.execute(f"PRAGMA table_info({table})")}
            if column not in have:
                self.con.execute(f"ALTER TABLE {table} ADD COLUMN {column} TEXT")
        self.con.commit()

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
        df = df.copy()
        for c in cols:
            if c not in df:
                df[c] = 0 if c in ("first_timers", "new_converts", "testimonies", "pastors_present") else None
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

    def replace_report(self, table: str, df: pd.DataFrame, keys: list[str]) -> tuple[int, int]:
        """Report-style load: rows for the same week (+ category etc.) are replaced as a block.

        Transport reports have no per-row id, so a re-sent report replaces that
        week's report instead of adding to it."""
        if df.empty:
            return 0, 0
        cur = self.con.cursor()
        replaced = 0
        for combo in df[keys].drop_duplicates().itertuples(index=False):
            where = " AND ".join(f"{k} = ?" for k in keys)
            replaced += cur.execute(f"SELECT COUNT(*) FROM {table} WHERE {where}", tuple(combo)).fetchone()[0]
            cur.execute(f"DELETE FROM {table} WHERE {where}", tuple(combo))
        cols = list(df.columns)
        rows = df.astype(object).where(df.notna(), None).values.tolist()
        cur.executemany(f"INSERT INTO {table} ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})", rows)
        self.con.commit()
        return len(rows), replaced

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
    "areas": ("MASTER_AREAS.csv", ["Area_ID", "Area_Name", "Active", "Area_No", "Aliases", "In_Transport"]),
    "zones": ("MASTER_ZONES.csv", ["Zone_ID", "Zone_Name", "Area_Name", "Has_CHOP", "Active", "Zone_No", "Address",
                                   "Zone_Status", "Is_Community"]),
    "cells": ("MASTER_CELLS.csv", ["Cell", "Zone_Name", "Operational"]),
    "community": ("MASTER_COMMUNITY.csv", ["Church_ID", "Church_Name", "Location", "Active"]),
    "fleet": ("MASTER_FLEET.csv", ["Vehicle_ID", "Category", "Capacity", "Owner", "Status",
                                   "Area_Name", "Zone_Name", "Hired_By_Type"]),
}

# Columns that may be missing from older master files (filled with blanks)
OPTIONAL_COLUMNS = {"Area_No", "Aliases", "In_Transport", "Zone_No", "Address", "Zone_Status", "Is_Community"}


def norm_key(value) -> str:
    """Matching key for names from incoming files: letters and digits only, lower case.

    "IYANA- ODO", "Iyana Odo" and "IYANA-ODO" all become "iyanaodo"."""
    return re.sub(r"[^a-z0-9]", "", str(value).replace("\u2011", "-").lower())


def yes(value) -> bool:
    return str(value).strip().lower() in {"y", "yes", "true", "1", "active", "operational"}


def _similar(a: str, b: str) -> float:
    import difflib
    if (a in b or b in a) and min(len(a), len(b)) >= 10:
        return 1.0
    return difflib.SequenceMatcher(None, a, b).ratio()


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
        m.folder = Path(folder)
        m._index()
        return m

    @staticmethod
    def write_blank(folder: Path) -> None:
        folder.mkdir(parents=True, exist_ok=True)
        for fname, cols in MASTER_FILES.values():
            path = folder / fname
            if not path.exists():
                pd.DataFrame(columns=cols).to_csv(path, index=False)

    def zone_for_address(self, address, area: str) -> str | None:
        """Zone code whose master address clearly matches a loading-bay address in the same Area."""
        target = norm_key(address) if address is not None else ""
        if len(target) < 6 or not area:
            return None
        scored = [(_similar(target, norm_key(addr)), code) for code, addr in self.zone_addresses.get(area, [])]
        scored = [x for x in scored if x[0] >= 0.85]
        if not scored:
            return None
        scored.sort(reverse=True)
        # only when one zone clearly wins
        if len(scored) > 1 and scored[0][0] == scored[1][0] and scored[0][1] != scored[1][1]:
            return None
        return scored[0][1]

    def _index(self) -> None:
        z = self.zones
        self.area_by_key = {}
        for _, r in self.areas.iterrows():
            keys = [r["Area_Name"], r["Area_ID"], r["Area_No"]] + str(r["Aliases"]).split(";")
            for k in keys:
                if str(k).strip():
                    self.area_by_key[norm_key(k)] = r["Area_Name"]
        self.area_no = {r["Area_No"]: r["Area_Name"] for _, r in self.areas.iterrows() if str(r["Area_No"]).strip()}
        # Areas that take part in every service except transport (e.g. CANAANLAND 1)
        self.no_transport = set(self.areas.loc[self.areas["In_Transport"].str.strip().str.upper() == "N", "Area_Name"])
        self.zone_by_key = {norm_key(n): n for n in z["Zone_Name"]}
        self.zone_by_key.update({norm_key(i): n for i, n in zip(z["Zone_ID"], z["Zone_Name"])})
        self.zone_area = dict(zip(z["Zone_Name"], z["Area_Name"]))
        self.zone_address = dict(zip(z["Zone_Name"], z["Address"]))
        self.zone_addresses: dict[str, list] = {}
        for code, area, addr in zip(z["Zone_Name"], z["Area_Name"], z["Address"]):
            if str(addr).strip():
                self.zone_addresses.setdefault(area, []).append((code, addr))
        self.zone_has_chop = {n: yes(f) for n, f in zip(z["Zone_Name"], z["Has_CHOP"])}
        # Community Churches are zones in the reporting system but never part of the zonal universe
        self.community_zones = set(z.loc[z["Is_Community"].map(yes), "Zone_Name"])
        self.active_zones = z.loc[z["Active"].map(yes) & ~z["Zone_Name"].isin(self.community_zones),
                                  ["Zone_Name", "Area_Name"]]
        c = self.cells
        self.cell_by_key = {norm_key(n): n for n in c["Cell"]}
        self.cell_zone = dict(zip(c["Cell"], c["Zone_Name"]))
        self.operational_cells = c.loc[c["Operational"].map(yes)]
        cc = self.community
        self.church_by_key = {norm_key(n): n for n in cc["Church_Name"]}
        self.church_by_key.update({norm_key(i): n for i, n in zip(cc["Church_ID"], cc["Church_Name"])})
        f = self.fleet
        self.vehicle_by_key = {norm_key(v): v for v in f["Vehicle_ID"]}
        self.vehicle = f.set_index("Vehicle_ID").to_dict("index")
