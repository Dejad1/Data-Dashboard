"""Load weekly report files from the inbox into the databank, then rebuild
the workbook.

    python ingest.py                 # use this folder as the data root
    python ingest.py --root demo     # use another data root
    python ingest.py --no-build      # load only, don't rebuild the workbook

Files can be .xlsx, .xlsm or .csv, with any sheet names, header rows below
title rows and columns in any order.  Columns are matched by name using
column_aliases.yaml.  Every rejected or suspicious row is written to the
DATA QUALITY sheet with the reason.
"""
from __future__ import annotations

import argparse
import datetime as dt
import re
import shutil
import sys
from dataclasses import dataclass, field
from pathlib import Path

import numpy as np
import pandas as pd
import yaml

import service_reports
import transport_reports
import zone_master
from config import ALIASES_FILE, COMMUNITY_SERVICE_TYPES, HERE, Paths, load_settings, week_ending
from databank import Databank, Masters, norm_key, yes

COUNT_FIELDS = ["male", "female", "children"]

# --------------------------------------------------------------------------
# Header handling
# --------------------------------------------------------------------------


def squash(text) -> str:
    """Header comparison key: lower case, letters and digits only."""
    return re.sub(r"[^a-z0-9]", "", str(text).lower())


def load_aliases(path: Path = ALIASES_FILE) -> dict[str, str]:
    with open(path, encoding="utf-8") as fh:
        raw = yaml.safe_load(fh)
    lookup = {}
    for canonical, aliases in raw.items():
        lookup[squash(canonical)] = canonical
        for a in aliases or []:
            lookup[squash(a)] = canonical
    return lookup


def find_header(raw: pd.DataFrame, aliases: dict[str, str], scan_rows: int = 30):
    """Return (header row index, {column index: field}) for the best header row."""
    best = (None, {})
    for r in range(min(scan_rows, len(raw))):
        mapping = {}
        for c, v in enumerate(raw.iloc[r].tolist()):
            if v is None or (isinstance(v, float) and np.isnan(v)):
                continue
            f = aliases.get(squash(v))
            if f and f not in mapping.values():
                mapping[c] = f
        if len(mapping) > len(best[1]):
            best = (r, mapping)
    if len(best[1]) < 3:
        return None, {}
    return best


def unknown_headers(raw: pd.DataFrame, header_row: int, mapping: dict) -> list[str]:
    out = []
    for c, v in enumerate(raw.iloc[header_row].tolist()):
        if c not in mapping and v is not None and not (isinstance(v, float) and np.isnan(v)):
            if str(v).strip():
                out.append(str(v).strip())
    return out


SKIP_SHEETS = ("how to", "instruction", "readme", "read me", "notes", "example")


def read_sheets(path: Path) -> list[tuple[str, pd.DataFrame]]:
    """Every non-empty sheet, except instruction sheets such as the templates' 'How to fill'."""
    if path.suffix.lower() == ".csv":
        return [("csv", pd.read_csv(path, header=None, dtype=object, keep_default_na=False))]
    book = pd.read_excel(path, sheet_name=None, header=None, dtype=object)
    return [(name, df) for name, df in book.items()
            if not df.empty and not any(w in name.lower() for w in SKIP_SHEETS)]


# --------------------------------------------------------------------------
# Stream detection
# --------------------------------------------------------------------------


def hint_from_text(text: str) -> str | None:
    t = text.lower()
    if "chop" in t or "covenant" in t:
        return "CHOP"
    if "midweek" in t or "mid week" in t or "mid-week" in t:
        return "MIDWEEK"
    if "community" in t:
        return "COMMUNITY"
    if "wsf" in t or "saturday" in t:
        return "WSF"
    return None


def detect_stream(fields: set[str], file_name: str, sheet: str) -> str | None:
    if "church" in fields:
        return "COMMUNITY"
    if "cell" in fields or {"cells_total", "cells_reported"} <= fields:
        return "WSF"
    if "zone" in fields:
        if "stream" in fields:
            return "ZONAL_MIXED"
        for text in (sheet, file_name):
            h = hint_from_text(text)
            if h in ("CHOP", "MIDWEEK", "WSF"):
                return h
    return None


# --------------------------------------------------------------------------
# Value parsing
# --------------------------------------------------------------------------

DATE_FORMATS = ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%Y-%m-%d", "%Y/%m/%d", "%d %b %Y", "%d-%b-%Y",
                "%d %B %Y", "%b %d, %Y", "%d/%m/%y", "%Y-%m-%d %H:%M:%S"]


def parse_date(v):
    if v is None or (isinstance(v, float) and np.isnan(v)) or v == "":
        return None
    if isinstance(v, (pd.Timestamp, dt.datetime)):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, (int, float, np.integer, np.floating)) and 20000 < float(v) < 80000:
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(v))).date()
    s = str(v).strip()
    if re.fullmatch(r"\d{5}(\.0+)?", s):
        return parse_date(float(s))
    for fmt in DATE_FORMATS:
        try:
            return dt.datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    return None


def parse_dates(series: pd.Series) -> pd.Series:
    uniq = {v: parse_date(v) for v in pd.unique(series.astype(object))} if len(series) else {}
    return series.astype(object).map(uniq)


def to_number(series: pd.Series) -> pd.Series:
    s = series.astype(object).map(lambda v: str(v).replace(",", "").replace("₦", "").strip()
                                  if v is not None and not (isinstance(v, float) and np.isnan(v)) else "")
    s = s.replace({"": "0", "-": "0"})
    return pd.to_numeric(s, errors="coerce")


# --------------------------------------------------------------------------
# Row validation
# --------------------------------------------------------------------------


@dataclass
class Batch:
    """A parsed block of rows plus its accumulated rejection reasons."""
    df: pd.DataFrame
    source_file: str
    sheet: str
    stream: str
    reasons: pd.Series = None
    issues: list = field(default_factory=list)

    def __post_init__(self):
        self.reasons = pd.Series([""] * len(self.df), index=self.df.index, dtype=object)

    def reject(self, mask: pd.Series, reason: str) -> None:
        mask = mask.fillna(False) & (self.reasons == "")
        self.reasons[mask] = reason

    def warn(self, mask: pd.Series, issue: str, detail_fn) -> None:
        for idx in self.df.index[mask.fillna(False)]:
            row = self.df.loc[idx]
            self.issues.append(self._issue("Warning", issue, detail_fn(row), row))

    def _issue(self, severity, issue, detail, row):
        we = row.get("week_ending")
        return {"source_file": self.source_file, "sheet": self.sheet, "source_row": int(row["_row"]),
                "stream": self.stream, "severity": severity, "issue": issue, "detail": detail,
                "week_ending": we.isoformat() if isinstance(we, dt.date) else None}

    def accepted(self) -> pd.DataFrame:
        bad = self.reasons != ""
        for idx in self.df.index[bad]:
            row = self.df.loc[idx]
            self.issues.append(self._issue("Rejected", self.reasons[idx], describe(row), row))
        return self.df.loc[~bad].copy()


def describe(row) -> str:
    parts = []
    for k in ("area", "zone", "cell", "church", "service_type", "vehicle_id", "category", "cost_type",
              "date_raw"):
        v = row.get(k)
        raw = row.get(k + "_raw")
        if (v is None or (isinstance(v, float) and np.isnan(v))) and raw not in (None, "", "None", "nan"):
            v = raw
        if v is not None and not (isinstance(v, float) and np.isnan(v)) and str(v) != "":
            parts.append(f"{k}={v}")
    return "; ".join(parts)[:250]


def common_checks(b: Batch, settings: dict, cal_start: dt.date, cal_end: dt.date,
                  counts: list[str]) -> None:
    df = b.df
    if "date" not in df:
        b.reject(pd.Series(True, index=df.index), "No date column")
        df["report_date"] = None
    else:
        df["date_raw"] = df["date"].astype(str).str.replace(" 00:00:00", "", regex=False)
        df["report_date"] = parse_dates(df["date"])
    b.reject(df["report_date"].isna(), "Date missing or not recognised")
    wd = settings["week_end_day"]
    df["week_ending"] = df["report_date"].map(lambda d: week_ending(d, wd) if d else None)
    out_of_range = df["week_ending"].map(lambda w: w is not None and not (cal_start <= w <= cal_end))
    b.reject(out_of_range, "Date outside the Calendar range")
    for c in counts:
        raw = df[c] if c in df else pd.Series([0] * len(df), index=df.index)
        num = to_number(raw)
        b.reject(num.isna(), f"{c} is not a number")
        b.reject(num < 0, f"{c} is negative")
        df[c] = num.fillna(0).clip(lower=0).round().astype("int64")


def total_checks(b: Batch) -> None:
    df = b.df
    adult = df["male"] + df["female"]
    if "adult_total" in df:
        given = to_number(df["adult_total"])
        mism = given.notna() & (df["adult_total"].astype(str).str.strip() != "") & (given != adult)
        b.warn(mism & (b.reasons == ""), "Adult total mismatch",
               lambda r: f"file says {r['adult_total']}, Male+Female = {r['male'] + r['female']}; computed value used")
    if "grand_total" in df:
        given = to_number(df["grand_total"])
        mism = given.notna() & (df["grand_total"].astype(str).str.strip() != "") & (given != adult + df["children"])
        b.warn(mism & (b.reasons == ""), "Grand total mismatch",
               lambda r: f"file says {r['grand_total']}, computed = {r['male'] + r['female'] + r['children']}; computed value used")


def dedupe_in_file(b: Batch, df: pd.DataFrame, key: list[str]) -> pd.DataFrame:
    dup = df.duplicated(subset=key, keep="last")
    for _, row in df[dup].iterrows():
        b.issues.append(b._issue("Warning", "Duplicate row in file",
                                 describe(row) + " (the last copy was kept)", row))
    return df[~dup]


def match(series: pd.Series, lookup: dict) -> pd.Series:
    return series.astype(object).map(lambda v: lookup.get(norm_key(v)) if v is not None else None)


def resolve_zone(b: Batch, m: Masters) -> None:
    df = b.df
    df["zone_raw"] = df["zone"].astype(str)
    df["zone"] = match(df["zone"], m.zone_by_key)
    is_church = df["zone"].isna() & match(df["zone_raw"], m.church_by_key).notna()
    b.reject(is_church, "Community Church row in a zonal file (kept separate by design)")
    b.reject(df["zone"].isna(), "Unknown Zone (not in MASTER_ZONES)")
    master_area = df["zone"].map(m.zone_area)
    if "area" in df:
        given = match(df["area"], m.area_by_key)
        diff = df["zone"].notna() & given.notna() & (given != master_area)
        b.warn(diff & (b.reasons == ""), "Area does not match Zone master",
               lambda r: f"file Area '{r['area']}' but {r['zone']} belongs to {m.zone_area.get(r['zone'])}; master used")
        unknown_area = given.isna() & df["area"].astype(str).str.strip().ne("") & df["zone"].notna()
        b.warn(unknown_area & (b.reasons == ""), "Unknown Area name",
               lambda r: f"Area '{r['area']}' not in MASTER_AREAS; Area taken from the Zone master")
    df["area"] = master_area


# --------------------------------------------------------------------------
# Stream normalisers
# --------------------------------------------------------------------------


def normalise_zonal(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, COUNT_FIELDS)
    resolve_zone(b, m)
    if b.stream == "CHOP":
        no_chop = b.df["zone"].notna() & ~b.df["zone"].map(m.zone_has_chop).fillna(False).astype(bool)
        b.reject(no_chop, "CHOP report from a zone without CHOP (Has_CHOP = N)")
    total_checks(b)
    df = b.accepted()
    return dedupe_in_file(b, df, ["week_ending", "zone"])


def normalise_wsf_cells(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, COUNT_FIELDS)
    df = b.df
    df["cell_raw"] = df["cell"].astype(str)
    df["cell"] = match(df["cell"], m.cell_by_key)
    b.reject(df["cell"].isna(), "Unknown Cell (not in MASTER_CELLS)")
    master_zone = df["cell"].map(m.cell_zone)
    if "zone" in df:
        given = match(df["zone"], m.zone_by_key)
        diff = df["cell"].notna() & given.notna() & (given != master_zone)
        b.warn(diff & (b.reasons == ""), "Zone does not match Cell master",
               lambda r: f"file Zone '{r['zone']}' but {r['cell']} belongs to {m.cell_zone.get(r['cell'])}; master used")
    df["zone"] = master_zone
    df["area"] = df["zone"].map(m.zone_area)
    op = set(m.operational_cells["Cell"])
    b.warn(df["cell"].notna() & ~df["cell"].isin(op) & (b.reasons == ""), "Report from a non-operational cell",
           lambda r: f"{r['cell']} is marked non-operational in MASTER_CELLS; loaded and counted")
    total_checks(b)
    out = b.accepted()
    return dedupe_in_file(b, out, ["week_ending", "cell"])


def normalise_wsf_zone(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, COUNT_FIELDS + ["cells_total", "cells_reported"])
    resolve_zone(b, m)
    b.reject(b.df["cells_reported"] > b.df["cells_total"], "Cells reported is more than cells total")
    total_checks(b)
    out = b.accepted()
    return dedupe_in_file(b, out, ["week_ending", "zone"])


SERVICE_WORDS = [("sun", "Sunday"), ("chop", "CHOP"), ("covenant", "CHOP"), ("mid", "Midweek"),
                 ("wsf", "WSF"), ("saturday", "WSF")]


def service_of(v) -> str | None:
    t = str(v).lower()
    for w, s in SERVICE_WORDS:
        if w in t:
            return s
    return None


def normalise_community(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, COUNT_FIELDS)
    df = b.df
    df["church_raw"] = df["church"].astype(str)
    df["church"] = match(df["church"], m.church_by_key)
    is_zone = df["church"].isna() & match(df["church_raw"], m.zone_by_key).notna()
    b.reject(is_zone, "Zonal row in a Community Church file")
    b.reject(df["church"].isna(), "Unknown Community Church (not in MASTER_COMMUNITY)")
    if "service_type" in df:
        df["service_type"] = df["service_type"].map(service_of)
    else:
        guess = service_of(b.sheet) or service_of(b.source_file)
        df["service_type"] = guess
    b.reject(df["service_type"].isna(), "Service type missing or not one of " + ", ".join(COMMUNITY_SERVICE_TYPES))
    total_checks(b)
    out = b.accepted()
    return dedupe_in_file(b, out, ["week_ending", "church", "service_type"])


NORMALISERS = {
    "WSF_CELLS": normalise_wsf_cells,
    "WSF_ZONE": normalise_wsf_zone,
    "MIDWEEK": normalise_zonal,
    "CHOP": normalise_zonal,
    "COMMUNITY": normalise_community,
}

# --------------------------------------------------------------------------
# WSF zone aggregation
# --------------------------------------------------------------------------


def rebuild_wsf_zone_weeks(db: Databank, m: Masters, weeks: set[str], loaded_on: str) -> None:
    """Aggregate cell rows to one row per active zone per week.

    Every active zone gets a row, so a zone where no cell reported shows
    Cells_Reported = 0 instead of disappearing.  Rows that came from a
    zone-level file are kept when no cell data exists for that zone."""
    op = m.operational_cells.groupby("Zone_Name").size()
    for week in sorted(weeks):
        cells = db.read("wsf_cells", "week_ending = ?", (week,))
        zone_rows = db.read("wsf_zone", "week_ending = ? AND origin = 'zone'", (week,))
        g = cells.groupby("zone").agg(cells_reported=("cell", "nunique"), male=("male", "sum"),
                                      female=("female", "sum"), children=("children", "sum"),
                                      report_date=("report_date", "max"),
                                      source_file=("source_file", "max"))
        extra_nonop = cells[~cells["cell"].isin(set(m.operational_cells["Cell"]))].groupby("zone").size()
        z = m.active_zones.rename(columns={"Zone_Name": "zone", "Area_Name": "area"}).set_index("zone")
        z = z.join(g, how="left")
        z["cells_total"] = op.reindex(z.index).fillna(0) + extra_nonop.reindex(z.index).fillna(0)
        for c in ("cells_reported", "male", "female", "children"):
            z[c] = z[c].fillna(0).astype("int64")
        z["cells_total"] = z["cells_total"].astype("int64")
        z["report_date"] = z["report_date"].fillna(week)
        z["source_file"] = z["source_file"].fillna("(no cell reports)")
        z = z.reset_index()
        keep_zone_level = set(zone_rows["zone"]) - set(g.index)
        z = z[~z["zone"].isin(keep_zone_level)]
        z["week_ending"] = week
        z["origin"] = "cells"
        z["loaded_on"] = loaded_on
        db.execute("DELETE FROM wsf_zone WHERE week_ending = ? AND origin = 'cells'", (week,))
        db.upsert("WSF_ZONE", z)


def archive_wsf_cells(db: Databank, paths: Paths, years: set[str]) -> None:
    for y in sorted(years):
        df = db.read("wsf_cells", "substr(week_ending, 1, 4) = ?", (y,))
        df.sort_values(["week_ending", "area", "zone", "cell"]).to_csv(
            paths.archive / f"WSF_cells_{y}.csv", index=False)


# --------------------------------------------------------------------------
# Driver
# --------------------------------------------------------------------------


def calendar_range(settings) -> tuple[dt.date, dt.date]:
    start = week_ending(dt.date(settings["calendar_start_year"], 1, 1), settings["week_end_day"])
    end = start + dt.timedelta(weeks=52 * settings["calendar_years"] + settings["calendar_years"] // 5 + 1)
    return start, end


def to_iso(df: pd.DataFrame) -> pd.DataFrame:
    for c in ("week_ending", "report_date"):
        df[c] = df[c].map(lambda d: d.isoformat() if isinstance(d, dt.date) else d)
    return df


def split_zonal_mixed(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    s = df["stream"].map(lambda v: hint_from_text(str(v)))
    return {k: df[s == k].copy() for k in ("MIDWEEK", "CHOP", "WSF") if (s == k).any()} | (
        {"UNKNOWN": df[~s.isin(["MIDWEEK", "CHOP", "WSF"])].copy()} if (~s.isin(["MIDWEEK", "CHOP", "WSF"])).any() else {})


TRANSPORT_TABLES = {
    # a re-sent report replaces the same report for the same week
    "transport_runs": (transport_reports.RUN_COLS, ["week_ending", "report"]),
    "transport_costs": (transport_reports.COST_COLS, ["week_ending", "report"]),
    "transport_budget": (transport_reports.BUDGET_COLS, ["week_ending", "report"]),
}


SERVICE_TABLE = {"WSF": "wsf_zone", "MIDWEEK": "midweek", "CHOP": "chop"}
SERVICE_TYPE = {"WSF": "WSF", "MIDWEEK": "Midweek", "CHOP": "CHOP"}


def _replace_week(db: Databank, table: str, where: str, params: tuple, df: pd.DataFrame, key: str) -> tuple:
    """Report-style load: this report replaces the same report for the same week."""
    replaced = db.execute(f"SELECT COUNT(*) FROM {table} WHERE {where}", params)[0][0]
    db.execute(f"DELETE FROM {table} WHERE {where}", params)
    db.con.commit()
    loaded, _ = db.upsert(key, df) if key else (0, 0)
    return loaded, replaced


def load_service_export(res, db: Databank, m: Masters, file_name: str, sheet: str, settings, loaded_on: str,
                        summary: dict) -> None:
    """Load a reporting-system export (WSF cell service, Midweek, CHOP) for one week."""
    label = res.stream
    issues = [dict(i, logged_on=loaded_on, source_file=file_name, sheet=sheet, stream=label) for i in res.issues]
    if res.report_date is None or res.zones is None:
        db.log_issues(issues)
        return
    we = week_ending(res.report_date, settings["week_end_day"]).isoformat()
    for i in issues:
        i["week_ending"] = we
    # refresh the zone master from the system's flags (new zones, CHOP zones, Community Churches)
    master, flag_issues = zone_master.apply_export_flags(m.zones, res.flags, res.stream)
    master.to_csv(m.folder / "MASTER_ZONES.csv", index=False)
    summary["_masters_changed"] = True
    issues += [dict(i, logged_on=loaded_on, source_file=file_name, sheet=sheet, stream="ZONE MASTER",
                    week_ending=we, source_row=None) for i in flag_issues]
    base = dict(week_ending=we, report_date=res.report_date.isoformat(), source_file=file_name, loaded_on=loaded_on)
    z = res.zones.assign(**base)
    table = SERVICE_TABLE[res.stream]
    if res.stream == "WSF":
        z["origin"] = "zone"
        loaded, replaced = _replace_week(db, table, "week_ending = ?", (we,), z, "WSF_ZONE")
    else:
        loaded, replaced = _replace_week(db, table, "week_ending = ?", (we,), z, res.stream)
    c = res.community.rename(columns={"zone": "church"}).assign(service_type=SERVICE_TYPE[res.stream], **base)
    _replace_week(db, "community", "week_ending = ? AND service_type = ?", (we, SERVICE_TYPE[res.stream]), c,
                  "COMMUNITY")
    a = res.areas.assign(stream=res.stream, **base)
    db.execute("DELETE FROM service_areas WHERE stream = ? AND week_ending = ?", (res.stream, we))
    cols = ["stream", "week_ending", "report_date", "area", "area_no", "total_zones", "zones_with_report",
            "chop_zones", "community_zones", "cells_total", "cells_reported", "male", "female", "children",
            "source_file", "loaded_on"]
    db.con.executemany(f"INSERT INTO service_areas ({','.join(cols)}) VALUES ({','.join('?' * len(cols))})",
                       a[cols].astype(object).values.tolist())
    db.con.commit()
    db.log_issues(issues)
    rejected = sum(1 for i in issues if i["severity"] == "Rejected")
    db.log_load(loaded_on=loaded_on, source_file=file_name, sheet=sheet, stream=label, rows_loaded=loaded,
                rows_replaced=replaced, rows_rejected=rejected)
    s = summary.setdefault(label, [0, 0, 0])
    s[0] += loaded + len(c)
    s[1] += replaced
    s[2] += rejected


def load_sunday(result, db: Databank, m: Masters, file_name: str, sheet: str, settings, loaded_on: str,
                summary: dict) -> None:
    """Community Church Sunday workbook sheet: every Sunday in it replaces that Sunday's rows."""
    rows, issues = result
    issues = [dict(i, logged_on=loaded_on, source_file=file_name, sheet=sheet, stream="COMMUNITY SUNDAY")
              for i in issues]
    if len(rows):
        unknown = rows[rows.area.isna()]
        for _, r in unknown.iterrows():
            issues.append({"logged_on": loaded_on, "source_file": file_name, "sheet": sheet,
                           "source_row": int(r.source_row), "stream": "COMMUNITY SUNDAY", "severity": "Warning",
                           "issue": "Church Area not found", "detail": r.church})
        community = zone_master.refresh_community(m.community, rows)
        community.to_csv(m.folder / "MASTER_COMMUNITY.csv", index=False)
        summary["_masters_changed"] = True
        rows = rows.assign(week_ending=rows.report_date.map(lambda d: week_ending(d, settings["week_end_day"])
                                                           .isoformat()),
                           report_date=rows.report_date.map(lambda d: d.isoformat()), service_type="Sunday",
                           source_file=file_name, loaded_on=loaded_on)
        replaced = 0
        for we in sorted(rows.week_ending.unique()):
            replaced += db.execute("SELECT COUNT(*) FROM community WHERE week_ending = ? AND service_type = "
                                   "'Sunday'", (we,))[0][0]
            db.execute("DELETE FROM community WHERE week_ending = ? AND service_type = 'Sunday'", (we,))
        db.con.commit()
        loaded, _ = db.upsert("COMMUNITY", rows)
        s = summary.setdefault("COMMUNITY SUNDAY", [0, 0, 0])
        s[0] += loaded
        s[1] += replaced
    db.log_issues(issues)


def update_zone_master(zone_sheets: dict, db: Databank, m: Masters, file_name: str, loaded_on: str,
                       summary: dict) -> None:
    """Rebuild masters/MASTER_ZONES.csv from a ZONE DETAILS workbook."""
    seen = {z: a for z, a in db.execute("SELECT DISTINCT zone_code, area FROM transport_runs WHERE zone_code <> ''")
            if z}
    master, issues = zone_master.build(zone_sheets, m, m.zones, seen)
    path = m.folder / "MASTER_ZONES.csv"
    master.to_csv(path, index=False)
    db.log_issues([dict(i, logged_on=loaded_on, source_file=file_name, stream="ZONE MASTER", week_ending=None)
                   for i in issues])
    summary["ZONE MASTER"] = [len(master), 0, sum(1 for i in issues if i["severity"] == "Rejected")]
    summary["_masters_changed"] = True


def tag_coaster_locations(db: Databank, week: str) -> None:
    """Coasters run from Hubs unless the fuel schedule marks the coaster ZONE (a loading bay)."""
    db.execute("""
        UPDATE transport_runs SET location_type = CASE WHEN EXISTS (
            SELECT 1 FROM transport_costs c
            WHERE c.week_ending = transport_runs.week_ending AND c.hub_status = 'ZONE'
              AND transport_runs.vehicle_id <> ''
              AND (c.vehicle_id = transport_runs.vehicle_id OR c.vehicle_reg = transport_runs.vehicle_id))
            THEN 'Loading Bay' ELSE 'Hub' END
        WHERE week_ending = ? AND category = 'Church Coaster'""", (week,))
    db.con.commit()


def load_transport_report(report, db: Databank, file_name: str, sheet: str, settings, loaded_on: str,
                          summary: dict) -> None:
    """Load one transport report sheet; a re-sent report replaces that week's rows for the category."""
    label = f"TRANSPORT {report.layout}"
    issues = [dict(i, logged_on=loaded_on, source_file=file_name, sheet=sheet, stream=label) for i in report.issues]
    if report.table is None:
        note = {"logged_on": loaded_on, "source_file": file_name, "sheet": sheet, "source_row": None,
                "stream": label, "severity": "Info", "issue": "Reference sheet, not loaded",
                "detail": "Recognised as a reference/summary sheet; nothing to load weekly"}
        if report.checks and report.report_date:
            # Hub Payment Summary: compare with the HUB coaster fuel loaded from the schedule sheet
            we = week_ending(report.report_date, settings["week_end_day"]).isoformat()
            paid = db.execute("SELECT COALESCE(SUM(amount), 0) FROM transport_costs WHERE week_ending = ? "
                              "AND payable = 'Y'", (we,))[0][0]
            _, file_total, _ = report.checks[0]
            same = abs(float(file_total) - float(paid)) <= 0.5
            note.update(week_ending=we, severity="Info" if same else "Warning",
                        issue="Hub Payment Summary agrees with the schedule" if same else
                        "Hub Payment Summary differs from the schedule",
                        detail=f"Summary says ₦{file_total:,.0f}; HUB coasters on the schedule add up to ₦{paid:,.0f}")
        db.log_issues([note])
        return
    if report.report_date is None:
        db.log_issues(issues)
        summary.setdefault(label, [0, 0, 0])[2] += len(report.rows)
        return
    cols, keys = TRANSPORT_TABLES[report.table]
    df = report.frame(cols)
    we = week_ending(report.report_date, settings["week_end_day"]).isoformat()
    df["week_ending"] = we
    df["report_date"] = report.report_date.isoformat()
    df["source_file"] = file_name
    df["loaded_on"] = loaded_on
    for i in issues:
        i["week_ending"] = we
    for label_, file_total, loaded_total in report.checks:
        if file_total is not None and loaded_total is not None and abs(float(file_total) - float(loaded_total)) > 0.5:
            issues.append({"logged_on": loaded_on, "source_file": file_name, "sheet": sheet, "source_row": None,
                           "stream": label, "severity": "Warning", "issue": "File total differs from its rows",
                           "detail": f"{label_}: file says {file_total:,.0f}, rows add up to {loaded_total:,.0f}",
                           "week_ending": we})
    loaded, replaced = db.replace_report(report.table, df, keys)
    if report.layout in ("COASTER_REPORT", "COASTER_FUEL"):
        tag_coaster_locations(db, we)
    rejected = sum(1 for i in issues if i["severity"] == "Rejected")
    db.log_issues(issues)
    db.log_load(loaded_on=loaded_on, source_file=file_name, sheet=sheet, stream=label, rows_loaded=loaded,
                rows_replaced=replaced, rows_rejected=rejected)
    s = summary.setdefault(label, [0, 0, 0])
    s[0] += loaded
    s[1] += replaced
    s[2] += rejected


def process_file(path: Path, db: Databank, m: Masters, settings, aliases, loaded_on: str, summary: dict,
                 wsf_weeks: set) -> bool:
    cal = calendar_range(settings)
    recognised = False
    sheets = read_sheets(path)
    zone_sheets = {name: raw for name, raw in sheets if zone_master.is_zone_details(raw)}
    if zone_sheets:
        update_zone_master(zone_sheets, db, m, path.name, loaded_on, summary)
        return True
    for sheet, raw in sheets:
        export = service_reports.read_export(raw, sheet, path.name, m)
        if export is not None:
            load_service_export(export, db, m, path.name, sheet, settings, loaded_on, summary)
            recognised = True
            continue
        sunday = service_reports.read_sunday(raw, sheet, m)
        if sunday is not None:
            load_sunday(sunday, db, m, path.name, sheet, settings, loaded_on, summary)
            recognised = True
            continue
        report = transport_reports.read(raw, sheet, path.name, m, settings)
        if report is not None:
            load_transport_report(report, db, path.name, sheet, settings, loaded_on, summary)
            recognised = True
            continue
        hdr, mapping = find_header(raw, aliases)
        if hdr is None:
            continue
        fields = set(mapping.values())
        stream = detect_stream(fields, path.name, sheet)
        if stream is None:
            db.log_issues([{"logged_on": loaded_on, "source_file": path.name, "sheet": sheet, "source_row": hdr + 1,
                            "stream": "?", "severity": "Rejected", "issue": "Stream not recognised",
                            "detail": "Headers: " + ", ".join(sorted(fields)) +
                                      ". Add a stream name (e.g. 'CHOP') to the file or sheet name."}])
            continue
        body = raw.iloc[hdr + 1:, list(mapping)].copy()
        body.columns = [mapping[c] for c in mapping]
        body["_row"] = body.index + 1
        body = body[~body.drop(columns="_row").apply(
            lambda r: all(v is None or str(v).strip() in ("", "nan", "NaT") for v in r), axis=1)]
        extra = unknown_headers(raw, hdr, mapping)
        parts = split_zonal_mixed(body) if stream == "ZONAL_MIXED" else {stream: body}
        for part_stream, part in parts.items():
            if part_stream == "UNKNOWN":
                db.log_issues([{"logged_on": loaded_on, "source_file": path.name, "sheet": sheet,
                                "source_row": int(r), "stream": "?", "severity": "Rejected",
                                "issue": "Stream value not recognised", "detail": ""} for r in part["_row"]])
                continue
            key = part_stream
            if part_stream == "WSF":
                key = "WSF_CELLS" if "cell" in part else "WSF_ZONE"
            b = Batch(part.reset_index(drop=True), path.name, sheet, key)
            clean = NORMALISERS[key](b, m, settings, cal)
            clean = to_iso(clean)
            clean["source_file"] = path.name
            clean["loaded_on"] = loaded_on
            if key == "WSF_ZONE":
                clean["origin"] = "zone"
            loaded, replaced = db.upsert(key, clean)
            if key == "WSF_CELLS":
                wsf_weeks.update(clean["week_ending"].unique())
            rejected = sum(1 for i in b.issues if i["severity"] == "Rejected")
            for i in b.issues:
                i["logged_on"] = loaded_on
            if extra:
                b.issues.append({"logged_on": loaded_on, "source_file": path.name, "sheet": sheet,
                                 "source_row": hdr + 1, "stream": key, "severity": "Info",
                                 "issue": "Columns ignored", "detail": ", ".join(extra)[:250]})
            db.log_issues(b.issues)
            db.log_load(loaded_on=loaded_on, source_file=path.name, sheet=sheet, stream=key,
                        rows_loaded=loaded, rows_replaced=replaced, rows_rejected=rejected)
            label = "WSF" if key.startswith("WSF") else key
            s = summary.setdefault(label, [0, 0, 0])
            s[0] += loaded
            s[1] += replaced
            s[2] += rejected
            recognised = True
    return recognised


def run(root: Path, build: bool = True, quiet: bool = False) -> dict:
    paths = Paths(root)
    paths.ensure()
    settings = load_settings(paths)
    Masters.write_blank(paths.masters)
    m = Masters.load(paths.masters)
    aliases = load_aliases()
    db = Databank(paths.databank)
    loaded_on = dt.datetime.now().replace(microsecond=0).isoformat(sep=" ")
    summary: dict[str, list[int]] = {}
    wsf_weeks: set[str] = set()
    # zone lists go first so the reports in the same batch can use the new zones
    files = sorted((p for p in paths.inbox.iterdir()
                    if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".csv") and not p.name.startswith("~$")),
                   key=lambda p: ("zone" not in p.name.lower(), p.name))
    for path in files:
        try:
            ok = process_file(path, db, m, settings, aliases, loaded_on, summary, wsf_weeks)
            if summary.pop("_masters_changed", False):
                m = Masters.load(paths.masters)
        except Exception as exc:  # a broken file must not stop the others
            db.log_issues([{"logged_on": loaded_on, "source_file": path.name, "sheet": "", "source_row": None,
                            "stream": "?", "severity": "Rejected", "issue": "File could not be read",
                            "detail": str(exc)[:250]}])
            ok = False
        dest = paths.processed if ok else paths.rejected_files
        stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
        shutil.move(str(path), dest / f"{stamp}_{path.name}")
        if not quiet:
            print(f"  {'OK ' if ok else 'SKIPPED'} {path.name}")
    if wsf_weeks:
        rebuild_wsf_zone_weeks(db, m, wsf_weeks, loaded_on)
        archive_wsf_cells(db, paths, {w[:4] for w in wsf_weeks})
    db.close()
    if not quiet:
        print("\nStream               Loaded   Replaced   Rejected")
        for s, (a, r, j) in sorted(summary.items()):
            print(f"{s:<18}{a:>9,}{r:>11,}{j:>11,}")
        if not files:
            print("(inbox was empty)")
    if build:
        import build_workbook
        out = build_workbook.build(root)
        if not quiet:
            print(f"\nWorkbook rebuilt: {out}")
    return summary


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=str(HERE), help="data root folder (default: this folder)")
    ap.add_argument("--no-build", action="store_true", help="load data only; don't rebuild the workbook")
    args = ap.parse_args(argv)
    run(Path(args.root).resolve(), build=not args.no_build)


if __name__ == "__main__":
    sys.exit(main())
