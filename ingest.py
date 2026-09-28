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

from config import (ALIASES_FILE, COMMUNITY_SERVICE_TYPES, COST_TYPES, FLEET_CATEGORIES, HIRED_BY_TYPES,
                    HERE, Paths, load_settings, week_ending)
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
    if "finance" in t or "expense" in t or "cost" in t:
        return "TRANSPORT_FINANCE"
    if "transport" in t or "fleet" in t or "bus" in t:
        return "TRANSPORT_OPS"
    return None


def detect_stream(fields: set[str], file_name: str, sheet: str) -> str | None:
    if "cost_type" in fields and "amount" in fields:
        return "TRANSPORT_FINANCE"
    if "vehicle_id" in fields and ({"trips", "ridership"} & fields):
        return "TRANSPORT_OPS"
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


def category_of(v) -> str | None:
    t = squash(v)
    if not t:
        return None
    for key, cat in [("wsf", "WSF Procured"), ("hire", "WSF Procured"), ("coaster", "Church Coaster"),
                     ("electric", "Electric Bus"), ("big", "Big Bus"), ("ft", "FT Procured"),
                     ("lt", "FT Procured")]:
        if key in t:
            return cat
    return None


def normalise_transport_ops(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, ["trips", "ridership"])
    df = b.df
    df["vehicle_raw"] = df["vehicle_id"].astype(str)
    df["vehicle_id"] = match(df["vehicle_id"], m.vehicle_by_key)
    b.reject(df["vehicle_id"].isna(), "Unknown vehicle (not in MASTER_FLEET)")
    fleet = df["vehicle_id"].map(lambda v: m.vehicle.get(v, {}))
    master_cat = fleet.map(lambda f: f.get("Category"))
    if "category" in df:
        given = df["category"].map(category_of)
        diff = df["vehicle_id"].notna() & given.notna() & (given != master_cat)
        b.warn(diff & (b.reasons == ""), "Category does not match fleet master",
               lambda r: f"file says '{r['category']}' for {r['vehicle_id']}; master category used")
    df["category"] = master_cat
    for f, master_col in (("area", "Area_Name"), ("zone", "Zone_Name"), ("hired_by_type", "Hired_By_Type")):
        master_val = fleet.map(lambda x, c=master_col: x.get(c, ""))
        if f in df:
            given = df[f].astype(object).map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip())
            if f == "area":
                given = given.map(lambda v: m.area_by_key.get(norm_key(v), v) if v else "")
            elif f == "zone":
                given = given.map(lambda v: m.zone_by_key.get(norm_key(v), v) if v else "")
            elif f == "hired_by_type":
                given = given.map(lambda v: next((h for h in HIRED_BY_TYPES if h.lower() == v.lower()), "") if v else "")
            df[f] = given.where(given != "", master_val)
        else:
            df[f] = master_val
    if "operational" in df:
        df["operational"] = df["operational"].map(lambda v: "Y" if yes(v) or str(v).strip().lower() in {"op", "running"} else "N")
    else:
        df["operational"] = np.where(df["trips"] > 0, "Y", "N")
    b.warn((df["operational"] == "N") & (df["ridership"] > 0) & (b.reasons == ""),
           "Riders on a non-operational bus", lambda r: f"{r['vehicle_id']} marked not operational but carried {r['ridership']}")
    out = b.accepted()
    return dedupe_in_file(b, out, ["week_ending", "vehicle_id"])


def normalise_transport_fin(b: Batch, m: Masters, settings, cal) -> pd.DataFrame:
    common_checks(b, settings, *cal, [])
    df = b.df
    amt = to_number(df["amount"])
    b.reject(amt.isna(), "Amount is not a number")
    b.reject(amt < 0, "Amount is negative")
    df["amount"] = amt.fillna(0).clip(lower=0).round(2)
    df["category"] = df["category"].map(category_of) if "category" in df else None
    b.reject(df["category"].isna(), "Fleet category missing or not recognised")

    def ctype(v):
        t = squash(v)
        for c in COST_TYPES:
            if squash(c) == t:
                return c
        for key, c in [("fuel", "Fuel"), ("diesel", "Fuel"), ("hire", "Hire Fee"), ("maint", "Maintenance"),
                       ("repair", "Maintenance"), ("driver", "Driver Allowance"), ("allow", "Driver Allowance"),
                       ("member", "Member Payment")]:
            if key in t:
                return c
        return str(v).strip().title() if t else None

    df["cost_type"] = df["cost_type"].map(ctype)
    b.reject(df["cost_type"].isna(), "Cost type missing")
    known = set(COST_TYPES)
    b.warn(df["cost_type"].notna() & ~df["cost_type"].isin(known) & (b.reasons == ""), "New cost type",
           lambda r: f"'{r['cost_type']}' is not in the standard list; loaded as its own cost type")
    area_raw = df["area"].astype(object).map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()) if "area" in df else pd.Series("", index=df.index)
    df["area"] = area_raw.map(lambda v: "Central" if v == "" or v.lower() == "central" else m.area_by_key.get(norm_key(v)))
    b.reject(df["area"].isna(), "Unknown Area (not in MASTER_AREAS)")
    zone_raw = df["zone"].astype(object).map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip()) if "zone" in df else pd.Series("", index=df.index)
    df["zone"] = zone_raw.map(lambda v: "" if v == "" else m.zone_by_key.get(norm_key(v)))
    b.reject(df["zone"].isna(), "Unknown Zone (not in MASTER_ZONES)")
    default_payer = {k: v[2].split(" ")[0] for k, v in FLEET_CATEGORIES.items()}
    if "paid_by" in df:
        pb = df["paid_by"].astype(object).map(lambda v: "" if v is None or (isinstance(v, float) and np.isnan(v)) else str(v).strip().title())
        df["paid_by"] = pb.where(pb != "", df["category"].map(default_payer))
    else:
        df["paid_by"] = df["category"].map(default_payer)
    out = b.accepted()
    return dedupe_in_file(b, out, ["week_ending", "category", "area", "zone", "cost_type"])


NORMALISERS = {
    "WSF_CELLS": normalise_wsf_cells,
    "WSF_ZONE": normalise_wsf_zone,
    "MIDWEEK": normalise_zonal,
    "CHOP": normalise_zonal,
    "COMMUNITY": normalise_community,
    "TRANSPORT_OPS": normalise_transport_ops,
    "TRANSPORT_FINANCE": normalise_transport_fin,
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


def process_file(path: Path, db: Databank, m: Masters, settings, aliases, loaded_on: str, summary: dict,
                 wsf_weeks: set) -> bool:
    cal = calendar_range(settings)
    recognised = False
    for sheet, raw in read_sheets(path):
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
    files = sorted(p for p in paths.inbox.iterdir()
                   if p.is_file() and p.suffix.lower() in (".xlsx", ".xlsm", ".csv") and not p.name.startswith("~$"))
    for path in files:
        try:
            ok = process_file(path, db, m, settings, aliases, loaded_on, summary, wsf_weeks)
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
