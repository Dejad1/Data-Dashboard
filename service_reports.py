"""Readers for the church reporting system's exports and the Community Church Sunday workbook.

1. "Area By Area" service exports (CELL SERVICE = WSF, MID-WEEK SERVICE, CHOP SERVICE)
   One ZoneSummary row per zone and one AreaSummary row per Area:
     * zone code = Area code + zone code, e.g. Area 004 zone 12 -> LFC0412
     * zones flagged forCommunity are Community Churches: their figures go to the
       Community tables, never into zonal totals
     * a zone registered twice in the system is added up once
     * attendance an Area reports above its zones (e.g. CHOP held at the Area
       facility) is kept as an "AREA LEVEL" row so Area totals stay right
     * minister names and phone numbers are never read
     * the report date comes from the file name ("..._-2026-09-19-2026-09-19")

2. Community Church Sunday workbook: one sheet per month, one block of columns
   per Sunday (the date sits above each block), one row per church (ZINCODE).
"""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

import pandas as pd

from databank import Masters
from transport_reports import AreaMatcher, blank, hkey, num, zone_code

EXPORT_KEYS = {"type", "areacode", "areaname", "zonecode", "zonename", "zoneswithreport", "cellswithreport"}
COUNTS = ["male", "female", "children", "first_timers", "new_converts", "testimonies"]
EXPORT_COLS = {"male": "totalmale", "female": "totalfemale", "children": "totalchildren", "total": "total",
               "first_timers": "totalfirsttimer", "new_converts": "totalnewconvert", "testimonies": "totaltestimony",
               "cells_total": "totalcells", "cells_reported": "cellswithreport"}


def export_stream(sheet: str, file_name: str) -> str | None:
    t = f"{sheet} {file_name}".lower()
    if "chop" in t:
        return "CHOP"
    if "mid-week" in t or "midweek" in t or "mid week" in t:
        return "MIDWEEK"
    if "cell service" in t or "cell_service" in t or "wsf" in t:
        return "WSF"
    return None


def is_export(raw: pd.DataFrame) -> bool:
    return len(raw) > 0 and EXPORT_KEYS <= {hkey(v) for v in raw.iloc[0].tolist()}


def export_dates(file_name: str) -> tuple[dt.date, dt.date] | None:
    m = re.findall(r"(\d{4})-(\d{2})-(\d{2})", file_name)
    if not m:
        return None
    dates = [dt.date(int(y), int(mo), int(d)) for y, mo, d in m]
    return min(dates), max(dates)


@dataclass
class ExportResult:
    stream: str
    report_date: dt.date | None
    zones: pd.DataFrame = None            # zonal rows (reported only for Midweek/CHOP; all zones for WSF)
    community: pd.DataFrame = None        # community churches' rows for this service
    areas: pd.DataFrame = None            # Area summaries (compliance figures from the system)
    flags: pd.DataFrame = None            # zone code, area, name, forChop, forCommunity (to refresh the masters)
    issues: list = field(default_factory=list)

    def issue(self, severity, issue, detail, row=None):
        self.issues.append({"severity": severity, "issue": issue, "detail": str(detail)[:250], "source_row": row})


def read_export(raw: pd.DataFrame, sheet: str, file_name: str, masters: Masters) -> ExportResult | None:
    stream = export_stream(sheet, file_name)
    if stream is None or not is_export(raw):
        return None
    dates = export_dates(file_name)
    res = ExportResult(stream, dates[1] if dates else None)
    if dates is None:
        res.issue("Rejected", "Report date not found", "Expected dates like '-2026-09-19-2026-09-19' in the file name")
        return res
    if dates[0] != dates[1]:
        res.issue("Info", "Report covers several days", f"{dates[0]} to {dates[1]}; filed under the week of {dates[1]}")
    head = [hkey(v) for v in raw.iloc[0].tolist()]
    col = {k: head.index(v) for k, v in EXPORT_COLS.items()}
    col.update({k: head.index(k) for k in ("type", "areacode", "areaname", "zonecode", "zonename", "forchop",
                                           "forcommunity", "totalzones", "zoneswithreport", "chopcount",
                                           "communitycount")})
    match = AreaMatcher(masters)
    zone_rows, area_rows = [], []
    for i in range(1, len(raw)):
        r = raw.iloc[i].tolist()
        kind = str(r[col["type"]] or "")
        if kind not in ("ZoneSummary", "AreaSummary"):
            continue
        area_no = int(num(r[col["areacode"]]) or 0)
        area = match(r[col["areaname"]], area_no)
        if area is None:
            res.issue("Rejected", "Unknown Area", f"'{r[col['areaname']]}' (code {r[col['areacode']]})", i + 1)
            continue
        vals = {k: int(num(r[col[k]]) or 0) for k in EXPORT_COLS}
        if kind == "AreaSummary":
            area_rows.append(dict(area=area, area_no=area_no, total_zones=int(num(r[col["totalzones"]]) or 0),
                                  zones_with_report=int(num(r[col["zoneswithreport"]]) or 0),
                                  chop_zones=int(num(r[col["chopcount"]]) or 0),
                                  community_zones=int(num(r[col["communitycount"]]) or 0), **vals))
            continue
        zc = r[col["zonecode"]]
        code = f"LFC{area_no:02d}{int(num(zc) or 0):02d}" if not blank(zc) else None
        if code is None:
            res.issue("Rejected", "Zone without a code", f"{area}: '{r[col['zonename']]}'", i + 1)
            continue
        zone_rows.append(dict(zone=code, area=area, zone_name=str(r[col["zonename"]] or "").strip(),
                              for_chop=r[col["forchop"]] is True, for_community=r[col["forcommunity"]] is True,
                              row=i + 1, **vals))
    z = pd.DataFrame(zone_rows)
    a = pd.DataFrame(area_rows)
    if z.empty:
        res.issue("Rejected", "No zone rows", "The export has no ZoneSummary rows")
        return res
    # a zone registered twice in the system: add attendance once, count its cells once
    dup = z[z.zone.duplicated(keep=False)]
    for code, g in dup.groupby("zone"):
        res.issue("Info", "Zone registered twice in the reporting system",
                  f"{code} ({g.area.iloc[0]}): " + " / ".join(f"'{n}'" for n in g.zone_name) + "; added up once",
                  int(g.row.iloc[0]))
    agg = {k: "sum" for k in COUNTS + ["total"]}
    agg.update(cells_total="max", cells_reported="sum", area="first", zone_name="first", for_chop="max",
               for_community="max", row="first")
    z = z.groupby("zone", as_index=False).agg(agg)
    z["cells_reported"] = z[["cells_reported", "cells_total"]].min(axis=1)
    community_codes = set(masters.community["Church_ID"]) if len(masters.community) else set()
    z["for_community"] = z.for_community | z.zone.isin(community_codes)
    res.flags = z[["zone", "area", "zone_name", "for_chop", "for_community"]].copy()
    for c in ("male", "female", "children"):
        if (z[c] < 0).any():
            res.issue("Warning", f"{c} is negative", ", ".join(z.loc[z[c] < 0, "zone"]))
    given = z.male + z.female + z.children
    for _, row in z[given != z.total].iterrows():
        res.issue("Warning", "Grand total mismatch",
                  f"{row.zone}: total {row.total}, Male+Female+Children = {row.male + row.female + row.children}; "
                  "computed value used", int(row.row))
    comm = z[z.for_community]
    zonal = z[~z.for_community].copy()
    # attendance an Area reports above its zones (CHOP at the Area facility, corrections...)
    extra = []
    for _, ar in a.iterrows():
        in_zones = z[z.area == ar.area]
        diff = {k: ar[k] - int(in_zones[k].sum()) for k in COUNTS}
        people = [diff[k] for k in ("male", "female", "children")]
        if any(people):
            if all(v >= 0 for v in people):
                diff = {k: max(0, v) for k, v in diff.items()}
                extra.append(dict(zone=f"AREA LEVEL {int(ar.area_no):02d}", area=ar.area, zone_name="Area level",
                                  cells_total=0, cells_reported=0, **diff))
                if stream != "CHOP":
                    res.issue("Info", "Area total above its zones",
                              f"{ar.area}: {sum(diff[k] for k in ('male', 'female', 'children')):,} attendance "
                              "reported at Area level (kept as an AREA LEVEL row)")
            else:
                res.issue("Warning", "Area total differs from its zones",
                          f"{ar.area}: Area row {int(ar.total):,}, zones add up to {int(in_zones.total.sum()):,}; "
                          "zone figures used")
    if stream == "CHOP" and extra:
        res.issue("Info", "CHOP reported at Area level",
                  f"{sum(e['male'] + e['female'] + e['children'] for e in extra):,} CHOP attendance across "
                  f"{len(extra)} Areas is reported for the Area, not a zone (kept as AREA LEVEL rows)")
    zonal = pd.concat([zonal, pd.DataFrame(extra)], ignore_index=True) if extra else zonal
    reported = (zonal.male + zonal.female + zonal.children) > 0
    res.zones = zonal if stream == "WSF" else zonal[reported]
    res.community = comm[(comm.male + comm.female + comm.children) > 0]
    res.areas = a
    return res


# --------------------------------------------------------------------------
# Community Church Sunday workbook
# --------------------------------------------------------------------------

SUNDAY_SIGNATURE = ("communitychurchlocation", "zincode", "waspastor1present")
SUNDAY_FIELDS = {"male": "male", "female": "female", "children": "children", "total": "total",
                 "first_timers": "firsttimers", "new_converts": "newconvert", "testimonies": "testimonies",
                 "pastor1": "waspastor1present", "pastor2": "waspastor2present",
                 "cells": "currenttotalnumberofcellsineachlocation"}


def is_sunday_sheet(raw: pd.DataFrame) -> int | None:
    for r in range(min(6, len(raw))):
        keys = {hkey(v) for v in raw.iloc[r].tolist()}
        if all(any(k.startswith(s) for k in keys) for s in SUNDAY_SIGNATURE):
            return r
    return None


def _sheet_month(sheet: str):
    m = re.match(r"(?i)\s*([a-z]{3,9})\s*(\d{4})", sheet)
    months = ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"]
    if m and m.group(1)[:3].lower() in months:
        return int(m.group(2)), months.index(m.group(1)[:3].lower()) + 1
    return None


def _present(v) -> int:
    return 1 if str(v or "").strip().lower() in ("yes", "y", "true", "1") else 0


def read_sunday(raw: pd.DataFrame, sheet: str, masters: Masters):
    """Return (rows dataframe, issues).  One row per church per Sunday that has figures."""
    hdr = is_sunday_sheet(raw)
    if hdr is None:
        return None
    issues = []
    header = [hkey(v) for v in raw.iloc[hdr].tolist()]
    date_row = raw.iloc[hdr - 1].tolist() if hdr > 0 else [None] * len(header)
    dated = [(j, pd.Timestamp(v).date()) for j, v in enumerate(date_row)
             if isinstance(v, (dt.date, dt.datetime, pd.Timestamp))]
    # blocks end where the next block (dated or not, e.g. "SUMMARY") begins
    edges = sorted(j for j, v in enumerate(date_row) if not blank(v) and j > 0) + [len(header)]
    sheet_month = _sheet_month(sheet)
    starts = []
    for j, day in dated:
        if sheet_month and (day.year, day.month) != sheet_month:
            issues.append({"severity": "Info", "issue": "Block outside the sheet's month skipped",
                           "detail": f"Sheet '{sheet}' has a block dated {day} (left over from an old template)",
                           "source_row": hdr})
            continue
        starts.append((j, day))
    zc = header.index("zincode")
    area_c = next((j for j, k in enumerate(header) if k.startswith("hostarea")), None)
    loc_c = next((j for j, k in enumerate(header) if k.startswith("communitychurchlocation")), None)
    no_c = next((j for j, k in enumerate(header) if k.startswith("areano")), None)
    fixed_codes = set()
    match = AreaMatcher(masters)
    rows = []
    for start, day in starts:
        end = next(e for e in edges if e > start)
        block = {k: next((j for j in range(start, end) if header[j].startswith(v)), None)
                 for k, v in SUNDAY_FIELDS.items()}
        for i in range(hdr + 1, len(raw)):
            r = raw.iloc[i].tolist()
            code = zone_code(r[zc])
            if code is None:
                continue
            area_no = num(r[no_c]) if no_c is not None else None
            if area_no is not None and int(code[3:5]) != int(area_no):
                # e.g. 'LFC320' typed under AREA NO 35: the Area number decides -> LFC3520
                fixed = f"LFC{int(area_no):02d}{code[-2:]}"
                if (code, fixed) not in fixed_codes:
                    fixed_codes.add((code, fixed))
                    issues.append({"severity": "Warning", "issue": "Church code corrected to its Area",
                                   "detail": f"'{r[zc]}' under Area {int(area_no)} read as {fixed}",
                                   "source_row": i + 1})
                code = fixed
            counts = {k: num(r[block[k]]) if block[k] is not None else None for k in
                      ("male", "female", "children", "first_timers", "new_converts", "testimonies", "total")}
            if all(counts[k] is None for k in ("male", "female", "children")):
                continue                                              # no report that Sunday
            m, f, c = (int(counts[k] or 0) for k in ("male", "female", "children"))
            if counts["total"] is not None and int(counts["total"]) != m + f + c:
                issues.append({"severity": "Warning", "issue": "Grand total mismatch", "source_row": i + 1,
                               "detail": f"{code} {day}: total {int(counts['total'])}, M+F+C = {m + f + c}"})
            rows.append(dict(church=code, area=match(None, int(code[3:5])) or (match(r[area_c]) if area_c is not None
                                                                               else None),
                             location=str(r[loc_c] or "").strip() if loc_c is not None else "",
                             report_date=day, male=m, female=f, children=c,
                             first_timers=int(counts["first_timers"] or 0),
                             new_converts=int(counts["new_converts"] or 0),
                             testimonies=int(counts["testimonies"] or 0),
                             pastors_present=_present(r[block["pastor1"]] if block["pastor1"] is not None else None)
                             + _present(r[block["pastor2"]] if block["pastor2"] is not None else None),
                             source_row=i + 1))
    return pd.DataFrame(rows), issues
