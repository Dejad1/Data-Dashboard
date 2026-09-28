"""Build Operations_Dashboard.xlsx from the databank and the master lists.

    python build_workbook.py                  # data root = this folder
    python build_workbook.py --root demo
    python build_workbook.py --empty out.xlsx # an empty-data workbook (masters only)

ingest.py calls this automatically after every load, so the workbook is always
a fresh, consistent picture of the databank.  SETTINGS values that staff
changed in the previous workbook are carried over.
"""
from __future__ import annotations

import argparse
import datetime as dt
from pathlib import Path

import pandas as pd
from openpyxl import Workbook

from builder import data as D
from builder.home import build_home, build_placeholder
from builder.wsf import build_calc as build_wsf_calc, build_dashboard as build_wsf_dashboard
from config import COLOURS, HERE, Paths, load_settings
from databank import Databank, Masters

SHEET_ORDER = [
    "HOME", "MASTER DASHBOARD", "WSF DASHBOARD", "MIDWEEK DASHBOARD", "CHOP DASHBOARD",
    "COMMUNITY CHURCH DASHBOARD", "TRANSPORT DASHBOARD", "AREA SCORECARD", "ZONE SCORECARD", "DATA QUALITY",
    "DATA_WSF", "WSF_CELLS_CURRENT", "DATA_MIDWEEK", "DATA_CHOP", "DATA_COMMUNITY", "DATA_TRANSPORT_OPS",
    "DATA_TRANSPORT_FINANCE",
    "MASTER_AREAS", "MASTER_ZONES", "MASTER_CELLS", "MASTER_COMMUNITY", "MASTER_FLEET",
    "CALENDAR", "SETTINGS", "CALC_LISTS", "CALC_SEARCH", "CALC_WSF",
]

TABLE_SOURCES = {
    "WSF": ("wsf_zone", ["week_ending", "area", "zone"]),
    "MIDWEEK": ("midweek", ["week_ending", "area", "zone"]),
    "CHOP": ("chop", ["week_ending", "area", "zone"]),
    "COMMUNITY": ("community", ["week_ending", "church", "service_type"]),
    "TRANSPORT_OPS": ("transport_ops", ["week_ending", "category", "vehicle_id"]),
    "TRANSPORT_FINANCE": ("transport_fin", ["week_ending", "category", "area", "cost_type"]),
}

PLACEHOLDER_NOTE = ("Stage 2: this dashboard is built after the WSF DASHBOARD is reviewed, using the same control "
                    "strip, KPI cards, time analysis, rankings and non-reporting layout. Its data is already loaded "
                    "in the {data} sheet.")


def cells_current(db: Databank, masters: Masters, weeks_kept: int) -> pd.DataFrame:
    """All operational cells x the latest N weeks, with Reported = Y/N."""
    weeks = [w for (w,) in db.execute(
        "SELECT DISTINCT week_ending FROM wsf_cells ORDER BY week_ending DESC LIMIT ?", (weeks_kept,))]
    if not weeks:
        return pd.DataFrame(columns=["week_ending", "report_date", "area", "zone", "cell", "reported", "male",
                                     "female", "children", "source_file", "loaded_on"])
    placeholders = ",".join("?" * len(weeks))
    rep = db.read("wsf_cells", f"week_ending IN ({placeholders})", tuple(weeks))
    cells = masters.operational_cells[["Cell", "Zone_Name"]].rename(columns={"Cell": "cell", "Zone_Name": "zone"})
    cells["area"] = cells["zone"].map(masters.zone_area)
    frames = []
    for w in weeks:
        r = rep[rep["week_ending"] == w]
        grid = cells.merge(r.drop(columns=["area", "zone"]), on="cell", how="outer")
        grid["zone"] = grid["zone"].fillna(grid["cell"].map(masters.cell_zone))
        grid["area"] = grid["area"].fillna(grid["zone"].map(masters.zone_area))
        grid["reported"] = grid["week_ending"].notna().map({True: "Y", False: "N"})
        grid["week_ending"] = w
        grid["report_date"] = grid["report_date"].fillna(w)
        grid["source_file"] = grid["source_file"].fillna("")
        frames.append(grid)
    out = pd.concat(frames, ignore_index=True)
    return out.sort_values(["week_ending", "area", "zone", "cell"], ascending=[False, True, True, True])


def coverage_and_gaps(tables: dict[str, pd.DataFrame]):
    cov, gaps = [], []
    for key, df in tables.items():
        if df.empty:
            cov.append({"stream": key, "first": None, "latest": None, "weeks": 0, "rows": 0})
            continue
        weeks = sorted(set(df["week_ending"]))
        cov.append({"stream": key, "first": weeks[0], "latest": weeks[-1], "weeks": len(weeks), "rows": len(df)})
        d = dt.date.fromisoformat(weeks[0])
        last = dt.date.fromisoformat(weeks[-1])
        have = set(weeks)
        while d <= last:
            if d.isoformat() not in have:
                gaps.append({"stream": key, "week_ending": d.isoformat()})
            d += dt.timedelta(days=7)
    return pd.DataFrame(cov), pd.DataFrame(gaps, columns=["stream", "week_ending"])


def build(root: Path, out: Path | None = None, empty: bool = False) -> Path:
    paths = Paths(Path(root))
    settings = load_settings(paths)
    masters = Masters.load(paths.masters)
    out = Path(out) if out else paths.workbook
    existing = D.read_existing_settings(out) if out.exists() else {}

    db = Databank(paths.databank)
    tables = {}
    for key, (table, order) in TABLE_SOURCES.items():
        df = db.read(table) if not empty else db.read(table, "0")
        tables[key] = df.sort_values(order).reset_index(drop=True)
    dq = db.read("dq_log") if not empty else db.read("dq_log", "0")
    dq = dq.iloc[::-1].head(20000)
    cur = cells_current(db, masters, 0 if empty else settings["wsf_cells_weeks_in_workbook"])
    db.close()

    wb = Workbook()
    wb.remove(wb.active)
    colour_of = {"WSF": "WSF", "MIDWEEK": "MIDWEEK", "CHOP": "CHOP", "COMMUNITY": "COMMUNITY",
                 "TRANSPORT_OPS": "TRANSPORT", "TRANSPORT_FINANCE": "TRANSPORT"}

    D.write_settings(wb, settings, existing)
    D.write_calendar(wb, settings)
    for key, spec in D.SPECS.items():
        D.write_table(wb, spec, tables[key], COLOURS[colour_of[key]])
    D.write_table(wb, D.CELLS_CURRENT, cur, COLOURS["WSF"])
    sizes = D.write_masters(wb, masters)

    build_wsf_calc(wb, sizes["MASTER_AREAS"])
    build_wsf_dashboard(wb, sizes["MASTER_ZONES"])

    placeholders = [
        ("MASTER DASHBOARD", "MASTER", "Global view of every stream side by side", "all DATA_*"),
        ("MIDWEEK DASHBOARD", "MIDWEEK", "Zonal midweek services", "DATA_MIDWEEK"),
        ("CHOP DASHBOARD", "CHOP", "Covenant Hour of Prayer (CHOP zones only)", "DATA_CHOP"),
        ("COMMUNITY CHURCH DASHBOARD", "COMMUNITY", "35 Community Churches · Sunday service first",
         "DATA_COMMUNITY"),
        ("TRANSPORT DASHBOARD", "TRANSPORT", "Fleet operations, finance and optimisation",
         "DATA_TRANSPORT_OPS / DATA_TRANSPORT_FINANCE"),
        ("AREA SCORECARD", "MASTER", "One Area across every stream, with rank among all Areas", "DATA_*"),
        ("ZONE SCORECARD", "MASTER", "One Zone across every stream", "DATA_*"),
    ]
    for sheet, key, sub, data_sheet in placeholders:
        build_placeholder(wb, sheet, key, sub, PLACEHOLDER_NOTE.format(data=data_sheet))
    cov, gaps = coverage_and_gaps(tables)
    D.write_data_quality(wb, dq, cov, gaps)
    build_home(wb, stage_note="Review build (stage 1): the WSF DASHBOARD, the data tables, masters, Calendar, "
                              "SETTINGS and DATA QUALITY are complete. The other dashboards follow once the WSF "
                              "layout is approved.")

    wb._sheets = [wb[name] for name in SHEET_ORDER] + [s for s in wb._sheets if s.title not in SHEET_ORDER]
    wb.active = 0
    for ws in wb.worksheets:
        ws.sheet_view.tabSelected = ws.title == "HOME"
    wb.calculation.fullCalcOnLoad = True
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default=str(HERE))
    ap.add_argument("--out", default=None, help="output path (default: <root>/Operations_Dashboard.xlsx)")
    ap.add_argument("--empty", action="store_true", help="build without any fact rows (masters only)")
    args = ap.parse_args(argv)
    print(build(Path(args.root).resolve(), Path(args.out) if args.out else None, empty=args.empty))


if __name__ == "__main__":
    main()
