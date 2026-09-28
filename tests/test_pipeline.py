"""End-to-end tests: generate data -> ingest -> build -> LibreOffice recalc -> check numbers.

Run with:  python -m pytest tests -q        (needs LibreOffice for the workbook tests)

The expected values are recomputed independently with pandas from the
databank and compared with what the workbook formulas produce.
"""
from __future__ import annotations

import datetime as dt
import shutil
import sqlite3
import sys
from pathlib import Path

import pandas as pd
import pytest
from openpyxl import load_workbook

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent))
sys.path.insert(0, str(HERE))

import build_workbook  # noqa: E402
import generate_sample_data  # noqa: E402
import ingest  # noqa: E402
from lo_tools import recalc, scan  # noqa: E402

TODAY = dt.date(2026, 9, 28)
N_WEEKS = 16
HAS_LO = shutil.which("soffice") is not None


@pytest.fixture(scope="module")
def root(tmp_path_factory):
    r = tmp_path_factory.mktemp("ops")
    generate_sample_data.generate(r, n_weeks=N_WEEKS, seed=11, today=TODAY, scale=0.1)
    ingest.run(r, build=False, quiet=True)
    build_workbook.build(r)
    return r


def db(root) -> sqlite3.Connection:
    return sqlite3.connect(root / "databank.sqlite")


# --------------------------------------------------------------------------
# Ingestion
# --------------------------------------------------------------------------


def test_all_streams_loaded(root):
    con = db(root)
    for table in ("wsf_cells", "wsf_zone", "midweek", "chop", "community", "transport_ops", "transport_fin"):
        weeks = con.execute(f"SELECT COUNT(DISTINCT week_ending) FROM {table}").fetchone()[0]
        assert weeks == N_WEEKS, table


def test_injected_errors_are_logged(root):
    issues = pd.read_sql("SELECT stream, severity, issue FROM dq_log", db(root))
    got = set(zip(issues.stream, issues.issue))
    for expected in [("WSF_CELLS", "Grand total mismatch"), ("WSF_CELLS", "male is negative"),
                     ("WSF_CELLS", "Unknown Cell (not in MASTER_CELLS)"), ("WSF_CELLS", "Duplicate row in file"),
                     ("WSF_CELLS", "Date missing or not recognised"),
                     ("MIDWEEK", "Community Church row in a zonal file (kept separate by design)"),
                     ("MIDWEEK", "Unknown Zone (not in MASTER_ZONES)"),
                     ("CHOP", "CHOP report from a zone without CHOP (Has_CHOP = N)"),
                     ("CHOP", "Adult total mismatch"),
                     ("TRANSPORT_OPS", "Unknown vehicle (not in MASTER_FLEET)")]:
        assert expected in got, expected


def test_chop_only_from_chop_zones(root):
    con = db(root)
    zones = pd.read_csv(root / "masters" / "MASTER_ZONES.csv", dtype=str)
    chop_zones = set(zones.loc[zones.Has_CHOP == "Y", "Zone_Name"])
    loaded = {z for (z,) in con.execute("SELECT DISTINCT zone FROM chop")}
    assert loaded <= chop_zones


def test_community_never_in_zonal_tables(root):
    con = db(root)
    churches = set(pd.read_csv(root / "masters" / "MASTER_COMMUNITY.csv", dtype=str)["Church_Name"])
    for table in ("wsf_zone", "midweek", "chop"):
        zones = {z for (z,) in con.execute(f"SELECT DISTINCT zone FROM {table}")}
        assert not zones & churches, table


def test_wsf_zone_rows_cover_every_active_zone(root):
    con = db(root)
    n_active = len(pd.read_csv(root / "masters" / "MASTER_ZONES.csv"))
    per_week = con.execute("SELECT COUNT(*) FROM wsf_zone GROUP BY week_ending").fetchall()
    assert {n for (n,) in per_week} == {n_active}
    agg = con.execute("SELECT SUM(male), SUM(female), SUM(children) FROM wsf_zone").fetchone()
    raw = con.execute("SELECT SUM(male), SUM(female), SUM(children) FROM wsf_cells").fetchone()
    assert agg == raw


def test_reingest_same_file_does_not_double_count(root, tmp_path):
    con = db(root)
    before = {t: con.execute(f"SELECT COUNT(*), SUM(male) FROM {t}").fetchone()
              for t in ("wsf_cells", "wsf_zone", "midweek", "chop")}
    fin_before = con.execute("SELECT COUNT(*), SUM(amount) FROM transport_fin").fetchone()
    con.close()
    processed = root / "inbox" / "processed"
    resend = [p for p in processed.iterdir() if "Midweek" in p.name or "WSF cell" in p.name
              or "Covenant" in p.name or "finance" in p.name][:8]
    for p in resend:
        shutil.copy(p, root / "inbox" / p.name.split("_", 1)[1])
    summary = ingest.run(root, build=False, quiet=True)
    con = db(root)
    after = {t: con.execute(f"SELECT COUNT(*), SUM(male) FROM {t}").fetchone() for t in before}
    assert after == before
    assert con.execute("SELECT COUNT(*), SUM(amount) FROM transport_fin").fetchone() == fin_before
    # everything re-sent was counted as a replacement, not a new row
    for stream, (loaded, replaced, _rej) in summary.items():
        assert loaded == replaced, stream


def test_header_detection_and_aliases(tmp_path):
    aliases = ingest.load_aliases()
    raw = pd.DataFrame([["MIDWEEK RETURNS", None, None, None, None], [None] * 5,
                        ["No. of Females", "zone name", "KIDS", "Meeting-Date", "MEN"],
                        [10, "X Zone 01", 3, "02/09/2026", 7]])
    hdr, mapping = ingest.find_header(raw, aliases)
    assert hdr == 2
    assert sorted(mapping.values()) == ["children", "date", "female", "male", "zone"]
    assert ingest.detect_stream(set(mapping.values()), "Zonal Midweek wk36.xlsx", "Sheet1") == "MIDWEEK"
    assert ingest.parse_date("02/09/2026") == dt.date(2026, 9, 2)   # day first (Nigeria)


# --------------------------------------------------------------------------
# Workbook scenarios (need LibreOffice)
# --------------------------------------------------------------------------

def expected_wsf(root, scope: str, sel: str, week: dt.date | None, n: int) -> dict:
    z = pd.read_sql("SELECT * FROM wsf_zone", db(root))
    z["week"] = pd.to_datetime(z.week_ending).dt.date
    if scope == "Area":
        z = z[z.area == sel]
    elif scope == "Zone":
        z = z[z.zone == sel]
    z["grand"] = z.male + z.female + z.children
    wk = z.groupby("week").agg(grand=("grand", "sum"), cells=("cells_total", "sum"), rep=("cells_reported", "sum"))
    latest = max(pd.read_sql("SELECT week_ending FROM wsf_zone", db(root)).week_ending.map(dt.date.fromisoformat))
    sel_w = week or latest
    prev = sel_w - dt.timedelta(days=7)
    window = [sel_w - dt.timedelta(days=7 * k) for k in range(1, n + 1)]
    win = wk[wk.index.isin(window)]
    mtd = wk[[(w.month == sel_w.month and w.year == sel_w.year and w <= sel_w) for w in wk.index]]
    ytd = wk[[(w.year == sel_w.year and w <= sel_w) for w in wk.index]]
    return {"grand": wk.grand.get(sel_w), "prev": wk.grand.get(prev),
            "navg": win.grand.mean() if len(win) else None, "mtd": mtd.grand.sum(), "ytd": ytd.grand.sum(),
            "pct": wk.rep.get(sel_w) / wk.cells.get(sel_w),
            "pct_navg": win.rep.sum() / win.cells.sum() if len(win) else None,
            "week": sel_w}


def expected_nonreporting(root, scope, sel, week) -> int:
    con = db(root)
    cells = pd.read_csv(root / "masters" / "MASTER_CELLS.csv", dtype=str)
    zones = pd.read_csv(root / "masters" / "MASTER_ZONES.csv", dtype=str)
    cells = cells[cells.Operational == "Y"].merge(zones[["Zone_Name", "Area_Name"]], on="Zone_Name")
    rep = {c for (c,) in con.execute("SELECT cell FROM wsf_cells WHERE week_ending = ?", (week.isoformat(),))}
    if scope == "Area":
        cells = cells[cells.Area_Name == sel]
    elif scope == "Zone":
        cells = cells[cells.Zone_Name == sel]
    return int((~cells.Cell.isin(rep)).sum())


def run_scenario(root, tmp_path, name, scope="Global", select=None, week=None, window=4):
    wb = load_workbook(root / "Operations_Dashboard.xlsx")
    ws = wb["WSF DASHBOARD"]
    ws["B4"] = scope
    ws["E4"] = select
    ws["I4"] = week if week else "Latest"
    ws["K4"] = window
    out = tmp_path / f"{name}.xlsx"
    wb.save(out)
    recalc(out)
    return out


def calc_values(path):
    wb = load_workbook(path, data_only=True)
    c, d = wb["CALC_WSF"], wb["WSF DASHBOARD"]
    vals = {"grand": c["B90"].value, "prev": c["C90"].value, "navg": c["F90"].value, "mtd": c["H90"].value,
            "ytd": c["J90"].value, "pct": c["B98"].value, "pct_navg": c["F98"].value, "week": c["B10"].value,
            "nonrep": c["AX1"].value, "rank_head": d["B46"].value, "card": d["B8"].value, "status": d["B5"].value,
            "rank_box": d["O46"].value, "top_area": c["AJ3"].value}
    wb.close()
    return vals


def close(a, b, tol=1e-6):
    if b is None or (isinstance(b, float) and pd.isna(b)):
        return a in ("", None)
    return a not in ("", None) and abs(float(a) - float(b)) <= tol * max(1, abs(float(b)))


SCENARIOS = [
    ("global_w4", "Global", None, None, 4),
    ("global_w1", "Global", None, None, 1),
    ("global_w13", "Global", None, None, 13),
    ("area_w4", "Area", "AREA", None, 4),
    ("zone_chop_w13", "Zone", "ZONE_CHOP", None, 13),
    ("zone_nochop_w4", "Zone", "ZONE_NOCHOP", None, 4),
    ("older_week_w4", "Global", None, "PREV3", 4),
]


@pytest.mark.skipif(not HAS_LO, reason="LibreOffice not installed")
@pytest.mark.parametrize("name,scope,select,week,window", SCENARIOS)
def test_wsf_dashboard_scenarios(root, tmp_path, name, scope, select, week, window):
    zones = pd.read_csv(root / "masters" / "MASTER_ZONES.csv", dtype=str)
    pick = {"AREA": zones.Area_Name.iloc[len(zones) // 2],
            "ZONE_CHOP": zones[zones.Has_CHOP == "Y"].Zone_Name.iloc[3],
            "ZONE_NOCHOP": zones[zones.Has_CHOP == "N"].Zone_Name.iloc[3]}
    sel = pick.get(select, select)
    latest = max(pd.read_sql("SELECT week_ending FROM wsf_zone", db(root)).week_ending.map(dt.date.fromisoformat))
    wk = latest - dt.timedelta(weeks=3) if week == "PREV3" else None
    out = run_scenario(root, tmp_path, name, scope, sel, wk, window)
    got = calc_values(out)
    exp = expected_wsf(root, scope, sel, wk, window)
    for k in ("grand", "prev", "navg", "mtd", "ytd", "pct", "pct_navg"):
        assert close(got[k], exp[k]), (k, got[k], exp[k])
    assert got["week"].date() == exp["week"]
    if (latest - exp["week"]).days < 7 * 8:
        assert got["nonrep"] == expected_nonreporting(root, scope, sel, exp["week"])
    if scope == "Zone":
        assert got["rank_head"] == "n/a at Zone scope"
    else:
        assert got["rank_head"].startswith("1. ")
    if scope == "Area":
        assert got["rank_box"] != "n/a"
    assert scan(out)["total_errors"] == 0


@pytest.mark.skipif(not HAS_LO, reason="LibreOffice not installed")
def test_empty_workbook_has_no_errors(root, tmp_path):
    out = tmp_path / "empty.xlsx"
    build_workbook.build(root, out, empty=True)
    recalc(out)
    assert scan(out)["total_errors"] == 0
    got = calc_values(out)
    assert got["card"] == "–"
    assert got["status"].startswith("No WSF data loaded yet")
