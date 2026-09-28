"""HOME navigation page and the placeholder pages for dashboards still to come."""
from __future__ import annotations

from config import COLOURS
from .common import (BOX, CENTER, FMT_DATE, FMT_DATETIME, FMT_INT0, LEFT, LEFT_WRAP, MUTED, WHITE,
                     dashboard_canvas, fill, font, put, q, section_title, tint)
from .data import SPECS

NAV = [
    ("MASTER DASHBOARD", "MASTER", "All streams side by side"),
    ("WSF DASHBOARD", "WSF", "Saturday cell services"),
    ("MIDWEEK DASHBOARD", "MIDWEEK", "Zonal midweek services"),
    ("CHOP DASHBOARD", "CHOP", "Covenant Hour of Prayer"),
    ("COMMUNITY CHURCH DASHBOARD", "COMMUNITY", "Sunday service first"),
    ("TRANSPORT DASHBOARD", "TRANSPORT", "Operations, finance, optimisation"),
    ("AREA SCORECARD", "MASTER", "One Area across every stream"),
    ("ZONE SCORECARD", "MASTER", "One Zone across every stream"),
    ("DATA QUALITY", "QUALITY", "Rejected rows and missing weeks"),
]

HOW_TO = [
    "Choose a dashboard above. Every dashboard has the same control strip at the top.",
    "SCOPE: Global for everything, Area for one Area, Zone for one Zone.",
    "SEARCH: type part of a name (e.g. 'ikeja'), then open SELECT: it lists only the matching Areas or Zones.",
    "WEEK ENDING: 'Latest' follows new data automatically; pick an older week to look back.",
    "COMPARE (WEEKS): the averaging window, 1 to 52 weeks. '–' means there is no data for that figure.",
    "New data never goes into these sheets by hand: drop the files in the inbox folder and run ingest.py (see README).",
]

STREAM_STATUS = [("WSF", "WSF"), ("Midweek", "MIDWEEK"), ("CHOP", "CHOP"), ("Community Church", "COMMUNITY"),
                 ("Transport operations", "TRANSPORT_OPS"), ("Transport finance", "TRANSPORT_FINANCE")]


def build_home(wb, stage_note: str | None = None):
    ws = wb.create_sheet("HOME", 0)
    ws.sheet_properties.tabColor = COLOURS["MASTER"]
    dashboard_canvas(ws, "MASTER", "OPERATIONS REPORTING & ANALYSIS",
                     "Central · 92 Areas · Zones · Cells · Community Churches · Transport", n_cols=12,
                     widths={c: 13 for c in "BCDEFGHIJKLM"})
    ws["M1"].value = None
    section_title(ws, 4, "Dashboards", COLOURS["MASTER"], last="M")
    for i, (sheet, key, sub) in enumerate(NAV):
        r = 6 + (i // 3) * 4
        c1 = "BFJ"[i % 3]
        c2 = chr(ord(c1) + 3)
        colour = COLOURS[key]
        put(ws, f"{c1}{r}", sheet, f=font(11, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}{r}")
        put(ws, f"{c1}{r + 1}", sub, f=font(8, False, tint(colour, 0.8)), bg=colour, align=CENTER,
            merge_to=f"{c2}{r + 1}")
        ws[f"{c1}{r}"].hyperlink = f"#{q(sheet)}!A1"
        ws[f"{c1}{r + 1}"].hyperlink = f"#{q(sheet)}!A1"
        ws.row_dimensions[r].height = 26
        ws.row_dimensions[r + 1].height = 16

    section_title(ws, 19, "Data status", COLOURS["MASTER"], last="M")
    loaded = ",".join(SPECS[k].ref("Loaded_On") for k in SPECS)
    put(ws, "B20", "Last data loaded", f=font(10, True), merge_to="D20")
    put(ws, "E20", f'=IF(MAX({loaded})=0,"No data yet",MAX({loaded}))', f=font(10, True, COLOURS["MASTER"]),
        fmt=FMT_DATETIME, align=LEFT, merge_to="G20")
    for j, h in enumerate(["Stream", "", "", "Latest week", "", "Rows in workbook"]):
        if h:
            put(ws, f"{'BCDEFG'[j]}22", h, f=font(9, True, WHITE), bg=COLOURS["MASTER"])
    for c in "BCDEFGH":
        ws[f"{c}22"].fill = fill(COLOURS["MASTER"])
    for i, (label, key) in enumerate(STREAM_STATUS):
        r = 23 + i
        spec = SPECS[key]
        put(ws, f"B{r}", label, f=font(10), merge_to=f"D{r}")
        put(ws, f"E{r}", f'=IF(MAX({spec.ref("Week_Ending")})=0,"–",MAX({spec.ref("Week_Ending")}))', f=font(10, True),
            fmt=FMT_DATE, align=LEFT, merge_to=f"F{r}")
        put(ws, f"G{r}", f"=COUNT({spec.ref('Week_Ending')})", f=font(10), fmt=FMT_INT0, merge_to=f"H{r}")

    put(ws, "J20", "How to use", f=font(10, True), merge_to="M20")
    for i, text in enumerate(HOW_TO):
        put(ws, f"J{21 + i}", f"{i + 1}.  {text}", f=font(9), align=LEFT_WRAP, merge_to=f"M{21 + i}")
        ws.row_dimensions[21 + i].height = 36
    if stage_note:
        put(ws, "B30", stage_note, f=font(9, True, "8A5A00"), bg="FFF1CC", align=LEFT_WRAP, merge_to="M31")
    ws.freeze_panes = None
    return ws


def build_placeholder(wb, sheet: str, key: str, subtitle: str, note: str):
    ws = wb.create_sheet(sheet)
    ws.sheet_properties.tabColor = COLOURS[key]
    dashboard_canvas(ws, key, sheet, subtitle, n_cols=14)
    put(ws, "B4", note, f=font(11, False, MUTED, italic=True), align=LEFT_WRAP, merge_to="O6")
    return ws
