"""CALC_WSF (hidden weekly aggregates) and the WSF DASHBOARD.

The dashboard never scans raw rows itself.  CALC_WSF does a small, fixed
number of SUMIFS against DATA_WSF (60 weeks x 8 measures for the selected
scope, plus 92 Areas x 5 for the rankings) and every visible number reads
from there.
"""
from __future__ import annotations

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border

from config import COLOURS
from .common import (BOX, CENTER, FMT_DEC1, FMT_INT, FMT_INT0, FMT_PCT, FMT_PCT_SIGNED, FMT_RATIO, INK, LEFT, MUTED,
                     RED_FG, RIGHT, WHITE, arrow_rules, dashboard_canvas, fill, font, put, q, rag_rules,
                     section_title, side, signed_rules, tint)
from .controls import SCOPE, SELECT, WEEK, WINDOW, control_strip, search_helper, week_list
from .data import CELLS_CURRENT, SPECS

DASH = "WSF DASHBOARD"
CALC = "CALC_WSF"
S0, S1 = 21, 80            # series rows: k = 0 (selected week) ... 59
KPI0 = 90                  # first KPI row in CALC
AREA0 = 3                  # first Area row in CALC
TOP0 = 3                   # first row of the Top/Bottom lists
NR0 = 3                    # first row of the non-reporting list
NR_ROWS = 50

# series columns in CALC_WSF
SER = dict(k="A", week="B", month="C", qtr="D", year="E", rows="F", male="G", female="H", children="I",
           adult="J", grand="K", cells="L", rep="M", notrep="N", pct="O", avgcell="P", has="Q", win="R",
           mtd="S", qtd="T", ytd="U")

# KPI table: label, kind, source (series col or (numerator, denominator)), number format
KPIS = [
    ("Grand total attendance", "add", "grand", FMT_INT),
    ("Adult total (Male + Female)", "add", "adult", FMT_INT),
    ("Male", "add", "male", FMT_INT),
    ("Female", "add", "female", FMT_INT),
    ("Children", "add", "children", FMT_INT),
    ("Operational cells", "add", "cells", FMT_INT),
    ("Cells that reported", "add", "rep", FMT_INT),
    ("Cells with no report", "add", "notrep", FMT_INT),
    ("% cells reporting", "ratio", ("rep", "cells"), FMT_PCT),
    ("Reporting : non-reporting", "ratio", ("rep", "notrep"), FMT_RATIO),
    ("Avg attendance per reporting cell", "ratio", ("grand", "rep"), FMT_DEC1),
]
KPI_ROW = {label: KPI0 + i for i, (label, *_) in enumerate(KPIS)}
MEASURES = ["Selected week", "Previous week", "Change", "% change", "Avg last N wks", "vs N-wk avg",
            "MTD total", "QTD total", "YTD total", "Wkly avg (month)", "Wkly avg (qtr)", "Wkly avg (year)"]


def cref(cell: str) -> str:
    return f"{CALC}!{cell}"


def build_calc(wb, n_areas: int):
    ws = wb.create_sheet(CALC)
    ws.sheet_state = "hidden"
    d = SPECS["WSF"]
    W, A, Z = d.ref("Week_Ending"), d.ref("Area"), d.ref("Zone")
    D = q(DASH)
    put(ws, "A1", "CALC_WSF — helper calculations for the WSF DASHBOARD (hidden; do not edit)", f=font(11, True))

    # ---- inputs ---------------------------------------------------------------
    inputs = [
        ("Scope", f"={D}!{SCOPE}"),
        ("Selection", f'=IF({D}!{SELECT}="","",{D}!{SELECT})'),
        ("Selection valid", f'=IF(B3="Global",TRUE,IF(B3="Area",COUNTIF(MASTER_AREAS!$B:$B,B4)>0,'
                            f'COUNTIF(MASTER_ZONES!$B:$B,B4)>0))'),
        ("Area criterion", '=IF(B3="Area",B4,"*")'),
        ("Zone criterion", '=IF(B3="Zone",B4,"*")'),
        ("Zone's Area", '=IF(B3="Zone",IFERROR(INDEX(MASTER_ZONES!$C:$C,MATCH(B4,MASTER_ZONES!$B:$B,0)),""),"")'),
        ("Latest week with data", f"=MAX({W})"),
        ("Selected week", f'=IF(ISNUMBER({D}!{WEEK}),{D}!{WEEK},B9)'),
        ("Window N (weeks)", f"=IF(ISNUMBER({D}!{WINDOW}),{D}!{WINDOW},DefaultWindow)"),
        ("Selected month", "=IF(B10=0,0,MONTH(B10))"),
        ("Selected quarter", "=IF(B10=0,0,ROUNDUP(MONTH(B10)/3,0))"),
        ("Selected year", "=IF(B10=0,0,YEAR(B10))"),
        ("Any WSF data", "=B9>0"),
        ("Scope label", '=IF(B3="Global","All Areas",IF(B3="Area","Area: "&IF(B4="","(none selected)",B4),'
                        '"Zone: "&IF(B4="","(none selected)",B4&IF(B8="",""," ("&B8&")"))))'),
        ("Zones with zero cells reporting", f'=IF(B10=0,0,COUNTIFS({W},B10,{A},B6,{Z},B7,{d.ref("Cells_Reported")},0))'),
    ]
    for i, (label, formula) in enumerate(inputs, start=3):
        ws[f"A{i}"] = label
        ws[f"B{i}"] = formula
    ws["B10"].number_format = ws["B9"].number_format = "dd-mmm-yyyy"

    # ---- weekly series for the selected scope ---------------------------------
    heads = ["k", "Week", "Month", "Quarter", "Year", "Rows", "Male", "Female", "Children", "Adult", "Grand",
             "Cells total", "Cells reported", "Not reported", "% reporting", "Avg / cell", "Has data",
             "In window", "In MTD", "In QTD", "In YTD"]
    for j, h in enumerate(heads):
        ws.cell(S0 - 1, j + 1, h).font = font(9, True)
    crit = f"{W},B{{r}},{A},$B$6,{Z},$B$7"
    for k in range(S1 - S0 + 1):
        r = S0 + k
        c = crit.format(r=r)
        sums = lambda col_name: f"=IF(F{r}=0,0,SUMIFS({d.ref(col_name)},{c}))"
        row = {
            "A": k,
            "B": f'=IF($B$10=0,"",$B$10-7*A{r})',
            "C": f'=IF(B{r}="","",MONTH(B{r}))',
            "D": f'=IF(B{r}="","",ROUNDUP(MONTH(B{r})/3,0))',
            "E": f'=IF(B{r}="","",YEAR(B{r}))',
            "F": f'=IF(B{r}="",0,COUNTIFS({c}))',
            "G": sums("Male"), "H": sums("Female"), "I": sums("Children"),
            "J": f"=G{r}+H{r}", "K": f"=J{r}+I{r}",
            "L": sums("Cells_Total"), "M": sums("Cells_Reported"),
            "N": f"=L{r}-M{r}",
            "O": f"=IF(L{r}=0,0,M{r}/L{r})",
            "P": f"=IF(M{r}=0,0,K{r}/M{r})",
            "Q": f"=IF(F{r}>0,1,0)",
            "R": f"=IF(AND(A{r}>=1,A{r}<=$B$11),Q{r},0)",
            "S": f"=IF(AND(C{r}=$B$12,E{r}=$B$14),Q{r},0)",
            "T": f"=IF(AND(D{r}=$B$13,E{r}=$B$14),Q{r},0)",
            "U": f"=IF(E{r}=$B$14,Q{r},0)",
        }
        for colL, v in row.items():
            ws[f"{colL}{r}"] = v
        ws[f"B{r}"].number_format = "dd-mmm-yyyy"

    # ---- KPI x time-measure table ---------------------------------------------
    for j, h in enumerate(["KPI"] + MEASURES):
        ws.cell(KPI0 - 1, j + 1, h).font = font(9, True)
    rng = lambda L: f"${L}${S0}:${L}${S1}"
    R, S, T, U = rng("R"), rng("S"), rng("T"), rng("U")
    for i, (label, kind, src, _fmt) in enumerate(KPIS):
        r = KPI0 + i
        ws[f"A{r}"] = label
        if kind == "add":
            X = SER[src]
            f = {
                "B": f'=IF($Q${S0}=0,"",{X}{S0})',
                "C": f'=IF($Q${S0 + 1}=0,"",{X}{S0 + 1})',
                "F": f'=IF(SUM({R})=0,"",SUMPRODUCT({rng(X)},{R})/SUM({R}))',
                "H": f'=IF(SUM({S})=0,"",SUMPRODUCT({rng(X)},{S}))',
                "I": f'=IF(SUM({T})=0,"",SUMPRODUCT({rng(X)},{T}))',
                "J": f'=IF(SUM({U})=0,"",SUMPRODUCT({rng(X)},{U}))',
                "K": f'=IF(H{r}="","",H{r}/SUM({S}))',
                "L": f'=IF(I{r}="","",I{r}/SUM({T}))',
                "M": f'=IF(J{r}="","",J{r}/SUM({U}))',
            }
        else:
            N_, D_ = SER[src[0]], SER[src[1]]
            ratio = lambda flags: (f'=IF(SUMPRODUCT({rng(D_)},{flags})=0,"",'
                                   f'SUMPRODUCT({rng(N_)},{flags})/SUMPRODUCT({rng(D_)},{flags}))')
            f = {
                "B": f'=IF(OR($Q${S0}=0,{D_}{S0}=0),"",{N_}{S0}/{D_}{S0})',
                "C": f'=IF(OR($Q${S0 + 1}=0,{D_}{S0 + 1}=0),"",{N_}{S0 + 1}/{D_}{S0 + 1})',
                "F": ratio(R), "H": ratio(S), "I": ratio(T), "J": ratio(U),
                "K": f"=H{r}", "L": f"=I{r}", "M": f"=J{r}",
            }
        f["D"] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
        f["E"] = f'=IF(OR(B{r}="",C{r}=""),"",IF(C{r}=0,"",B{r}/C{r}-1))'
        f["G"] = f'=IF(OR(B{r}="",F{r}=""),"",IF(F{r}=0,"",B{r}/F{r}-1))'
        for colL, v in f.items():
            ws[f"{colL}{r}"] = v

    # ---- chart feeds (oldest -> newest) ---------------------------------------
    # 13-week composition: zero on weeks without data (bars simply don't show)
    ws["A104"], ws["B104"], ws["C104"], ws["D104"], ws["E104"] = "Week", "Male", "Female", "Children", "% reporting"
    for j in range(13):
        r, src = 105 + j, S0 + 12 - j
        ws[f"A{r}"] = f'=IF(B{src}="","",TEXT(B{src},"dd mmm"))'
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = f"=G{src}", f"=H{src}", f"=I{src}"
        ws[f"E{r}"] = f"=O{src}"
    # 52-week trend as columns: a week without data is simply an empty slot (no #N/A needed)
    ws["A120"], ws["B120"] = "Week", "Grand total"
    for j in range(52):
        r, src = 121 + j, S0 + 51 - j
        ws[f"A{r}"] = f'=IF(B{src}="","",TEXT(B{src},"dd mmm yy"))'
        ws[f"B{r}"] = f"=K{src}"

    # ---- per-Area block for rankings -------------------------------------------
    heads = ["Area", "Rows", "Grand (sel wk)", "Grand (prev wk)", "Cells total", "Cells reported", "% reporting",
             "Growth", "Score att", "Score pct", "Score growth"]
    for j, h in enumerate(heads):
        ws.cell(AREA0 - 1, 23 + j, h).font = font(9, True)     # starts at column W
    last_area = AREA0 + max(n_areas, 1) - 1
    sw = f"{W},$B$10,{A},W{{r}}"
    for i in range(max(n_areas, 1)):
        r = AREA0 + i
        c = sw.format(r=r)
        ws[f"W{r}"] = f'=IFERROR(INDEX(MASTER_AREAS!$B:$B,{i + 2})&"","")'
        ws[f"X{r}"] = f'=IF(OR(W{r}="",$B$10=0),0,COUNTIFS({c}))'
        ws[f"Y{r}"] = f"=IF(X{r}=0,0,SUMIFS({d.ref('Grand_Total')},{c}))"
        ws[f"Z{r}"] = (f'=IF(OR(W{r}="",$B$10=0),0,SUMIFS({d.ref("Grand_Total")},{W},$B$10-7,{A},W{r}))')
        ws[f"AA{r}"] = f"=IF(X{r}=0,0,SUMIFS({d.ref('Cells_Total')},{c}))"
        ws[f"AB{r}"] = f"=IF(X{r}=0,0,SUMIFS({d.ref('Cells_Reported')},{c}))"
        ws[f"AC{r}"] = f'=IF(AA{r}=0,"",AB{r}/AA{r})'
        ws[f"AD{r}"] = f'=IF(OR(X{r}=0,Z{r}=0),"",Y{r}/Z{r}-1)'
        ws[f"AE{r}"] = f'=IF(X{r}=0,"",Y{r}+ROW()/1000000)'
        ws[f"AF{r}"] = f'=IF(AC{r}="","",AC{r}+ROW()/1000000000)'
        ws[f"AG{r}"] = f'=IF(AD{r}="","",AD{r}+ROW()/1000000000)'
    area_rng = lambda L: f"${L}${AREA0}:${L}${last_area}"

    # ---- Top 10 / Bottom 10 ------------------------------------------------------
    lists = [("AJ", "AK", "LARGE", "AE", "Y"), ("AL", "AM", "SMALL", "AE", "Y"),
             ("AN", "AO", "LARGE", "AF", "AC"), ("AP", "AQ", "SMALL", "AF", "AC"),
             ("AR", "AS", "LARGE", "AG", "AD"), ("AT", "AU", "SMALL", "AG", "AD")]
    ws["AI2"] = "k"
    for name_col, val_col, fn, score, value in lists:
        ws[f"{name_col}2"] = f"{fn} {score} name"
        ws[f"{val_col}2"] = "value"
        for k in range(1, 11):
            r = TOP0 + k - 1
            ws[f"AI{r}"] = k
            m = f"MATCH({fn}({area_rng(score)},{k}),{area_rng(score)},0)"
            ws[f"{name_col}{r}"] = f'=IFERROR(INDEX({area_rng("W")},{m}),"")'
            ws[f"{val_col}{r}"] = f'=IFERROR(INDEX({area_rng(value)},{m}),"")'
    # rank of the selected Area (Area scope only)
    ws["AI15"], ws["AI16"], ws["AI17"], ws["AI18"] = ("Selected area rank: attendance", "% reporting", "growth",
                                                      "Areas ranked")
    for r, score in ((15, "AE"), (16, "AF"), (17, "AG")):
        ws[f"AJ{r}"] = (f'=IF($B$3<>"Area","",IFERROR(COUNTIF({area_rng(score)},">"&'
                        f'INDEX({area_rng(score)},MATCH($B$4,{area_rng("W")},0)))+1,""))')
    ws["AJ18"] = f"=COUNT({area_rng('AE')})"

    # ---- non-reporting cells list ------------------------------------------------
    cc = CELLS_CURRENT
    helper = cc.ref("List_Helper")
    ws["AW1"] = "Non-reporting cells in scope"
    ws["AX1"] = f"=COUNT({helper})"
    ws["AW2"] = "Cell detail held for selected week"
    ws["AX2"] = f'=COUNTIF({cc.ref("Week_Ending")},$B$10)>0'
    for k in range(1, NR_ROWS + 1):
        r = NR0 + k
        ws[f"AW{r}"] = f'=IFERROR(SMALL({helper},{k}),"")'
        for colL, src in (("AX", "Area"), ("AY", "Zone"), ("AZ", "Cell")):
            ws[f"{colL}{r}"] = f'=IF($AW{r}="","",INDEX({cc.ref(src)},$AW{r}))'
    return ws


# --------------------------------------------------------------------------
# Dashboard
# --------------------------------------------------------------------------

def dash_value(src: str) -> str:
    """Show an en dash instead of blank / missing values."""
    return f'=IF({src}="","–",{src})'


def build_dashboard(wb, n_zones: int):
    colour = COLOURS["WSF"]
    ws = wb.create_sheet(DASH)
    ws.sheet_properties.tabColor = colour
    dashboard_canvas(ws, "WSF", "WSF DASHBOARD",
                     "Worship Service Fellowship · Saturday cell services · attendance and cell reporting compliance",
                     n_cols=14, widths={"B": 15})
    select_list = search_helper(wb, "WSF", DASH, 0, n_zones + 200)
    week_name, _, _ = week_list(wb, "WSF", SPECS["WSF"].ref("Week_Ending"), 0)
    control_strip(wb, ws, "WSF", colour, "WSF", select_list, week_name)
    C = lambda cell: f"{CALC}!{cell}"
    ws["L4"] = (f'={C("$B$16")}&IF({C("$B$10")}=0,"","  ·  w/e "&TEXT({C("$B$10")},"dd mmm yyyy"))'
                f'&"  ·  "&{C("$B$11")}&"-wk window"')
    ws.merge_cells("L4:O4")
    ws["L4"].font = font(10, True, INK)
    ws["B5"] = (f'=IF({C("$B$9")}=0,"No WSF data loaded yet. Put files in the inbox folder and run ingest.py.",'
                f'IF(NOT({C("$B$5")}),"⚠  Pick "&IF({C("$B$3")}="Area","an Area","a Zone")&'
                f'" from SELECT (type part of the name in SEARCH first to narrow the list).",'
                f'IF({C("$Q$21")}=0,"No WSF reports in this scope for the selected week.",'
                f'"Numbers compare the selected week with the previous week and with the average of the "'
                f'&{C("$B$11")}&" weeks before it.")))')
    ws.merge_cells("B5:O5")
    ws["B5"].font = font(9, False, tint(colour, 0.2), italic=True)
    ws["B5"].alignment = LEFT

    # ---- KPI cards ------------------------------------------------------------
    cards = [
        ("GRAND TOTAL ATTENDANCE", "Grand total attendance", FMT_INT, "pct"),
        ("ADULTS (M + F)", "Adult total (Male + Female)", FMT_INT, "pct"),
        ("CHILDREN", "Children", FMT_INT, "pct"),
        ("% CELLS REPORTING", "% cells reporting", FMT_PCT, "pts"),
        ("CELLS REPORTED", "Cells that reported", FMT_INT, "pct"),
        ("AVG PER REPORTING CELL", "Avg attendance per reporting cell", FMT_DEC1, "pct"),
        ("ZONES WITH NO CELL REPORT", None, FMT_INT0, None),
    ]
    first_cols = ["B", "D", "F", "H", "J", "L", "N"]
    card_bg = tint(colour, 0.92)
    for (title, kpi, fmt, delta), c1 in zip(cards, first_cols):
        c2 = chr(ord(c1) + 1)
        for r in range(7, 11):
            for cc in (c1, c2):
                cell = ws[f"{cc}{r}"]
                cell.fill = fill(card_bg)
                cell.border = Border(left=side(WHITE, "thick") if cc == c1 else None,
                                     right=side(WHITE, "thick") if cc == c2 else None,
                                     top=side(colour, "thick") if r == 7 else None)
        put(ws, f"{c1}7", title, f=font(8, True, tint(colour, 0.15)), align=CENTER, merge_to=f"{c2}7")
        if kpi:
            r = KPI_ROW[kpi]
            val = dash_value(C(f"$B${r}"))
            if delta == "pct":
                dtext = (f'=IF({C(f"$E${r}")}="","– vs prev wk",IF({C(f"$E${r}")}>=0,"▲ ","▼ ")&'
                         f'TEXT(ABS({C(f"$E${r}")}),"0.0%")&" vs prev wk")')
            else:
                dtext = (f'=IF({C(f"$D${r}")}="","– vs prev wk",IF({C(f"$D${r}")}>=0,"▲ ","▼ ")&'
                         f'TEXT(ABS({C(f"$D${r}")})*100,"0.0")&" pts vs prev wk")')
            avg_fmt = {"0.0%": "0.0%", FMT_DEC1: "#,##0.0"}.get(fmt, "#,##0")
            atext = (f'="Avg last "&{C("$B$11")}&" wks: "&IF({C(f"$F${r}")}="","–",'
                     f'TEXT({C(f"$F${r}")},"{avg_fmt}"))')
        else:
            val = f'=IF({C("$Q$21")}=0,"–",{C("$B$17")})'
            dtext = f'=IF({C("$Q$21")}=0,"","of "&TEXT(COUNTIFS({SPECS["WSF"].ref("Week_Ending")},{C("$B$10")},' \
                    f'{SPECS["WSF"].ref("Area")},{C("$B$6")},{SPECS["WSF"].ref("Zone")},{C("$B$7")}),"#,##0")&" zones")'
            atext = '="Zones where not one cell reported"'
        put(ws, f"{c1}8", val, f=font(20, True, colour), fmt=fmt, align=CENTER, merge_to=f"{c2}8")
        put(ws, f"{c1}9", dtext, f=font(9, True, MUTED), align=CENTER, merge_to=f"{c2}9")
        put(ws, f"{c1}10", atext, f=font(8, False, MUTED), align=CENTER, merge_to=f"{c2}10")
        arrow_rules(ws, f"{c1}9", f"{c1}9")
        if kpi == "% cells reporting":
            rag_rules(ws, f"{c1}8", f"{c1}8")
    ws.row_dimensions[8].height = 34

    # ---- time analysis table ------------------------------------------------------
    section_title(ws, 12, "Time analysis — selected week against previous week, the N-week average and "
                          "period-to-date", colour, last="O")
    put(ws, "B13", "KPI", f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to="C13")
    heads = MEASURES[:]
    for j, h in enumerate(heads):
        cell = ws.cell(13, 4 + j)
        if h == "Avg last N wks":
            cell.value = f'="Avg last "&{C("$B$11")}&" wks"'
        elif h == "vs N-wk avg":
            cell.value = f'="vs "&{C("$B$11")}&"-wk avg"'
        else:
            cell.value = h
        cell.font = font(9, True, WHITE)
        cell.fill = fill(colour)
        cell.alignment = CENTER
    ws.row_dimensions[13].height = 28
    for i, (label, kind, src, fmt) in enumerate(KPIS):
        r = 14 + i
        cr = KPI0 + i
        band = tint(colour, 0.95) if i % 2 == 0 else None
        put(ws, f"B{r}", label, f=font(9, True if i in (0, 8) else False), align=LEFT, bg=band, merge_to=f"C{r}")
        for j, L in enumerate("BCDEFGHIJKLM"):
            cell = ws.cell(r, 4 + j)
            cell.value = dash_value(C(f"${L}${cr}"))
            if L in "EG":
                cell.number_format = FMT_PCT_SIGNED
            elif L == "D" and fmt == FMT_PCT:
                cell.number_format = '+0.0" pts";-0.0" pts";0.0" pts"'
                cell.value = f'=IF({C(f"$D${cr}")}="","–",{C(f"$D${cr}")}*100)'
            elif L == "D":
                cell.number_format = {FMT_INT: '+#,##0;-#,##0;0', FMT_RATIO: '+0.0;-0.0;0.0',
                                      FMT_DEC1: '+#,##0.0;-#,##0.0;0.0'}[fmt]
            elif L in "KLM" and fmt == FMT_INT:
                cell.number_format = "#,##0"
            else:
                cell.number_format = fmt
            cell.font = font(9, L == "B")
            cell.alignment = RIGHT
            if band:
                cell.fill = fill(band)
        signed_rules(ws, f"F{r}:F{r}", f"F{r}")
        signed_rules(ws, f"G{r}:G{r}", f"G{r}")
        signed_rules(ws, f"I{r}:I{r}", f"I{r}")
        if fmt == FMT_PCT and kind == "ratio" and src == ("rep", "cells"):
            rag_rules(ws, f"D{r}:E{r}", f"D{r}")
            rag_rules(ws, f"H{r}", f"H{r}")
            rag_rules(ws, f"J{r}:O{r}", f"J{r}")
    put(ws, "B25", "Averages use only weeks that have data. MTD/QTD/YTD run to the selected week; for ratio rows the "
                   "period columns show the ratio of the period's totals. Cell counts in period columns are "
                   "cell-weeks.", f=font(8, False, MUTED, italic=True), align=LEFT, merge_to="O25")

    # ---- charts ----------------------------------------------------------------------
    section_title(ws, 27, "Trends", colour, last="O")
    calc = wb[CALC]
    shades = [colour, tint(colour, 0.45), "E0A526"]
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap = "col", "stacked", 100
    _title(bar, "Attendance mix — last 13 weeks")
    for idx, colL in enumerate("BCD"):
        bar.add_data(Reference(calc, min_col=ord(colL) - 64, min_row=104, max_row=117), titles_from_data=True)
    bar.set_categories(Reference(calc, min_col=1, min_row=105, max_row=117))
    for s, shade in zip(bar.series, shades):
        s.graphicalProperties.solidFill = shade
        s.graphicalProperties.line.solidFill = shade
    bar.gapWidth = 60
    bar.legend.position = "b"
    bar.y_axis.numFmt = "#,##0"
    _style_chart(bar)
    ws.add_chart(bar, "B28")

    trend = BarChart()
    trend.type = "col"
    _title(trend, "Grand total — last 52 weeks")
    trend.add_data(Reference(calc, min_col=2, min_row=120, max_row=172), titles_from_data=True)
    trend.set_categories(Reference(calc, min_col=1, min_row=121, max_row=172))
    trend.series[0].graphicalProperties.solidFill = colour
    trend.series[0].graphicalProperties.line.solidFill = colour
    trend.gapWidth = 40
    trend.y_axis.numFmt = "#,##0"
    trend.legend = None
    _style_chart(trend)
    ws.add_chart(trend, "G28")

    pct = LineChart()
    _title(pct, "% cells reporting — last 13 weeks")
    pct.add_data(Reference(calc, min_col=5, min_row=104, max_row=117), titles_from_data=True)
    pct.set_categories(Reference(calc, min_col=1, min_row=105, max_row=117))
    pct.series[0].graphicalProperties.line.solidFill = colour
    pct.series[0].graphicalProperties.line.width = 28000
    pct.series[0].marker.symbol = "circle"
    pct.series[0].marker.size = 6
    pct.series[0].marker.graphicalProperties.solidFill = colour
    pct.series[0].marker.graphicalProperties.line.solidFill = colour
    pct.series[0].smooth = False
    pct.y_axis.numFmt = "0%"
    pct.y_axis.scaling.min, pct.y_axis.scaling.max = 0, 1
    pct.legend = None
    _style_chart(pct)
    ws.add_chart(pct, "L28")
    for chart, w in ((bar, 13.0), (trend, 13.0), (pct, 10.4)):
        chart.width, chart.height = w, 7.4

    # ---- Top / Bottom 10 Areas ------------------------------------------------------
    section_title(ws, 44, "Area rankings — selected week", colour, last="O",
                  note="Growth = selected week vs previous week")
    tables = [("B", "Top 10 · attendance", "AJ", "AK", FMT_INT), ("D", "Bottom 10 · attendance", "AL", "AM", FMT_INT),
              ("F", "Top 10 · % reporting", "AN", "AO", FMT_PCT), ("H", "Bottom 10 · % reporting", "AP", "AQ", FMT_PCT),
              ("J", "Top 10 · growth", "AR", "AS", FMT_PCT_SIGNED), ("L", "Bottom 10 · growth", "AT", "AU", FMT_PCT_SIGNED)]
    zone_scope = f'{C("$B$3")}="Zone"'
    for c1, title, ncol, vcol, fmt in tables:
        c2 = chr(ord(c1) + 1)
        put(ws, f"{c1}45", title, f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}45")
        for k in range(10):
            r, cr = 46 + k, TOP0 + k
            nm = C(f"${ncol}${cr}")
            if k == 0:
                ws[f"{c1}{r}"] = f'=IF({zone_scope},"n/a at Zone scope",IF({nm}="","–",{k + 1}&". "&{nm}))'
            else:
                ws[f"{c1}{r}"] = f'=IF({zone_scope},"",IF({nm}="","",{k + 1}&". "&{nm}))'
            ws[f"{c2}{r}"] = f'=IF(OR({zone_scope},{nm}=""),"",{C(f"${vcol}${cr}")})'
            ws[f"{c1}{r}"].font = font(9)
            ws[f"{c2}{r}"].font = font(9, True)
            ws[f"{c2}{r}"].number_format = fmt
            ws[f"{c2}{r}"].alignment = RIGHT
            for cc in (c1, c2):
                ws[f"{cc}{r}"].border = Border(bottom=side())
        if fmt == FMT_PCT:
            rag_rules(ws, f"{c2}46:{c2}55", f"{c2}46")
        elif fmt == FMT_INT:
            ws.conditional_formatting.add(f"{c2}46:{c2}55", DataBarRule(start_type="num", start_value=0,
                                                                        end_type="max", color=tint(colour, 0.35)))
        else:
            signed_rules(ws, f"{c2}46:{c2}55", f"{c2}46")
    # selected area rank box
    put(ws, "N45", "Selected Area rank", f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to="O45")
    rank_rows = [("Attendance", "$AJ$15"), ("% reporting", "$AJ$16"), ("Growth", "$AJ$17")]
    for i, (label, ref) in enumerate(rank_rows):
        r = 46 + i
        put(ws, f"N{r}", label, f=font(9))
        put(ws, f"O{r}", f'=IF({C("$B$3")}<>"Area","n/a",IF({C(ref)}="","–",{C(ref)}&" of "&{C("$AJ$18")}))',
            f=font(9, True, colour), align=RIGHT)
    put(ws, "N50", "Shown when Scope = Area. Rank 1 = highest.", f=font(8, False, MUTED, italic=True),
        align=Alignment(wrap_text=True, vertical="top"), merge_to="O52")

    # ---- non-reporting list ----------------------------------------------------------
    section_title(ws, 57, "Who didn't report — cells with no WSF return for the selected week", colour, last="O")
    ws["B58"] = (f'=IF(NOT({C("$AX$2")}),"Cell-level detail is kept in the workbook for the latest '
                 f'"&{8}&" weeks only; older weeks are in the archive folder (WSF_cells_YYYY.csv).",'
                 f'TEXT({C("$AX$1")},"#,##0")&" cells did not report in this scope"&IF({C("$AX$1")}>{NR_ROWS},'
                 f'" — showing the first {NR_ROWS}. Narrow with Scope = Area or Zone.",".")&'
                 f'"  Zones where no cell reported: "&TEXT({C("$B$17")},"#,##0")&".")')
    ws.merge_cells("B58:O58")
    ws["B58"].font = font(9, True, RED_FG)
    for c1, c2, h in (("B", "D", "Area"), ("E", "H", "Zone"), ("I", "J", "Cell"), ("K", "L", "Week ending")):
        put(ws, f"{c1}59", h, f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to=f"{c2}59")
    for k in range(NR_ROWS):
        r, cr = 60 + k, NR0 + 1 + k
        for c1, c2, src in (("B", "D", "AX"), ("E", "H", "AY"), ("I", "J", "AZ")):
            put(ws, f"{c1}{r}", f"={C(f'${src}${cr}')}", f=font(9), merge_to=f"{c2}{r}")
        put(ws, f"K{r}", f'=IF({C(f"$AW${cr}")}="","",{C("$B$10")})', f=font(9), fmt="dd-mmm-yyyy",
            merge_to=f"L{r}")
        if k % 2 == 0:
            for cc in "BCDEFGHIJKL":
                ws[f"{cc}{r}"].fill = fill(tint(colour, 0.95))
    ws.print_area = f"A1:P{60 + NR_ROWS}"
    ws.print_title_rows = "1:5"
    return ws


def _title(chart, text: str, size: int = 1000):
    """Chart title at a readable size (the default is very large)."""
    from openpyxl.chart.text import RichText, Text
    from openpyxl.chart.title import Title
    from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
    cp = CharacterProperties(sz=size, b=True, solidFill="1F2933")
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    chart.title = Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def _style_chart(chart):
    chart.style = 2
    chart.roundedCorners = False
    try:
        from openpyxl.chart.text import RichText
        from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties
        cp = CharacterProperties(sz=800, latin=None)
        for ax in (chart.x_axis, chart.y_axis):
            ax.txPr = RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])
            ax.delete = False
        chart.y_axis.majorGridlines.spPr = None
    except Exception:
        pass
