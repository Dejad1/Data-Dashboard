"""CALC_TRANSPORT (hidden weekly aggregates) and the TRANSPORT DASHBOARD.

Transport comes from the transport office's weekly reports (see
transport_reports.py): FT Procured hired LT buses, church coasters, EV/TATA
buses and WSF Procured (member-paid) buses, plus coaster fuel and the FT
budget.  As on the WSF dashboard, every visible number reads from a small
fixed set of SUMIFS in CALC_TRANSPORT.
"""
from __future__ import annotations

from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border

from config import COLOURS
from .common import (CENTER, FMT_DEC1, FMT_INT, FMT_NGN, FMT_PCT, FMT_PCT_SIGNED, INK, LEFT, MUTED, RED_FG, RIGHT,
                     WHITE, arrow_rules, dashboard_canvas, fill, font, put, q, rag_rules, section_title, side,
                     signed_rules, tint)
from .controls import SCOPE, SELECT, WEEK, WINDOW, control_strip, search_helper, week_list
from .data import SPECS
from .wsf import _style_chart, _title

DASH = "TRANSPORT DASHBOARD"
CALC = "CALC_TRANSPORT"
S0, S1 = 21, 80
KPI0 = 90
CAT0 = 110            # category table rows
AREA0 = 3             # per-Area block (columns AE onwards)
NOTES = 40            # rows in the operational notes list
CATEGORIES = ["FT Procured", "Church Coaster", "EV/TATA", "WSF Procured"]
LOCATIONS = ["Hub", "Loading Bay"]      # TATA/EV/coasters load at Hubs; FT & WSF Procured at Loading Bays
FMT_NGN_INT = '"₦"#,##0'
FMT_NGN_DEC = '"₦"#,##0.00'

# series columns: key -> (letter, header)
SER = {}
_order = ["k", "week", "month", "qtr", "year", "rows", "riders", "male", "female", "children", "listed", "ran",
          "seats", "rws", "excess", "hire", "fuel", "spend", "central", "members", "ft_budget", "ft_cost", "breakdowns",
          "util", "has", "win", "mtd", "qtd", "ytd"]
for _i, _k in enumerate(_order):
    from openpyxl.utils import get_column_letter as _gcl
    SER[_k] = _gcl(_i + 1)

KPIS = [
    ("Riders carried", "add", "riders", FMT_INT),
    ("Buses that ran", "add", "ran", FMT_INT),
    ("Buses listed in reports", "add", "listed", FMT_INT),
    ("% buses that ran", "ratio", ("ran", "listed"), FMT_PCT),
    ("Seats offered", "add", "seats", FMT_INT),
    ("Seat utilisation %", "ratio", ("rws", "seats"), FMT_PCT),
    ("Riders beyond seats (unmet demand)", "add", "excess", FMT_INT),
    ("Total spend", "add", "spend", FMT_NGN_INT),
    ("Central spend (FT hire + coaster fuel)", "add", "central", FMT_NGN_INT),
    ("Member-paid (WSF Procured)", "add", "members", FMT_NGN_INT),
    ("Coaster fuel paid (hubs)", "add", "fuel", FMT_NGN_INT),
    ("Cost per rider", "ratio", ("spend", "riders"), FMT_NGN_DEC),
    ("Cost per bus that ran", "ratio", ("spend", "ran"), FMT_NGN_INT),
    ("Cost per seat offered", "ratio", ("spend", "seats"), FMT_NGN_DEC),
    ("FT spend vs FT budget", "ratio", ("ft_cost", "ft_budget"), FMT_PCT),
    ("Breakdowns", "add", "breakdowns", FMT_INT),
]
KPI_ROW = {label: KPI0 + i for i, (label, *_) in enumerate(KPIS)}
MEASURES = ["Selected week", "Previous week", "Change", "% change", "Avg last N wks", "vs N-wk avg",
            "MTD total", "QTD total", "YTD total", "Wkly avg (month)", "Wkly avg (qtr)", "Wkly avg (year)"]


def build_calc(wb, n_areas: int):
    ws = wb.create_sheet(CALC)
    ws.sheet_state = "hidden"
    T, CO, BU = SPECS["TRANSPORT"], SPECS["TRANSPORT_COSTS"], SPECS["TRANSPORT_BUDGET"]
    D = q(DASH)
    put(ws, "A1", "CALC_TRANSPORT — helper calculations for the TRANSPORT DASHBOARD (hidden; do not edit)",
        f=font(11, True))
    inputs = [
        ("Scope", f"={D}!{SCOPE}"),
        ("Selection", f'=IF({D}!{SELECT}="","",{D}!{SELECT})'),
        ("Selection valid", '=IF(B3="Global",TRUE,IF(B3="Area",COUNTIF(MASTER_AREAS!$B:$B,B4)>0,'
                            'COUNTIF(MASTER_ZONES!$B:$B,B4)>0))'),
        ("Area criterion", '=IF(B3="Area",B4,"*")'),
        ("Zone criterion", '=IF(B3="Zone",B4,"*")'),
        ("Zone's Area", '=IF(B3="Zone",IFERROR(INDEX(MASTER_ZONES!$C:$C,MATCH(B4,MASTER_ZONES!$B:$B,0)),""),"")'),
        ("Latest week with data", f"=MAX({T.ref('Week_Ending')})"),
        ("Selected week", f"=IF(ISNUMBER({D}!{WEEK}),{D}!{WEEK},B9)"),
        ("Window N (weeks)", f"=IF(ISNUMBER({D}!{WINDOW}),{D}!{WINDOW},DefaultWindow)"),
        ("Selected month", "=IF(B10=0,0,MONTH(B10))"),
        ("Selected quarter", "=IF(B10=0,0,ROUNDUP(MONTH(B10)/3,0))"),
        ("Selected year", "=IF(B10=0,0,YEAR(B10))"),
        ("Any transport data", "=B9>0"),
        ("Scope label", '=IF(B3="Global","All Areas",IF(B3="Area","Area: "&IF(B4="","(none selected)",B4),'
                        '"Zone: "&IF(B4="","(none selected)",B4&IF(B8="",""," ("&B8&")"))))'),
        ("Zone scope (costs not split by zone)", '=B3="Zone"'),
    ]
    for i, (label, formula) in enumerate(inputs, start=3):
        ws[f"A{i}"], ws[f"B{i}"] = label, formula
    ws["B9"].number_format = ws["B10"].number_format = "dd-mmm-yyyy"

    # ---- weekly series ------------------------------------------------------------
    for key, L in SER.items():
        ws[f"{L}{S0 - 1}"] = key
    tcrit = lambda wk: f"{T.ref('Week_Ending')},{wk},{T.ref('Area')},$B$6,{T.ref('Zone_Key')},$B$7"
    ccrit = lambda wk: f"{CO.ref('Week_Ending')},{wk},{CO.ref('Area')},$B$6"
    bcrit = lambda wk: f"{BU.ref('Week_Ending')},{wk},{BU.ref('Area')},$B$6"
    for k in range(S1 - S0 + 1):
        r = S0 + k
        S = {key: f"{L}{r}" for key, L in SER.items()}
        wk = S["week"]
        tsum = lambda col, extra="": f"=IF({S['rows']}=0,0,SUMIFS({T.ref(col)},{tcrit(wk)}{extra}))"
        f = {
            "k": k,
            "week": f'=IF($B$10=0,"",$B$10-7*{S["k"]})',
            "month": f'=IF({wk}="","",MONTH({wk}))',
            "qtr": f'=IF({wk}="","",ROUNDUP(MONTH({wk})/3,0))',
            "year": f'=IF({wk}="","",YEAR({wk}))',
            "rows": f'=IF({wk}="",0,COUNTIFS({tcrit(wk)}))',
            "riders": tsum("Grand_Total"), "male": tsum("Male"), "female": tsum("Female"),
            "children": tsum("Children"), "listed": tsum("Buses"),
            "ran": tsum("Buses", f',{T.ref("Status")},"Ran"'),
            "seats": tsum("Seats_Offered"), "rws": tsum("Riders_Seated"), "excess": tsum("Excess_Riders"),
            "hire": tsum("Cost"),
            "fuel": f'=IF(OR({wk}="",$B$17),0,SUMIFS({CO.ref("Paid_Amount")},{ccrit(wk)}))',
            "spend": f"={S['hire']}+{S['fuel']}",
            "central": f"={S['fuel']}+" + tsum("Cost", f',{T.ref("Paid_By")},"Central"')[1:],
            "members": tsum("Cost", f',{T.ref("Paid_By")},"Members"'),
            "ft_budget": f'=IF(OR({wk}="",$B$17),0,SUMIFS({BU.ref("Expected_Spend")},{bcrit(wk)}))',
            "ft_cost": tsum("Cost", f',{T.ref("Category")},"FT Procured"'),
            "breakdowns": f'=IF({S["rows"]}=0,0,COUNTIFS({tcrit(wk)},{T.ref("Status")},"Breakdown"))',
            "util": f"=IF({S['seats']}=0,0,{S['rws']}/{S['seats']})",
            "has": f"=IF({S['rows']}>0,1,0)",
            "win": f"=IF(AND({S['k']}>=1,{S['k']}<=$B$11),{S['has']},0)",
            "mtd": f"=IF(AND({S['month']}=$B$12,{S['year']}=$B$14),{S['has']},0)",
            "qtd": f"=IF(AND({S['qtr']}=$B$13,{S['year']}=$B$14),{S['has']},0)",
            "ytd": f"=IF({S['year']}=$B$14,{S['has']},0)",
        }
        for key, v in f.items():
            ws[S[key]] = v
        ws[S["week"]].number_format = "dd-mmm-yyyy"

    # ---- KPI x time measures --------------------------------------------------------
    for j, h in enumerate(["KPI"] + MEASURES):
        ws.cell(KPI0 - 1, j + 1, h)
    rng = lambda key: f"${SER[key]}${S0}:${SER[key]}${S1}"
    flags = {x: rng(x) for x in ("win", "mtd", "qtd", "ytd")}
    has0, has1 = f"${SER['has']}${S0}", f"${SER['has']}${S0 + 1}"
    for i, (label, kind, src, _fmt) in enumerate(KPIS):
        r = KPI0 + i
        ws[f"A{r}"] = label
        if kind == "add":
            X = SER[src]
            f = {"B": f'=IF({has0}=0,"",{X}{S0})', "C": f'=IF({has1}=0,"",{X}{S0 + 1})'}
            for col, fl in zip("FHIJ", ("win", "mtd", "qtd", "ytd")):
                if col == "F":
                    f[col] = f'=IF(SUM({flags[fl]})=0,"",SUMPRODUCT({rng(src)},{flags[fl]})/SUM({flags[fl]}))'
                else:
                    f[col] = f'=IF(SUM({flags[fl]})=0,"",SUMPRODUCT({rng(src)},{flags[fl]}))'
            for col, base, fl in (("K", "H", "mtd"), ("L", "I", "qtd"), ("M", "J", "ytd")):
                f[col] = f'=IF({base}{r}="","",{base}{r}/SUM({flags[fl]}))'
        else:
            N_, D_ = SER[src[0]], SER[src[1]]
            ratio = lambda fl: (f'=IF(SUMPRODUCT({rng(src[1])},{flags[fl]})=0,"",'
                                f'SUMPRODUCT({rng(src[0])},{flags[fl]})/SUMPRODUCT({rng(src[1])},{flags[fl]}))')
            f = {"B": f'=IF(OR({has0}=0,{D_}{S0}=0),"",{N_}{S0}/{D_}{S0})',
                 "C": f'=IF(OR({has1}=0,{D_}{S0 + 1}=0),"",{N_}{S0 + 1}/{D_}{S0 + 1})',
                 "F": ratio("win"), "H": ratio("mtd"), "I": ratio("qtd"), "J": ratio("ytd"),
                 "K": f"=H{r}", "L": f"=I{r}", "M": f"=J{r}"}
        f["D"] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
        f["E"] = f'=IF(OR(B{r}="",C{r}=""),"",IF(C{r}=0,"",B{r}/C{r}-1))'
        f["G"] = f'=IF(OR(B{r}="",F{r}=""),"",IF(F{r}=0,"",B{r}/F{r}-1))'
        for col, v in f.items():
            ws[f"{col}{r}"] = v

    # ---- by category, selected week (row CAT0 ...) ------------------------------------
    heads = ["Category", "Listed", "Ran", "% ran", "Riders", "Seats", "Riders seated", "Utilisation", "Hire cost",
             "Fuel paid", "Spend", "Cost/rider", "Cost/bus", "Cost/seat", "Riders prev wk", "Riders WoW", "Paid by",
             "Excess riders"]
    for j, h in enumerate(heads):
        ws.cell(CAT0 - 1, j + 1, h)
    for i, cat in enumerate(CATEGORIES + ["All categories"] + LOCATIONS):
        r = CAT0 + i
        if cat in LOCATIONS:
            cc = f',{T.ref("Location_Type")},"{cat}"'
        else:
            cc = f',{T.ref("Category")},"{cat}"' if cat != "All categories" else ""
        base = f"{tcrit('$B$10')}{cc}"
        prev = f"{tcrit('$B$10-7')}{cc}"
        ws[f"A{r}"] = cat
        ws[f"B{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Buses')},{base}))"
        ws[f"C{r}"] = f'=IF($B$10=0,0,SUMIFS({T.ref("Buses")},{base},{T.ref("Status")},"Ran"))'
        ws[f"D{r}"] = f'=IF(B{r}=0,"",C{r}/B{r})'
        ws[f"E{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Grand_Total')},{base}))"
        ws[f"F{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Seats_Offered')},{base}))"
        ws[f"G{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Riders_Seated')},{base}))"
        ws[f"R{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Excess_Riders')},{base}))"
        ws[f"H{r}"] = f'=IF(F{r}=0,"",G{r}/F{r})'
        ws[f"I{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Cost')},{base}))"
        if cat in ("Church Coaster", "All categories"):
            ws[f"J{r}"] = f"=IF(OR($B$10=0,$B$17),0,SUMIFS({CO.ref('Paid_Amount')},{ccrit('$B$10')}))"
        elif cat in LOCATIONS:
            status = "HUB" if cat == "Hub" else "ZONE"
            ws[f"J{r}"] = (f'=IF(OR($B$10=0,$B$17),0,SUMIFS({CO.ref("Paid_Amount")},{ccrit("$B$10")},'
                           f'{CO.ref("Hub_Status")},"{status}"))')
        else:
            ws[f"J{r}"] = 0
        ws[f"K{r}"] = f"=I{r}+J{r}"
        ws[f"L{r}"] = f'=IF(OR(E{r}=0,K{r}=0),"",K{r}/E{r})'
        ws[f"M{r}"] = f'=IF(OR(C{r}=0,K{r}=0),"",K{r}/C{r})'
        ws[f"N{r}"] = f'=IF(OR(F{r}=0,K{r}=0),"",K{r}/F{r})'
        ws[f"O{r}"] = f"=IF($B$10=0,0,SUMIFS({T.ref('Grand_Total')},{prev}))"
        ws[f"P{r}"] = f'=IF(O{r}=0,"",E{r}/O{r}-1)'
        ws[f"Q{r}"] = {"FT Procured": "Central (hire)", "Church Coaster": "Central (fuel)",
                       "EV/TATA": "Central (no cost)", "WSF Procured": "Members",
                       "Hub": "Mixed", "Loading Bay": "Mixed"}.get(cat, "")

    # ---- chart feeds: 13 weeks oldest -> newest ------------------------------------------
    ws["A125"] = "Week"
    for j, cat in enumerate(CATEGORIES):
        ws.cell(125, 2 + j, cat)                       # riders by category
        ws.cell(125, 7 + j, cat)                       # spend by category
    ws["K125"] = "Seat utilisation"
    for j in range(13):
        r, src = 126 + j, S0 + 12 - j
        wk = f"{SER['week']}{src}"
        ws[f"A{r}"] = f'=IF({wk}="","",TEXT({wk},"dd mmm"))'
        for c, cat in enumerate(CATEGORIES):
            crit = f'{tcrit(wk)},{T.ref("Category")},"{cat}"'
            ws.cell(r, 2 + c, f'=IF({wk}="",0,SUMIFS({T.ref("Grand_Total")},{crit}))')
            spend = f'SUMIFS({T.ref("Cost")},{crit})'
            if cat == "Church Coaster":
                spend += f'+IF($B$17,0,SUMIFS({CO.ref("Paid_Amount")},{ccrit(wk)}))'
            ws.cell(r, 7 + c, f'=IF({wk}="",0,{spend})')
        ws[f"K{r}"] = f"={SER['util']}{src}"

    # ---- per-Area block (columns AE..) ------------------------------------------------------
    heads = ["Area", "Riders", "Buses ran", "Seats", "Riders seated", "Excess riders", "Utilisation", "Spend",
             "Cost/rider",
             "FT cost", "FT budget", "FT over budget", "FT buses in church", "Buses allocated", "Breakdowns",
             "Util wk-1", "Util wk-2", "Util wk-3", "Weeks under UtilLow", "Flag", "Score util", "Score cpr",
             "Score over", "Score riders", "Flag row", "Score excess"]
    c0 = 31                                                  # column AE
    from openpyxl.utils import get_column_letter as L_
    A = {h: L_(c0 + j) for j, h in enumerate(heads)}
    for h, colL in A.items():
        ws[f"{colL}{AREA0 - 1}"] = h
    last = AREA0 + max(n_areas, 1) - 1
    for i in range(max(n_areas, 1)):
        r = AREA0 + i
        a = f"{A['Area']}{r}"
        ac = lambda wk: f"{T.ref('Week_Ending')},{wk},{T.ref('Area')},{a}"
        ws[a] = f'=IFERROR(INDEX(MASTER_AREAS!$B:$B,{i + 2})&"","")'
        g = lambda col, wk="$B$10", extra="": f'IF(OR({a}="",$B$10=0),0,SUMIFS({T.ref(col)},{ac(wk)}{extra}))'
        ws[f"{A['Riders']}{r}"] = "=" + g("Grand_Total")
        ws[f"{A['Buses ran']}{r}"] = "=" + g("Buses", extra=f',{T.ref("Status")},"Ran"')
        ws[f"{A['Seats']}{r}"] = "=" + g("Seats_Offered")
        ws[f"{A['Riders seated']}{r}"] = "=" + g("Riders_Seated")
        ws[f"{A['Excess riders']}{r}"] = "=" + g("Excess_Riders")
        ws[f"{A['Utilisation']}{r}"] = f'=IF({A["Seats"]}{r}=0,"",{A["Riders seated"]}{r}/{A["Seats"]}{r})'
        ws[f"{A['Spend']}{r}"] = ("=" + g("Cost") + f'+IF(OR({a}="",$B$10=0),0,SUMIFS({CO.ref("Paid_Amount")},'
                                                     f'{CO.ref("Week_Ending")},$B$10,{CO.ref("Area")},{a}))')
        ws[f"{A['Cost/rider']}{r}"] = f'=IF(OR({A["Riders"]}{r}=0,{A["Spend"]}{r}=0),"",{A["Spend"]}{r}/{A["Riders"]}{r})'
        ws[f"{A['FT cost']}{r}"] = "=" + g("Cost", extra=f',{T.ref("Category")},"FT Procured"')
        ws[f"{A['FT budget']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,SUMIFS({BU.ref("Expected_Spend")},'
                                      f'{BU.ref("Week_Ending")},$B$10,{BU.ref("Area")},{a}))')
        ws[f"{A['FT over budget']}{r}"] = (f'=IF({A["FT budget"]}{r}=0,"",{A["FT cost"]}{r}-{A["FT budget"]}{r})')
        ws[f"{A['FT buses in church']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,COUNTIFS({ac("$B$10")},'
                                               f'{T.ref("Category")},"FT Procured",{T.ref("In_Church")},"Y"))')
        ws[f"{A['Buses allocated']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,SUMIFS({BU.ref("Buses_Allocated")},'
                                            f'{BU.ref("Week_Ending")},$B$10,{BU.ref("Area")},{a}))')
        ws[f"{A['Breakdowns']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,COUNTIFS({ac("$B$10")},'
                                       f'{T.ref("Status")},"Breakdown"))')
        for j, h in ((1, "Util wk-1"), (2, "Util wk-2"), (3, "Util wk-3")):
            wk = f"$B$10-{7 * j}"
            seats = g("Seats_Offered", wk)
            ws[f"{A[h]}{r}"] = f'=IF({seats}=0,"",{g("Riders_Seated", wk)}/{seats})'
        util_cells = [f"{A[h]}{r}" for h in ("Utilisation", "Util wk-1", "Util wk-2", "Util wk-3")]
        ws[f"{A['Weeks under UtilLow']}{r}"] = "=" + "+".join(
            f'IF(AND(ISNUMBER({c}),{c}<UtilLow),1,0)' for c in util_cells)
        ws[f"{A['Flag']}{r}"] = (
            f'=IF({A["Weeks under UtilLow"]}{r}>=UtilLowWeeks,{a}&": seats only "&TEXT({A["Utilisation"]}{r},"0%")'
            f'&" full, under "&TEXT(UtilLow,"0%")&" for "&{A["Weeks under UtilLow"]}{r}&" weeks running",'
            f'IF(AND({A["Seats"]}{r}>0,{A["Excess riders"]}{r}>={A["Seats"]}{r}*ExcessShare),{a}&": "'
            f'&TEXT({A["Excess riders"]}{r},"#,##0")&" riders beyond the seats — consider more buses",'
            f'IF(AND(ISNUMBER({A["FT over budget"]}{r}),{A["FT over budget"]}{r}>0),{a}&": FT spend ₦"'
            f'&TEXT({A["FT over budget"]}{r},"#,##0")&" over the weekly budget",'
            f'IF({A["Breakdowns"]}{r}>0,{a}&": "&{A["Breakdowns"]}{r}&" coaster breakdown(s) this week",""))))')
        ws[f"{A['Score util']}{r}"] = f'=IF(ISNUMBER({A["Utilisation"]}{r}),{A["Utilisation"]}{r}+ROW()/1000000000,"")'
        ws[f"{A['Score cpr']}{r}"] = f'=IF(ISNUMBER({A["Cost/rider"]}{r}),{A["Cost/rider"]}{r}+ROW()/1000000,"")'
        ws[f"{A['Score over']}{r}"] = (f'=IF(AND(ISNUMBER({A["FT over budget"]}{r}),{A["FT over budget"]}{r}>0),'
                                       f'{A["FT over budget"]}{r}+ROW()/1000000,"")')
        ws[f"{A['Score riders']}{r}"] = f'=IF({A["Riders"]}{r}>0,{A["Riders"]}{r}+ROW()/1000000,"")'
        ws[f"{A['Flag row']}{r}"] = f'=IF({A["Flag"]}{r}="","",ROW())'
        ws[f"{A['Score excess']}{r}"] = (f'=IF({A["Excess riders"]}{r}>0,{A["Excess riders"]}{r}+ROW()/1000000,"")')
    ar = lambda h: f"${A[h]}${AREA0}:${A[h]}${last}"

    # ---- Top 10 lists (columns BC..) ------------------------------------------------------------
    lists = [("demand", "LARGE", "Score excess", "Excess riders"), ("low", "SMALL", "Score util", "Utilisation"),
             ("cpr", "LARGE", "Score cpr", "Cost/rider"), ("over", "LARGE", "Score over", "FT over budget"),
             ("riders", "LARGE", "Score riders", "Riders")]
    TOP = {}
    col = c0 + len(heads) + 1
    for key, fn, score, value in lists:
        n_col, v_col = L_(col), L_(col + 1)
        TOP[key] = (n_col, v_col)
        ws[f"{n_col}2"], ws[f"{v_col}2"] = f"{key} name", f"{key} value"
        for k in range(1, 11):
            r = AREA0 + k - 1
            m = f"MATCH({fn}({ar(score)},{k}),{ar(score)},0)"
            ws[f"{n_col}{r}"] = f'=IFERROR(INDEX({ar("Area")},{m}),"")'
            ws[f"{v_col}{r}"] = f'=IFERROR(INDEX({ar(value)},{m}),"")'
        col += 2
    flag_col = L_(col)
    TOP["flags"] = (flag_col, None)
    ws[f"{flag_col}2"] = "flags"
    for k in range(1, 11):
        r = AREA0 + k - 1
        ws[f"{flag_col}{r}"] = f'=IFERROR(INDEX({ar("Flag")},SMALL({ar("Flag row")},{k})-{AREA0 - 1}),"")'
    col += 1
    # ---- operational notes list (breakdowns / extra passengers / "need more buses") ----------------
    note_cols = [L_(col + j) for j in range(6)]
    TOP["notes"] = note_cols
    for L, h in zip(note_cols, ["row", "Area", "Category", "Location / bus", "Note", "Details"]):
        ws[f"{L}2"] = h
    ws[f"{note_cols[0]}1"] = f"=COUNT({T.ref('Note_Helper')})"
    for k in range(1, NOTES + 1):
        r = AREA0 + k - 1
        rowc = f"{note_cols[0]}{r}"
        ws[rowc] = f'=IFERROR(SMALL({T.ref("Note_Helper")},{k}),"")'
        ws[f"{note_cols[1]}{r}"] = f'=IF({rowc}="","",INDEX({T.ref("Area")},{rowc}))'
        ws[f"{note_cols[2]}{r}"] = f'=IF({rowc}="","",INDEX({T.ref("Category")},{rowc}))'
        ws[f"{note_cols[3]}{r}"] = (f'=IF({rowc}="","",TRIM(INDEX({T.ref("Location")},{rowc})&" "&'
                                    f'IF(INDEX({T.ref("Vehicle_ID")},{rowc})="","","(bus "&INDEX({T.ref("Vehicle_ID")},{rowc})&")")))')
        ws[f"{note_cols[4]}{r}"] = (f'=IF({rowc}="","",IF(INDEX({T.ref("Status")},{rowc})="Breakdown","Breakdown",'
                                    f'IF(INDEX({T.ref("Remarks_Category")},{rowc})="","Needs more buses",'
                                    f'INDEX({T.ref("Remarks_Category")},{rowc}))))')
        ws[f"{note_cols[5]}{r}"] = f'=IF({rowc}="","",INDEX({T.ref("Remarks")},{rowc})&"")'
    return TOP


def dash_value(src: str) -> str:
    return f'=IF({src}="","–",{src})'


def build_dashboard(wb, n_zones: int, top: dict):
    colour = COLOURS["TRANSPORT"]
    ws = wb.create_sheet(DASH)
    ws.sheet_properties.tabColor = colour
    dashboard_canvas(ws, "TRANSPORT", "TRANSPORT DASHBOARD",
                     "FT Procured (hired LT) · church coasters · EV/TATA · WSF Procured (member-paid) · riders, seats, "
                     "spend and optimisation", n_cols=14, widths={"B": 15})
    select_list = search_helper(wb, "TRANSPORT", DASH, 1, n_zones + 200)
    week_name, _, _ = week_list(wb, "TRANSPORT", SPECS["TRANSPORT"].ref("Week_Ending"), 1)
    control_strip(wb, ws, "TRANSPORT", colour, "Transport", select_list, week_name)
    C = lambda cell: f"{CALC}!{cell}"
    ws["L4"] = (f'={C("$B$16")}&IF({C("$B$10")}=0,"","  ·  w/e "&TEXT({C("$B$10")},"dd mmm yyyy"))'
                f'&"  ·  "&{C("$B$11")}&"-wk window"')
    ws.merge_cells("L4:O4")
    ws["L4"].font = font(10, True, INK)
    ws["B5"] = (f'=IF({C("$B$9")}=0,"No transport reports loaded yet. Put the weekly transport files in the inbox '
                f'and run ingest.py.",IF(NOT({C("$B$5")}),"⚠  Pick "&IF({C("$B$3")}="Area","an Area","a Zone")&'
                f'" from SELECT (type part of the name in SEARCH first).",IF({C("$B$3")}="Zone",'
                f'"Zone scope shows FT Procured and WSF Procured only (coasters, EV/TATA, fuel and budget are by Area).",'
                f'IF({C("$" + SER["has"] + "$21")}=0,"No transport reports in this scope for the selected week.",'
                f'"Utilisation = riders ÷ seats offered, on buses whose capacity is known. Spend = hire cost + coaster '
                f'fuel paid (hubs only)."))))')
    ws.merge_cells("B5:O5")
    ws["B5"].font = font(9, False, tint(colour, 0.2), italic=True)
    ws["B5"].alignment = LEFT

    # ---- KPI cards --------------------------------------------------------------------------------
    cards = [("RIDERS CARRIED", "Riders carried", FMT_INT, "pct", True),
             ("BUSES THAT RAN", "Buses that ran", FMT_INT, "pct", True),
             ("SEAT UTILISATION", "Seat utilisation %", FMT_PCT, "pts", True),
             ("TOTAL SPEND", "Total spend", FMT_NGN_INT, "pct", False),
             ("COST PER RIDER", "Cost per rider", FMT_NGN_DEC, "pct", False),
             ("FT SPEND vs BUDGET", "FT spend vs FT budget", FMT_PCT, "pts", False),
             ("BREAKDOWNS", "Breakdowns", FMT_INT, "pct", False)]
    card_bg = tint(colour, 0.92)
    for (title, kpi, fmt, delta, good_up), c1 in zip(cards, "BDFHJLN"):
        c2 = chr(ord(c1) + 1)
        for r in range(7, 11):
            for cc in (c1, c2):
                ws[f"{cc}{r}"].fill = fill(card_bg)
                ws[f"{cc}{r}"].border = Border(left=side(WHITE, "thick") if cc == c1 else None,
                                               right=side(WHITE, "thick") if cc == c2 else None,
                                               top=side(colour, "thick") if r == 7 else None)
        r = KPI_ROW[kpi]
        put(ws, f"{c1}7", title, f=font(8, True, tint(colour, 0.15)), align=CENTER, merge_to=f"{c2}7")
        put(ws, f"{c1}8", dash_value(C(f"$B${r}")), f=font(18, True, colour), fmt=fmt, align=CENTER,
            merge_to=f"{c2}8")
        if delta == "pct":
            d = (f'=IF({C(f"$E${r}")}="","– vs prev wk",IF({C(f"$E${r}")}>=0,"▲ ","▼ ")&'
                 f'TEXT(ABS({C(f"$E${r}")}),"0.0%")&" vs prev wk")')
        else:
            d = (f'=IF({C(f"$D${r}")}="","– vs prev wk",IF({C(f"$D${r}")}>=0,"▲ ","▼ ")&'
                 f'TEXT(ABS({C(f"$D${r}")})*100,"0.0")&" pts vs prev wk")')
        avg = C(f"$F${r}")
        avg_text = {FMT_PCT: f'TEXT({avg},"0.0%")', FMT_NGN_INT: f'"₦"&TEXT({avg},"#,##0")',
                    FMT_NGN_DEC: f'"₦"&TEXT({avg},"#,##0.00")'}.get(fmt, f'TEXT({avg},"#,##0")')
        put(ws, f"{c1}9", d, f=font(9, True, MUTED), align=CENTER, merge_to=f"{c2}9")
        put(ws, f"{c1}10", f'="Avg last "&{C("$B$11")}&" wks: "&IF({avg}="","–",{avg_text})',
            f=font(8, False, MUTED), align=CENTER, merge_to=f"{c2}10")
        arrow_rules(ws, f"{c1}9", f"{c1}9", good_up=good_up)
        if kpi == "Seat utilisation %":
            rag_rules(ws, f"{c1}8", f"{c1}8")
    ws.row_dimensions[8].height = 34

    # ---- by category ------------------------------------------------------------------------------
    section_title(ws, 12, "By bus category — selected week", colour, last="O",
                  note="Cost per rider shows where money works hardest")
    heads = [("B", "C", "Category", "A", None), ("D", None, "Listed", "B", FMT_INT), ("E", None, "Ran", "C", FMT_INT),
             ("F", None, "% ran", "D", FMT_PCT), ("G", None, "Riders", "E", FMT_INT), ("H", None, "Seats", "F", FMT_INT),
             ("I", None, "Utilisation", "H", FMT_PCT), ("J", None, "Spend", "K", FMT_NGN_INT),
             ("K", None, "Cost / rider", "L", FMT_NGN_DEC), ("L", None, "Cost / bus", "M", FMT_NGN_INT),
             ("M", None, "Riders beyond seats", "R", FMT_INT), ("N", None, "Riders WoW", "P", FMT_PCT_SIGNED),
             ("O", None, "Paid by", "Q", None)]
    for c1, c2, h, _, _ in heads:
        put(ws, f"{c1}13", h, f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}13" if c2 else None)
    ws.row_dimensions[13].height = 26
    n_rows = len(CATEGORIES) + 1 + len(LOCATIONS)
    location_names = {"Hub": "Hubs (TATA/EV/coasters)", "Loading Bay": "Loading Bays (FT/WSF, zone coasters)"}
    for i in range(n_rows):
        r, cr = 14 + i, CAT0 + i
        total = i == len(CATEGORIES)
        loc = i > len(CATEGORIES)
        band = tint(colour, 0.93) if total else (tint(colour, 0.97) if loc else None)
        for c1, c2, _, src, fmt in heads:
            if loc and src == "A":
                value = location_names[LOCATIONS[i - len(CATEGORIES) - 1]]
            else:
                value = dash_value(C(f"${src}${cr}")) if fmt else f"={C(f'${src}${cr}')}"
            cell = put(ws, f"{c1}{r}", value, f=font(9, total, INK, italic=loc), bg=band, fmt=fmt,
                       align=RIGHT if fmt else LEFT, merge_to=f"{c2}{r}" if c2 else None)
            cell.border = Border(bottom=side(), top=side(colour) if i == len(CATEGORIES) + 1 else None)
    last_cat = 14 + n_rows - 1
    rag_rules(ws, f"I14:I{last_cat}", "I14")
    signed_rules(ws, f"N14:N{last_cat}", "N14")
    ws.conditional_formatting.add("K14:K17", DataBarRule(start_type="num", start_value=0, end_type="max",
                                                         color=tint(colour, 0.35)))

    # ---- time analysis --------------------------------------------------------------------------------
    ta = last_cat + 2
    section_title(ws, ta, "Time analysis — selected week against previous week, the N-week average and "
                          "period-to-date", colour, last="O")
    put(ws, f"B{ta + 1}", "KPI", f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to=f"C{ta + 1}")
    for j, h in enumerate(MEASURES):
        cell = ws.cell(ta + 1, 4 + j)
        cell.value = (f'="Avg last "&{C("$B$11")}&" wks"' if h == "Avg last N wks" else
                      f'="vs "&{C("$B$11")}&"-wk avg"' if h == "vs N-wk avg" else h)
        cell.font, cell.fill, cell.alignment = font(9, True, WHITE), fill(colour), CENTER
    ws.row_dimensions[ta + 1].height = 28
    for i, (label, kind, src, fmt) in enumerate(KPIS):
        r, cr = ta + 2 + i, KPI0 + i
        band = tint(colour, 0.95) if i % 2 == 0 else None
        put(ws, f"B{r}", label, f=font(9, i in (0, 5, 7, 11)), align=LEFT, bg=band, merge_to=f"C{r}")
        for j, L in enumerate("BCDEFGHIJKLM"):
            cell = ws.cell(r, 4 + j)
            cell.value = dash_value(C(f"${L}${cr}"))
            if L in "EG":
                cell.number_format = FMT_PCT_SIGNED
            elif L == "D" and fmt == FMT_PCT:
                cell.number_format = '+0.0" pts";-0.0" pts";0.0" pts"'
                cell.value = f'=IF({C(f"$D${cr}")}="","–",{C(f"$D${cr}")}*100)'
            elif L == "D":
                cell.number_format = {FMT_INT: "+#,##0;-#,##0;0", FMT_NGN_INT: '+"₦"#,##0;-"₦"#,##0;"₦"0',
                                      FMT_NGN_DEC: '+"₦"#,##0.00;-"₦"#,##0.00;"₦"0.00'}[fmt]
            elif L in "KLM" and fmt == FMT_INT:
                cell.number_format = "#,##0"
            else:
                cell.number_format = fmt
            cell.font = font(9, L == "B")
            cell.alignment = RIGHT
            if band:
                cell.fill = fill(band)
        for colL in "FGI":
            signed_rules(ws, f"{colL}{r}", f"{colL}{r}")
        if label == "Seat utilisation %":
            rag_rules(ws, f"D{r}:E{r}", f"D{r}")
    last_ta = ta + 2 + len(KPIS)
    put(ws, f"B{last_ta}", "Averages use only weeks that have data. MTD/QTD/YTD run to the selected week; ratio rows "
                           "show the ratio of the period's totals.", f=font(8, False, MUTED, italic=True), align=LEFT,
        merge_to=f"O{last_ta}")

    # ---- charts ------------------------------------------------------------------------------------------
    ch_row = last_ta + 2
    section_title(ws, ch_row, "Trends — last 13 weeks", colour, last="O")
    calc = wb[CALC]
    shades = [colour, tint(colour, 0.45), "1F5FA8", "7FA7D9"]
    for n, (title, first_col, fmt) in enumerate((("Riders by category", 2, "#,##0"),
                                                  ("Spend by category (₦)", 7, "#,##0"))):
        bar = BarChart()
        bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 60
        _title(bar, title)
        for j in range(4):
            bar.add_data(Reference(calc, min_col=first_col + j, min_row=125, max_row=138), titles_from_data=True)
        bar.set_categories(Reference(calc, min_col=1, min_row=126, max_row=138))
        for s_, shade in zip(bar.series, shades):
            s_.graphicalProperties.solidFill = shade
            s_.graphicalProperties.line.solidFill = shade
        bar.legend.position = "b"
        bar.y_axis.numFmt = fmt
        _style_chart(bar)
        bar.width, bar.height = 13.0, 7.4
        ws.add_chart(bar, "B" + str(ch_row + 1) if n == 0 else "G" + str(ch_row + 1))
    line = BarChart()
    line.type, line.gapWidth = "col", 60
    _title(line, "Seat utilisation")
    line.add_data(Reference(calc, min_col=11, min_row=125, max_row=138), titles_from_data=True)
    line.set_categories(Reference(calc, min_col=1, min_row=126, max_row=138))
    line.series[0].graphicalProperties.solidFill = tint(colour, 0.2)
    line.series[0].graphicalProperties.line.solidFill = tint(colour, 0.2)
    line.y_axis.numFmt = "0%"
    line.y_axis.scaling.min = 0
    line.legend = None
    _style_chart(line)
    line.width, line.height = 10.4, 7.4
    ws.add_chart(line, f"L{ch_row + 1}")

    # ---- optimisation panel ------------------------------------------------------------------------------
    op = ch_row + 17
    section_title(ws, op, "Optimisation — selected week, by Area", colour, last="O",
                  note="Thresholds (UtilLow, UtilLowWeeks, ExcessShare) are on SETTINGS")
    zone_scope = f'{C("$B$3")}="Zone"'
    tables = [("B", "Unmet demand (riders beyond seats)", "demand", FMT_INT),
              ("D", "Lowest utilisation", "low", FMT_PCT),
              ("F", "Highest cost per rider", "cpr", FMT_NGN_DEC),
              ("H", "FT spend over budget", "over", FMT_NGN_INT),
              ("J", "Most riders", "riders", FMT_INT)]
    for c1, title, key, fmt in tables:
        c2 = chr(ord(c1) + 1)
        n_col, v_col = top[key]
        put(ws, f"{c1}{op + 1}", title, f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}{op + 1}")
        for k in range(10):
            r, cr = op + 2 + k, AREA0 + k
            nm = C(f"${n_col}${cr}")
            ws[f"{c1}{r}"] = (f'=IF({zone_scope},IF({k}=0,"n/a at Zone scope",""),IF({nm}="",IF({k}=0,"–",""),'
                              f'{k + 1}&". "&{nm}))')
            ws[f"{c2}{r}"] = f'=IF(OR({zone_scope},{nm}=""),"",{C(f"${v_col}${cr}")})'
            ws[f"{c1}{r}"].font, ws[f"{c2}{r}"].font = font(9), font(9, True)
            ws[f"{c2}{r}"].number_format = fmt
            ws[f"{c2}{r}"].alignment = RIGHT
            for cc in (c1, c2):
                ws[f"{cc}{r}"].border = Border(bottom=side())
        if fmt == FMT_PCT:
            rag_rules(ws, f"{c2}{op + 2}:{c2}{op + 11}", f"{c2}{op + 2}")
        else:
            ws.conditional_formatting.add(f"{c2}{op + 2}:{c2}{op + 11}", DataBarRule(
                start_type="num", start_value=0, end_type="max", color=tint(colour, 0.35)))
    put(ws, f"L{op + 1}", "Plain-language flags", f=font(9, True, WHITE), bg=colour, align=CENTER,
        merge_to=f"O{op + 1}")
    flag_col = top["flags"][0]
    for k in range(10):
        r = op + 2 + k
        put(ws, f"L{r}", f'=IF({zone_scope},IF({k}=0,"n/a at Zone scope",""),IF({C(f"${flag_col}${AREA0 + k}")}="",'
                         f'IF({k}=0,"No flags this week",""),"• "&{C(f"${flag_col}${AREA0 + k}")}))',
            f=font(9, False, RED_FG), align=Alignment(wrap_text=True, vertical="center"), merge_to=f"O{r}")
    ws.row_dimensions[op + 1].height = 20

    # ---- operational notes -------------------------------------------------------------------------------
    nt = op + 14
    section_title(ws, nt, "Operational notes — breakdowns, extra passengers and requests for more buses", colour,
                  last="O")
    note_cols = top["notes"]
    ws[f"B{nt + 1}"] = (f'=IF({C(f"${note_cols[0]}$1")}=0,"None reported for the selected week in this scope.",'
                        f'TEXT({C(f"${note_cols[0]}$1")},"#,##0")&" note(s)"&IF({C(f"${note_cols[0]}$1")}>{NOTES},'
                        f'" — showing the first {NOTES}.","."))')
    ws.merge_cells(f"B{nt + 1}:O{nt + 1}")
    ws[f"B{nt + 1}"].font = font(9, True, RED_FG)
    layout = [("B", "C", "Area", 1), ("D", "E", "Category", 2), ("F", "I", "Location / bus", 3), ("J", "K", "Note", 4),
              ("L", "O", "Details", 5)]
    for c1, c2, h, _ in layout:
        put(ws, f"{c1}{nt + 2}", h, f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to=f"{c2}{nt + 2}")
    for k in range(NOTES):
        r = nt + 3 + k
        for c1, c2, _, j in layout:
            put(ws, f"{c1}{r}", f"={C(f'${note_cols[j]}${AREA0 + k}')}", f=font(9), merge_to=f"{c2}{r}",
                bg=tint(colour, 0.95) if k % 2 == 0 else None)
    ws.print_area = f"A1:P{nt + 3 + NOTES}"
    ws.print_title_rows = "1:5"
    return ws
