"""CALC_COMMUNITY and the COMMUNITY CHURCH DASHBOARD.

The 35 Community Churches are reported on their own and never added into zonal
totals.  Sunday service comes first; their WSF, Midweek and CHOP figures (taken
out of the zonal exports) follow.  Scope Area = the churches hosted by that
Area; scope Zone = one church (its zincode).
"""
from __future__ import annotations

from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Border
from openpyxl.utils import get_column_letter as L_

from config import COLOURS
from .common import (CENTER, FMT_DEC1, FMT_INT, FMT_PCT, FMT_PCT_SIGNED, INK, LEFT, MUTED, RED_FG, RIGHT, WHITE,
                     arrow_rules, dashboard_canvas, fill, font, put, q, rag_rules, section_title, side,
                     signed_rules, tint)
from .controls import SCOPE, SELECT, WEEK, WINDOW, control_strip, search_helper, week_list
from .data import SPECS
from .wsf import _style_chart, _title

DASH, CALC = "COMMUNITY CHURCH DASHBOARD", "CALC_COMMUNITY"
S0, S1 = 21, 80
KPI0 = 90
CH0 = 3                 # church table rows (columns W..)
MAX_CHURCHES = 60
SERVICES = ["WSF", "Midweek", "CHOP"]

SER = {}
for _i, _k in enumerate(["k", "week", "month", "qtr", "year", "rows", "male", "female", "children", "adult",
                         "grand", "first", "converts", "testimonies", "pastors", "universe", "reported", "notrep",
                         "has", "win", "mtd", "qtd", "ytd"]):
    SER[_k] = L_(_i + 1)

KPIS = [
    ("Sunday attendance", "add", "grand", FMT_INT),
    ("Adults (Male + Female)", "add", "adult", FMT_INT),
    ("Male", "add", "male", FMT_INT),
    ("Female", "add", "female", FMT_INT),
    ("Children", "add", "children", FMT_INT),
    ("First timers", "add", "first", FMT_INT),
    ("New converts", "add", "converts", FMT_INT),
    ("Testimonies", "add", "testimonies", FMT_INT),
    ("Churches", "add", "universe", FMT_INT),
    ("Churches that reported", "add", "reported", FMT_INT),
    ("Churches with no report", "add", "notrep", FMT_INT),
    ("% churches reporting", "ratio", ("reported", "universe"), FMT_PCT),
    ("Avg attendance per reporting church", "ratio", ("grand", "reported"), FMT_DEC1),
    ("Pastors present (of 2 per church)", "ratio", ("pastors", "reported"), '0.00" of 2"'),
]
KPI_ROW = {label: KPI0 + i for i, (label, *_) in enumerate(KPIS)}
MEASURES = ["Selected week", "Previous week", "Change", "% change", "Avg last N wks", "vs N-wk avg",
            "MTD total", "QTD total", "YTD total", "Wkly avg (month)", "Wkly avg (qtr)", "Wkly avg (year)"]


def build_calc(wb, n_churches: int):
    ws = wb.create_sheet(CALC)
    ws.sheet_state = "hidden"
    d = SPECS["COMMUNITY"]
    W, A, CH, SV = d.ref("Week_Ending"), d.ref("Area"), d.ref("Church"), d.ref("Service_Type")
    D = q(DASH)
    put(ws, "A1", "CALC_COMMUNITY — helper calculations for the COMMUNITY CHURCH DASHBOARD (hidden)", f=font(11, True))
    inputs = [
        ("Scope", f"={D}!{SCOPE}"),
        ("Selection", f'=IF({D}!{SELECT}="","",{D}!{SELECT})'),
        ("Selection valid", '=IF(B3="Global",TRUE,IF(B3="Area",COUNTIF(MASTER_AREAS!$B:$B,B4)>0,'
                            'COUNTIF(MASTER_COMMUNITY!$A:$A,B4)>0))'),
        ("Area criterion", '=IF(B3="Area",B4,"*")'),
        ("Church criterion", '=IF(B3="Zone",B4,"*")'),
        ("Church name", '=IF(B3="Zone",IFERROR(INDEX(MASTER_COMMUNITY!$B:$B,MATCH(B4,MASTER_COMMUNITY!$A:$A,0)),'
                        '""),"")'),
        ("Latest week with data", f'=MAX({W})'),
        ("Selected week", f"=IF(ISNUMBER({D}!{WEEK}),{D}!{WEEK},B9)"),
        ("Window N (weeks)", f"=IF(ISNUMBER({D}!{WINDOW}),{D}!{WINDOW},DefaultWindow)"),
        ("Selected month", "=IF(B10=0,0,MONTH(B10))"),
        ("Selected quarter", "=IF(B10=0,0,ROUNDUP(MONTH(B10)/3,0))"),
        ("Selected year", "=IF(B10=0,0,YEAR(B10))"),
        ("Any data", "=B9>0"),
        ("Scope label", '=IF(B3="Global","All 35 Community Churches",IF(B3="Area","Churches in "&'
                        'IF(B4="","(no Area selected)",B4),IF(B4="","(no church selected)",B4&" "&B8)))'),
        ("Churches in scope", '=IF(B3="Zone",IF(B5,1,0),COUNTIFS(MASTER_COMMUNITY!$C:$C,B6,'
                              'MASTER_COMMUNITY!$D:$D,"Y"))'),
        ("Selection is not a Community Church", '=AND(B3="Zone",B4<>"",NOT(B5))'),
    ]
    for i, (label, formula) in enumerate(inputs, start=3):
        ws[f"A{i}"], ws[f"B{i}"] = label, formula
    ws["B9"].number_format = ws["B10"].number_format = "dd-mmm-yyyy"
    for key, L in SER.items():
        ws[f"{L}{S0 - 1}"] = key
    crit = lambda wk, svc="Sunday": f'{W},{wk},{A},$B$6,{CH},$B$7,{SV},"{svc}"'
    for k in range(S1 - S0 + 1):
        r = S0 + k
        S = {key: f"{L}{r}" for key, L in SER.items()}
        wk = S["week"]
        sm = lambda col: f"=IF({S['rows']}=0,0,SUMIFS({d.ref(col)},{crit(wk)}))"
        f = {"k": k, "week": f'=IF($B$10=0,"",$B$10-7*{S["k"]})',
             "month": f'=IF({wk}="","",MONTH({wk}))', "qtr": f'=IF({wk}="","",ROUNDUP(MONTH({wk})/3,0))',
             "year": f'=IF({wk}="","",YEAR({wk}))', "rows": f'=IF({wk}="",0,COUNTIFS({crit(wk)}))',
             "male": sm("Male"), "female": sm("Female"), "children": sm("Children"),
             "adult": f"={S['male']}+{S['female']}", "grand": f"={S['adult']}+{S['children']}",
             "first": sm("First_Timers"), "converts": sm("New_Converts"), "testimonies": sm("Testimonies"),
             "pastors": sm("Pastors_Present"), "universe": f'=IF({wk}="",0,$B$17)',
             "reported": f"={S['rows']}", "notrep": f"=MAX(0,{S['universe']}-{S['reported']})",
             "has": f"=IF({S['rows']}>0,1,0)",
             "win": f"=IF(AND({S['k']}>=1,{S['k']}<=$B$11),{S['has']},0)",
             "mtd": f"=IF(AND({S['month']}=$B$12,{S['year']}=$B$14),{S['has']},0)",
             "qtd": f"=IF(AND({S['qtr']}=$B$13,{S['year']}=$B$14),{S['has']},0)",
             "ytd": f"=IF({S['year']}=$B$14,{S['has']},0)"}
        for key, v in f.items():
            ws[S[key]] = v
        ws[S["week"]].number_format = "dd-mmm-yyyy"
    for j, h in enumerate(["KPI"] + MEASURES):
        ws.cell(KPI0 - 1, j + 1, h)
    rng = lambda key: f"${SER[key]}${S0}:${SER[key]}${S1}"
    fl = {x: rng(x) for x in ("win", "mtd", "qtd", "ytd")}
    has0, has1 = f"${SER['has']}${S0}", f"${SER['has']}${S0 + 1}"
    for i, (label, kind, src, _fmt) in enumerate(KPIS):
        r = KPI0 + i
        ws[f"A{r}"] = label
        if kind == "add":
            X = SER[src]
            f = {"B": f'=IF({has0}=0,"",{X}{S0})', "C": f'=IF({has1}=0,"",{X}{S0 + 1})',
                 "F": f'=IF(SUM({fl["win"]})=0,"",SUMPRODUCT({rng(src)},{fl["win"]})/SUM({fl["win"]}))'}
            for col, key in (("H", "mtd"), ("I", "qtd"), ("J", "ytd")):
                f[col] = f'=IF(SUM({fl[key]})=0,"",SUMPRODUCT({rng(src)},{fl[key]}))'
            for col, base, key in (("K", "H", "mtd"), ("L", "I", "qtd"), ("M", "J", "ytd")):
                f[col] = f'=IF({base}{r}="","",{base}{r}/SUM({fl[key]}))'
        else:
            N_, D_ = SER[src[0]], SER[src[1]]
            ratio = lambda key: (f'=IF(SUMPRODUCT({rng(src[1])},{fl[key]})=0,"",'
                                 f'SUMPRODUCT({rng(src[0])},{fl[key]})/SUMPRODUCT({rng(src[1])},{fl[key]}))')
            f = {"B": f'=IF(OR({has0}=0,{D_}{S0}=0),"",{N_}{S0}/{D_}{S0})',
                 "C": f'=IF(OR({has1}=0,{D_}{S0 + 1}=0),"",{N_}{S0 + 1}/{D_}{S0 + 1})',
                 "F": ratio("win"), "H": ratio("mtd"), "I": ratio("qtd"), "J": ratio("ytd"),
                 "K": f"=H{r}", "L": f"=I{r}", "M": f"=J{r}"}
        f["D"] = f'=IF(OR(B{r}="",C{r}=""),"",B{r}-C{r})'
        f["E"] = f'=IF(OR(B{r}="",C{r}=""),"",IF(C{r}=0,"",B{r}/C{r}-1))'
        f["G"] = f'=IF(OR(B{r}="",F{r}=""),"",IF(F{r}=0,"",B{r}/F{r}-1))'
        for col, v in f.items():
            ws[f"{col}{r}"] = v
    # chart feeds: 13-week mix, 52-week total (oldest -> newest)
    ws["A106"], ws["B106"], ws["C106"], ws["D106"] = "Week", "Male", "Female", "Children"
    for j in range(13):
        r, src = 107 + j, S0 + 12 - j
        ws[f"A{r}"] = f'=IF({SER["week"]}{src}="","",TEXT({SER["week"]}{src},"dd mmm"))'
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = (f"={SER['male']}{src}", f"={SER['female']}{src}",
                                                 f"={SER['children']}{src}")
    ws["F106"], ws["G106"] = "Week", "Sunday attendance"
    for j in range(52):
        r, src = 107 + j, S0 + 51 - j
        ws[f"F{r}"] = f'=IF({SER["week"]}{src}="","",TEXT({SER["week"]}{src},"dd mmm yy"))'
        ws[f"G{r}"] = f"={SER['grand']}{src}"
    # other services this week (from the zonal exports)
    ws["A170"] = "Service"
    for i, svc in enumerate(SERVICES):
        r = 171 + i
        c = crit("$B$10", svc)
        ws[f"A{r}"] = svc
        ws[f"B{r}"] = f"=IF($B$10=0,0,COUNTIFS({c}))"
        ws[f"C{r}"] = f"=IF($B$10=0,0,SUMIFS({d.ref('Grand_Total')},{c}))"
        ws[f"D{r}"] = f"=IF($B$10=0,0,SUMIFS({d.ref('First_Timers')},{c}))"
        ws[f"E{r}"] = f"=IF($B$10=0,0,SUMIFS({d.ref('New_Converts')},{c}))"
        # the WSF / Midweek / CHOP of the same week: latest of each on or before the selected week
        ws[f"F{r}"] = f'=IFERROR(_xlfn.MAXIFS({W},{SV},"{svc}",{W},"<="&$B$10),"")'
    # per-church block (columns W..)
    heads = ["Church", "Name", "Area", "In scope", "Sunday", "Prev Sunday", "Growth", "Avg N wks", "vs avg",
             "First timers", "Pastors", "Reported", "Row", "Missing row"]
    C_ = {h: L_(27 + j) for j, h in enumerate(heads)}      # clear of the weekly series A..W
    for h, colL in C_.items():
        ws[f"{colL}{CH0 - 1}"] = h
    for i in range(MAX_CHURCHES):
        r = CH0 + i
        ch = f"{C_['Church']}{r}"
        ws[ch] = f'=IFERROR(INDEX(MASTER_COMMUNITY!$A:$A,{i + 2})&"","")'
        ws[f"{C_['Name']}{r}"] = f'=IF({ch}="","",INDEX(MASTER_COMMUNITY!$B:$B,{i + 2})&"")'
        ws[f"{C_['Area']}{r}"] = f'=IF({ch}="","",INDEX(MASTER_COMMUNITY!$C:$C,{i + 2})&"")'
        ws[f"{C_['In scope']}{r}"] = (f'=AND({ch}<>"",OR($B$6="*",{C_["Area"]}{r}=$B$6),OR($B$7="*",{ch}=$B$7))')
        one = lambda wk: f'{W},{wk},{CH},{ch},{SV},"Sunday"'
        ws[f"{C_['Sunday']}{r}"] = f'=IF(OR({ch}="",$B$10=0),"",IF(COUNTIFS({one("$B$10")})=0,"",SUMIFS({d.ref("Grand_Total")},{one("$B$10")})))'
        ws[f"{C_['Prev Sunday']}{r}"] = f'=IF(OR({ch}="",$B$10=0),"",IF(COUNTIFS({one("$B$10-7")})=0,"",SUMIFS({d.ref("Grand_Total")},{one("$B$10-7")})))'
        ws[f"{C_['Growth']}{r}"] = (f'=IF(OR({C_["Sunday"]}{r}="",{C_["Prev Sunday"]}{r}="",{C_["Prev Sunday"]}{r}=0),"",'
                                    f'{C_["Sunday"]}{r}/{C_["Prev Sunday"]}{r}-1)')
        win = f'{W},">="&($B$10-7*$B$11),{W},"<"&$B$10,{CH},{ch},{SV},"Sunday"'
        ws[f"{C_['Avg N wks']}{r}"] = (f'=IF(OR({ch}="",$B$10=0),"",IF(COUNTIFS({win})=0,"",'
                                       f'SUMIFS({d.ref("Grand_Total")},{win})/COUNTIFS({win})))')
        ws[f"{C_['vs avg']}{r}"] = (f'=IF(OR({C_["Sunday"]}{r}="",{C_["Avg N wks"]}{r}="",{C_["Avg N wks"]}{r}=0),"",'
                                    f'{C_["Sunday"]}{r}/{C_["Avg N wks"]}{r}-1)')
        ws[f"{C_['First timers']}{r}"] = f'=IF({C_["Sunday"]}{r}="","",SUMIFS({d.ref("First_Timers")},{one("$B$10")}))'
        ws[f"{C_['Pastors']}{r}"] = f'=IF({C_["Sunday"]}{r}="","",SUMIFS({d.ref("Pastors_Present")},{one("$B$10")}))'
        ws[f"{C_['Reported']}{r}"] = f'=IF({ch}="","",IF({C_["Sunday"]}{r}="","No","Yes"))'
        ws[f"{C_['Row']}{r}"] = f'=IF({C_["In scope"]}{r},ROW(),"")'
        ws[f"{C_['Missing row']}{r}"] = f'=IF(AND({C_["In scope"]}{r},{C_["Reported"]}{r}="No"),ROW(),"")'
    last = CH0 + MAX_CHURCHES - 1
    ar = lambda h: f"${C_[h]}${CH0}:${C_[h]}${last}"
    # churches in scope, in master order, then the ones missing a Sunday report
    out = L_(27 + len(heads) + 1)
    miss = L_(27 + len(heads) + 2)
    for k in range(1, MAX_CHURCHES + 1):
        r = CH0 + k - 1
        ws[f"{out}{r}"] = f'=IFERROR(SMALL({ar("Row")},{k}),"")'
        ws[f"{miss}{r}"] = f'=IFERROR(SMALL({ar("Missing row")},{k}),"")'
    return {"cols": C_, "out": out, "miss": miss}


def build_dashboard(wb, n_zones: int, info: dict):
    colour = COLOURS["COMMUNITY"]
    ws = wb.create_sheet(DASH)
    ws.sheet_properties.tabColor = colour
    dashboard_canvas(ws, "COMMUNITY", "COMMUNITY CHURCH DASHBOARD",
                     "35 Community Churches · Sunday service first · reported separately from zonal totals",
                     n_cols=14, widths={"B": 15})
    select_list = search_helper(wb, "COMMUNITY", DASH, 4, n_zones + 200)
    week_name, _, _ = week_list(wb, "COMMUNITY", SPECS["COMMUNITY"].ref("Week_Ending"), 4)
    control_strip(wb, ws, "COMMUNITY", colour, "Community", select_list, week_name)
    ws["E3"].value = "SELECT AREA / CHURCH (zincode)  ▼"
    C = lambda cell: f"{CALC}!{cell}"
    ws["L4"] = (f'={C("$B$16")}&IF({C("$B$10")}=0,"","  ·  w/e "&TEXT({C("$B$10")},"dd mmm yyyy"))'
                f'&"  ·  "&{C("$B$11")}&"-wk window"')
    ws.merge_cells("L4:O4")
    ws["L4"].font = font(10, True, INK)
    ws["B5"] = (f'=IF({C("$B$9")}=0,"No Community Church reports loaded yet.",IF({C("$B$18")},'
                f'"That zone is not a Community Church. Pick a church zincode (e.g. LFC0412) or use Scope = Area.",'
                f'IF(NOT({C("$B$5")}),"⚠  Pick "&IF({C("$B$3")}="Area","an Area","a church zincode")&" from SELECT.",'
                f'"Community Church figures are never added into zone, Area or global zonal totals.")))')
    ws.merge_cells("B5:O5")
    ws["B5"].font = font(9, False, tint(colour, 0.2), italic=True)
    ws["B5"].alignment = LEFT

    cards = [("SUNDAY ATTENDANCE", "Sunday attendance", FMT_INT, "pct"),
             ("CHURCHES REPORTED", "Churches that reported", FMT_INT, "pct"),
             ("% CHURCHES REPORTING", "% churches reporting", FMT_PCT, "pts"),
             ("AVG PER CHURCH", "Avg attendance per reporting church", FMT_DEC1, "pct"),
             ("FIRST TIMERS", "First timers", FMT_INT, "pct"), ("NEW CONVERTS", "New converts", FMT_INT, "pct"),
             ("PASTORS PRESENT", "Pastors present (of 2 per church)", '0.00" of 2"', "pct")]
    card_bg = tint(colour, 0.92)
    for (title, kpi, fmt, delta), c1 in zip(cards, "BDFHJLN"):
        c2 = chr(ord(c1) + 1)
        for r in range(7, 11):
            for cc in (c1, c2):
                ws[f"{cc}{r}"].fill = fill(card_bg)
                ws[f"{cc}{r}"].border = Border(left=side(WHITE, "thick") if cc == c1 else None,
                                               right=side(WHITE, "thick") if cc == c2 else None,
                                               top=side(colour, "thick") if r == 7 else None)
        r = KPI_ROW[kpi]
        put(ws, f"{c1}7", title, f=font(8, True, tint(colour, 0.15)), align=CENTER, merge_to=f"{c2}7")
        put(ws, f"{c1}8", f'=IF({C(f"$B${r}")}="","–",{C(f"$B${r}")})', f=font(20, True, colour), fmt=fmt,
            align=CENTER, merge_to=f"{c2}8")
        if delta == "pct":
            dtext = (f'=IF({C(f"$E${r}")}="","– vs prev Sunday",IF({C(f"$E${r}")}>=0,"▲ ","▼ ")&'
                     f'TEXT(ABS({C(f"$E${r}")}),"0.0%")&" vs prev Sunday")')
        else:
            dtext = (f'=IF({C(f"$D${r}")}="","– vs prev Sunday",IF({C(f"$D${r}")}>=0,"▲ ","▼ ")&'
                     f'TEXT(ABS({C(f"$D${r}")})*100,"0.0")&" pts vs prev Sunday")')
        avg = C(f"$F${r}")
        avg_text = {FMT_PCT: f'TEXT({avg},"0.0%")', FMT_DEC1: f'TEXT({avg},"#,##0.0")',
                    '0.00" of 2"': f'TEXT({avg},"0.00")'}.get(fmt, f'TEXT({avg},"#,##0")')
        put(ws, f"{c1}9", dtext, f=font(9, True, MUTED), align=CENTER, merge_to=f"{c2}9")
        put(ws, f"{c1}10", f'="Avg last "&{C("$B$11")}&" Sundays: "&IF({avg}="","–",{avg_text})',
            f=font(8, False, MUTED), align=CENTER, merge_to=f"{c2}10")
        arrow_rules(ws, f"{c1}9", f"{c1}9")
        if fmt == FMT_PCT:
            rag_rules(ws, f"{c1}8", f"{c1}8")
    ws.row_dimensions[8].height = 34

    section_title(ws, 12, "Sunday service — time analysis", colour, last="O")
    put(ws, "B13", "KPI", f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to="C13")
    for j, h in enumerate(MEASURES):
        cell = ws.cell(13, 4 + j)
        cell.value = (f'="Avg last "&{C("$B$11")}&" wks"' if h == "Avg last N wks" else
                      f'="vs "&{C("$B$11")}&"-wk avg"' if h == "vs N-wk avg" else h)
        cell.font, cell.fill, cell.alignment = font(9, True, WHITE), fill(colour), CENTER
    ws.row_dimensions[13].height = 28
    for i, (label, kind, src, fmt) in enumerate(KPIS):
        r, cr = 14 + i, KPI0 + i
        band = tint(colour, 0.95) if i % 2 == 0 else None
        put(ws, f"B{r}", label, f=font(9, i in (0, 11)), align=LEFT, bg=band, merge_to=f"C{r}")
        for j, L in enumerate("BCDEFGHIJKLM"):
            cell = ws.cell(r, 4 + j)
            cell.value = f'=IF({C(f"${L}${cr}")}="","–",{C(f"${L}${cr}")})'
            if L in "EG":
                cell.number_format = FMT_PCT_SIGNED
            elif L == "D" and fmt == FMT_PCT:
                cell.number_format = '+0.0" pts";-0.0" pts";0.0" pts"'
                cell.value = f'=IF({C(f"$D${cr}")}="","–",{C(f"$D${cr}")}*100)'
            elif L == "D":
                cell.number_format = "+#,##0.0;-#,##0.0;0.0" if fmt != FMT_INT else "+#,##0;-#,##0;0"
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
        if fmt == FMT_PCT:
            rag_rules(ws, f"D{r}:E{r}", f"D{r}")
    end_ta = 14 + len(KPIS)

    ch = end_ta + 1
    section_title(ws, ch, "Sunday trends", colour, last="O")
    calc = wb[CALC]
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 60
    _title(bar, "Sunday attendance mix — last 13 weeks")
    for c in (2, 3, 4):
        bar.add_data(Reference(calc, min_col=c, min_row=106, max_row=119), titles_from_data=True)
    bar.set_categories(Reference(calc, min_col=1, min_row=107, max_row=119))
    for s_, shade in zip(bar.series, [colour, tint(colour, 0.45), "E0A526"]):
        s_.graphicalProperties.solidFill = shade
        s_.graphicalProperties.line.solidFill = shade
    bar.legend.position = "b"
    bar.y_axis.numFmt = "#,##0"
    _style_chart(bar)
    bar.width, bar.height = 15.5, 7.4
    ws.add_chart(bar, f"B{ch + 1}")
    tr = BarChart()
    tr.type, tr.gapWidth = "col", 40
    _title(tr, "Sunday attendance — last 52 weeks")
    tr.add_data(Reference(calc, min_col=7, min_row=106, max_row=158), titles_from_data=True)
    tr.set_categories(Reference(calc, min_col=6, min_row=107, max_row=158))
    tr.series[0].graphicalProperties.solidFill = colour
    tr.series[0].graphicalProperties.line.solidFill = colour
    tr.legend = None
    tr.y_axis.numFmt = "#,##0"
    _style_chart(tr)
    tr.width, tr.height = 20.0, 7.4
    ws.add_chart(tr, f"H{ch + 1}")

    ot = ch + 17
    section_title(ws, ot, "Other services at the Community Churches (from the zonal reports)", colour, last="O",
                  note="Latest report on or before the selected week")
    for c1, c2, h in (("B", "C", "Service"), ("D", "E", "Week ending"), ("F", "G", "Churches reported"),
                      ("H", "I", "Attendance"), ("J", "K", "First timers"), ("L", "M", "New converts")):
        put(ws, f"{c1}{ot + 1}", h, f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}{ot + 1}")
    d = SPECS["COMMUNITY"]
    for i, svc in enumerate(SERVICES):
        r = ot + 2 + i
        wk = C(f"$F${171 + i}")
        c = (f'{d.ref("Week_Ending")},{wk},{d.ref("Area")},{C("$B$6")},{d.ref("Church")},{C("$B$7")},'
             f'{d.ref("Service_Type")},"{svc}"')
        put(ws, f"B{r}", svc, f=font(9, True), merge_to=f"C{r}")
        put(ws, f"D{r}", f'=IF(OR({wk}="",{wk}=0),"–",{wk})', f=font(9), fmt="dd-mmm-yyyy", align=CENTER,
            merge_to=f"E{r}")
        for (c1, c2), col in zip((("F", "G"), ("H", "I"), ("J", "K"), ("L", "M")),
                                 (None, "Grand_Total", "First_Timers", "New_Converts")):
            f_ = (f'=IF(OR({wk}="",{wk}=0),"–",COUNTIFS({c}))' if col is None else
                  f'=IF(OR({wk}="",{wk}=0),"–",SUMIFS({d.ref(col)},{c}))')
            put(ws, f"{c1}{r}", f_, f=font(9), fmt="#,##0", align=RIGHT, merge_to=f"{c2}{r}")

    tb = ot + 6
    section_title(ws, tb, "Churches — selected Sunday", colour, last="O",
                  note="Pastors = pastors present (0–2)")
    cols = info["cols"]
    layout = [("B", "C", "Church", "Church", None), ("D", "G", "Location", "Name", None),
              ("H", "I", "Host Area", "Area", None), ("J", None, "This Sunday", "Sunday", "#,##0"),
              ("K", None, "Prev Sunday", "Prev Sunday", "#,##0"), ("L", None, "Change", "Growth", FMT_PCT_SIGNED),
              ("M", None, "vs N-wk avg", "vs avg", FMT_PCT_SIGNED), ("N", None, "First timers", "First timers",
                                                                     "#,##0"),
              ("O", None, "Pastors", "Pastors", "0")]
    for c1, c2, h, _, _ in layout:
        put(ws, f"{c1}{tb + 1}", h, f=font(9, True, WHITE), bg=colour, align=CENTER,
            merge_to=f"{c2}{tb + 1}" if c2 else None)
    out = info["out"]
    for k in range(MAX_CHURCHES):
        r = tb + 2 + k
        rowref = C(f"${out}${CH0 + k}")
        for c1, c2, _, src, fmt in layout:
            srccol = cols[src]
            v = f'=IF({rowref}="","",IF(INDEX({CALC}!${srccol}:${srccol},{rowref})="","–",' \
                f'INDEX({CALC}!${srccol}:${srccol},{rowref})))'
            put(ws, f"{c1}{r}", v, f=font(9), fmt=fmt, align=RIGHT if fmt else LEFT,
                merge_to=f"{c2}{r}" if c2 else None, bg=tint(colour, 0.95) if k % 2 == 0 else None)
    ws.conditional_formatting.add(f"J{tb + 2}:J{tb + 1 + MAX_CHURCHES}", DataBarRule(
        start_type="num", start_value=0, end_type="max", color=tint(colour, 0.35)))
    signed_rules(ws, f"L{tb + 2}:M{tb + 1 + MAX_CHURCHES}", f"L{tb + 2}")
    last_row = tb + 1 + MAX_CHURCHES

    nr = last_row + 2
    section_title(ws, nr, "Who didn't report — churches with no Sunday report for the selected week", colour,
                  last="O")
    miss = info["miss"]
    ws[f"B{nr + 1}"] = (f'=IF({C("$B$10")}=0,"",IF(COUNT({CALC}!${miss}${CH0}:${miss}${CH0 + MAX_CHURCHES - 1})=0,'
                        f'"Every church in scope reported.",COUNT({CALC}!${miss}${CH0}:${miss}'
                        f'${CH0 + MAX_CHURCHES - 1})&" church(es) without a Sunday report."))')
    ws.merge_cells(f"B{nr + 1}:O{nr + 1}")
    ws[f"B{nr + 1}"].font = font(9, True, RED_FG)
    for k in range(MAX_CHURCHES // 2):
        r = nr + 2 + k
        rowref = C(f"${miss}${CH0 + k}")
        put(ws, f"B{r}", f'=IF({rowref}="","",INDEX({CALC}!${cols["Church"]}:${cols["Church"]},{rowref}))',
            f=font(9), merge_to=f"C{r}")
        put(ws, f"D{r}", f'=IF({rowref}="","",INDEX({CALC}!${cols["Name"]}:${cols["Name"]},{rowref}))',
            f=font(9), merge_to=f"I{r}")
        put(ws, f"J{r}", f'=IF({rowref}="","",INDEX({CALC}!${cols["Area"]}:${cols["Area"]},{rowref}))',
            f=font(9), merge_to=f"L{r}")
    ws.print_area = f"A1:P{nr + 2 + MAX_CHURCHES // 2}"
    ws.print_title_rows = "1:5"
    return ws
