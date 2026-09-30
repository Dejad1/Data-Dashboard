"""Zonal service dashboards (MIDWEEK and CHOP) and their hidden CALC sheets.

Same layout as the WSF DASHBOARD: control strip, KPI cards, time analysis,
trends, Area rankings and a non-reporting list.  Reporting compliance comes
from the reporting system's own Area summaries (DATA_SERVICE_AREAS):

    Midweek: zones with a report / all zones
    CHOP:    zones with a report / CHOP zones (a zone without CHOP shows
             "No CHOP in this zone" instead of zeros)

Community Church zones are reported on their own dashboard and never counted
here.  CHOP is held at selected zones; the export gives the CHOP zones' attendance
as Area totals only, kept on "AREA nn CHOP ZONES" rows so Area and Global totals
include it.
"""
from __future__ import annotations

from dataclasses import dataclass

from openpyxl.chart import BarChart, Reference
from openpyxl.formatting.rule import DataBarRule
from openpyxl.styles import Alignment, Border
from openpyxl.utils import get_column_letter as L_

from config import COLOURS
from .common import (CENTER, FMT_DEC1, FMT_INT, FMT_PCT, FMT_PCT_SIGNED, INK, LEFT, MUTED, RED_FG, RIGHT, WHITE,
                     arrow_rules, dashboard_canvas, fill, font, put, q, rag_rules, section_title, side,
                     signed_rules, tint)
from .controls import SCOPE, SELECT, WEEK, WINDOW, control_strip, search_helper, week_list
from .data import SPECS
from .wsf import _style_chart, _title

S0, S1 = 21, 80
KPI0 = 90
AREA0 = 3
LIST_ROWS = 50
BLOCK0 = 27             # first column of the per-Area block (clear of the weekly series A..W)


@dataclass
class Zonal:
    key: str                 # MIDWEEK / CHOP (also the DATA_SERVICE_AREAS stream)
    dash: str
    calc: str
    colour_key: str
    title: str
    subtitle: str
    universe_col: str        # DATA_SERVICE_AREAS column holding the universe of zones
    universe_label: str      # "zones" / "CHOP zones"
    block: int               # position of the helper blocks in CALC_SEARCH / CALC_LISTS
    chop_only: bool = False


MIDWEEK = Zonal("MIDWEEK", "MIDWEEK DASHBOARD", "CALC_MIDWEEK", "MIDWEEK", "MIDWEEK DASHBOARD",
                "Zonal Midweek service · attendance and zone reporting compliance (Community Churches excluded)",
                "Total_Zones", "zones", 2)
CHOP = Zonal("CHOP", "CHOP DASHBOARD", "CALC_CHOP", "CHOP", "CHOP DASHBOARD",
             "Covenant Hour of Prayer · held at selected CHOP zones · reported by zone",
             "CHOP_Zones", "CHOP zones", 3, chop_only=True)

SER = {}
for _i, _k in enumerate(["k", "week", "month", "qtr", "year", "rows", "male", "female", "children", "adult",
                         "grand", "first", "converts", "universe", "reported", "notrep", "pct", "avg", "has",
                         "win", "mtd", "qtd", "ytd"]):
    SER[_k] = L_(_i + 1)


def kpis(z: Zonal):
    u = z.universe_label
    return [
        ("Grand total attendance", "add", "grand", FMT_INT),
        ("Adult total (Male + Female)", "add", "adult", FMT_INT),
        ("Male", "add", "male", FMT_INT),
        ("Female", "add", "female", FMT_INT),
        ("Children", "add", "children", FMT_INT),
        ("First timers", "add", "first", FMT_INT),
        ("New converts", "add", "converts", FMT_INT),
        (u[0].upper() + u[1:], "add", "universe", FMT_INT),
        (f"{u[0].upper() + u[1:]} that reported", "add", "reported", FMT_INT),
        (f"{u[0].upper() + u[1:]} with no report", "add", "notrep", FMT_INT),
        (f"% {u} reporting", "ratio", ("reported", "universe"), FMT_PCT),
        ("Reporting : non-reporting", "ratio", ("reported", "notrep"), '0.0" : 1"'),
        ("Avg attendance per reporting zone", "ratio", ("grand", "reported"), FMT_DEC1),
    ]


MEASURES = ["Selected week", "Previous week", "Change", "% change", "Avg last N wks", "vs N-wk avg",
            "MTD total", "QTD total", "YTD total", "Wkly avg (month)", "Wkly avg (qtr)", "Wkly avg (year)"]


def build_calc(wb, z: Zonal, n_areas: int, n_zones: int):
    ws = wb.create_sheet(z.calc)
    ws.sheet_state = "hidden"
    d, sa = SPECS[z.key], SPECS["SERVICE_AREAS"]
    W, A, Z = d.ref("Week_Ending"), d.ref("Area"), d.ref("Zone")
    D = q(z.dash)
    put(ws, "A1", f"{z.calc} — helper calculations for the {z.dash} (hidden; do not edit)", f=font(11, True))
    chop_test = ('IFERROR(INDEX(MASTER_ZONES!$D:$D,MATCH(B4,MASTER_ZONES!$B:$B,0))="Y",FALSE)'
                 if z.chop_only else "TRUE")
    inputs = [
        ("Scope", f"={D}!{SCOPE}"),
        ("Selection", f'=IF({D}!{SELECT}="","",{D}!{SELECT})'),
        ("Selection valid", '=IF(B3="Global",TRUE,IF(B3="Area",COUNTIF(MASTER_AREAS!$B:$B,B4)>0,'
                            'COUNTIF(MASTER_ZONES!$B:$B,B4)>0))'),
        ("Area criterion", '=IF(B3="Area",B4,IF(B3="Zone",B8,"*"))'),
        ("Zone criterion", '=IF(B3="Zone",B4,"*")'),
        ("Zone's Area", '=IF(B3="Zone",IFERROR(INDEX(MASTER_ZONES!$C:$C,MATCH(B4,MASTER_ZONES!$B:$B,0)),""),"")'),
        ("Latest week with data", f"=MAX({W})"),
        ("Selected week", f"=IF(ISNUMBER({D}!{WEEK}),{D}!{WEEK},B9)"),
        ("Window N (weeks)", f"=IF(ISNUMBER({D}!{WINDOW}),{D}!{WINDOW},DefaultWindow)"),
        ("Selected month", "=IF(B10=0,0,MONTH(B10))"),
        ("Selected quarter", "=IF(B10=0,0,ROUNDUP(MONTH(B10)/3,0))"),
        ("Selected year", "=IF(B10=0,0,YEAR(B10))"),
        ("Any data", "=B9>0"),
        ("Scope label", '=IF(B3="Global","All Areas",IF(B3="Area","Area: "&IF(B4="","(none selected)",B4),'
                        '"Zone: "&IF(B4="","(none selected)",B4&IF(B8="",""," ("&B8&")"))))'),
        ("Zone is in this service", f'=IF(B3<>"Zone",TRUE,{chop_test})'),
        ("Zone is a Community Church",
         '=IF(B3<>"Zone",FALSE,IFERROR(INDEX(MASTER_ZONES!$I:$I,MATCH(B4,MASTER_ZONES!$B:$B,0))="Y",FALSE))'),
    ]
    for i, (label, formula) in enumerate(inputs, start=3):
        ws[f"A{i}"], ws[f"B{i}"] = label, formula
    ws["B9"].number_format = ws["B10"].number_format = "dd-mmm-yyyy"

    for key, L in SER.items():
        ws[f"{L}{S0 - 1}"] = key
    crit = lambda wk: f"{W},{wk},{A},$B$6,{Z},$B$7"
    cm = SPECS["COMMUNITY"]
    service_type = "CHOP" if z.key == "CHOP" else "Midweek"
    community_universe = ('COUNTIFS(MASTER_ZONES!$C:$C,$B$6,MASTER_ZONES!$I:$I,"Y"'
                          + (',MASTER_ZONES!$D:$D,"Y")' if z.chop_only else ")"))
    sacrit = lambda wk: f'{sa.ref("Week_Ending")},{wk},{sa.ref("Stream")},"{z.key}",{sa.ref("Area")},$B$6'
    for k in range(S1 - S0 + 1):
        r = S0 + k
        S = {key: f"{L}{r}" for key, L in SER.items()}
        wk = S["week"]
        sm = lambda col: f"=IF({S['rows']}=0,0,SUMIFS({d.ref(col)},{crit(wk)}))"
        f = {
            "k": k, "week": f'=IF($B$10=0,"",$B$10-7*{S["k"]})',
            "month": f'=IF({wk}="","",MONTH({wk}))', "qtr": f'=IF({wk}="","",ROUNDUP(MONTH({wk})/3,0))',
            "year": f'=IF({wk}="","",YEAR({wk}))',
            "rows": f'=IF({wk}="",0,COUNTIFS({crit(wk)}))',
            "male": sm("Male"), "female": sm("Female"), "children": sm("Children"),
            "adult": f"={S['male']}+{S['female']}", "grand": f"={S['adult']}+{S['children']}",
            "first": sm("First_Timers"), "converts": sm("New_Converts"),
            # compliance: the reporting system's counts by Area; at Zone scope, 1 zone
            # the system's counts include the Community Church zones: take them out
            "universe": (f'=IF({wk}="",0,IF($B$3="Zone",IF(AND($B$17,NOT($B$18)),1,0),'
                         f'MAX(0,SUMIFS({sa.ref(z.universe_col)},{sacrit(wk)})-{community_universe})))'),
            "reported": (f'=IF({wk}="",0,IF($B$3="Zone",IF(AND($B$17,NOT($B$18),{S["rows"]}>0),1,0),'
                         f'MAX(0,SUMIFS({sa.ref("Zones_With_Report")},{sacrit(wk)})-COUNTIFS({cm.ref("Week_Ending")},'
                         f'{wk},{cm.ref("Service_Type")},"{service_type}",{cm.ref("Area")},$B$6))))'),
            "notrep": f"=MAX(0,{S['universe']}-{S['reported']})",
            "pct": f"=IF({S['universe']}=0,0,{S['reported']}/{S['universe']})",
            "avg": f"=IF({S['reported']}=0,0,{S['grand']}/{S['reported']})",
            "has": f"=IF(OR({S['rows']}>0,{S['universe']}>0),1,0)",
            "win": f"=IF(AND({S['k']}>=1,{S['k']}<=$B$11),{S['has']},0)",
            "mtd": f"=IF(AND({S['month']}=$B$12,{S['year']}=$B$14),{S['has']},0)",
            "qtd": f"=IF(AND({S['qtr']}=$B$13,{S['year']}=$B$14),{S['has']},0)",
            "ytd": f"=IF({S['year']}=$B$14,{S['has']},0)",
        }
        for key, v in f.items():
            ws[S[key]] = v
        ws[S["week"]].number_format = "dd-mmm-yyyy"

    K = kpis(z)
    for j, h in enumerate(["KPI"] + MEASURES):
        ws.cell(KPI0 - 1, j + 1, h)
    rng = lambda key: f"${SER[key]}${S0}:${SER[key]}${S1}"
    fl = {x: rng(x) for x in ("win", "mtd", "qtd", "ytd")}
    has0, has1 = f"${SER['has']}${S0}", f"${SER['has']}${S0 + 1}"
    for i, (label, kind, src, _fmt) in enumerate(K):
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

    # chart feed: 13 weeks, oldest -> newest
    ws["A104"], ws["B104"], ws["C104"], ws["D104"], ws["E104"] = "Week", "Male", "Female", "Children", "% reporting"
    for j in range(13):
        r, src = 105 + j, S0 + 12 - j
        ws[f"A{r}"] = f'=IF({SER["week"]}{src}="","",TEXT({SER["week"]}{src},"dd mmm"))'
        ws[f"B{r}"], ws[f"C{r}"], ws[f"D{r}"] = (f"={SER['male']}{src}", f"={SER['female']}{src}",
                                                 f"={SER['children']}{src}")
        ws[f"E{r}"] = f"={SER['pct']}{src}"

    # per-Area block (columns W..) for rankings and the Areas-with-gaps list
    heads = ["Area", "Rows", "Grand", "Grand prev", "Universe", "Reported", "% reporting", "Growth", "Score att",
             "Score pct", "Score growth", "Missing", "Missing row"]
    A_ = {h: L_(BLOCK0 + j) for j, h in enumerate(heads)}
    for h, colL in A_.items():
        ws[f"{colL}{AREA0 - 1}"] = h
    last = AREA0 + max(n_areas, 1) - 1
    for i in range(max(n_areas, 1)):
        r = AREA0 + i
        a = f"{A_['Area']}{r}"
        c = f"{W},$B$10,{A},{a}"
        sa_c = f'{sa.ref("Week_Ending")},$B$10,{sa.ref("Stream")},"{z.key}",{sa.ref("Area")},{a}'
        ws[a] = f'=IFERROR(INDEX(MASTER_AREAS!$B:$B,{i + 2})&"","")'
        ws[f"{A_['Rows']}{r}"] = f'=IF(OR({a}="",$B$10=0),0,COUNTIFS({c}))'
        ws[f"{A_['Grand']}{r}"] = f"=IF({A_['Rows']}{r}=0,0,SUMIFS({d.ref('Grand_Total')},{c}))"
        ws[f"{A_['Grand prev']}{r}"] = f'=IF(OR({a}="",$B$10=0),0,SUMIFS({d.ref("Grand_Total")},{W},$B$10-7,{A},{a}))'
        comm_u = (f'COUNTIFS(MASTER_ZONES!$C:$C,{a},MASTER_ZONES!$I:$I,"Y"'
                  + (',MASTER_ZONES!$D:$D,"Y")' if z.chop_only else ")"))
        comm_r = (f'COUNTIFS({cm.ref("Week_Ending")},$B$10,{cm.ref("Service_Type")},"{service_type}",'
                  f'{cm.ref("Area")},{a})')
        ws[f"{A_['Universe']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,MAX(0,SUMIFS({sa.ref(z.universe_col)},{sa_c})'
                                      f'-{comm_u}))')
        ws[f"{A_['Reported']}{r}"] = (f'=IF(OR({a}="",$B$10=0),0,MAX(0,SUMIFS({sa.ref("Zones_With_Report")},'
                                      f'{sa_c})-{comm_r}))')
        ws[f"{A_['% reporting']}{r}"] = (f'=IF({A_["Universe"]}{r}=0,"",'
                                         f'MIN(1,{A_["Reported"]}{r}/{A_["Universe"]}{r}))')
        ws[f"{A_['Growth']}{r}"] = f'=IF(OR({A_["Rows"]}{r}=0,{A_["Grand prev"]}{r}=0),"",{A_["Grand"]}{r}/{A_["Grand prev"]}{r}-1)'
        ws[f"{A_['Score att']}{r}"] = f'=IF({A_["Rows"]}{r}=0,"",{A_["Grand"]}{r}+ROW()/1000000)'
        ws[f"{A_['Score pct']}{r}"] = f'=IF({A_["% reporting"]}{r}="","",{A_["% reporting"]}{r}+ROW()/1000000000)'
        ws[f"{A_['Score growth']}{r}"] = f'=IF({A_["Growth"]}{r}="","",{A_["Growth"]}{r}+ROW()/1000000000)'
        ws[f"{A_['Missing']}{r}"] = f"=MAX(0,{A_['Universe']}{r}-{A_['Reported']}{r})"
        ws[f"{A_['Missing row']}{r}"] = (f'=IF(AND({A_["Missing"]}{r}>0,OR($B$6="*",{a}=$B$6)),ROW(),"")')
    ar = lambda h: f"${A_[h]}${AREA0}:${A_[h]}${last}"
    lists = [("LARGE", "Score att", "Grand"), ("SMALL", "Score att", "Grand"),
             ("LARGE", "Score pct", "% reporting"), ("SMALL", "Score pct", "% reporting"),
             ("LARGE", "Score growth", "Growth"), ("SMALL", "Score growth", "Growth")]
    top = []
    col = BLOCK0 + len(heads) + 1
    for fn, score, value in lists:
        n_col, v_col = L_(col), L_(col + 1)
        top.append((n_col, v_col))
        for k in range(1, 11):
            r = AREA0 + k - 1
            m = f"MATCH({fn}({ar(score)},{k}),{ar(score)},0)"
            ws[f"{n_col}{r}"] = f'=IFERROR(INDEX({ar("Area")},{m}),"")'
            ws[f"{v_col}{r}"] = f'=IFERROR(INDEX({ar(value)},{m}),"")'
        col += 2
    # rank of the selected Area
    rank_col = L_(col)
    for r, score in ((AREA0, "Score att"), (AREA0 + 1, "Score pct"), (AREA0 + 2, "Score growth")):
        ws[f"{rank_col}{r}"] = (f'=IF($B$3<>"Area","",IFERROR(COUNTIF({ar(score)},">"&'
                                f'INDEX({ar(score)},MATCH($B$4,{ar("Area")},0)))+1,""))')
    ws[f"{rank_col}{AREA0 + 3}"] = f"=COUNT({ar('Score att')})"
    col += 1
    # Areas with zones missing a report
    gap = [L_(col + j) for j in range(5)]
    for k in range(1, LIST_ROWS + 1):
        r = AREA0 + k - 1
        rowc = f"{gap[0]}{r}"
        ws[rowc] = f'=IFERROR(SMALL({ar("Missing row")},{k}),"")'
        for L, h in zip(gap[1:], ("Area", "Universe", "Reported", "Missing")):
            ws[f"{L}{r}"] = f'=IF({rowc}="","",INDEX(${A_[h]}:${A_[h]},{rowc}))'
    ws[f"{gap[0]}1"] = f"=COUNT({ar('Missing row')})"
    ws[f"{gap[1]}1"] = f"=SUM({ar('Missing')})"
    return {"top": top, "rank": rank_col, "gap": gap}


def dash_value(src: str) -> str:
    return f'=IF({src}="","–",{src})'


def build_dashboard(wb, z: Zonal, n_zones: int, calc_info: dict):
    colour = COLOURS[z.colour_key]
    ws = wb.create_sheet(z.dash)
    ws.sheet_properties.tabColor = colour
    dashboard_canvas(ws, z.colour_key, z.title, z.subtitle, n_cols=14, widths={"B": 15})
    select_list = search_helper(wb, z.key, z.dash, z.block, n_zones + 200)
    week_name, _, _ = week_list(wb, z.key, SPECS[z.key].ref("Week_Ending"), z.block)
    control_strip(wb, ws, z.key, colour, z.key.title() if z.key != "CHOP" else "CHOP", select_list, week_name)
    C = lambda cell: f"{z.calc}!{cell}"
    ws["L4"] = (f'={C("$B$16")}&IF({C("$B$10")}=0,"","  ·  w/e "&TEXT({C("$B$10")},"dd mmm yyyy"))'
                f'&"  ·  "&{C("$B$11")}&"-wk window"')
    ws.merge_cells("L4:O4")
    ws["L4"].font = font(10, True, INK)
    no_service = "No CHOP in this zone." if z.chop_only else "This zone is not in this service."
    zone_note = (f'IF({C("$B$3")}="Zone","The CHOP export gives CHOP-zone attendance as Area totals only, so a '
                 f'single CHOP zone may show no figures; select its Area.",' if z.chop_only else "")
    ws["B5"] = (f'=IF({C("$B$9")}=0,"No {z.key.title()} data loaded yet. Put the export in the inbox and run '
                f'ingest.py.",IF(NOT({C("$B$5")}),"⚠  Pick "&IF({C("$B$3")}="Area","an Area","a Zone")&'
                f'" from SELECT (type part of the name in SEARCH first).",IF({C("$B$18")},"This zone is a '
                f'Community Church: see the COMMUNITY CHURCH DASHBOARD.",IF(NOT({C("$B$17")}),"{no_service}",'
                f'{zone_note}"Compliance uses the reporting system\'s count of {z.universe_label} with a report. '
                f'Community Churches are excluded."{")" if z.chop_only else ""}))))')
    ws.merge_cells("B5:O5")
    ws["B5"].font = font(9, False, tint(colour, 0.2), italic=True)
    ws["B5"].alignment = LEFT
    K = kpis(z)
    krow = {label: KPI0 + i for i, (label, *_) in enumerate(K)}
    u = z.universe_label
    cards = [("GRAND TOTAL ATTENDANCE", K[0][0], FMT_INT, "pct"), ("ADULTS (M + F)", K[1][0], FMT_INT, "pct"),
             ("CHILDREN", K[4][0], FMT_INT, "pct"), (f"% {u.upper()} REPORTING", K[10][0], FMT_PCT, "pts"),
             (f"{u.upper()} REPORTED", K[8][0], FMT_INT, "pct"), ("FIRST TIMERS", K[5][0], FMT_INT, "pct"),
             ("AVG PER REPORTING ZONE", K[12][0], FMT_DEC1, "pct")]
    blocked = f'OR(NOT({C("$B$17")}),{C("$B$18")})'
    card_bg = tint(colour, 0.92)
    for (title, kpi, fmt, delta), c1 in zip(cards, "BDFHJLN"):
        c2 = chr(ord(c1) + 1)
        for r in range(7, 11):
            for cc in (c1, c2):
                ws[f"{cc}{r}"].fill = fill(card_bg)
                ws[f"{cc}{r}"].border = Border(left=side(WHITE, "thick") if cc == c1 else None,
                                               right=side(WHITE, "thick") if cc == c2 else None,
                                               top=side(colour, "thick") if r == 7 else None)
        r = krow[kpi]
        put(ws, f"{c1}7", title, f=font(8, True, tint(colour, 0.15)), align=CENTER, merge_to=f"{c2}7")
        put(ws, f"{c1}8", f'=IF({blocked},"n/a",IF({C(f"$B${r}")}="","–",{C(f"$B${r}")}))',
            f=font(20, True, colour), fmt=fmt, align=CENTER, merge_to=f"{c2}8")
        if delta == "pct":
            dtext = (f'=IF({C(f"$E${r}")}="","– vs prev wk",IF({C(f"$E${r}")}>=0,"▲ ","▼ ")&'
                     f'TEXT(ABS({C(f"$E${r}")}),"0.0%")&" vs prev wk")')
        else:
            dtext = (f'=IF({C(f"$D${r}")}="","– vs prev wk",IF({C(f"$D${r}")}>=0,"▲ ","▼ ")&'
                     f'TEXT(ABS({C(f"$D${r}")})*100,"0.0")&" pts vs prev wk")')
        avg = C(f"$F${r}")
        avg_text = {FMT_PCT: f'TEXT({avg},"0.0%")', FMT_DEC1: f'TEXT({avg},"#,##0.0")'}.get(
            fmt, f'TEXT({avg},"#,##0")')
        put(ws, f"{c1}9", dtext, f=font(9, True, MUTED), align=CENTER, merge_to=f"{c2}9")
        put(ws, f"{c1}10", f'="Avg last "&{C("$B$11")}&" wks: "&IF({avg}="","–",{avg_text})',
            f=font(8, False, MUTED), align=CENTER, merge_to=f"{c2}10")
        arrow_rules(ws, f"{c1}9", f"{c1}9")
        if fmt == FMT_PCT:
            rag_rules(ws, f"{c1}8", f"{c1}8")
    ws.row_dimensions[8].height = 34

    section_title(ws, 12, "Time analysis — selected week against previous week, the N-week average and "
                          "period-to-date", colour, last="O")
    put(ws, "B13", "KPI", f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to="C13")
    for j, h in enumerate(MEASURES):
        cell = ws.cell(13, 4 + j)
        cell.value = (f'="Avg last "&{C("$B$11")}&" wks"' if h == "Avg last N wks" else
                      f'="vs "&{C("$B$11")}&"-wk avg"' if h == "vs N-wk avg" else h)
        cell.font, cell.fill, cell.alignment = font(9, True, WHITE), fill(colour), CENTER
    ws.row_dimensions[13].height = 28
    for i, (label, kind, src, fmt) in enumerate(K):
        r, cr = 14 + i, KPI0 + i
        band = tint(colour, 0.95) if i % 2 == 0 else None
        put(ws, f"B{r}", label, f=font(9, i in (0, 10)), align=LEFT, bg=band, merge_to=f"C{r}")
        for j, L in enumerate("BCDEFGHIJKLM"):
            cell = ws.cell(r, 4 + j)
            cell.value = f'=IF({blocked},"n/a",IF({C(f"${L}${cr}")}="","–",{C(f"${L}${cr}")}))'
            if L in "EG":
                cell.number_format = FMT_PCT_SIGNED
            elif L == "D" and fmt == FMT_PCT:
                cell.number_format = '+0.0" pts";-0.0" pts";0.0" pts"'
                cell.value = (f'=IF({blocked},"n/a",IF({C(f"$D${cr}")}="","–",{C(f"$D${cr}")}*100))')
            elif L == "D":
                cell.number_format = {FMT_INT: "+#,##0;-#,##0;0", FMT_DEC1: "+#,##0.0;-#,##0.0;0.0"}.get(
                    fmt, "+0.0;-0.0;0.0")
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
            rag_rules(ws, f"H{r}", f"H{r}")
    end_ta = 14 + len(K)
    put(ws, f"B{end_ta}", "Averages use only weeks that have data. MTD/QTD/YTD run to the selected week; ratio rows "
                          "show the ratio of the period's totals.", f=font(8, False, MUTED, italic=True), align=LEFT,
        merge_to=f"O{end_ta}")

    ch = end_ta + 2
    section_title(ws, ch, "Trends — last 13 weeks", colour, last="O")
    calc = wb[z.calc]
    bar = BarChart()
    bar.type, bar.grouping, bar.overlap, bar.gapWidth = "col", "stacked", 100, 60
    _title(bar, "Attendance mix")
    for c in (2, 3, 4):
        bar.add_data(Reference(calc, min_col=c, min_row=104, max_row=117), titles_from_data=True)
    bar.set_categories(Reference(calc, min_col=1, min_row=105, max_row=117))
    for s_, shade in zip(bar.series, [colour, tint(colour, 0.45), "E0A526"]):
        s_.graphicalProperties.solidFill = shade
        s_.graphicalProperties.line.solidFill = shade
    bar.legend.position = "b"
    bar.y_axis.numFmt = "#,##0"
    _style_chart(bar)
    bar.width, bar.height = 18.5, 7.4
    ws.add_chart(bar, f"B{ch + 1}")
    pct = BarChart()
    pct.type, pct.gapWidth = "col", 60
    _title(pct, f"% {u} reporting")
    pct.add_data(Reference(calc, min_col=5, min_row=104, max_row=117), titles_from_data=True)
    pct.set_categories(Reference(calc, min_col=1, min_row=105, max_row=117))
    pct.series[0].graphicalProperties.solidFill = tint(colour, 0.25)
    pct.series[0].graphicalProperties.line.solidFill = tint(colour, 0.25)
    pct.y_axis.numFmt = "0%"
    pct.y_axis.scaling.min, pct.y_axis.scaling.max = 0, 1
    pct.legend = None
    _style_chart(pct)
    pct.width, pct.height = 17.0, 7.4
    ws.add_chart(pct, f"I{ch + 1}")

    rk = ch + 17
    section_title(ws, rk, "Area rankings — selected week", colour, last="O",
                  note="Growth = selected week vs previous week")
    zone_scope = f'{C("$B$3")}="Zone"'
    tables = [("B", "Top 10 · attendance", FMT_INT), ("D", "Bottom 10 · attendance", FMT_INT),
              ("F", "Top 10 · % reporting", FMT_PCT), ("H", "Bottom 10 · % reporting", FMT_PCT),
              ("J", "Top 10 · growth", FMT_PCT_SIGNED), ("L", "Bottom 10 · growth", FMT_PCT_SIGNED)]
    for (c1, title, fmt), (n_col, v_col) in zip(tables, calc_info["top"]):
        c2 = chr(ord(c1) + 1)
        put(ws, f"{c1}{rk + 1}", title, f=font(9, True, WHITE), bg=colour, align=CENTER, merge_to=f"{c2}{rk + 1}")
        for k in range(10):
            r, cr = rk + 2 + k, AREA0 + k
            nm = C(f"${n_col}${cr}")
            ws[f"{c1}{r}"] = (f'=IF({zone_scope},IF({k}=0,"n/a at Zone scope",""),IF({nm}="",IF({k}=0,"–",""),'
                              f'{k + 1}&". "&{nm}))')
            ws[f"{c2}{r}"] = f'=IF(OR({zone_scope},{nm}=""),"",{C(f"${v_col}${cr}")})'
            ws[f"{c1}{r}"].font, ws[f"{c2}{r}"].font = font(9), font(9, True)
            ws[f"{c2}{r}"].number_format, ws[f"{c2}{r}"].alignment = fmt, RIGHT
            for cc in (c1, c2):
                ws[f"{cc}{r}"].border = Border(bottom=side())
        if fmt == FMT_PCT:
            rag_rules(ws, f"{c2}{rk + 2}:{c2}{rk + 11}", f"{c2}{rk + 2}")
        elif fmt == FMT_INT:
            ws.conditional_formatting.add(f"{c2}{rk + 2}:{c2}{rk + 11}", DataBarRule(
                start_type="num", start_value=0, end_type="max", color=tint(colour, 0.35)))
        else:
            signed_rules(ws, f"{c2}{rk + 2}:{c2}{rk + 11}", f"{c2}{rk + 2}")
    put(ws, f"N{rk + 1}", "Selected Area rank", f=font(9, True, WHITE), bg=colour, align=CENTER,
        merge_to=f"O{rk + 1}")
    rc = calc_info["rank"]
    for i, label in enumerate(("Attendance", "% reporting", "Growth")):
        r = rk + 2 + i
        put(ws, f"N{r}", label, f=font(9))
        put(ws, f"O{r}", f'=IF({C("$B$3")}<>"Area","n/a",IF({C(f"${rc}${AREA0 + i}")}="","–",'
                         f'{C(f"${rc}${AREA0 + i}")}&" of "&{C(f"${rc}${AREA0 + 3}")}))', f=font(9, True, colour),
            align=RIGHT)

    nr = rk + 14
    gap = calc_info["gap"]
    section_title(ws, nr, f"Who didn't report — Areas with {u} missing a report for the selected week", colour,
                  last="O")
    ws[f"B{nr + 1}"] = (f'=IF({zone_scope},IF({C("$B$17")},IF({C("$" + SER["reported"] + "$21")}=1,'
                        f'"This zone reported.","This zone has no report for the selected week."),""),'
                        f'IF({C(f"${gap[0]}$1")}=0,"Every {u[:-1]} reported.",TEXT({C(f"${gap[1]}$1")},"#,##0")'
                        f'&" {u} without a report, across "&{C(f"${gap[0]}$1")}&" Area(s)."))')
    ws.merge_cells(f"B{nr + 1}:O{nr + 1}")
    ws[f"B{nr + 1}"].font = font(9, True, RED_FG)
    for c1, c2, h in (("B", "E", "Area"), ("F", "G", u[0].upper() + u[1:]), ("H", "I", "Reported"),
                      ("J", "K", "Missing")):
        put(ws, f"{c1}{nr + 2}", h, f=font(9, True, WHITE), bg=colour, align=LEFT, merge_to=f"{c2}{nr + 2}")
    for k in range(LIST_ROWS):
        r = nr + 3 + k
        for (c1, c2), L in zip((("B", "E"), ("F", "G"), ("H", "I"), ("J", "K")), gap[1:]):
            put(ws, f"{c1}{r}", f'=IF({zone_scope},"",{C(f"${L}${AREA0 + k}")})', f=font(9), merge_to=f"{c2}{r}",
                bg=tint(colour, 0.95) if k % 2 == 0 else None, fmt=None if c1 == "B" else "#,##0")
    ws.print_area = f"A1:P{nr + 3 + LIST_ROWS}"
    ws.print_title_rows = "1:5"
    return ws
