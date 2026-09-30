"""Readers for the weekly transport reports, in the layouts the transport office uses.

These reports can't go through the generic column matcher in ingest.py: they
repeat header names ("TOTAL SPENT PER AREA" three times), leave headers blank
(WSF Procured zone code and cost), swap a header with its data (the EV sheet)
and mix Area-total rows in with bus rows.  Each layout is recognised by its
header signature and read by its own function.

Layouts read into the databank
    FT_PROCURED     one row per hired LT bus (plus an Area-total row per Area)
    COASTER_REPORT  one row per church coaster (some EV BRT rows mixed in)
    EV_TATA_REPORT  one row per EV / TATA bus run
    WSF_PROCURED    one row per zone (buses hired by members)
    FT_ALLOCATION   buses allocated and approved cost per bus, per Area
    COASTER_FUEL    fuel payable per coaster (only HUB coasters are paid)

Reference sheets recognised and skipped
    EV/TATA allocation, Hiace bus list, hub payment summary
"""
from __future__ import annotations

import datetime as dt
import re
from collections import Counter
from dataclasses import dataclass, field

import numpy as np
import pandas as pd

from databank import Masters, norm_key

MONTHS = {m: i for i, m in enumerate(
    ["jan", "feb", "mar", "apr", "may", "jun", "jul", "aug", "sep", "oct", "nov", "dec"], start=1)}


def hkey(v) -> str:
    return re.sub(r"[^a-z0-9]", "", str(v).lower()) if v is not None else ""


def blank(v) -> bool:
    return v is None or (isinstance(v, float) and np.isnan(v)) or str(v).strip() in ("", "nan", "NaT")


def num(v):
    """Parse messy numbers: '12$', '$18.00', '170-,000', '20`', '50000'. None when empty or not a number."""
    if blank(v):
        return None
    if isinstance(v, (int, float, np.integer, np.floating)):
        return float(v)
    s = str(v).strip()
    if s.startswith("#"):          # #DIV/0!, #VALUE!
        return None
    s = s.replace(",", "")
    m = re.search(r"\d+(\.\d+)?", s.replace("-", "").replace("`", "").replace("$", ""))
    return float(m.group(0)) if m else None


def is_messy(v) -> bool:
    """True when a number had to be cleaned (so the source cell deserves a warning)."""
    if blank(v) or isinstance(v, (int, float, np.integer, np.floating)):
        return False
    return not re.fullmatch(r"\s*\d[\d,]*(\.\d+)?\s*", str(v))


def zone_code(v) -> str | None:
    """'lfc0112', 'LFC 0701', 'LFC701', '313', 'LFC06010', 'L;FC6705', '77015' -> 'LFC0112' style."""
    if blank(v):
        return None
    first = re.split(r"[/&]", str(v))[0]
    digits = "".join(re.findall(r"\d", first))
    if len(digits) == 3:
        area, zone = digits[0], digits[1:]
    elif len(digits) == 4:
        area, zone = digits[:2], digits[2:]
    elif len(digits) == 5:
        area, zone = digits[:2], digits[2:]
    else:
        return None
    return f"LFC{int(area):02d}{int(zone):02d}"


def date_in_text(*texts) -> dt.date | None:
    for t in texts:
        if blank(t):
            continue
        if isinstance(t, (dt.date, dt.datetime, pd.Timestamp)):
            return pd.Timestamp(t).date()
        m = re.search(r"(\d{1,2})\s*(?:st|nd|rd|th)?[\s_\-,.]*([A-Za-z]{3,9})[\s_\-,.]*(\d{4})", str(t), re.I)
        if m and m.group(2)[:3].lower() in MONTHS:
            try:
                return dt.date(int(m.group(3)), MONTHS[m.group(2)[:3].lower()], int(m.group(1)))
            except ValueError:
                continue
    return None


# --------------------------------------------------------------------------
# Layout detection
# --------------------------------------------------------------------------

SIGNATURES = {
    "FT_PROCURED": ["zincode", "sightedbusesamount"],
    "COASTER_REPORT": ["bustagnumber", "capacityfullyutilized"],
    "EV_TATA_REPORT": ["hublocations", "adulttotal", "serviceattended"],
    "WSF_PROCURED": ["totalnoofbusesdeliveredperzone"],
    "FT_ALLOCATION": ["approvedcostbus", "newtotalnumberofbuses"],
    "COASTER_FUEL": ["totalpayablefuelamount"],
    # reference sheets: recognised so they are skipped quietly
    "REF_EV_ALLOCATION": ["totalallocatedelectric"],
    "REF_HIACE_LIST": ["platenumber", "conditionofthevehicle"],
    "REF_HUB_SUMMARY": ["noofpayablecoasters"],
}
LOADED_LAYOUTS = {"FT_PROCURED", "COASTER_REPORT", "EV_TATA_REPORT", "WSF_PROCURED", "FT_ALLOCATION",
                  "COASTER_FUEL"}


def detect(raw: pd.DataFrame, scan_rows: int = 15):
    """Return (layout, header row index) or (None, None)."""
    for r in range(min(scan_rows, len(raw))):
        keys = [hkey(v) for v in raw.iloc[r].tolist()]
        for layout, sig in SIGNATURES.items():
            if all(any(k.startswith(s) for k in keys) for s in sig):
                if layout == "EV_TATA_REPORT" and any(k.startswith("bustagnumber") for k in keys):
                    continue
                return layout, r
    return None, None


class Cols:
    """Column lookup on a header row: by exact key, by prefix, with an occurrence index for repeats."""

    def __init__(self, header: list):
        self.keys = [hkey(v) for v in header]

    def find(self, key, prefix=False, nth=0):
        hits = [i for i, k in enumerate(self.keys) if (k.startswith(key) if prefix else k == key)]
        return hits[nth] if len(hits) > nth else None


# --------------------------------------------------------------------------
# Result container
# --------------------------------------------------------------------------

RUN_COLS = ["week_ending", "report_date", "report", "category", "vehicle_type", "area", "zone_code", "location",
            "vehicle_id", "buses", "capacity", "trips", "male", "female", "children", "cost", "paid_by",
            "status", "service", "remarks_category", "remarks", "in_church", "source_row"]
COST_COLS = ["week_ending", "report_date", "report", "category", "area", "vehicle_id", "cost_type", "hub_status",
             "trips", "quantity", "unit_price", "amount", "payable", "paid_by", "source_row"]
BUDGET_COLS = ["week_ending", "report_date", "report", "area_no", "area", "buses_allocated", "cost_per_bus",
               "expected_spend", "payable_this_week", "source_row"]


@dataclass
class Parsed:
    layout: str
    table: str | None                       # transport_runs / transport_costs / transport_budget
    category: str | None
    report_date: dt.date | None
    rows: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    checks: list = field(default_factory=list)   # (label, file total, loaded total)

    def issue(self, severity, issue, detail, row=None):
        self.issues.append({"severity": severity, "issue": issue, "detail": str(detail)[:250],
                            "source_row": row})

    def frame(self, cols) -> pd.DataFrame:
        return pd.DataFrame(self.rows, columns=cols)


class AreaMatcher:
    def __init__(self, m: Masters):
        self.m = m
        self._no = {name: int(no) for no, name in m.area_no.items() if str(no).strip().isdigit()}

    def number_of(self, area) -> int | None:
        return self._no.get(area)

    def __call__(self, name=None, number=None):
        for v in (name,):
            if not blank(v):
                clean = re.sub(r"\(.*?\)", "", str(v))
                hit = self.m.area_by_key.get(norm_key(v)) or self.m.area_by_key.get(norm_key(clean))
                if hit:
                    return hit
        if not blank(number):
            n = num(number)
            if n is not None:
                return self.m.area_no.get(str(int(n)))
        return None


def _people(row, c_m, c_f, c_c, p: Parsed, excel_row: int):
    out = []
    for c, label in ((c_m, "Male"), (c_f, "Female"), (c_c, "Children")):
        v = row[c] if c is not None else None
        n = num(v)
        if is_messy(v):
            p.issue("Warning", "Number cleaned", f"{label} '{v}' read as {int(n) if n is not None else 0}", excel_row)
        out.append(int(n) if n else 0)
    return out


# --------------------------------------------------------------------------
# Layout readers
# --------------------------------------------------------------------------

def read_ft_procured(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("FT_PROCURED", "transport_runs", "FT Procured", when)
    c = Cols(raw.iloc[hdr].tolist())
    ci = dict(area_no=c.find("areano"), area=c.find("areaname"), addr=c.find("loadingbay", True),
              zone=c.find("zincode"), cap_a=c.find("capacityofadditionalbusesapproved", True),
              cap=c.find("capacity"), sighted=c.find("sightedbusesamount"), in_church=c.find("numberofbusesinchurch"),
              total=c.find("total"),
              a_buses=c.find("totalspentperarea", nth=0), a_cost=c.find("totalspentperarea", nth=1),
              a_riders=c.find("totalspentperarea", nth=2), m=c.find("male"), f=c.find("female"),
              ch=c.find("children"), remarks=c.find("remarks", True))
    area_totals = {}
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        excel_row = i + 1
        name = row[ci["area"]]
        if blank(name) and blank(row[ci["area_no"]]):
            if any(num(row[k]) for k in (ci["a_buses"], ci["a_cost"], ci["sighted"]) if k is not None):
                p.checks.append(("File grand total: buses", num(row[ci["a_buses"]]), None))
                p.checks.append(("File grand total: cost", num(row[ci["sighted"]]), None))
            continue
        area = match(name, row[ci["area_no"]])
        if area is None:
            p.issue("Rejected", "Unknown Area", f"Area '{name}' / no. {row[ci['area_no']]} not in MASTER_AREAS",
                    excel_row)
            continue
        # a row at an Area boundary sometimes carries the neighbouring Area's name: when the zone code and
        # the Area number agree with each other but not with the name, trust the code
        z0, n0 = zone_code(row[ci["zone"]]), num(row[ci["area_no"]])
        if z0 and n0 and int(z0[3:5]) == int(n0) and match.number_of(area) not in (None, int(n0)):
            fixed = match(None, n0)
            if fixed:
                p.issue("Warning", "Area name corrected from zone code",
                        f"row named {area} but zone {z0} and Area no. {int(n0)} are {fixed}; {fixed} used", excel_row)
                area = fixed
        if num(row[ci["a_buses"]]) is not None or num(row[ci["a_cost"]]) is not None:
            area_totals[area] = (num(row[ci["a_buses"]]) or 0, num(row[ci["a_cost"]]) or 0,
                                 num(row[ci["a_riders"]]) or 0, excel_row)
        cap = num(row[ci["cap"]]) if ci["cap"] is not None else None
        if cap is None and ci["cap_a"] is not None:
            cap = num(row[ci["cap_a"]])
        cost = num(row[ci["sighted"]])
        male, female, children = _people(row, ci["m"], ci["f"], ci["ch"], p, excel_row)
        riders = male + female + children
        if not (cap or cost or riders):
            continue                                  # loading bay listed, no bus
        if is_messy(row[ci["sighted"]]):
            p.issue("Warning", "Number cleaned", f"Bus cost '{row[ci['sighted']]}' read as {cost}", excel_row)
        given_total = num(row[ci["total"]]) if ci["total"] is not None else None
        if given_total is not None and int(given_total) != riders:
            p.issue("Warning", "Total mismatch", f"TOTAL says {int(given_total)}, Male+Female+Children = {riders}; "
                                                 "computed value used", excel_row)
        in_church = "Y" if num(row[ci["in_church"]]) else "N"
        z = zone_code(row[ci["zone"]])
        if z and match.number_of(area) and int(z[3:5]) != match.number_of(area):
            p.issue("Warning", "Zone code outside its Area", f"{row[ci['zone']]} listed under {area}", excel_row)
        status = "Ran" if riders else "No riders reported"
        if cap and cap > 22:
            vtype = "Bigger than LT"
        elif cap and cap < 14:
            vtype = "Smaller than LT"
        else:
            vtype = "LT 22-seater"
        p.rows.append([None, when, p.layout, "FT Procured", vtype, area, z, row[ci["addr"]] if not blank(row[ci["addr"]]) else "",
                       "", 1, cap, 1, male, female, children, cost or 0, "Central", status, "",
                       "", "" if blank(row[ci["remarks"]]) else str(row[ci["remarks"]]), in_church, excel_row])
    # cross-check against each Area's own total row
    df = p.frame(RUN_COLS)
    for area, (buses, cost, riders, r) in area_totals.items():
        got = df[df.area == area]
        in_ch = int((got.in_church == "Y").sum())
        got_riders = int((got.male + got.female + got.children).sum())
        if int(buses) != in_ch or abs(cost - got.cost.sum()) > 0.5 or int(riders) != got_riders:
            p.issue("Warning", "Area total differs from its bus rows",
                    f"{area}: Area row says {int(buses)} buses in church / ₦{cost:,.0f} / {int(riders)} riders; "
                    f"bus rows give {in_ch} / ₦{got.cost.sum():,.0f} / {got_riders}", r)
    for k, (label, file_v, _) in enumerate(p.checks):
        loaded = int((df.in_church == "Y").sum()) if "buses" in label else df.cost.sum()
        p.checks[k] = (label.replace("buses", "buses marked in church"), file_v, loaded)
    p.checks.append(("Bus rows (all hired buses)", None, len(df)))
    return p


def read_coaster_report(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("COASTER_REPORT", "transport_runs", "Church Coaster", when)
    c = Cols(raw.iloc[hdr].tolist())
    ci = dict(hub=c.find("hublocations"), host=c.find("fthubhostarea"), park=c.find("busparkingstation", True),
              tag=c.find("bustagnumber"), bus=c.find("bus"), full=c.find("capacityfullyutilized"),
              trips=c.find("numberoftrips"), m=c.find("male", True), f=c.find("female", True),
              ch=c.find("children", True), service=c.find("serviceattended", True),
              rcat=c.find("operationalremarkscategory", True), rdet=c.find("operationalremarksdetails", True))
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        excel_row = i + 1
        host, park, tag = row[ci["host"]], row[ci["park"]], row[ci["tag"]]
        if blank(host) and blank(park) and blank(tag):
            continue
        area = match(host) or match(park)
        if area is None:
            p.issue("Rejected", "Unknown Area", f"host '{host}', parking '{park}'", excel_row)
            continue
        male, female, children = _people(row, ci["m"], ci["f"], ci["ch"], p, excel_row)
        riders = male + female + children
        trips = num(row[ci["trips"]]) or (1 if riders else 0)
        kind = str(row[ci["bus"]] or "").upper()
        rcat = "" if blank(row[ci["rcat"]]) else str(row[ci["rcat"]]).strip()
        if "breakdown" in rcat.lower():
            status = "Breakdown"
        elif riders:
            status = "Ran"
        else:
            status = "Not run / no report"
        if "EV" in kind:
            category, vtype, cap = "EV/TATA", "Electric", settings.get("cap_electric", 70)
        else:
            category, vtype, cap = "Church Coaster", "Coaster", settings.get("cap_coaster", 30)
        p.rows.append([None, when, p.layout, category, vtype, area, None, "" if blank(row[ci["hub"]]) else str(row[ci["hub"]]),
                       vehicle_key(tag), 1, cap, int(trips), male, female, children, 0, "Central", status,
                       "" if blank(row[ci["service"]]) else str(row[ci["service"]]), rcat,
                       "" if blank(row[ci["rdet"]]) else str(row[ci["rdet"]]), "", excel_row])
    df = p.frame(RUN_COLS)
    p.checks.append(("Coasters listed", None, len(df)))
    p.checks.append(("Riders", None, int((df.male + df.female + df.children).sum())))
    return p


def read_ev_tata_report(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("EV_TATA_REPORT", "transport_runs", "EV/TATA", when)
    c = Cols(raw.iloc[hdr].tolist())
    a, b = c.find("hublocations"), c.find("fthubhostarea")
    body = raw.iloc[hdr + 1:]
    # the "hub" and "host area" headers are sometimes swapped: use whichever column holds Area names
    hits = lambda col: body.iloc[:, col].map(lambda v: match(v) is not None).sum()
    area_col, loc_col = (a, b) if hits(a) >= hits(b) else (b, a)
    ci = dict(park=c.find("busparkingstation", True), m=c.find("male"), f=c.find("female"),
              ch=c.find("children"), service=c.find("serviceattended"), rem=c.find("remarks"))
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        excel_row = i + 1
        if blank(row[area_col]) and blank(row[loc_col]):
            continue
        area = match(row[area_col]) or match(row[loc_col])
        if area is None:
            p.issue("Rejected", "Unknown Area", f"'{row[area_col]}' / '{row[loc_col]}'", excel_row)
            continue
        male, female, children = _people(row, ci["m"], ci["f"], ci["ch"], p, excel_row)
        riders = male + female + children
        park = "" if blank(row[ci["park"]]) else str(row[ci["park"]])
        vtype = "TATA" if "canaan" in park.lower() else "Electric"
        cap = settings.get("cap_tata", 70) if vtype == "TATA" else settings.get("cap_electric", 70)
        remarks = "" if blank(row[ci["rem"]]) else str(row[ci["rem"]])
        p.rows.append([None, when, p.layout, "EV/TATA", vtype, area, None, "" if blank(row[loc_col]) else str(row[loc_col]),
                       "", 1, cap, 1 if riders else 0, male, female, children, 0, "Central",
                       "Ran" if riders else "Not run / no report",
                       "" if blank(row[ci["service"]]) else str(row[ci["service"]]), "", remarks, "", excel_row])
    return p


def read_wsf_procured(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("WSF_PROCURED", "transport_runs", "WSF Procured", when)
    header = raw.iloc[hdr].tolist()
    above = raw.iloc[hdr - 1].tolist() if hdr > 0 else [None] * len(header)
    # blank header cells take the group label above them (e.g. COST)
    header = [h if not blank(h) else above[j] for j, h in enumerate(header)]
    c = Cols(header)
    body = raw.iloc[hdr + 1:]
    zcol = None
    for j in range(raw.shape[1]):
        vals = body.iloc[:, j].dropna().astype(str)
        if len(vals) and (vals.str.contains(r"(?i)l\s*f\s*c|^\s*\d{3,5}\s*$").mean() > 0.6):
            zcol = j
            break
    ci = dict(area_no=c.find("areano"), area=c.find("areaname"), buses=c.find("totalnoofbusesdelivered", True),
              cap=c.find("totalcapacityofbuses", True), m=c.find("m"), f=c.find("f"), ch=c.find("c"),
              cost=c.find("cost"), rem=c.find("remark"))
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        excel_row = i + 1
        male, female, children = _people(row, ci["m"], ci["f"], ci["ch"], p, excel_row)
        riders = male + female + children
        buses, cap = num(row[ci["buses"]]), num(row[ci["cap"]])
        if blank(row[ci["area"]]) and blank(row[ci["area_no"]]):
            if riders or buses:
                p.checks.append(("File grand total: buses", buses, None))
                p.checks.append(("File grand total: riders", riders, None))
                p.checks.append(("File grand total: cost", num(row[ci["cost"]]), None))
            continue
        if not (riders or buses):
            continue                                   # zone listed, no WSF procured bus
        area = match(row[ci["area"]], row[ci["area_no"]])
        if area is None:
            p.issue("Rejected", "Unknown Area", f"Area '{row[ci['area']]}' / no. {row[ci['area_no']]}", excel_row)
            continue
        if not buses and cap and cap <= 10 and riders > cap * 5:
            p.issue("Warning", "Columns shifted", f"bus count found in the capacity column ({int(cap)}); "
                                                  "capacity unknown", excel_row)
            buses, cap = cap, None
        elif not buses:
            p.issue("Warning", "Bus count missing", f"{riders} riders but no bus count; 1 bus assumed", excel_row)
            buses = 1
        if not cap:
            p.issue("Warning", "Capacity missing", f"{int(buses)} bus(es), capacity not given; "
                                                   "utilisation not calculated for this zone", excel_row)
        z = zone_code(row[zcol]) if zcol is not None else None
        p.rows.append([None, when, p.layout, "WSF Procured", "Hired by zone", area, z, "", "", int(buses), cap, 1,
                       male, female, children, num(row[ci["cost"]]) or 0, "Members", "Ran", "", "",
                       "" if ci["rem"] is None or blank(row[ci["rem"]]) else str(row[ci["rem"]]), "", excel_row])
    df = p.frame(RUN_COLS)
    loaded = {"buses": df.buses.sum(), "riders": int((df.male + df.female + df.children).sum()),
              "cost": df.cost.sum()}
    p.checks = [(lbl, v, loaded[lbl.split(": ")[1]]) for lbl, v, _ in p.checks]
    return p


def read_ft_allocation(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("FT_ALLOCATION", "transport_budget", None, when)
    c = Cols(raw.iloc[hdr].tolist())
    payable = [j for j, k in enumerate(c.keys) if k.startswith("payableforftprocured")]
    ci = dict(no=c.find("areano"), name=c.find("areaname"), buses=c.find("newtotalnumberofbuses"),
              cost=c.find("approvedcostbus"), exp=c.find("totalexpectedspendperweek"),
              pay=payable[-1] if payable else None)
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        if blank(row[ci["name"]]):
            continue
        area = match(row[ci["name"]], row[ci["no"]])
        if area is None:
            p.issue("Rejected", "Unknown Area", f"'{row[ci['name']]}'", i + 1)
            continue
        buses, cost = num(row[ci["buses"]]) or 0, num(row[ci["cost"]]) or 0
        p.rows.append([None, when, p.layout, int(num(row[ci["no"]]) or 0), area, int(buses), cost, buses * cost,
                       num(row[ci["pay"]]) if ci["pay"] is not None else None, i + 1])
    return p


def read_coaster_fuel(raw, hdr, match, settings, when) -> Parsed:
    p = Parsed("COASTER_FUEL", "transport_costs", "Church Coaster", when)
    c = Cols(raw.iloc[hdr].tolist())
    ci = dict(host=c.find("fthubhostarea"), status=c.find("hubstatus"), trips=c.find("totaltrips"),
              no=c.find("areanumber"), serial=c.find("busserialnumber"), reg=c.find("coasterregistrationnumber"),
              litres=c.find("totalpayablefuellitres"), price=c.find("pumpprice", True),
              amount=c.find("totalpayablefuelamount", True))
    prices = Counter(num(v) for v in raw.iloc[hdr + 1:, ci["price"]] if num(v))
    usual = prices.most_common(1)[0][0] if prices else None
    for i in range(hdr + 1, len(raw)):
        row = raw.iloc[i].tolist()
        excel_row = i + 1
        if blank(row[ci["host"]]) and blank(row[ci["no"]]):
            continue
        area = match(row[ci["host"]], row[ci["no"]])
        if area is None:
            p.issue("Rejected", "Unknown Area", f"'{row[ci['host']]}' / no. {row[ci['no']]}", excel_row)
            continue
        by_no = match(None, row[ci["no"]])
        if by_no and by_no != area:
            p.issue("Warning", "Area number does not match Area name",
                    f"'{row[ci['host']]}' has Area number {row[ci['no']]} ({by_no}); name used", excel_row)
        status = str(row[ci["status"]] or "").strip().upper()
        amount, litres, price = num(row[ci["amount"]]) or 0, num(row[ci["litres"]]) or 0, num(row[ci["price"]])
        if usual and price and price != usual and litres:
            p.issue("Warning", "Pump price differs from the rest of the sheet",
                    f"₦{price:,.0f}/L instead of ₦{usual:,.0f}/L: ₦{litres * (price - usual):,.0f} extra", excel_row)
        vid = vehicle_key(row[ci["serial"]]) or vehicle_key(row[ci["reg"]])
        p.rows.append([None, when, p.layout, "Church Coaster", area, vid, "Fuel", status, num(row[ci["trips"]]) or 0,
                       litres, price, amount, "Y" if status == "HUB" and amount else "N", "Central", excel_row])
    df = p.frame(COST_COLS)
    p.checks.append(("Fuel payable, HUB coasters only", None, df.loc[df.payable == "Y", "amount"].sum()))
    p.checks.append(("Fuel on ZONE coasters (not paid)", None, df.loc[(df.payable == "N"), "amount"].sum()))
    return p


READERS = {"FT_PROCURED": read_ft_procured, "COASTER_REPORT": read_coaster_report,
           "EV_TATA_REPORT": read_ev_tata_report, "WSF_PROCURED": read_wsf_procured,
           "FT_ALLOCATION": read_ft_allocation, "COASTER_FUEL": read_coaster_fuel}


def vehicle_key(v) -> str:
    """Bus serial '154' / 154.0 -> '154'; plate 'AGL 352 KS' -> 'AGL352KS'."""
    if blank(v):
        return ""
    n = num(v)
    if n is not None and re.fullmatch(r"\s*\d+(\.0+)?\s*", str(v)):
        return str(int(n))
    return re.sub(r"[^A-Z0-9]", "", str(v).upper())


def read(raw: pd.DataFrame, sheet: str, file_name: str, masters: Masters, settings: dict) -> Parsed | None:
    """Read a sheet if it is one of the transport layouts; None if it isn't."""
    layout, hdr = detect(raw)
    if layout is None:
        return None
    if layout not in LOADED_LAYOUTS:
        return Parsed(layout, None, None, None)
    title_cells = [v for r in range(hdr + 1) for v in raw.iloc[r].tolist() if not blank(v)]
    when = date_in_text(*title_cells, sheet, file_name)
    parsed = READERS[layout](raw, hdr, AreaMatcher(masters), settings, when)
    parsed.report_date = when
    if when is None:
        parsed.issue("Rejected", "Report date not found",
                     "No date in the title rows, sheet name or file name (e.g. '27th Sept 2026')")
    return parsed
