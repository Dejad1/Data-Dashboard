"""Generate realistic synthetic masters and weekly inbox files at full scale.

    python generate_sample_data.py --root demo [--weeks 26] [--seed 7]

Creates <root>/masters/*.csv and <root>/inbox/* in deliberately varied
layouts (title rows above the header, shuffled columns, different header
spellings, CSV and Excel), with a handful of injected errors so the DATA
QUALITY sheet has something to show.
"""
from __future__ import annotations

import argparse
import datetime as dt
import random
from pathlib import Path

import numpy as np
import pandas as pd

from config import HERE, Paths, load_settings, week_ending

AREA_PLACES = """Ikeja Ikorodu Agege Alimosho Surulere Yaba Mushin Oshodi Isolo Ejigbo Ikotun Igando Idimu Egbeda
Ipaja Ayobo Abule-Egba Iyana-Ipaja Ogba Ojodu Berger Magodo Ketu Mile-12 Ojota Maryland Gbagada Bariga Shomolu
Palmgrove Onipanu Ilupeju Anthony Obanikoro Ebute-Metta Oyingbo Costain Ijora Apapa Ajegunle Amukoko Festac
Satellite-Town Ojo Alaba Iba Okokomaiko Badagry Agbara Ikoyi Victoria-Island Lekki Ajah Sangotedo Ibeju-Lekki
Epe Ikota Oniru Obalende Lagos-Island Idumota Isale-Eko Ogudu Alapere Agboyi Isheri Akowonjo Dopemu Iju Ifako
Ojokoro Alagbado Akute Ajuwon Mowe Ibafo Arepo Magboro Ota Sango Ifo Abeokuta Ibadan Ilorin Osogbo Akure Ife
Ijebu-Ode Sagamu Ilesa Ogijo Ikenne""".split()

COMMUNITY_PLACES = """Abuja Port-Harcourt Kano Kaduna Enugu Benin Warri Calabar Uyo Jos Owerri Asaba Onitsha Aba
Makurdi Lokoja Minna Bauchi Maiduguri Yola Sokoto Zaria Katsina Gombe Jalingo Lafia Awka Umuahia Abakaliki
Yenagoa Ado-Ekiti Ogbomosho Oyo Ondo Ilesha-East""".split()


def base_areas(n_areas: int) -> list[dict]:
    """The real Area list from masters/MASTER_AREAS.csv when present, else invented names."""
    real = HERE / "masters" / "MASTER_AREAS.csv"
    if real.exists():
        df = pd.read_csv(real, dtype=str, keep_default_na=False)
        if len(df) >= n_areas and "Area_No" in df:
            return df.head(n_areas).to_dict("records")
    return [{"Area_ID": f"A{i + 1:02d}", "Area_Name": AREA_PLACES[i].replace("-", " ").upper(), "Active": "Y",
             "Area_No": str(i + 1), "Aliases": ""} for i in range(n_areas)]


def build_masters(rng: random.Random, scale: float = 1.0):
    n_areas = max(3, round(92 * scale))
    target_zones, target_cells = 1250 * scale, 16000 * scale
    areas = base_areas(n_areas)
    # spread zones unevenly across areas (big areas have more zones); zones use the LFC<area><zone> codes
    weights = np.array([rng.uniform(0.5, 1.8) for _ in areas])
    per_area = np.maximum(5, np.round(weights / weights.sum() * target_zones)).astype(int)
    zones = []
    for a, n in zip(areas, per_area):
        for z in range(n):
            code = f"LFC{int(a['Area_No']):02d}{z + 1:02d}"
            zones.append({"Zone_ID": code, "Zone_Name": code, "Area_Name": a["Area_Name"],
                          "Has_CHOP": "Y" if rng.random() < 0.6 else "N", "Active": "Y"})
    cells_per_zone = target_cells / len(zones)
    cells = []
    for z in zones:
        n = max(4, round(rng.gauss(cells_per_zone + 0.4, 3)))
        for c in range(n):
            cells.append({"Cell": f"{z['Zone_ID']}-C{c + 1:02d}", "Zone_Name": z["Zone_Name"],
                          "Operational": "Y" if rng.random() > 0.02 else "N"})
    community = [{"Church_ID": f"CC{i + 1:02d}", "Church_Name": f"Community Church {p.replace('-', ' ')}",
                  "Location": p.replace("-", " "), "Active": "Y"} for i, p in enumerate(COMMUNITY_PLACES)]
    fleet = []
    for i in range(max(4, round(216 * scale))):
        a = rng.choice(areas)
        fleet.append({"Vehicle_ID": str(i + 1), "Category": "Church Coaster", "Capacity": 30, "Owner": "Church",
                      "Status": "Active", "Area_Name": a["Area_Name"], "Zone_Name": "", "Hired_By_Type": ""})
    return [pd.DataFrame(x) for x in (areas, zones, cells, community, fleet)]


def weeks_back(n: int, week_end_day: int, today: dt.date):
    last = week_ending(today, week_end_day)
    if last >= today:
        last -= dt.timedelta(days=7)
    return [last - dt.timedelta(weeks=k) for k in range(n)][::-1]


def seasonal(week_idx: int) -> float:
    return 1 + 0.06 * np.sin(week_idx / 4.0) + 0.002 * week_idx


def write_messy(df: pd.DataFrame, path: Path, rng: random.Random, sheet="Sheet1", title=None, csv=False):
    """Write with a random column order and optional title rows above the header."""
    cols = list(df.columns)
    rng.shuffle(cols)
    df = df[cols]
    if csv:
        df.to_csv(path.with_suffix(".csv"), index=False)
        return
    with pd.ExcelWriter(path.with_suffix(".xlsx"), engine="openpyxl") as xw:
        start = 0
        if title:
            pd.DataFrame([[title], [f"Generated {dt.date.today():%d %b %Y}"]]).to_excel(
                xw, sheet_name=sheet, index=False, header=False)
            start = 3
        df.to_excel(xw, sheet_name=sheet, index=False, startrow=start)


def fmt_date(d: dt.date, style: int):
    return [d, d.strftime("%d/%m/%Y"), d.isoformat(), d.strftime("%d-%b-%Y")][style % 4]


def generate(root: Path, n_weeks=26, seed=7, today: dt.date | None = None, scale: float = 1.0):
    rng = random.Random(seed)
    nrng = np.random.default_rng(seed)
    paths = Paths(root)
    paths.ensure()
    settings = load_settings(paths)
    wd = settings["week_end_day"]
    areas, zones, cells, community, fleet = build_masters(rng, scale)
    for name, df in zip(("MASTER_AREAS", "MASTER_ZONES", "MASTER_CELLS", "MASTER_COMMUNITY", "MASTER_FLEET"),
                        (areas, zones, cells, community, fleet)):
        df.to_csv(paths.masters / f"{name}.csv", index=False)

    weeks = weeks_back(n_weeks, wd, today or dt.date.today())
    zone_size = {z: rng.uniform(0.6, 1.6) for z in zones["Zone_Name"]}
    area_report_bias = {a: rng.uniform(0.0, 0.15) for a in areas["Area_Name"]}
    op_cells = cells[cells["Operational"] == "Y"].merge(zones[["Zone_Name", "Area_Name"]], on="Zone_Name")
    chop_zones = zones[zones["Has_CHOP"] == "Y"]
    header_variants = [
        {"male": "Male", "female": "Female", "children": "Children"},
        {"male": "MEN", "female": "WOMEN", "children": "Kids"},
        {"male": "No. of Males", "female": "No. of Females", "children": "No. of Children"},
    ]

    for wi, we in enumerate(weeks):
        s = seasonal(wi)
        non_report = rng.uniform(0.05, 0.25)
        hv = header_variants[wi % 3]
        sat = we - dt.timedelta(days=(we.isoweekday() - 6) % 7)  # the Saturday in that week
        wed = we - dt.timedelta(days=(we.isoweekday() - 3) % 7)
        fri = we - dt.timedelta(days=(we.isoweekday() - 5) % 7)
        sun = we

        # ---- WSF, cell level ------------------------------------------------
        p_miss = non_report + op_cells["Area_Name"].map(area_report_bias).to_numpy() - 0.07
        rep = op_cells[nrng.random(len(op_cells)) > p_miss].copy()
        base = rep["Zone_Name"].map(zone_size).to_numpy() * s
        rep["male"] = nrng.poisson(4.2 * base)
        rep["female"] = nrng.poisson(6.1 * base)
        rep["children"] = nrng.poisson(3.3 * base)
        out = pd.DataFrame({"Date": [fmt_date(sat, wi)] * len(rep), "Area": rep["Area_Name"].values,
                            "Zone": rep["Zone_Name"].values, "Cell": rep["Cell"].values,
                            hv["male"]: rep["male"].values, hv["female"]: rep["female"].values,
                            hv["children"]: rep["children"].values})
        out["Grand Total"] = rep["male"].values + rep["female"].values + rep["children"].values
        if wi == n_weeks - 1:  # injected errors in the latest week
            out.loc[0, "Grand Total"] = out.loc[0, "Grand Total"] + 5         # total mismatch
            out.loc[1, hv["male"]] = -3                                         # negative count
            out.loc[2, "Cell"] = "XX-UNKNOWN-C01"                               # unknown cell
            out = pd.concat([out, out.iloc[[3]]])                               # duplicate row
            out.iloc[4, out.columns.get_loc("Date")] = "31/02/2026"             # impossible date
        write_messy(out, paths.inbox / f"WSF cell returns w-e {we:%Y-%m-%d}", rng, sheet=f"WSF {we:%d.%m}",
                    title="WORSHIP SERVICE FELLOWSHIP - CELL RETURNS" if wi % 2 else None, csv=(wi % 5 == 4))

        # ---- Midweek, zone level -------------------------------------------
        zr = zones[nrng.random(len(zones)) > non_report * 0.8].copy()
        base = zr["Zone_Name"].map(zone_size).to_numpy() * s
        mid = pd.DataFrame({"Zone Name": zr["Zone_Name"].values, "Area": zr["Area_Name"].values,
                            "Meeting Date": fmt_date(wed, wi + 1),
                            hv["male"]: nrng.poisson(38 * base), hv["female"]: nrng.poisson(55 * base),
                            hv["children"]: nrng.poisson(21 * base)})
        if wi == n_weeks - 1:
            mid.loc[mid.index[0], "Zone Name"] = "Community Church Abuja"       # community row in zonal file
            mid.loc[mid.index[1], "Zone Name"] = "Nowhere Zone 99"              # unknown zone
        write_messy(mid, paths.inbox / f"Zonal Midweek report {we:%d-%m-%Y}", rng, sheet="Midweek",
                    title="ZONAL MIDWEEK SERVICE" if wi % 3 == 0 else None)

        # ---- CHOP, only CHOP zones -----------------------------------------
        cz = chop_zones[nrng.random(len(chop_zones)) > non_report].copy()
        base = cz["Zone_Name"].map(zone_size).to_numpy() * s
        chop = pd.DataFrame({"Date": fmt_date(fri, wi + 2), "Zone": cz["Zone_Name"].values,
                             "Male": nrng.poisson(17 * base), "Female": nrng.poisson(29 * base),
                             "Children": nrng.poisson(6 * base)})
        chop["Adult Total"] = chop["Male"] + chop["Female"]
        if wi == n_weeks - 1:
            nc = zones[zones["Has_CHOP"] == "N"].iloc[0]["Zone_Name"]
            chop = pd.concat([chop, pd.DataFrame([{"Date": fmt_date(fri, 0), "Zone": nc, "Male": 5, "Female": 9,
                                                   "Children": 1, "Adult Total": 14}])])
            chop.iloc[0, chop.columns.get_loc("Adult Total")] += 7              # adult total mismatch
        write_messy(chop, paths.inbox / f"Covenant Hour of Prayer {we:%Y%m%d}", rng, sheet="CHOP",
                    csv=(wi % 4 == 1))

        # ---- Community churches (all four services) --------------------------
        rows = []
        for _, c in community.iterrows():
            size = zone_size.get(c["Church_Name"]) or rng.uniform(1.5, 4)
            for svc, d, mult in (("Sunday Service", sun, 130), ("WSF", sat, 45), ("Midweek", wed, 60),
                                 ("CHOP", fri, 35)):
                if rng.random() < non_report * 0.6:
                    continue
                rows.append({"Church": c["Location"] if rng.random() < 0.3 else c["Church_Name"],
                             "Service Type": svc, "Date": fmt_date(d, wi),
                             "Men": int(nrng.poisson(mult * 0.36 * size * s)),
                             "Women": int(nrng.poisson(mult * 0.5 * size * s)),
                             "Children": int(nrng.poisson(mult * 0.3 * size * s))})
        write_messy(pd.DataFrame(rows), paths.inbox / f"Community Churches {we:%Y-%m-%d}", rng,
                    sheet="All services", title="COMMUNITY CHURCHES - WEEKLY RETURNS")

        # ---- Transport: the transport office's own report layouts -------------
        write_transport_week(paths, we, wi, n_weeks, areas, zones, fleet, rng, nrng, s)

    return {"areas": len(areas), "zones": len(zones), "cells": len(cells), "chop_zones": len(chop_zones),
            "community": len(community), "fleet": len(fleet), "weeks": [weeks[0], weeks[-1]]}


FT_HEADERS = ["AREA NO", "AREA NAME", "LOADING BAY FULL ADDRESSES FOR LOCATIONS THAT NEEDS ADDITIONAL BUSES",
              "ZINCODE", "NO. OF ADDITIONAL BUSES NEEDED TO COMPLIMENT THE AREA BUSES AVAILABLE AT YOUR AREA FOR "
              "SUNDAY SERVICES", "CAPACITY OF ADDITIONAL BUSES APPROVED (22-SEATER LT)", "ONGOING SERVICE AT ARRIVAL",
              "NUMBER OF BUSES APPROVED", "APPROVED COST", "CAPACITY", "NUMBER OF BUSES IN CHURCH",
              "TOTAL SPENT PER AREA ", "SIGHTED BUSES AMOUNT", "TOTAL SPENT PER AREA ", "MALE", "FEMALE", "CHILDREN",
              "TOTAL", "TOTAL SPENT PER AREA ", "BUS OPTIMIZATION", "UTILIZATION",
              "REMARKS (PAID BEFORE SERVICE/ NOT PAID BEFORE SERVICE)", "TOTAL FULL UTILIZATION", "TOTAL UNDERUTILIZED"]
COASTER_HEADERS = ["S/N", "HUB LOCATIONS", "FT-HUB HOST AREA", "BUS PARKING STATION (AREA FACILITY)",
                   "BUS TAG NUMBER", "BUS", "CAPACITY FULLY UTILIZED", "NUMBER OF TRIPS",
                   "MALE (for multiple trips, pls sum up all trips)", "FEMALE (for multiple trips, pls sum up all trips)",
                   "CHILDREN (for multiple trips, pls sum up all trips)", "TOTAL",
                   "SERVICE ATTENDED (for multiple trips, please select all services)", "OPERATIONAL REMARKS (category)",
                   "OPERATIONAL REMARKS (details here)"]
EV_HEADERS = ["S/N", "HUB LOCATIONS", "FT-HUB HOST AREA", "BUS PARKING STATION (AREA FACILITY)", "MALE", "FEMALE",
              "CHILDREN", "ADULT TOTAL", "TOTAL", "SERVICE ATTENDED", "REMARKS"]
FUEL_HEADERS = ["S/N", "HUB (EXACT ADDRESS) LOCATIONS", "FT-HUB HOST AREA", "HUB STATUS", "TOTAL TRIPS", "AREA NUMBER",
                "BUS SERIAL NUMBER", "COASTER REGISTRATION NUMBER", "MISSION PASTOR NAME",
                "DISTANCE TO CANAANLAND (KM)", "FUEL LITRES PER TO & FRO TRIP", "EXTRA TRIP (NOT USED)",
                "TOTAL PAYABLE FUEL LITRES", "PUMP PRICE (₦/L)", "TOTAL PAYABLE FUEL AMOUNT (₦)"]
SERVICES = ["1ST SERVICE", "2ND SERVICE", "3RD SERVICE", "2ND SERVICE, 3RD SERVICE"]


def _spelling(area: dict, rng) -> str:
    """Report spelling of an Area: usually the master name, sometimes one of its aliases."""
    aliases = [a for a in str(area.get("Aliases", "")).split(";") if a.strip()]
    return rng.choice(aliases) if aliases and rng.random() < 0.3 else area["Area_Name"]


def _messy_zone(code: str, rng) -> str:
    n = code[3:]
    return rng.choice([code, code.lower(), f"LFC {n}", n.lstrip("0"), code])


def write_transport_week(paths, we, wi, n_weeks, areas, zones, fleet, rng, nrng, season):
    title_date = f"{we.day}th Sept, {we.year}" if we.month == 9 else we.strftime("%d %B %Y")
    last = wi == n_weeks - 1
    zones_by_area = zones.groupby("Area_Name")["Zone_ID"].apply(list).to_dict()

    # FT Procured: an Area-total row, then one row per hired bus
    ft_rows, alloc = [], []
    for a in areas.to_dict("records"):
        no = int(a["Area_No"])
        allocated = rng.randint(4, 20)
        cost_bus = rng.choice([18000, 30000, 40000, 45000, 50000, 55000, 60000, 130000])
        alloc.append((no, a["Area_Name"], allocated, cost_bus))
        n_bus = max(1, int(allocated * rng.uniform(0.55, 1.0)))
        buses = []
        for _ in range(n_bus):
            cap = rng.choice([22, 22, 22, 18, 14, 30, 7])
            riders = int(min(cap * 1.1, max(3, nrng.normal(cap * 0.95 * season, 3))))
            m_, f_ = int(riders * 0.38), int(riders * 0.45)
            buses.append([no, _spelling(a, rng), "", _messy_zone(rng.choice(zones_by_area[a["Area_Name"]]), rng),
                          None, cap, None, None, None, cap, 1 if cap >= 14 else None, None,
                          int(cost_bus * cap / 22 / 1000) * 1000, None, m_, f_, riders - m_ - f_, riders, None,
                          min(1, riders / cap), "Full Utilization" if riders >= cap else "Underutilized",
                          None, None, None])
        in_church = sum(1 for b in buses if b[10])
        ft_rows.append([no, a["Area_Name"], "AREA FACILITY", rng.choice(zones_by_area[a["Area_Name"]]), None, None,
                        None, None, None, None, None, in_church, None, sum(b[12] for b in buses), None, None, None, 0,
                        sum(b[17] for b in buses), "#DIV/0!", "#DIV/0!", None, None, None])
        ft_rows.extend(buses)
    if last:
        ft_rows[1][16] = f"{ft_rows[1][16]}$"                 # messy number
        ft_rows.append([99, "NOWHERE AREA", "", "LFC9901", None, 22, None, None, None, 22, 1, None, 40000, None,
                        5, 6, 7, 18, None, None, None, None, None, None])
    ft = pd.DataFrame(ft_rows, columns=FT_HEADERS)

    # Coasters: one row per coaster; a few break down or don't run
    co_rows, fuel_rows = [], []
    for k, v in enumerate(fleet.to_dict("records"), start=1):
        area = areas[areas.Area_Name == v["Area_Name"]].iloc[0].to_dict()
        state = rng.random()
        trips = 0 if state < 0.12 else rng.choice([1, 1, 1, 2])
        riders = [int(nrng.poisson(12 * trips * season)), int(nrng.poisson(14 * trips * season)),
                  int(nrng.poisson(6 * trips * season))] if trips else [None, None, None]
        remark = "Breakdown" if 0.04 < state < 0.12 else ("Extra passengers" if rng.random() < 0.15 else
                                                           "Smooth Operation")
        co_rows.append([k, "HUB BUS STOP", _spelling(area, rng), area["Area_Name"], int(v["Vehicle_ID"]), 1,
                        "YES" if trips else None, trips or None, *riders, sum(r or 0 for r in riders) or None,
                        rng.choice(SERVICES) if trips else None, remark if (trips or remark == "Breakdown") else None,
                        None])
        hub = "HUB" if rng.random() < 0.75 else "ZONE"
        litres_trip = rng.choice([20, 25, 35, 40, 47, 65])
        price = 1215 + (k - len(fleet) + 3 if last and k > len(fleet) - 3 else 0)   # creeping price, last week
        fuel_rows.append([k, "HUB", area["Area_Name"], hub, 1 if trips else 0, int(area["Area_No"]), v["Vehicle_ID"],
                          f"REG{k:03d}", None, litres_trip, litres_trip, litres_trip, litres_trip if trips else 0,
                          price, (litres_trip if trips else 0) * price])
    co = pd.DataFrame(co_rows, columns=COASTER_HEADERS)
    fuel = pd.DataFrame(fuel_rows, columns=FUEL_HEADERS)

    # EV / TATA: area names sit under the "HUB LOCATIONS" header, as in the real sheet
    ev_rows = []
    for k in range(1, max(4, round(len(areas) * 0.3)) + 1):
        a = areas.iloc[k % len(areas)]
        ran = rng.random() > 0.25
        mfc = [int(nrng.poisson(25 * season)), int(nrng.poisson(27 * season)), int(nrng.poisson(10 * season))]
        mfc = mfc if ran else [None, None, None]
        ev_rows.append([k, a["Area_Name"], "BRT BUS STOP", "CANAANLAND" if k % 4 == 0 else "CAPERNAUM", *mfc,
                        (mfc[0] or 0) + (mfc[1] or 0), sum(x or 0 for x in mfc),
                        rng.choice(SERVICES) if ran else None, "SMOOTH OPERATION" if ran else None])
    ev = pd.DataFrame(ev_rows, columns=EV_HEADERS)

    # WSF Procured: every zone listed; only some hired buses
    wsf_rows = [[None] * 10 + ["WSF PR"] + [None] * 4,
                ["WSF PROCURED TRANSPORT REPORT ON ZONAL LOADING BUSES FOR SUNDAY SERVICES"] + [None] * 14,
                [None, None, None, we.strftime("%dTH %B %Y").upper()] + [None] * 11,
                ["ZONAL DETAILS"] + [None] * 6 + ["CARRIAGE"] + [None] * 4 + ["COST", "REMARK", None],
                ["Area No", None, "AREA NAME", None, "ADDRESS LOCATIONS", "TOTAL NO OF BUSES DELIVERED PER ZONE",
                 "TOTAL CAPACITY OF BUSES PER ZONE ²²", "M", "F", "C", "ADULT TOTAL", "TOTAL", None, None, None]]
    for z in zones.to_dict("records"):
        a = areas[areas.Area_Name == z["Area_Name"]].iloc[0]
        if rng.random() < 0.25:
            b = rng.choice([1, 1, 2, 3])
            cap = b * rng.choice([14, 18, 22])
            mm, ff, cc = int(cap * 0.35), int(cap * 0.45), int(cap * 0.2)
            wsf_rows.append([int(a["Area_No"]), None, a["Area_Name"], z["Zone_ID"], None, b, cap, mm, ff, cc,
                             mm + ff, mm + ff + cc, b * 35000, None, None])
        else:
            wsf_rows.append([int(a["Area_No"]), None, a["Area_Name"], z["Zone_ID"], None, None, None, None, None, None,
                             0, 0, None, None, None])
    body = [r for r in wsf_rows[5:] if r[5]]
    wsf_rows.append([None, None, None, 0, None, sum(r[5] for r in body), None, sum(r[7] for r in body),
                     sum(r[8] for r in body), sum(r[9] for r in body), sum(r[10] for r in body),
                     sum(r[11] for r in body), sum(r[12] for r in body), None, None])

    raw_path = paths.inbox / f"Transport raw reports for {we:%d %b %Y}.xlsx"
    with pd.ExcelWriter(raw_path, engine="openpyxl") as xw:
        pd.DataFrame([[f"ELECTRIC BUS - NEW BUSES ({we.strftime('%dTH %B, %Y').upper()})"]]).to_excel(
            xw, sheet_name="TATA EV REPORT ", index=False, header=False)
        ev.to_excel(xw, sheet_name="TATA EV REPORT ", index=False, startrow=1)
        pd.DataFrame([[f"CANAANLAND HUB COASTER BUSES {we.strftime('%dTH %B %Y').upper()}"]]).to_excel(
            xw, sheet_name="FT COASTERS", index=False, header=False)
        co.to_excel(xw, sheet_name="FT COASTERS", index=False, startrow=1)
        pd.DataFrame([[title_date]]).to_excel(xw, sheet_name="FT PROCURED ", index=False, header=False)
        ft.to_excel(xw, sheet_name="FT PROCURED ", index=False, startrow=1)
        pd.DataFrame(wsf_rows).to_excel(xw, sheet_name="WSF PROCURED", index=False, header=False)
    fuel.to_excel(paths.inbox / f"FT COASTERS FUELING FOR HUBS SUNDAY {we.strftime('%dTH %B %Y').upper()}.xlsx",
                  sheet_name="Coaster Fueling Schedule", index=False)
    al = pd.DataFrame([[i + 1, no, name, 22, n, c, n * c, None, None, None, None, None, n * c]
                       for i, (no, name, n, c) in enumerate(alloc)],
                      columns=["S/NO", "AREA NO", "AREA NAME", "MINIMUM CAPACITY OF  BUSES APPROVED (22-SEATER LT)",
                               "NEW TOTAL NUMBER OF BUSES", "APPROVED COST/BUS", "TOTAL EXPECTED SPEND PER WEEK",
                               "Total Spent", "ACCUMULATED BALANCE", "PAYABLE FOR FT PROCURED PREVIOUS", "Total Spent 2",
                               "ACCUMULATED BALANCE 2", f"PAYABLE FOR FT PROCURED SUNDAY {we:%d%m%Y}"])
    al.to_excel(paths.inbox / f"FT PROCURED FOR {we.day}th {we:%b %Y}.xlsx", sheet_name="ACCUMULATED BALANCE",
                index=False)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--root", default="demo")
    ap.add_argument("--weeks", type=int, default=26)
    ap.add_argument("--seed", type=int, default=7)
    ap.add_argument("--scale", type=float, default=1.0, help="fraction of full size (tests use 0.1)")
    args = ap.parse_args(argv)
    info = generate(Path(args.root).resolve(), args.weeks, args.seed, scale=args.scale)
    print(info)


if __name__ == "__main__":
    main()
