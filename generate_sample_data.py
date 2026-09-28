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

from config import FLEET_CATEGORIES, Paths, load_settings, week_ending

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


def build_masters(rng: random.Random, scale: float = 1.0):
    n_areas = max(3, round(92 * scale))
    target_zones, target_cells = 1250 * scale, 16000 * scale
    areas = [{"Area_ID": f"A{i + 1:03d}", "Area_Name": AREA_PLACES[i].replace("-", " "), "Active": "Y"}
             for i in range(n_areas)]
    # spread zones unevenly across areas (big areas have more zones)
    weights = np.array([rng.uniform(0.5, 1.8) for _ in areas])
    per_area = np.maximum(5, np.round(weights / weights.sum() * target_zones)).astype(int)
    zones = []
    for a, n in zip(areas, per_area):
        for z in range(n):
            zones.append({"Zone_ID": f"{a['Area_ID']}-Z{z + 1:02d}",
                          "Zone_Name": f"{a['Area_Name']} Zone {z + 1:02d}",
                          "Area_Name": a["Area_Name"],
                          "Has_CHOP": "Y" if rng.random() < 0.6 else "N",
                          "Active": "Y"})
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
    counts = {k: max(2, round(v * scale)) for k, v in
              {"FT Procured": 1200, "Church Coaster": 216, "Electric Bus": 50, "Big Bus": 20,
               "WSF Procured": 160}.items()}
    prefix = {"FT Procured": "LT", "Church Coaster": "CST", "Electric Bus": "EB", "Big Bus": "BB",
              "WSF Procured": "WP"}
    for cat, n in counts.items():
        cap, owner, _ = FLEET_CATEGORIES[cat]
        for i in range(n):
            z = rng.choice(zones)
            hired = rng.choice(["Area", "Zone", "Individual"]) if cat == "WSF Procured" else ""
            fleet.append({"Vehicle_ID": f"{prefix[cat]}-{i + 1:04d}", "Category": cat,
                          "Capacity": cap if cap else rng.choice([14, 18, 26, 30, 33]),
                          "Owner": owner, "Status": "Active" if rng.random() > 0.03 else "Retired",
                          "Area_Name": z["Area_Name"], "Zone_Name": "" if hired == "Area" else z["Zone_Name"],
                          "Hired_By_Type": hired})
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
    active_fleet = fleet[fleet["Status"] == "Active"].reset_index(drop=True)
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

        # ---- Transport operations --------------------------------------------
        f = active_fleet
        op = nrng.random(len(f)) > np.where(f["Category"] == "Church Coaster", 0.12, 0.07)
        trips = np.where(op, nrng.integers(1, 5, len(f)) * np.where(f["Category"] == "WSF Procured", 1, 2), 0)
        util = np.clip(nrng.normal(0.72, 0.18, len(f)), 0.15, 1.15)
        util = np.where(f["Category"] == "Church Coaster", util * 0.8, util)
        ridership = np.round(f["Capacity"].astype(int).to_numpy() * trips * util).astype(int)
        ops = pd.DataFrame({"Week Ending": fmt_date(we, wi), "Bus ID": f["Vehicle_ID"], "Bus Type": f["Category"],
                            "Area": f["Area_Name"], "Zone": f["Zone_Name"], "Trips": trips,
                            "Riders Carried": ridership, "Operational": np.where(op, "Yes", "No")})
        if wi == n_weeks - 1:
            ops.loc[0, "Bus ID"] = "LT-9999"                                    # unknown vehicle
        write_messy(ops, paths.inbox / f"Transport ops {we:%Y-%m-%d}", rng, sheet="Fleet")

        # ---- Transport finance --------------------------------------------------
        fin = []
        by_area = f.assign(op=op, trips=trips, riders=ridership).groupby(["Area_Name", "Category"])
        for (area, cat), g in by_area:
            buses, t = int(g["op"].sum()), int(g["trips"].sum())
            if buses == 0:
                continue
            if cat in ("FT Procured", "Church Coaster", "Electric Bus", "Big Bus"):
                fuel_rate = {"FT Procured": 9500, "Church Coaster": 14000, "Electric Bus": 2500, "Big Bus": 26000}[cat]
                fin.append((cat, area, "Fuel", t * fuel_rate * rng.uniform(0.9, 1.1),
                            "Central" if cat in ("FT Procured", "Church Coaster") else "Church"))
                fin.append((cat, area, "Driver Allowance", buses * 5000, "Central" if cat == "FT Procured" else "Church"))
                if rng.random() < 0.35:
                    fin.append((cat, area, "Maintenance", buses * rng.uniform(8000, 60000), "Church"))
            else:
                fee = int(g["riders"].sum()) * rng.uniform(300, 500)
                fin.append((cat, area, "Hire Fee", fee, "Members"))
                fin.append((cat, area, "Member Payment", fee * rng.uniform(0.92, 1.0), "Members"))
        fdf = pd.DataFrame(fin, columns=["Category", "Area", "Cost Type", "Amount (NGN)", "Paid By"])
        fdf["Amount (NGN)"] = fdf["Amount (NGN)"].round(0)
        fdf.insert(0, "Date", fmt_date(sun, wi))
        fdf.insert(3, "Zone", "")
        write_messy(fdf, paths.inbox / f"Transport finance {we:%Y-%m-%d}", rng, sheet="Expenses",
                    title="TRANSPORT EXPENDITURE" if wi % 2 else None)

    return {"areas": len(areas), "zones": len(zones), "cells": len(cells), "chop_zones": len(chop_zones),
            "community": len(community), "fleet": len(fleet), "weeks": [weeks[0], weeks[-1]]}


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
