"""Build MASTER_ZONES.csv from the ZONE DETAILS workbook (sheets ZONES and NEW ZONES).

The zone code ("zincode", e.g. LFC0112 = Area 01, Zone 12) is the zone's
identity and must be unique.  Rules applied, each logged with its reason:

  * codes are normalised: 'LFC 0701', 'LFCO702', 'lfc0112', 'LFC92016' ...
  * a code whose Area digits disagree with the row's Area number is corrected
    to that Area ('LFC6714' under Area 87 -> 'LFC8714')
  * a missing code is inferred from "ZONE n" when that code is free
  * the same zone listed twice (same code, same Area) is kept once
  * a NEW ZONES code already used by an existing zone is not given to the new
    zone; it is listed for the Area to choose another number
  * codes seen in transport reports but missing from the list are kept and
    marked "Seen in reports only"
  * Has_CHOP flags already in MASTER_ZONES.csv are kept
"""
from __future__ import annotations

import re

import pandas as pd

from databank import Masters, _similar, norm_key
from transport_reports import AreaMatcher, blank, hkey, zone_code

SIGNATURE_EXISTING = ("zincodes",)
SIGNATURE_NEW = ("proposedzonalno",)
COLUMNS = ["Zone_ID", "Zone_Name", "Area_Name", "Has_CHOP", "Active", "Zone_No", "Address", "Zone_Status",
           "Is_Community"]


def is_zone_details(raw: pd.DataFrame) -> str | None:
    for r in range(min(5, len(raw))):
        keys = [hkey(v) for v in raw.iloc[r].tolist()]
        # the zone list has a plain "Zone" column; transport reports with a ZINCODE column don't
        if any(k.startswith("zincode") for k in keys) and "zone" in keys and "areaname" in keys:
            return "ZONES"
        if any(k.startswith("proposedzonalno") for k in keys):
            return "NEW ZONES"
    return None


def _split_zone_text(text) -> tuple[str, str]:
    """'Zone 12 - 22 JELIL AREMU STREET' -> ('12', '22 JELIL AREMU STREET')."""
    s = "" if blank(text) else str(text).strip()
    m = re.match(r"(?i)^\s*zone\s*(\d+)\s*[-–:.,]*\s*(.*)$", s)
    if m:
        return m.group(1), m.group(2).strip()
    return "", s


def _rows(raw: pd.DataFrame, kind: str, match: AreaMatcher):
    header = [hkey(v) for v in raw.iloc[0].tolist()]
    col = lambda *keys: next(i for i, h in enumerate(header) if any(h.startswith(k) for k in keys))
    c_no, c_area = col("areano", "areanumber"), col("areaname")
    c_zone = col("zone", "zonaladdress")
    c_code = col("zincode", "proposedzonalno")
    for i in range(1, len(raw)):
        r = raw.iloc[i].tolist()
        if all(blank(v) for v in r):
            continue
        yield dict(row=i + 1, kind=kind, area_raw=r[c_area], area_no=r[c_no], zone_text=r[c_zone],
                   code_raw=r[c_code], area=match(r[c_area], r[c_no]))


def build(sheets: dict[str, pd.DataFrame], masters: Masters, existing: pd.DataFrame, seen_codes: dict[str, str]):
    """Return (master dataframe, issues).  `seen_codes` = {zone code: Area} from loaded reports."""
    match = AreaMatcher(masters)
    issues = []
    note = lambda sev, issue, detail, row=None, sheet="": issues.append(
        {"severity": sev, "issue": issue, "detail": str(detail)[:250], "source_row": row, "sheet": sheet})
    records = []
    for sheet, raw in sheets.items():
        kind = is_zone_details(raw)
        if kind:
            records.extend(_rows(raw, kind, match))
    zones: dict[str, dict] = {}
    pending_new, missing_code = [], []
    # existing zones first, then new ones (so an existing zone keeps a contested code)
    for rec in sorted(records, key=lambda x: x["kind"] != "ZONES"):
        sheet = rec["kind"]
        if rec["area"] is None:
            note("Rejected", "Unknown Area", f"'{rec['area_raw']}' / no. {rec['area_no']}", rec["row"], sheet)
            continue
        area_no = match.number_of(rec["area"])
        zone_no, address = _split_zone_text(rec["zone_text"])
        code = zone_code(rec["code_raw"])
        if code and area_no and int(code[3:5]) != area_no:
            fixed = f"LFC{area_no:02d}{code[5:]}"
            note("Warning", "Zone code corrected to its Area",
                 f"{rec['code_raw']} listed under {rec['area']} (Area {area_no}); read as {fixed}", rec["row"], sheet)
            code = fixed
        if code is None and zone_no and area_no:
            guess = f"LFC{area_no:02d}{int(zone_no):02d}"
            missing_code.append((rec, guess, zone_no, address))
            continue
        if code is None:
            note("Warning", "Zone without a code", f"{rec['area']}: '{rec['zone_text']}' has no zincode; not added",
                 rec["row"], sheet)
            continue
        entry = {"Zone_ID": code, "Zone_Name": code, "Area_Name": rec["area"], "Has_CHOP": "", "Active": "Y",
                 "Zone_No": zone_no or str(int(code[5:])), "Address": address,
                 "Zone_Status": "Existing" if rec["kind"] == "ZONES" else "New (proposed)"}
        if code in zones:
            first = zones[code]
            if first["Area_Name"] == entry["Area_Name"] and rec["kind"] == "ZONES" and \
                    first["Zone_Status"] == "Existing":
                note("Info", "Zone listed twice", f"{code} ({rec['area']}) appears again; first entry kept",
                     rec["row"], sheet)
            elif rec["kind"] == "NEW ZONES" and first["Zone_Status"] == "Existing" and \
                    _similar(norm_key(address), norm_key(first["Address"])) >= 0.7:
                note("Info", "New zone already listed",
                     f"'{address}' ({code}) is already in ZONES as '{first['Address']}'", rec["row"], sheet)
            elif rec["kind"] == "NEW ZONES" and first["Zone_Status"] == "Existing":
                pending_new.append((rec, code, address))
                note("Warning", "Proposed code already in use",
                     f"New zone '{address}' ({rec['area']}) proposed {code}, which is already "
                     f"{first['Area_Name']} zone '{first['Address']}'. Please give it another number.",
                     rec["row"], sheet)
            else:
                note("Warning", "Same code on two zones",
                     f"{code}: '{first['Address']}' and '{address}' ({rec['area']}); first kept", rec["row"], sheet)
            continue
        zones[code] = entry
    for rec, guess, zone_no, address in missing_code:
        if guess in zones:
            note("Warning", "Zone without a code", f"{rec['area']} zone {zone_no} '{address}' has no zincode and "
                                                   f"{guess} is taken; not added", rec["row"], rec["kind"])
            continue
        zones[guess] = {"Zone_ID": guess, "Zone_Name": guess, "Area_Name": rec["area"], "Has_CHOP": "",
                        "Active": "Y", "Zone_No": zone_no, "Address": address, "Zone_Status": "Existing"}
        note("Info", "Zone code inferred", f"{rec['area']} zone {zone_no} '{address}' had no zincode; "
                                           f"given {guess} from its zone number", rec["row"], rec["kind"])
    for code, area in sorted(seen_codes.items()):
        if code not in zones and area:
            zones[code] = {"Zone_ID": code, "Zone_Name": code, "Area_Name": area, "Has_CHOP": "", "Active": "Y",
                           "Zone_No": str(int(code[5:])), "Address": "", "Zone_Status": "Seen in reports only"}
            note("Warning", "Zone code not in ZONE DETAILS",
                 f"{code} ({area}) appears in transport reports but not in the zone list; kept")
    df = pd.DataFrame(zones.values(), columns=COLUMNS).fillna("")
    # keep Has_CHOP / Active choices already made in the master
    if len(existing):
        prev = existing.set_index("Zone_ID")
        for c in ("Has_CHOP", "Active", "Is_Community"):
            if c in prev:
                kept = df["Zone_ID"].map(prev[c]).fillna("")
                df[c] = kept.where(kept != "", df[c])
    df["_a"] = df["Zone_ID"].str[3:5].astype(int)
    df["_z"] = df["Zone_ID"].str[5:].astype(int)
    return df.sort_values(["_a", "_z"]).drop(columns=["_a", "_z"]).reset_index(drop=True), issues


def apply_export_flags(master: pd.DataFrame, flags: pd.DataFrame, stream: str):
    """Refresh MASTER_ZONES from a reporting-system export: add zones it doesn't have, and take the
    CHOP / Community flags from the system (the CHOP export itself leaves forChop blank, so it is ignored)."""
    issues = []
    df = master.copy()
    for c in COLUMNS:
        if c not in df:
            df[c] = ""
    df = df[COLUMNS].fillna("")
    known = set(df.Zone_ID)
    new = flags[~flags.zone.isin(known) & ~flags.zone.str.startswith("AREA")]
    if len(new):
        add = pd.DataFrame({"Zone_ID": new.zone, "Zone_Name": new.zone, "Area_Name": new.area, "Has_CHOP": "",
                            "Active": "Y", "Zone_No": new.zone.str[5:].astype(int).astype(str),
                            "Address": new.zone_name, "Zone_Status": "From reporting system", "Is_Community": ""})
        df = pd.concat([df, add], ignore_index=True)
        issues.append({"severity": "Info", "issue": "Zones added from the reporting system",
                       "detail": f"{len(new)} zones not in ZONE DETAILS: " + ", ".join(new.zone.head(15)) +
                                 (" ..." if len(new) > 15 else "")})
    f = flags.set_index("zone")
    if stream != "CHOP":
        chop = df.Zone_ID.map(f.for_chop)
        df.loc[chop.notna(), "Has_CHOP"] = chop[chop.notna()].map({True: "Y", False: "N"})
    comm = df.Zone_ID.map(f.for_community)
    df.loc[comm == True, "Is_Community"] = "Y"   # noqa: E712
    df.loc[(comm == False) & (df.Is_Community == ""), "Is_Community"] = "N"   # noqa: E712
    df["_a"] = df.Zone_ID.str[3:5].astype(int)
    df["_z"] = df.Zone_ID.str[5:].astype(int)
    return df.sort_values(["_a", "_z"]).drop(columns=["_a", "_z"]).reset_index(drop=True), issues


def refresh_community(community: pd.DataFrame, rows: pd.DataFrame) -> pd.DataFrame:
    """MASTER_COMMUNITY from the Community Church Sunday workbook: one church per zincode."""
    cols = ["Church_ID", "Church_Name", "Location", "Active"]
    have = community.copy() if len(community) else pd.DataFrame(columns=cols)
    latest = rows.sort_values("report_date").groupby("church").last().reset_index()
    for _, r in latest.iterrows():
        name = r.location.title() if r.location else r.church
        if r.church in set(have.Church_ID):
            have.loc[have.Church_ID == r.church, ["Church_Name", "Location"]] = [name, r.area or ""]
        else:
            have = pd.concat([have, pd.DataFrame([[r.church, name, r.area or "", "Y"]], columns=cols)],
                             ignore_index=True)
    return have[cols].sort_values("Church_ID").reset_index(drop=True)
