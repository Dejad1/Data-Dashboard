"""Shared configuration for the Operations Dashboard toolkit.

Everything that more than one script needs lives here: folder layout, stream
definitions, fleet defaults and the stream colour palette.  Operational
settings that staff may change (week-ending day, thresholds) are read from
``settings.yaml`` in the data root so no code edit is needed.
"""
from __future__ import annotations

import datetime as dt
from dataclasses import dataclass
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent

# --------------------------------------------------------------------------
# Folder layout (relative to a data root; the default root is this folder)
# --------------------------------------------------------------------------


@dataclass(frozen=True)
class Paths:
    root: Path

    @property
    def inbox(self) -> Path:
        return self.root / "inbox"

    @property
    def processed(self) -> Path:
        return self.root / "inbox" / "processed"

    @property
    def rejected_files(self) -> Path:
        return self.root / "inbox" / "unrecognised"

    @property
    def archive(self) -> Path:
        return self.root / "archive"

    @property
    def masters(self) -> Path:
        return self.root / "masters"

    @property
    def databank(self) -> Path:
        return self.root / "databank.sqlite"

    @property
    def workbook(self) -> Path:
        return self.root / "Operations_Dashboard.xlsx"

    @property
    def settings(self) -> Path:
        return self.root / "settings.yaml"

    def ensure(self) -> None:
        for p in (self.inbox, self.processed, self.rejected_files, self.archive, self.masters):
            p.mkdir(parents=True, exist_ok=True)


ALIASES_FILE = HERE / "column_aliases.yaml"

# --------------------------------------------------------------------------
# Settings (settings.yaml in the data root, created with defaults if missing)
# --------------------------------------------------------------------------

DEFAULT_SETTINGS = {
    # 1 = Monday ... 7 = Sunday (ISO numbering, same as Excel WEEKDAY(d, 2))
    "week_end_day": 7,
    "currency_symbol": "₦",
    # Weeks of cell-level WSF detail kept inside the workbook
    "wsf_cells_weeks_in_workbook": 8,
    # First week-ending date the Calendar sheet covers, and how many years
    "calendar_start_year": 2024,
    "calendar_years": 7,
}


def load_settings(paths: Paths) -> dict:
    settings = dict(DEFAULT_SETTINGS)
    if paths.settings.exists():
        with open(paths.settings, encoding="utf-8") as fh:
            settings.update(yaml.safe_load(fh) or {})
    else:
        paths.root.mkdir(parents=True, exist_ok=True)
        with open(paths.settings, "w", encoding="utf-8") as fh:
            fh.write("# Operations Dashboard settings used by ingest.py and build_workbook.py\n")
            fh.write("# week_end_day: 1=Monday ... 6=Saturday, 7=Sunday\n")
            yaml.safe_dump(settings, fh, sort_keys=False, allow_unicode=True)
    return settings


def week_ending(d: dt.date, week_end_day: int) -> dt.date:
    """Same rule as the workbook formula: date + MOD(WeekEndDay - WEEKDAY(date, 2), 7)."""
    return d + dt.timedelta(days=(week_end_day - d.isoweekday()) % 7)


# --------------------------------------------------------------------------
# Streams
# --------------------------------------------------------------------------

STREAMS = ["WSF", "MIDWEEK", "CHOP", "COMMUNITY", "TRANSPORT_OPS", "TRANSPORT_FINANCE"]

STREAM_LABELS = {
    "WSF": "WSF",
    "MIDWEEK": "Midweek",
    "CHOP": "CHOP",
    "COMMUNITY": "Community Church",
    "TRANSPORT_OPS": "Transport (operations)",
    "TRANSPORT_FINANCE": "Transport (finance)",
}

COMMUNITY_SERVICE_TYPES = ["Sunday", "WSF", "Midweek", "CHOP"]

# One colour per stream, used everywhere (hex without '#')
COLOURS = {
    "MASTER": "1B2A41",
    "WSF": "1F5FA8",
    "MIDWEEK": "0F7C74",
    "CHOP": "6A3FA0",
    "COMMUNITY": "B5446E",
    "TRANSPORT": "D9731A",
    "QUALITY": "5B6770",
}

# --------------------------------------------------------------------------
# Transport
# --------------------------------------------------------------------------

FLEET_CATEGORIES = {
    # category: (default capacity per trip, owner, who pays)
    "FT Procured": (22, "Central", "Central"),
    "Church Coaster": (30, "Church", "Central (fuel)"),
    "Electric Bus": (70, "Church", "Church"),
    "Big Bus": (70, "Church", "Church"),
    "WSF Procured": (None, "Hired", "Members"),
}

HIRED_BY_TYPES = ["Area", "Zone", "Individual"]

COST_TYPES = ["Fuel", "Hire Fee", "Maintenance", "Driver Allowance", "Member Payment", "Other"]

PAID_BY = ["Central", "Church", "Area", "Zone", "Members", "Individual"]
