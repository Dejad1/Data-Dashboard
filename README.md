# Operations Reporting & Analysis Dashboard

An Excel databank and dashboard for five weekly report streams: WSF, Midweek, CHOP, Community Church and Transport. Python scripts load the weekly files and rebuild the workbook, so nobody types data into the workbook by hand.

```
inbox/ (weekly files)  ──ingest.py──►  databank.sqlite  ──build_workbook.py──►  Operations_Dashboard.xlsx
                             │                                                    (ingest.py runs this for you)
                             └── rejects ─► DATA QUALITY sheet
```

## Status

| Stage | Status |
|---|---|
| Masters, Calendar, SETTINGS | Done. The real 92 Areas and a draft of 1,164 zone codes are in `masters/` |
| Data tables and ingestion (alias map, de-duplication, data-quality log) | Done |
| **WSF DASHBOARD** | Done, ready for review |
| **TRANSPORT DASHBOARD** | Done, built on the transport office's real reports (from 27 Sept 2026) |
| Midweek, CHOP, Community dashboards; Area and Zone scorecards | Next, as their sample files arrive. Their data already loads into the `DATA_*` sheets. |
| Master dashboard | After the stream dashboards |

## One-time setup

1. Install Python 3.10 or later, then run `pip install openpyxl pandas pyyaml`.
2. Fill in the five master lists in `masters/`. You can open the CSV files in Excel.

   | File | Columns |
   |---|---|
   | `MASTER_AREAS.csv` | Area_ID, Area_Name, Active (Y/N), Area_No, Aliases (other spellings, separated by `;`) |
   | `MASTER_ZONES.csv` | Zone_ID, Zone_Name, Area_Name, Has_CHOP (Y/N), Active (Y/N). Zones use their codes, e.g. `LFC0112` = Area 01, Zone 12 |
   | `MASTER_CELLS.csv` | Cell, Zone_Name, Operational (Y/N) |
   | `MASTER_COMMUNITY.csv` | Church_ID, Church_Name, Location, Active (Y/N) |
   | `MASTER_FLEET.csv` | Vehicle_ID, Category, Capacity, Owner, Status, Area_Name, Zone_Name, Hired_By_Type (reference list only; transport is loaded from the weekly reports) |

   Area names are matched ignoring case, spaces and punctuation, and through the Aliases column. "IYANA- ODO", "Eti -Osa" and "PETRO" all find the right Area.
3. Check `settings.yaml`. `week_end_day: 7` means weeks end on Sunday; 6 means Saturday.

## Weekly routine (5 steps)

1. **Collect** the week's files: WSF, Midweek, CHOP and Community Church, plus the transport office's reports. Any layout works:
   - the header row can sit below title rows;
   - columns can be in any order;
   - `.xlsx` and `.csv` are both accepted.
   Transport reports are read exactly as the transport office sends them:
   - the raw reports workbook (FT Procured, FT Coasters, TATA EV and WSF Procured sheets);
   - the coaster fuelling schedule;
   - the FT Procured allocation file.

   The report date is read from each sheet's title or the file name.
2. **Drop** them into the `inbox/` folder. Put "CHOP" or "Midweek" in the file or sheet name, because those two streams share the same columns.
3. **Run** `python ingest.py` (double-click `ingest.py` on Windows if Python is associated). It prints how many rows were loaded, replaced and rejected per stream, and rebuilds `Operations_Dashboard.xlsx`.
4. **Check** the DATA QUALITY sheet. Fix any rejected rows in the source file, then drop the fixed file in the inbox again. A re-sent week replaces that week's rows; it never double counts.
5. **Open** the workbook and use HOME to reach a dashboard. The week-ending box defaults to "Latest", so it follows the new data.

Processed files move to `inbox/processed/`. Files that couldn't be recognised go to `inbox/unrecognised/`.

## Adding things

| To add | Do this, then run `python ingest.py` |
|---|---|
| a new **Area** | Add a row to `MASTER_AREAS.csv`. |
| a new **Zone** | Add a row to `MASTER_ZONES.csv` with its Area_Name, Has_CHOP and Active = Y. |
| a new **Cell** | Add a row to `MASTER_CELLS.csv` with its Zone_Name and Operational = Y. Set Operational = N to take a cell out of the reporting universe without deleting its history. |
| a new **CHOP zone** | Set Has_CHOP = Y on that zone in `MASTER_ZONES.csv`. CHOP rows from zones with Has_CHOP = N are rejected. |
| a new **bus** | Nothing to do. Buses are read from the weekly transport reports. To change an Area's FT allocation or approved cost per bus, update the FT Procured allocation file and send it with that week's reports. |
| a new **header spelling** | Add it to `column_aliases.yaml` under the right field. |

Thresholds (green/amber/red, utilisation limit, default window, seat capacities) are on the hidden SETTINGS sheet: right-click any tab, then Unhide. Values you change there are kept when the workbook is rebuilt.

## Using the dashboards

Every dashboard has the same control strip, frozen at the top:

| Control | What it does |
|---|---|
| SCOPE | Global, Area or Zone. |
| SEARCH | Type part of a name, for example `ikeja`. |
| SELECT | A drop-down that lists only the Areas or Zones matching SEARCH. |
| WEEK ENDING | "Latest" or any week that has data. |
| COMPARE (WEEKS) | The comparison window: 1, 2, 4, 8, 13, 26 or 52 weeks. |

The WSF DASHBOARD has these sections:

- **KPI cards**:
  - grand total;
  - adults (Male + Female);
  - children;
  - % cells reporting (green/amber/red);
  - cells reported;
  - average attendance per reporting cell;
  - zones where no cell reported.

  Each card shows the change against the previous week (▲/▼) and the N-week average.
- **Time analysis** for every KPI:
  - selected week and previous week;
  - change and % change;
  - average of the N weeks before the selected week, and the selected week against that average;
  - MTD, QTD and YTD totals;
  - weekly averages for the month, quarter and year.
- **Trends**:
  - attendance mix (Male, Female, Children) over 13 weeks;
  - grand total over 52 weeks;
  - % cells reporting over 13 weeks.
- **Area rankings**: Top and Bottom 10 by attendance, by % reporting and by week-on-week growth, plus the selected Area's rank. These show "n/a" at Zone scope.
- **Who didn't report**: cells with no WSF return for the selected week, filtered by the current scope.

A "–" means there is no data for that figure. A real zero shows as 0.

## Transport

The TRANSPORT DASHBOARD reads four bus categories:

| Category | Source | Who pays | Cost used |
|---|---|---|---|
| FT Procured (hired LT, 22-seater; sometimes smaller or bigger when no LT was available) | FT PROCURED sheet, one row per bus | Central | "SIGHTED BUSES AMOUNT" per bus |
| Church Coaster | FT COASTERS sheet, one row per coaster | Central pays fuel | Coaster fuelling schedule, **HUB coasters only** |
| EV/TATA (70 seats; TATA park at Canaanland) | TATA EV REPORT sheet (and EV BRT rows on the coaster sheet) | Central, charged directly | None: no EV/TATA cost is reported |
| Hiace (31 buses) | Weekly report from October 2026 | – | Reader to be added when the first report arrives |
| WSF Procured | WSF PROCURED sheet, one row per zone | Members | COST column |

**Where buses load.** TATA, EV and most coasters load at **Hubs**, the major express points where long buses can stop. FT Procured and WSF Procured load at **Loading Bays**, which can be down small streets. A coaster marked **ZONE** on the fuelling schedule also counts as a Loading Bay. The dashboard compares Hubs and Loading Bays side by side.

Rules and checks:
- **Utilisation** = riders seated ÷ seats offered, counting only buses whose capacity is known. A full bus counts as 100%.
- **Riders beyond seats** = riders above a bus's seats, counted separately as unmet demand: the signal to add buses.
- **FT budget** = buses allocated × approved cost per bus, from the allocation file.
- **FT bus count.** Every FT bus row counts as a bus. A blank "NUMBER OF BUSES IN CHURCH" only means the loading-bay line wasn't filled in.
- **Cross-checks.** Each FT Area-total row, each sheet's grand total and the coaster **Hub Payment Summary** are compared with the rows they summarise. Differences go to DATA QUALITY, for example a total that skipped a number typed as text.
- **Number clean-up.** Messy numbers such as `12$` and `170-,000` are cleaned and logged.
- **Area boundary rows.** When a row carries the neighbouring Area's name, the zone code decides which Area it belongs to.
- **Optimisation panel:**
  - most riders beyond seats (unmet demand);
  - lowest utilisation;
  - highest cost per rider;
  - FT spend over budget;
  - plain-language flags, such as "under 50% full for 4 weeks running";
  - every breakdown, "extra passengers" and "need more buses" note.

## Design notes

- **WSF grain.** `DATA_WSF` holds one row per zone per week, about 65,000 rows a year. That keeps it well under Excel's 1,048,576-row limit for about 16 years.
  - Cell-level returns are aggregated by `ingest.py`.
  - Full cell history stays in `databank.sqlite` and in `archive/WSF_cells_YYYY.csv`.
  - `WSF_CELLS_CURRENT` holds only the latest 8 weeks, and feeds the "who didn't report" list.
  - Every active zone gets a row each week, so a zone where no cell reported shows `Cells_Reported = 0` instead of disappearing.
- **The workbook is rebuilt, not appended to.** openpyxl can't safely append to a workbook that has charts, because it drops them. So data is appended to `databank.sqlite` and the workbook is regenerated, which takes about 2 minutes at full scale. Each data sheet is still one continuous Excel Table.
- **Totals are computed, not trusted.** Adult_Total, Grand_Total, Week_Ending, Cells_Not_Reported and Seats_Offered are formulas. Incoming totals are only compared with the computed ones, and any mismatch goes to DATA QUALITY.
- **No dynamic arrays, on purpose.** Files written by Python have no "spill" metadata, so FILTER, SORT, UNIQUE and XLOOKUP would open as `@FILTER(...)` and show only one value. The search drop-down therefore uses a helper column with `ISNUMBER(SEARCH())`, `SMALL()` to list the matches, and a match-sized defined name as the validation list. It behaves the same, and it works in any Excel version and in LibreOffice.
- **Performance.** Dashboards read only from `CALC_*` sheets:
  - about 500 SUMIFS for the weekly series;
  - about 550 for the Area rankings;
  - no volatile functions except two tiny OFFSET names used as drop-down sources.

  At full scale with 26 weeks of synthetic data, the workbook has about 940,000 formulas (mostly simple row formulas in the data tables) and zero errors. LibreOffice takes about 60 seconds to recalculate it, or about 80 seconds including load and save. Excel is normally much faster, but that has not been timed here.
- **Settings to confirm:**
  - weeks ending Sunday;
  - Coasters at 30 seats;
  - cost types Fuel, Hire Fee, Maintenance, Driver Allowance, Member Payment and Other;
  - currency NGN.

## Files

| Path | Purpose |
|---|---|
| `ingest.py` | Loads the inbox into the databank, logs issues and rebuilds the workbook. |
| `build_workbook.py` + `builder/` | Workbook generator (`data.py` for tables/masters/calendar/DQ, `controls.py` for the filter strip, `wsf.py` for CALC_WSF and the dashboard, `home.py` for navigation). |
| `databank.py` | SQLite storage with natural keys, and the master-list loader. |
| `column_aliases.yaml` | Header spellings for each field. Editable. |
| `templates/` | One blank input template per stream (`make_templates.py` regenerates them). |
| `generate_sample_data.py` | Full-scale synthetic data: 92 Areas, about 1,250 Zones, about 16,000 Cells, 35 churches and the full fleet, with injected errors. |
| `tests/` | End-to-end tests (`python -m pytest tests -q`) and `lo_tools.py`, which recalculates a workbook in LibreOffice and scans it for errors. |
| `output/Operations_Dashboard_EMPTY.xlsx` | The empty-data workbook. |

To make the demo workbook:

```
python generate_sample_data.py --root demo --weeks 26
python ingest.py --root demo          # -> demo/Operations_Dashboard.xlsx
```
