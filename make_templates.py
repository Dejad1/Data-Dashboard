"""Write one blank input template per stream into templates/.

    python make_templates.py

Each template has the exact headers ingest.py expects on the first sheet and a
"How to fill" sheet with a legend and one example row.  (Other layouts also
work: ingest.py matches columns by name, in any order, below title rows.)
"""
from __future__ import annotations

from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HERE = Path(__file__).resolve().parent

TEMPLATES = {
    "WSF_cell_returns": ("WSF", ["Date", "Area", "Zone", "Cell", "Male", "Female", "Children"],
                         ["27/09/2026", "Ikeja", "Ikeja Zone 03", "A001-Z03-C07", 5, 8, 4],
                         "One row per cell that met. A cell with no row counts as 'not reported'."),
    "WSF_zone_totals": ("WSF", ["Date", "Area", "Zone", "Cells_Total", "Cells_Reported", "Male", "Female", "Children"],
                        ["27/09/2026", "Ikeja", "Ikeja Zone 03", 14, 12, 61, 88, 40],
                        "Only if a zone sends totals instead of cell rows. Cell rows are preferred."),
    "Midweek": ("Midweek", ["Date", "Area", "Zone", "Male", "Female", "Children"],
                ["23/09/2026", "Ikeja", "Ikeja Zone 03", 40, 57, 22],
                "One row per zone. Every zone holds Midweek."),
    "CHOP": ("CHOP", ["Date", "Area", "Zone", "Male", "Female", "Children"],
             ["25/09/2026", "Ikeja", "Ikeja Zone 03", 18, 30, 6],
             "One row per CHOP zone. Rows from zones without CHOP are rejected."),
    "Community_Church": ("Community", ["Date", "Church", "Service_Type", "Male", "Female", "Children"],
                         ["28/09/2026", "Community Church Abuja", "Sunday", 150, 210, 95],
                         "Service_Type is Sunday, WSF, Midweek or CHOP. Never mixed into zone totals."),
    "Transport_Operations": ("Transport", ["Date", "Vehicle_ID", "Category", "Area", "Zone", "Hired_By_Type", "Trips",
                                           "Ridership", "Operational"],
                             ["27/09/2026", "LT-0001", "FT Procured", "Ikeja", "Ikeja Zone 03", "", 4, 70, "Y"],
                             "One row per vehicle per week. Category/Area/Zone default from MASTER_FLEET if blank. "
                             "Hired_By_Type (Area, Zone or Individual) is for WSF Procured vehicles."),
    "Transport_Finance": ("Transport finance", ["Date", "Category", "Area", "Zone", "Cost_Type", "Amount", "Paid_By"],
                          ["27/09/2026", "Church Coaster", "Ikeja", "", "Fuel", 145000, "Central"],
                          "Cost_Type: Fuel, Hire Fee, Maintenance, Driver Allowance, Member Payment, Other. "
                          "Leave Area blank for central costs. Amount in NGN."),
}


def main():
    out = HERE / "templates"
    out.mkdir(exist_ok=True)
    head_font = Font(name="Arial", bold=True, color="FFFFFF")
    head_fill = PatternFill("solid", start_color="1B2A41", end_color="1B2A41")
    for name, (sheet, headers, example, note) in TEMPLATES.items():
        wb = Workbook()
        ws = wb.active
        ws.title = sheet
        ws.append(headers)
        for c in ws[1]:
            c.font, c.fill = head_font, head_fill
        for i, h in enumerate(headers, start=1):
            ws.column_dimensions[ws.cell(1, i).column_letter].width = max(12, len(h) + 4)
        ws.freeze_panes = "A2"
        info = wb.create_sheet("How to fill")
        info["A1"] = f"{name.replace('_', ' ')} template"
        info["A1"].font = Font(name="Arial", bold=True, size=14)
        lines = [
            "Fill the first sheet: one row per unit per week, starting on row 2. Don't rename the headers.",
            "Date: the service date (dd/mm/yyyy). The system works out the week ending itself.",
            "Names must match the master lists (Area, Zone, Cell, Church, Vehicle). Case and extra spaces don't matter.",
            "Male, Female, Children: whole numbers, 0 or more. Totals are calculated by the system; "
            "if you add a total column it is only checked.",
            note,
            "Save the file into the inbox folder and run ingest.py. Re-sending a week replaces that week's rows.",
        ]
        for i, text in enumerate(lines, start=3):
            info[f"A{i}"] = f"• {text}"
            info[f"A{i}"].alignment = Alignment(wrap_text=True, vertical="top")
            info[f"A{i}"].font = Font(name="Arial", size=10)
        r = len(lines) + 4
        info[f"A{r}"] = "Example row (do not copy into the data sheet unless the names are real):"
        info[f"A{r}"].font = Font(name="Arial", bold=True, size=10)
        for j, (h, v) in enumerate(zip(headers, example), start=1):
            info.cell(r + 1, j, h).font = Font(name="Arial", bold=True, size=9)
            info.cell(r + 2, j, v).font = Font(name="Arial", size=9, color="0000FF")
        info.column_dimensions["A"].width = 90
        wb.save(out / f"TEMPLATE_{name}.xlsx")
        print(out / f"TEMPLATE_{name}.xlsx")


if __name__ == "__main__":
    main()
