"""
Builds the synthetic input workbook (client_performance_data.xlsx) that the
Excel-to-deck agent reads from. All company/client names and figures are
fictional, created for portfolio-demo purposes only.
"""
import datetime
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def d(s):
    return datetime.date.fromisoformat(s)

FONT_NAME = "Arial"
HEADER_FILL = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
HEADER_FONT = Font(name=FONT_NAME, bold=True, color="FFFFFF", size=10)
BODY_FONT = Font(name=FONT_NAME, size=10)
NOTE_FONT = Font(name=FONT_NAME, size=9, italic=True, color="808080")
THIN = Side(style="thin", color="D9D9D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

wb = openpyxl.Workbook()

# ---------------------------------------------------------------- Clients ---
ws = wb.active
ws.title = "Clients"

headers = [
    "Client", "Since", "Engagement type", "Contract end",
    "Revenue this year (projected, $)", "Revenue last year ($)",
    "GM this year", "GM last year",
    "Headcount", "Non-billable resources", "Utilization %",
    "Number of projects", "Delivery status", "Open risks",
    "NPS", "Last NPS date", "eNPS", "Attrition %",
]
ws.append(headers)
for col in range(1, len(headers) + 1):
    c = ws.cell(row=1, column=col)
    c.font = HEADER_FONT
    c.fill = HEADER_FILL
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    c.border = BORDER
ws.row_dimensions[1].height = 32
ws.freeze_panes = "A2"

# name, since, engagement type, contract end,
# rev this yr, rev last yr, gm this yr, gm last yr,
# headcount, non-billable, utilization (None=N/A),
# num projects, delivery status,
# open risks, nps, last csat date, enps, attrition
rows = [
    ["Solvix Retail Group", d("2022-03-01"), "Staff augmentation", d("2028-03-01"),
     2851700, 2477600, 0.32, 0.29,
     48, 3, 0.88,
     3, "N/A – client-managed",
     2, 42, d("2026-03-15"), 38, 0.09],
    ["Northfield Logistics", d("2023-06-01"), "Project-based", d("2027-12-31"),
     4945000, 4348400, 0.41, 0.38,
     32, 4, None,
     5, "On track",
     1, 55, d("2026-06-01"), 44, 0.04],
    ["Harborline Financial Services", d("2021-01-15"), "Staff augmentation", d("2027-01-15"),
     2042700, 2204500, 0.27, 0.30,
     25, 2, 0.79,
     2, "N/A – client-managed",
     3, 21, d("2025-09-10"), 30, 0.18],
    ["Prairie Manufacturing Co.", d("2024-02-01"), "Project-based", d("2028-02-01"),
     1284400, 940500, 0.36, 0.33,
     18, 2, None,
     2, "At risk",
     4, 48, d("2026-05-20"), 41, 0.06],
]

for r in rows:
    ws.append(r)

last_row = 1 + len(rows)
for row in range(2, last_row + 1):
    for col in range(1, len(headers) + 1):
        c = ws.cell(row=row, column=col)
        c.font = BODY_FONT
        c.border = BORDER
        c.alignment = Alignment(horizontal="center", vertical="center")

# number formats
date_cols = [2, 4, 16]          # Since, Contract end, Last NPS date
currency_cols = [5, 6]           # Revenue
percent_cols = [7, 8, 11, 18]    # GM x2, Utilization, Attrition
for row in range(2, last_row + 1):
    for col in date_cols:
        ws.cell(row=row, column=col).number_format = "yyyy-mm-dd"
    for col in currency_cols:
        ws.cell(row=row, column=col).number_format = "$#,##0"
    for col in percent_cols:
        cell = ws.cell(row=row, column=col)
        if cell.value is not None:
            cell.number_format = "0.0%"

ws.cell(row=2, column=1).alignment = Alignment(horizontal="left", vertical="center")
for row in range(2, last_row + 1):
    ws.cell(row=row, column=1).alignment = Alignment(horizontal="left", vertical="center")

col_widths = [26, 11, 17, 12, 16, 14, 10, 10, 10, 11, 11, 9, 20, 9, 10, 13, 8, 11]
for i, w in enumerate(col_widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# Notes row directly under the table
note_row = last_row + 2
ws.cell(row=note_row, column=1,
        value=("Note: all client names and figures on this sheet are fictional, "
               "created for portfolio-demo purposes. Utilization % applies only to "
               "staff-augmentation engagements (client owns delivery); Delivery status "
               "applies only to project-based engagements (we own delivery)."))
ws.cell(row=note_row, column=1).font = NOTE_FONT
ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(headers))
ws.row_dimensions[note_row].height = 28
ws.cell(row=note_row, column=1).alignment = Alignment(wrap_text=True, vertical="top")

# ---------------------------------------------------------- Company_Summary --
ws2 = wb.create_sheet("Company_Summary")
ws2.sheet_view.showGridLines = False

labels = [
    "Total revenue this year (projected)",
    "Total revenue last year",
    "YoY revenue growth",
    "Blended GM this year",
    "Blended GM last year",
    "Total headcount",
    "Total non-billable resources",
    "Average utilization % (staff-aug accounts)",
    "Total number of projects",
    "Project-based accounts on track",
    "Project-based accounts at risk",
    "Project-based accounts delayed",
    "Total open risks",
    "Average NPS",
    "Average eNPS",
    "Average attrition %",
]

ws2["A1"] = "Meridian Advisory Group — Portfolio Summary"
ws2["A1"].font = Font(name=FONT_NAME, bold=True, size=14, color="1F3864")
ws2.merge_cells("A1:B1")

for i, label in enumerate(labels, start=3):
    ws2.cell(row=i, column=1, value=label).font = BODY_FONT

rng = f"Clients!E2:E{last_row}"
rng_last = f"Clients!F2:F{last_row}"
gm_rng = f"Clients!G2:G{last_row}"
gm_last_rng = f"Clients!H2:H{last_row}"

formulas = {
    3: f"=SUM({rng})",
    4: f"=SUM({rng_last})",
    5: "=B3/B4-1",
    6: f"=SUMPRODUCT(Clients!E2:E{last_row},Clients!G2:G{last_row})/SUM({rng})",
    7: f"=SUMPRODUCT(Clients!F2:F{last_row},Clients!H2:H{last_row})/SUM({rng_last})",
    8: f"=SUM(Clients!I2:I{last_row})",
    9: f"=SUM(Clients!J2:J{last_row})",
    10: f"=AVERAGE(Clients!K2:K{last_row})",
    11: f"=SUM(Clients!L2:L{last_row})",
    12: f'=COUNTIF(Clients!M2:M{last_row},"On track")',
    13: f'=COUNTIF(Clients!M2:M{last_row},"At risk")',
    14: f'=COUNTIF(Clients!M2:M{last_row},"Delayed")',
    15: f"=SUM(Clients!N2:N{last_row})",
    16: f"=AVERAGE(Clients!O2:O{last_row})",
    17: f"=AVERAGE(Clients!Q2:Q{last_row})",
    18: f"=AVERAGE(Clients!R2:R{last_row})",
}
for row, formula in formulas.items():
    c = ws2.cell(row=row, column=2, value=formula)
    c.font = BODY_FONT

pct_rows = {5, 6, 7, 10, 18}
currency_rows = {3, 4}
for row in range(3, 19):
    c = ws2.cell(row=row, column=2)
    if row in pct_rows:
        c.number_format = "0.0%"
    elif row in currency_rows:
        c.number_format = "$#,##0"
    else:
        c.number_format = "0.0"
    c.alignment = Alignment(horizontal="left")

ws2.column_dimensions["A"].width = 42
ws2.column_dimensions["B"].width = 16

note2 = ws2.cell(row=20, column=1,
                  value=("All figures computed live from the Clients sheet via formulas "
                         "(SUM/AVERAGE/SUMPRODUCT/COUNTIF) — edit any client row and "
                         "these totals recalculate."))
note2.font = NOTE_FONT
ws2.merge_cells("A20:D20")
ws2.cell(row=20, column=1).alignment = Alignment(wrap_text=True)

wb.save("client_performance_data.xlsx")
print("saved client_performance_data.xlsx")
