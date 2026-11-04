"""Build output/attendance_report.xlsx (live Excel formulas) from the cleaned punches and the issues log."""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---- Settings (safe for a client to change; also editable on the workbook's Settings sheet) ----
COMPANY_NAME = "Northwind Logistics"
WORK_START = "08:30"  # an IN punch after this time counts as late
BREAK_MINUTES = 60  # unpaid break, subtracted from hours on complete days

HEADER_FILL = PatternFill("solid", fgColor="1F3864")
STATUS_FILLS = {  # Daily Detail status colors (applied as conditional formatting)
    "Late": "F8CBAD",
    "Incomplete": "F4B6B6",
    "Weekend - review": "FFF2CC",
}
DD = "'Daily Detail'"  # sheet reference used in formulas
EH = "'Employee Hours'"

# Load the cleaned punches and the issues log
punches = pd.read_csv("output/clean_timeclock.csv", dtype=str, keep_default_na=False)
issues = pd.read_csv("output/issues_log.csv", dtype=str, keep_default_na=False)
issues["RowNumber"] = issues["RowNumber"].astype(int)

# Daily structure: first IN and last OUT per employee/day (hours and status are formulas in Excel)
first_in = punches[punches["PunchType"] == "IN"].groupby(["EmpID", "Date"])["Time"].min().rename("IN")
last_out = punches[punches["PunchType"] == "OUT"].groupby(["EmpID", "Date"])["Time"].max().rename("OUT")
daily = pd.concat([first_in, last_out], axis=1).reset_index().sort_values(["EmpID", "Date"])
people = punches.drop_duplicates("EmpID").set_index("EmpID")[["Name", "Department"]]
daily["Name"] = daily["EmpID"].map(people["Name"])
dd_last = len(daily) + 1  # last data row on Daily Detail


def to_time(text):
    return None if pd.isna(text) else datetime.strptime(text, "%H:%M").time()


def add_table(ws, header, rows, numeric=()):
    """Append a styled header row plus data rows; return (header row number, last row number).
    Headers named in `numeric` are right-aligned to line up with their numbers."""
    ws.append(header)
    header_row = ws.max_row
    for cell in ws[header_row]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = HEADER_FILL
        if cell.value in numeric:
            cell.alignment = Alignment(horizontal="right")
    for row in rows:
        ws.append([None if pd.isna(v) else (v.item() if hasattr(v, "item") else v) for v in row])
    return header_row, ws.max_row


def finish(ws, header_row, last_row, skip_rows=0):
    """Freeze the header, add filters, and auto-fit column widths (formulas get a fixed estimate)."""
    ws.freeze_panes = ws.cell(row=header_row + 1, column=1)
    ws.auto_filter.ref = f"A{header_row}:{get_column_letter(ws.max_column)}{last_row}"
    for col in ws.iter_cols(min_row=skip_rows + 1):
        lengths = [16 if str(c.value).startswith("=") else len(str(c.value)) for c in col if c.value is not None]
        ws.column_dimensions[col[0].column_letter].width = min(max(lengths, default=0) + 2, 60)


wb = Workbook()

# Sheet 1: Summary (title, key numbers, issue counts by type), all numbers are formulas
ws = wb.active
ws.title = "Summary"
month = pd.to_datetime(daily["Date"]).min().strftime("%B %Y")
ws.append([f"{COMPANY_NAME} - Attendance Report, {month}"])
ws["A1"].font = Font(bold=True, size=14)
ws.append(['="Hours are net of a "&Settings!$B$3&"-minute break"'])
ws["A2"].font = Font(italic=True)
ws.append([])
key_numbers = [  # table starts on row 4, so Total hours lands on row 7
    ("Employees", f"=COUNTA({EH}!$A:$A)-1"),
    ("Days covered", f"=SUMPRODUCT(1/COUNTIF({DD}!$C$2:$C${dd_last},{DD}!$C$2:$C${dd_last}))"),
    ("Total hours", f"=ROUND(SUM({DD}!$F:$F),2)"),
    ("Average hours per day", f'=IFERROR(ROUND(B7/SUM({EH}!$E:$E),2),"")'),
    ("Late arrivals", f'=COUNTIF({DD}!$G:$G,"Late")'),
    ("Incomplete days", f'=COUNTIF({DD}!$G:$G,"Incomplete")'),
]
add_table(ws, ["Metric", "Value"], key_numbers, numeric=("Value",))
ws["B7"].number_format = ws["B8"].number_format = "0.00"  # Total hours, Average hours per day
ws.append([])
problems = sorted(issues["Problem"].unique())
header_row, last_row = add_table(ws, ["Issue Type", "Count"], [(p, None) for p in problems], numeric=("Count",))
for r in range(header_row + 1, last_row + 1):
    ws[f"B{r}"] = f"=COUNTIF('Issues Log'!$E:$E,A{r})"
finish(ws, header_row, last_row, skip_rows=2)  # filter covers the issue table; title and note ignored for widths
ws.freeze_panes = "A2"  # keep just the title visible; freezing down to the issue table would eat the screen

# Sheet 2: Employee Hours (COUNTIFS / SUMIFS on Daily Detail)
ws = wb.create_sheet("Employee Hours")
rows = []
for r, emp_id in enumerate(daily["EmpID"].unique(), start=2):
    rows.append([
        emp_id, people.at[emp_id, "Name"], people.at[emp_id, "Department"],
        f"=COUNTIFS({DD}!$A:$A,$A{r})",
        f'=COUNTIFS({DD}!$A:$A,$A{r},{DD}!$D:$D,">=0",{DD}!$E:$E,">=0")',
        f"=ROUND(SUMIFS({DD}!$F:$F,{DD}!$A:$A,$A{r}),2)",
        f'=IF(E{r}=0,"",ROUND(F{r}/E{r},2))',
        f'=COUNTIFS({DD}!$A:$A,$A{r},{DD}!$G:$G,"Late")',
        f'=COUNTIFS({DD}!$A:$A,$A{r},{DD}!$G:$G,"Incomplete")',
    ])
header_row, last_row = add_table(
    ws, ["EmpID", "Name", "Department", "Days Present", "Complete Days", "Total Hours",
         "Avg Hours/Day", "Late Count", "Incomplete Days"], rows,
    numeric=("Days Present", "Complete Days", "Total Hours", "Avg Hours/Day", "Late Count", "Incomplete Days"))
finish(ws, header_row, last_row)
for r in range(2, last_row + 1):
    ws[f"F{r}"].number_format = ws[f"G{r}"].number_format = "0.00"  # Total Hours, Avg Hours/Day

# Sheet 3: Daily Detail (real times, Hours and Status formulas, conditional colors on Status)
ws = wb.create_sheet("Daily Detail")
rows = []
for r, d in enumerate(daily.itertuples(index=False), start=2):
    rows.append([
        d.EmpID, d.Name, datetime.strptime(d.Date, "%Y-%m-%d").date(), to_time(d.IN), to_time(d.OUT),
        f'=IF(OR(D{r}="",E{r}=""),"",ROUND((E{r}-D{r})*24-Settings!$B$3/60,2))',
        f'=IF(WEEKDAY(C{r},2)>=6,"Weekend - review",IF(OR(D{r}="",E{r}=""),"Incomplete",'
        f'IF(D{r}>Settings!$B$2,"Late","OK")))',
    ])
finish(ws, *add_table(ws, ["EmpID", "Name", "Date", "IN", "OUT", "Hours", "Status"], rows,
                      numeric=("Date", "IN", "OUT", "Hours")))
for r in range(2, dd_last + 1):
    ws[f"C{r}"].number_format = "yyyy-mm-dd"
    ws[f"D{r}"].number_format = ws[f"E{r}"].number_format = "hh:mm"
    ws[f"F{r}"].number_format = "0.00"
for status, color in STATUS_FILLS.items():
    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    ws.conditional_formatting.add(f"G2:G{dd_last}", CellIsRule(operator="equal", formula=[f'"{status}"'], fill=fill))

# Sheet 4: Issues Log (plain values)
ws = wb.create_sheet("Issues Log")
log_rows = issues.values.tolist()
for row in log_rows:  # real dates where possible; unparseable dates keep their original text
    try:
        row[3] = datetime.strptime(row[3], "%Y-%m-%d").date()
    except ValueError:
        pass
header_row, last_row = add_table(ws, list(issues.columns), log_rows, numeric=("RowNumber", "Date"))
finish(ws, header_row, last_row)
for r in range(2, last_row + 1):
    if not isinstance(ws[f"D{r}"].value, str):
        ws[f"D{r}"].number_format = "yyyy-mm-dd"

# Sheet 5: Settings (the formulas above read these two cells)
ws = wb.create_sheet("Settings")
add_table(ws, ["Setting", "Value"], [("Work start (IN after this is late)", to_time(WORK_START)),
                                     ("Break (minutes)", BREAK_MINUTES)], numeric=("Value",))
ws["B2"].number_format = "hh:mm"
finish(ws, 1, 3)

# Save; if the file is locked (e.g. open in Excel) explain and exit without a traceback
Path("output").mkdir(exist_ok=True)
try:
    wb.save("output/attendance_report.xlsx")
except PermissionError:
    print("Can't save attendance_report.xlsx - please close it in Excel and run again.")
    sys.exit(1)
print(f"Saved output/attendance_report.xlsx ({daily['EmpID'].nunique()} employees, {len(daily)} employee-days)")
