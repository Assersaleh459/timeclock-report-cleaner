# Time-Clock Cleaner & Excel Attendance Report

Turns a messy employee time-clock export into a clean file you can trust, and builds an Excel attendance report from it. It also writes a log of every problem it found and what it did about it.

Payroll errors often start with bad punch data: duplicate swipes, missing clock-outs, typos in names, and dates in three different formats. This tool fixes what can be fixed safely and flags the rest for a person to review. Nothing is changed silently.

## Screenshots

![Summary](docs/summary.png)

![Daily Detail](docs/daily_detail.png)

## Before and after

**Before** (`data/raw_timeclock.csv`): the same employee, on the same day, written two different ways.

```
EmpID,Name,Department,Date,Time,PunchType
E1005,CRISTIAN SANTOS  ,Drivers,Oct 1 2026,08:36,IN
E1005,cristian   santos,Drivers,2026-10-01,16:45,OUT
```

**After** (`output/clean_timeclock.csv`): one name spelling, one date format.

```
EmpID,Name,Department,Date,Time,PunchType
E1005,Cristian Santos,Drivers,2026-10-01,08:36,IN
E1005,Cristian Santos,Drivers,2026-10-01,16:45,OUT
```

**Issues log** (`output/issues_log.csv`): each entry points to the line in the original file.

```
RowNumber,EmpID,Name,Date,Problem,ActionTaken
6,E1002,Noah Rhodes,2026-10-01,Duplicate punch,Dropped duplicate row
120,E1021,Margaret Hawkins Dds,2026-10-02,Missing OUT punch,Kept IN punch
```

## What it checks

| Problem | What the cleaner does |
|---|---|
| Names with extra spaces or mixed capitals | Trims and fixes them, for example `  RYAN MUNOZ` becomes `Ryan Munoz` |
| Blank employee ID | Fills it in from the name if the name belongs to exactly one employee. Otherwise drops the row and logs it |
| Dates in different formats | Converts `2026-10-01`, `01/10/2026` and `Oct 1 2026` to `2026-10-01`. Drops dates it cannot read |
| Impossible times such as `25:10` or `13:99` | Drops the row and logs it |
| Duplicate punches | Keeps the first one and drops exact repeats |
| Weekend punches | Keeps them and flags them for review |
| Clock-in with no clock-out, or the reverse | Keeps the punch and flags the day |

## The Excel report

`report.py` turns the cleaned data into `output/attendance_report.xlsx`, with five sheets:

| Sheet | What it shows |
|---|---|
| Summary | Headline numbers for the month: employee-days, total hours, late arrivals and incomplete days |
| Employee Hours | One row per employee with their total hours, days worked, late arrivals and incomplete days |
| Daily Detail | One row per employee per day, with clock-in, clock-out and hours worked |
| Issues Log | Every problem the cleaner found, with the action taken |
| Settings | The work start time and break length that the report uses |

The totals are live Excel formulas, not fixed numbers. If you change the work start time or the break length on the Settings sheet, the late arrivals and hours update across the whole workbook.

## Results on the sample data

The sample file covers 40 employees over October 2026.

**Cleaner**

| | Count |
|---|---|
| Rows read | 1,741 |
| Rows written | 1,687 |
| Problems logged | 81 |
| Duplicate punches removed | 50 |
| Missing clock-outs flagged | 18 |
| Blank IDs recovered | 5 |
| Impossible times removed | 4 |
| Weekend punches flagged | 3 |
| Missing clock-ins flagged | 1 |

**Excel report**

| | Value |
|---|---|
| Employee-days | 842 |
| Total hours | 6,917.46 |
| Late arrivals | 130 |
| Incomplete days | 16 |

## How to run it

You need Python 3.10 or newer.

```bash
python -m venv venv
venv\Scripts\activate        # on macOS or Linux: source venv/bin/activate
pip install -r requirements.txt

python generate_data.py      # optional: rebuilds the sample data
python clean.py
python report.py
```

The cleaner reads `data/raw_timeclock.csv` and writes the clean file and the issues log to `output/`. The report script then reads those files and writes the Excel workbook to `output/`. Each script prints a short summary when it finishes.

If an output file is open in Excel, the script asks you to close it instead of crashing.

## About the sample data

All names and records are fake. `generate_data.py` creates them for a fictional company with a fixed random seed, so every run produces the same file. It adds the problems listed above on purpose, so the cleaner has something to find.

## Files

| File | Purpose |
|---|---|
| `clean.py` | The cleaner |
| `report.py` | Builds the Excel attendance report |
| `generate_data.py` | Builds the messy sample file |
| `data/raw_timeclock.csv` | Sample input |
| `output/clean_timeclock.csv` | Cleaned result |
| `output/issues_log.csv` | Every problem found, with the action taken |
| `output/attendance_report.xlsx` | The Excel attendance report |

## License

MIT. See [LICENSE](LICENSE).
