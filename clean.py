"""Clean data/raw_timeclock.csv and log every problem found."""
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd

# ---- Settings: adapt these to a new client file, nothing below needs to change ----
INPUT_FILE = "data/raw_timeclock.csv"
OUTPUT_DIR = "output"
# Map the names this script uses (left) to the column names in the client's file (right)
COLUMNS = {
    "EmpID": "EmpID",
    "Name": "Name",
    "Department": "Department",
    "Date": "Date",
    "Time": "Time",
    "PunchType": "PunchType",
}
PUNCH_IN, PUNCH_OUT = "IN", "OUT"  # the values used for clock-in and clock-out
DATE_FORMATS = ["%Y-%m-%d", "%d/%m/%Y", "%b %d %Y"]  # every date format seen in the file
# -----------------------------------------------------------------------------------

issues = []


def log(rows, problem, action):
    """Add one issues_log entry per row. RowNumber is the line number in the raw CSV."""
    for row_number, r in rows.iterrows():
        issues.append({"RowNumber": row_number, "EmpID": r["EmpID"], "Name": r["Name"],
                       "Date": r["Date"], "Problem": problem, "ActionTaken": action})


def save_csv(frame, path):
    """Write a CSV; if the file is locked (e.g. open in Excel), explain and exit without a traceback."""
    try:
        frame.to_csv(path, index=False)
    except PermissionError:
        print(f"Can't save {Path(path).name} - please close it in Excel and run again.")
        sys.exit(1)


def parse_date(text):
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(text.strip(), fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass
    return None


# Load everything as text; index = line number in the file (header is line 1)
df = pd.read_csv(INPUT_FILE, dtype=str, keep_default_na=False)
missing = [src for src in COLUMNS.values() if src not in df.columns]
if missing:
    print(f"Columns not found in {INPUT_FILE}: {missing}. Update COLUMNS at the top of this script.")
    sys.exit(1)
df = df[list(COLUMNS.values())].rename(columns={src: dst for dst, src in COLUMNS.items()})
df.index = df.index + 2
total_in = len(df)

# 1. Names: trim, collapse repeated spaces, Title Case
raw_names = df["Name"]
df["Name"] = df["Name"].str.strip().str.replace(r"\s+", " ", regex=True).str.title()
names_normalized = int((df["Name"] != raw_names).sum())

# 2. Dates: parse the three known formats into YYYY-MM-DD (unparseable rows are logged with their original text)
parsed = df["Date"].map(parse_date)
bad_date = parsed.isna()
log(df[bad_date], "Unparseable date", "Dropped row")
df = df[~bad_date].copy()
df["Date"] = parsed[~bad_date]

# 3. Blank EmpID: recover from the name if it maps to exactly one ID, otherwise drop
ids_by_name = df[df["EmpID"] != ""].groupby("Name")["EmpID"].agg(lambda s: set(s)).to_dict()
blank = df[df["EmpID"] == ""]
candidates = blank["Name"].map(lambda n: ids_by_name.get(n, set()))
recoverable = candidates.map(len) == 1
df.loc[recoverable[recoverable].index, "EmpID"] = candidates[recoverable].map(lambda s: next(iter(s)))
log(df.loc[recoverable[recoverable].index], "Blank EmpID", "Recovered EmpID from name")
unrecoverable = recoverable[~recoverable].index
log(df.loc[unrecoverable], "Blank EmpID", "Dropped row - name matches no single employee")
df = df.drop(unrecoverable)

# 4. Times: hour must be 0-23 and minute 0-59 (malformed times fail the same check)
hm = df["Time"].str.split(":", expand=True).apply(pd.to_numeric, errors="coerce")
bad_time = ~((hm[0] <= 23) & (hm[1] <= 59))
log(df[bad_time], "Impossible time", "Dropped row")
df = df[~bad_time]

# 5. Duplicates: exact repeats of an earlier row (after the cleaning above)
dup = df.duplicated(keep="first")
log(df[dup], "Duplicate punch", "Dropped duplicate row")
df = df[~dup]

# 6. Weekend punches: keep, but flag for review
weekend = pd.to_datetime(df["Date"]).dt.day_of_week >= 5
log(df[weekend], "Weekend punch - needs review", "Kept")

# 7. Missing OUT: employee/day with an IN but no OUT (one log entry per employee/day)
has_out = df.groupby(["EmpID", "Date"])["PunchType"].transform(lambda s: (s == PUNCH_OUT).any())
missing_out = df[(df["PunchType"] == PUNCH_IN) & ~has_out].drop_duplicates(["EmpID", "Date"])
log(missing_out, "Missing OUT punch", "Kept IN punch")

# 8. Missing IN: employee/day with an OUT but no IN (one log entry per employee/day)
has_in = df.groupby(["EmpID", "Date"])["PunchType"].transform(lambda s: (s == PUNCH_IN).any())
missing_in = df[(df["PunchType"] == PUNCH_OUT) & ~has_in].drop_duplicates(["EmpID", "Date"])
log(missing_in, "Missing IN punch", "Kept OUT punch")

# Write outputs (punch words become IN/OUT so report.py works whatever the client's system calls them)
df["PunchType"] = df["PunchType"].replace({PUNCH_IN: "IN", PUNCH_OUT: "OUT"})
Path(OUTPUT_DIR).mkdir(exist_ok=True)
save_csv(df, f"{OUTPUT_DIR}/clean_timeclock.csv")
log_df = pd.DataFrame(issues, columns=["RowNumber", "EmpID", "Name", "Date", "Problem", "ActionTaken"])
log_df = log_df.sort_values("RowNumber", kind="stable")
save_csv(log_df, f"{OUTPUT_DIR}/issues_log.csv")

# Summary
print(f"Rows read: {total_in} | Rows written: {len(df)} | Issues logged: {len(log_df)}")
for problem, n in log_df.groupby("Problem").size().items():
    print(f"  {problem}: {n}")
print(f"Names normalized: {names_normalized}")
