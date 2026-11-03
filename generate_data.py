"""Generate a deliberately messy time-clock export for the fictional Northwind Logistics."""
import csv
import random
from datetime import date, datetime, timedelta
from pathlib import Path

from faker import Faker

SEED = 42
random.seed(SEED)
fake = Faker()
Faker.seed(SEED)

DEPARTMENTS = ["Warehouse", "Dispatch", "Drivers", "Finance", "HR", "Customer Service"]
FIELDS = ["EmpID", "Name", "Department", "Date", "Time", "PunchType"]
START, END = date(2026, 10, 1), date(2026, 10, 31)
stats = {}


def count(problem, n=1):
    stats[problem] = stats.get(problem, 0) + n


employees = [
    {"EmpID": f"E{1001 + i}", "Name": fake.name(), "Department": random.choice(DEPARTMENTS)}
    for i in range(40)
]


def punch(emp, day, hour, minute, kind):
    return {**emp, "Date": day, "Time": f"{hour:02d}:{minute:02d}", "PunchType": kind}


# Clean data: weekday IN/OUT pairs, with the odd absence
rows = []
days = [START + timedelta(n) for n in range((END - START).days + 1)]
for day in (d for d in days if d.weekday() < 5):
    for emp in employees:
        if random.random() < 0.05:
            continue
        rows.append(punch(emp, day, random.randint(6, 9), random.randint(0, 59), "IN"))
        rows.append(punch(emp, day, random.randint(16, 18), random.randint(0, 59), "OUT"))

# Missing OUT punches
outs = [r for r in rows if r["PunchType"] == "OUT"]
for r in random.sample(outs, 12):
    rows.remove(r)
count("missing OUT punches", 12)

# Weekend punches
weekend = [d for d in days if d.weekday() >= 5]
for _ in range(3):
    rows.append(punch(random.choice(employees), random.choice(weekend), random.randint(7, 9), random.randint(0, 59), "IN"))
count("weekend punches", 3)

# Impossible times
for r in random.sample(rows, 4):
    r["Time"] = random.choice(["25:10", "24:75", "27:00", "13:99"])
count("impossible times", 4)

# Blank EmpIDs
for r in random.sample(rows, 5):
    r["EmpID"] = ""
count("blank EmpIDs", 5)

# Messy names: stray spaces and inconsistent capitalization
messy = random.sample(rows, int(len(rows) * 0.15))
for r in messy:
    name = random.choice([str.upper, str.lower, str.title, lambda s: s])(r["Name"])
    r["Name"] = random.choice(["  " + name, name + "  ", name.replace(" ", "   ")])
count("messy names", len(messy))

# Mixed date formats
formats = [
    lambda d: d.isoformat(),
    lambda d: d.strftime("%d/%m/%Y"),
    lambda d: f"{d:%b} {d.day} {d.year}",
]
for r in rows:
    r["Date"] = random.choice(formats)(r["Date"])

# Duplicate punches (adjacent copies, ~3%)
dupes = set(random.sample(range(len(rows)), int(len(rows) * 0.03)))
final = []
for i, r in enumerate(rows):
    final.append(r)
    if i in dupes:
        final.append(dict(r))
count("duplicate punches", len(dupes))

Path("data").mkdir(exist_ok=True)
with open("data/raw_timeclock.csv", "w", newline="", encoding="utf-8") as f:
    writer = csv.DictWriter(f, fieldnames=FIELDS, extrasaction="ignore")
    writer.writeheader()
    writer.writerows(final)

print(f"Wrote {len(final)} rows for {len(employees)} employees to data/raw_timeclock.csv")
for problem, n in stats.items():
    print(f"  {problem}: {n}")
