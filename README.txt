# Employee Attendance & Punctuality Dashboard

A Power BI dashboard that turns raw daily arrival and departure records into clear indicators of **attendance, punctuality, working hours, and overtime** for a team of 23 employees.

The project covers the full pipeline:

1. A **Python** script consolidates monthly Excel exports (May, June, July) into one file per employee.
2. A **Power BI** data model loads the result. It adds calculated columns, supporting tables, and DAX measures.
3. An interactive **report** shows trends by employee, by day of the week, and by month.

---

## Table of Contents

- [Features](#features)
- [Tech Stack](#tech-stack)
- [Repository Structure](#repository-structure)
- [Business Rules](#business-rules)
- [Data Pipeline](#data-pipeline)
- [Data Model](#data-model)
  - [Calculated Columns](#calculated-columns)
  - [Calculated Tables](#calculated-tables)
  - [Measures](#measures)
- [Dashboard Pages & Visuals](#dashboard-pages--visuals)
- [Getting Started](#getting-started)
- [Known Limitations & Roadmap](#known-limitations--roadmap)
- [Data Privacy](#data-privacy)
- [License](#license)

---

## Features

- Classifies each day as **Present**, **Incomplete**, or **Absent**
- Classifies each arrival as **On Time**, **Late Comer**, or **Very Late**
- Computes **net working hours** after deducting the lunch break
- Tracks **extra time** worked after the official closing hour
- Shows attendance, absence, and punctuality **rates** with trends over time
- Filters the whole report by **employee**, **punctuality status**, and **month**

---

## Tech Stack

| Component | Tool |
|---|---|
| Data consolidation | Python 3, pandas, openpyxl |
| Data modelling & calculations | Power BI Desktop, DAX |
| Source data | Monthly Excel workbooks (one sheet per employee) |

---

## Repository Structure

```
.
├── data/
│   ├── raw/                  # Monthly source files: MAY.xlsx, JUIN.xlsx, JUILLET.xlsx
│   └── merged/               # Output of the merge script: fichier_fusionne_<EMPLOYEE>.xlsx
├── scripts/
│   └── merge_monthly_files.py
├── powerbi/
│   └── attendance_dashboard.pbix
├── docs/
│   └── screenshots/          # Dashboard screenshots
└── README.md
```

> Adjust this tree if your folders differ.

---

## Business Rules

| Rule | Value |
|---|---|
| Scheduled arrival time | **08:15** |
| Late threshold | **08:30** |
| Official closing time | **18:00** |
| Lunch break deducted from the working day | **1 h 30 min** |
| Period covered | May – July 2024 |

**Punctuality**

| Arrival time | Status |
|---|---|
| before 08:15 | On Time |
| from 08:15 to before 08:30 | Late Comer |
| 08:30 or later | Very Late |

**Attendance**

| Arrival recorded | Closing recorded | Status |
|---|---|---|
| No | No | Absent |
| Only one of the two | | Incomplete |
| Yes | Yes | Present |

---

## Data Pipeline

### Source format

The data comes from three workbooks, one per month: `MAY.xlsx`, `JUIN.xlsx`, and `JUILLET.xlsx`. Each workbook has **one sheet per employee**, and every sheet has the same columns. These include at least `Name`, `Date`, `Arrival-Time`, and `Closing-Time`.

### Merge script (`scripts/merge_monthly_files.py`)

The script:

1. Loads every sheet of each monthly workbook into a pandas DataFrame. All values are read as text (`dtype=str`) so that time values are not altered.
2. For each sheet name (employee) found in **all three** workbooks, it concatenates the May, June, and July rows into a single DataFrame.
3. It exports each merged DataFrame to `fichier_fusionne_<sheet_name>.xlsx`.

```python
import pandas as pd

file_path1 = 'JUILLET.xlsx'
file_path2 = 'JUIN.xlsx'
file_path3 = 'MAY.xlsx'

def load_excel_to_dataframes(file_path):
    excel_file = pd.ExcelFile(file_path)
    dataframes = {}
    for sheet_name in excel_file.sheet_names:
        dataframes[sheet_name] = pd.read_excel(excel_file, sheet_name=sheet_name, dtype=str)
    return dataframes

dataframes1 = load_excel_to_dataframes(file_path1)
dataframes2 = load_excel_to_dataframes(file_path2)
dataframes3 = load_excel_to_dataframes(file_path3)

merged_dataframes = {}
for sheet_name in dataframes1.keys():
    if sheet_name in dataframes2 and sheet_name in dataframes3:
        merged_dataframes[sheet_name] = pd.concat(
            [dataframes1[sheet_name], dataframes2[sheet_name], dataframes3[sheet_name]],
            ignore_index=True
        )

for sheet_name, df in merged_dataframes.items():
    output_file_name = f'fichier_fusionne_{sheet_name}.xlsx'
    df.to_excel(output_file_name, index=False)
    print(f'DataFrame fusionné pour la feuille {sheet_name} exporté dans {output_file_name}')
```

> An employee's sheet is merged only if it exists in **all three** monthly files. If a sheet is missing from any one month, the script skips that employee without any warning.

---

## Data Model

Each merged employee file is loaded into Power BI as its own table. The data model then adds the columns, tables, and measures below.

### Calculated Columns

These columns are added to each employee table.

**`PunctualityStatus`**: classifies the arrival time.

```dax
PunctualityStatus =
VAR ScheduledArrival = TIME(8, 15, 0)
VAR LateThreshold    = TIME(8, 30, 0)
RETURN
    SWITCH(
        TRUE(),
        [Arrival-Time] < ScheduledArrival, "On Time",
        [Arrival-Time] < LateThreshold,    "Late Comer",
        "Very Late"
    )
```

**`AttendanceStatus`**: labels the day as present, incomplete, or absent.

```dax
AttendanceStatus =
IF (
    ISBLANK([Arrival-Time]) && ISBLANK([Closing-Time]),
    "Absent",
    IF (
        ISBLANK([Arrival-Time]) || ISBLANK([Closing-Time]),
        "Incomplete",
        "Present"
    )
)
```

**`ExtraTime`**: time worked after 18:00.

```dax
ExtraTime =
IF (
    [Closing-Time] > TIME(18, 0, 0),
    [Closing-Time] - TIME(18, 0, 0),
    BLANK()
)
```

**`WorkingHoursWithBreak`**: net working time after the 1 h 30 lunch break is deducted.

```dax
WorkingHoursWithBreak =
VAR TotalWorkingHours = [Closing-Time] - [Arrival-Time]
VAR BreakDuration     = TIME(1, 30, 0)
RETURN
    IF (
        TotalWorkingHours > BreakDuration,
        TotalWorkingHours - BreakDuration,
        BLANK()
    )
```

### Calculated Tables

**`Calendar`**: the date dimension used for month and day-of-week analysis.

```dax
Calendar =
ADDCOLUMNS (
    CALENDAR ( DATE(2024, 5, 1), DATE(2024, 5, 31) ),
    "Year",      YEAR([Date]),
    "Month",     MONTH([Date]),
    "Day",       DAY([Date]),
    "MonthName", FORMAT([Date], "MMMM")
)
```

**`PunctualitySlicer`**: a fixed list of statuses that drives the punctuality filter.

```dax
PunctualitySlicer =
DATATABLE (
    "PunctualityStatus", STRING,
    { {"On Time"}, {"Late Comer"}, {"Very Late"} }
)
```

**`MasterNameTable`**: one distinct list of employee names, built from all 23 employee tables. It drives the employee filter.

```dax
MasterNameTable =
UNION (
    DISTINCT ( '<EMPLOYEE_TABLE_1>'[Name] ),
    DISTINCT ( '<EMPLOYEE_TABLE_2>'[Name] ),
    -- ... one line per employee table (23 in total)
    DISTINCT ( '<EMPLOYEE_TABLE_23>'[Name] )
)
```

### Measures

**`TotalExtraTime`**: the sum of overtime.

```dax
TotalExtraTime = SUM ( sheet1[ExtraTime] )
```

**`TotalFormatted`**: displays the overtime total as `Xh Ym Zs`.

```dax
TotalFormatted =
    INT ( [TotalExtraTime] * 24 ) & "h " &
    INT ( MOD ( [TotalExtraTime] * 24 * 60, 60 ) ) & "m " &
    INT ( MOD ( [TotalExtraTime] * 24 * 3600, 60 ) ) & "s"
```

**`Total Working Hours`**: *to be finalised.* A suggested definition:

```dax
Total Working Hours = SUM ( sheet1[WorkingHoursWithBreak] ) * 24   -- in hours
```

---

## Dashboard Pages & Visuals

| Indicator | Visual | Axes / Fields |
|---|---|---|
| Average daily working hours (by weekday) | Line and stacked column chart | X: `DayName` · Y: Average Working Hours · Secondary Y: Closing Time |
| Average working hours (by employee) | Line chart | X: `Name` · Y: Average Working Hours · Secondary Y: Arrival Time |
| Total working hours | Card | `Total Working Hours` |
| Total absence rate | Donut chart | `AttendanceStatus` |
| Attendance trend | Line chart | X: Month · Y: Total Attendance Rate |
| Total punctuality rate | Donut chart | `PunctualityStatus` |
| Extra time tracking | Line and stacked column chart | X: `DayName` · Y: `ExtraTime` |

**Filters (slicers):** employee name (`MasterNameTable`), punctuality status (`PunctualitySlicer`), and month (`Calendar`).

---

## Getting Started

### Prerequisites

- Python 3.8+
- Power BI Desktop (Windows)

```bash
pip install pandas openpyxl
```

### Steps

1. Put the monthly workbooks (`MAY.xlsx`, `JUIN.xlsx`, `JUILLET.xlsx`) in the same folder as the script. You can also edit the paths in the script.
2. Run the merge:
   ```bash
   python scripts/merge_monthly_files.py
   ```
3. Open `powerbi/attendance_dashboard.pbix` in Power BI Desktop.
4. If the files have moved, update the source path under **Transform data → Data source settings**.
5. Click **Refresh** to reload the data and recompute the columns and measures.

---

## Known Limitations & Roadmap

- **Calendar range:** the `Calendar` table covers May 2024 only, but the data runs through July. Extend it to `DATE(2024, 7, 31)`, or generate it with `CALENDARAUTO()`, so that June and July appear in the monthly trend.
- **Blank arrival times:** in DAX, a blank value compares as less than `08:15`, so absent days are currently counted as **"On Time"**. Add a guard such as `ISBLANK([Arrival-Time]), BLANK()` as the first `SWITCH` condition.
- **Missing days off from lunch deduction:** the 1 h 30 break is deducted on every present day, including half days.
- **One table per employee:** this design needs a 23-line `UNION` and makes it harder to add new employees. A planned improvement is to append all employees into a **single fact table** (in Python or Power Query), with one `Name` column. The `MasterNameTable` union would then no longer be needed, and a measure such as `TotalExtraTime` would cover everyone instead of a single table.
- **Hard-coded rules:** the 08:15, 08:30, and 18:00 thresholds and the break length are written into the formulas. Moving them into a parameters table would make them easier to change.
- **Silent skips in the merge script:** add a warning when an employee's sheet is missing from one of the monthly files.
- **Pending measure:** finalise `Total Working Hours`, then add explicit rate measures for attendance, absence, and punctuality.

---

## Data Privacy

The source files contain the **names and daily schedules of real employees**. Do not commit raw or merged data files to a public repository. Keep them out of version control:

```gitignore
data/raw/
data/merged/
*.xlsx
```

For a public showcase, publish **anonymised sample data** (for example, `Employee 01`, `Employee 02`) and screenshots with the names blurred.

---

## License

This project is licensed under the terms described in the [`LICENSE`](LICENSE) file.

---

**Author:** Engolo Lucien