# Hexa Climate — Attendance vs Timesheet Reconciliation

A browser-based tool that compares the **Attendance Sheet** against the
**Timesheet** (uploaded in bulk, month-wise) and flags any employee/date
where the leave marking doesn't match.

## Leave codes recognized

| Code | Meaning |
| --- | --- |
| SL | Sick Leave |
| PL | Paid Leave |
| UL | Unpaid Leave |
| FL | Floater Leave |
| SPL | Special Leave |
| ML | Maternity Leave |
| PT | Paternity Leave |
| BL | Bereavement Leave |
| CL | Casual Leave |

Any other non-empty value in a date cell is treated as **Present**.

## Run it

Open `index.html` in a browser (or serve the folder with any static server, e.g. `python3 -m http.server`). No build step.

1. Drop the Attendance workbook into the left panel.
2. Drop the Timesheet workbook into the right panel.
3. Click **Reconcile & Flag Mismatches**.

To try it immediately, use the files in `samples/` (`attendance_sheet_sample.csv`
and `timesheet_sample.csv`). They are intentionally seeded with 3 leave
mismatches, and one `WFH` vs `P` cell that correctly does **not** flag.

You'll see:
- **Leave Mismatches** — same employee + date, but the two files disagree on the leave code (or one says Leave and the other says Present).
- **Missing Dates** — for matched employees, a date recorded in only one of the two files.
- **Timesheet employees not found in Attendance** — names that couldn't be matched.

Only the employees whose timesheets you upload are reconciled (not everyone in the attendance sheet). Names are matched leniently (small spelling differences and first-name-only file names are tolerated).

All tables download as CSV.

## Generate timesheets from Attendance

Upload the Attendance Sheet, then in **Generate timesheets from Attendance** pick one or more employees and click **Generate & Download**. For each selected employee you get a monthly timesheet `.xlsx` that replicates the company template (same layout, merged headers, borders), named like `Siddharth Sundriyal_December 2025.xlsx`.

Day types are taken from the attendance sheet: leave codes → the leave, `H`/holiday → Holiday, `WO`/week-off and weekends → W. Off, everything else → a standard 9:30 AM–6:30 AM working day. The styled `.xlsx` writer is vendored at `vendor/xlsx-style.js`.

## Supported sheet layouts

The parser auto-detects either shape:

- **Wide**: one row per employee, one column per date (dates as headers).
- **Long**: one row per employee-date pair with an explicit `Date` column and a `Status`/`Attendance`/`Leave`/`Mark`/`Type` column.

Multiple sheets in one workbook are merged (useful for month-per-sheet timesheets).

## What's next

This is v1 scaffolding. Drop in real Attendance + Timesheet files and we'll:
- Tune the header/column detection to the exact layout you use.
- Add employee-ID normalization if names don't match cleanly across the two files.
- Add half-day / partial-leave handling if your codes cover that.
