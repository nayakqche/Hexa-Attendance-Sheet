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

You'll see:
- **Leave Mismatches** — same employee + date, but the two files disagree on the leave code (or one says Leave and the other says Present).
- **Missing in One File** — an employee/date recorded in only one of the two files.

Both tables download as CSV.

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
