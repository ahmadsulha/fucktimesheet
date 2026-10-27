## How to run

```
uv run python process_timesheet.py "D:\timesheet\Individual Attendance Report.xls" 30 09,03
```

First param is .xls file.
Second param is the days oil is applied separated by comma. Must be in 2 digits. E.g. 02,09,18
Third param is the days working from home that is manually inserted to override the value in the original timesheet. Must be in 2 digits. E.g. 01,11,24