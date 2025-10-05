### 21. **Explain common date functions**

- **TODAY()**: Returns current date (updates daily)
- **NOW()**: Returns current date and time (updates constantly)
- **DATE(year, month, day)**: Creates date from numbers
Example: =DATE(2025, 10, 5) returns 10/5/2025
- **YEAR(date)**: Extracts year
- **MONTH(date)**: Extracts month (1-12)
- **DAY(date)**: Extracts day (1-31)
- **WEEKDAY(date, [return_type])**: Returns day of week (1-7)
- **EOMONTH(start_date, months)**: End of month
Example: =EOMONTH(TODAY(), 0) returns last day of current month
- **EDATE(start_date, months)**: Date months before/after
Example: =EDATE(TODAY(), 3) returns date 3 months from now
