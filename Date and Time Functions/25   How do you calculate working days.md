# How to Calculate Working Days in Excel

This guide explains how to calculate working days between two dates using Excel's NETWORKDAYS functions, which exclude weekends and optional holidays from date range calculations.

## Function Syntax

### NETWORKDAYS Function
```
=NETWORKDAYS(start_date, end_date, [holidays])
```
Calculates working days between two dates, excluding Saturdays, Sundays, and specified holidays.

**Parameters:**
- `start_date`: The beginning date of the period
- `end_date`: The ending date of the period
- `[holidays]`: Optional range containing holiday dates to exclude

### NETWORKDAYS.INTL Function
```
=NETWORKDAYS.INTL(start_date, end_date, [weekend], [holidays])
```
Calculates working days with customizable weekend days.

**Parameters:**
- `start_date`: The beginning date of the period
- `end_date`: The ending date of the period
- `[weekend]`: Optional number or string specifying weekend days
- `[holidays]`: Optional range containing holiday dates to exclude

## Weekend Codes for NETWORKDAYS.INTL

### Common Weekend Codes
```
1  = Saturday, Sunday (default)
2  = Sunday, Monday
3  = Monday, Tuesday
4  = Tuesday, Wednesday
5  = Wednesday, Thursday
6  = Thursday, Friday
7  = Friday, Saturday
11 = Sunday only
12 = Monday only
13 = Tuesday only
14 = Wednesday only
15 = Thursday only
16 = Friday only
17 = Saturday only
```

### String Method for Custom Weekends
Use 7-character string where:
- `1` = non-workday
- `0` = workday
- First character = Monday, last character = Sunday

Example: `"0000011"` = Saturday-Sunday weekend (same as code 1)

## Worked Examples

### Basic Working Days Calculation
**Start date (A1):** `March 1, 2024` (Friday)
**End date (B1):** `March 10, 2024` (Sunday)

```
=NETWORKDAYS(A1, B1)
```
Returns: `6` (Excludes March 2-3 and March 9-10 weekends)

### With Holiday Exclusion
**Holidays range (C1:C2):** `March 4, 2024` (Monday)

```
=NETWORKDAYS(A1, B1, C1:C2)
```
Returns: `5` (Excludes weekends plus March 4 holiday)

### Custom Weekend Schedule
**Sunday-only weekend:**
```
=NETWORKDAYS.INTL(A1, B1, 11)
```
Returns: `8` (Only Sundays excluded: March 3 and March 10)

**Friday-Saturday weekend:**
```
=NETWORKDAYS.INTL(A1, B1, 7)
```
Returns: `7` (Excludes March 1-2 and March 8-9)

**String method - Tuesday/Wednesday weekend:**
```
=NETWORKDAYS.INTL(A1, B1, "0011000")
```
Returns: `7` (Excludes Tuesdays and Wednesdays)

### Real-World Business Scenario
**Project timeline:** March 1-15, 2024
**Holidays:** March 4, March 11
**Weekend:** Saturday-Sunday

```
=NETWORKDAYS("3/1/2024", "3/15/2024", D1:D2)
```
Where D1:D2 contains the holiday dates.
Returns: `9` working days

> [!NOTE]
> NETWORKDAYS includes both the start_date and end_date in the calculation. If you want to exclude the end_date, use `=NETWORKDAYS(start_date, end_date-1, holidays)`

> [!IMPORTANT]
> The holidays parameter should be a range containing specific dates, not date ranges or recurring descriptions. Each holiday must be a separate date value in the range.

> [!TIP]
> Create a named range for your organization's holidays and reference it in your NETWORKDAYS formulas. This makes formulas easier to read and maintain: `=NETWORKDAYS(A1, B1, CompanyHolidays)`

## Practical Applications

### Project Management
**Days remaining:**
```
=NETWORKDAYS(TODAY(), project_deadline, holidays)
```

**Percentage complete:**
```
=NETWORKDAYS(project_start, TODAY(), holidays) / NETWORKDAYS(project_start, project_end, holidays)
```

### Human Resources
**Employee working days:**
```
=NETWORKDAYS(hire_date, TODAY(), holidays)
```

**Probation period:**
```
=NETWORKDAYS(hire_date, hire_date + 90, holidays)
```

### Service Level Agreements (SLAs)
**Response time in working days:**
```
=NETWORKDAYS(incident_date, response_date, holidays)
```

### Financial Calculations
**Payment terms:**
```
=NETWORKDAYS(invoice_date, invoice_date + 30, holidays)
```

## Advanced Techniques

### Dynamic Holiday Ranges
Use Excel Tables for automatic expansion:
```
=NETWORKDAYS(A1, B1, Table1[Holidays])
```

### Handling Empty Cells
```
=IF(OR(A1="", B1=""), "", NETWORKDAYS(A1, B1, holidays))
```

### Work Hours Calculation
Combine with time calculations for precise work hour tracking:
```
=(NETWORKDAYS(start_datetime, end_datetime, holidays)-1)*8 + (IF(WEEKDAY(end_datetime,2)<6, end_time, 0) - IF(WEEKDAY(start_datetime,2)<6, start_time, 0))
```

### Regional Weekend Patterns
**Middle East (Friday-Saturday):**
```
=NETWORKDAYS.INTL(A1, B1, 7)
```

**International with Sunday only:**
```
=NETWORKDAYS.INTL(A1, B1, 11)
```

## Common Issues and Solutions

### #VALUE! Errors
- Ensure all date parameters are valid Excel dates
- Verify holiday range contains only valid dates
- Check for text values masquerading as dates

### Incorrect Results
- Confirm weekend codes match your business schedule
- Validate holiday dates are in the calculation period
- Remember NETWORKDAYS includes both start and end dates

### Performance with Large Holiday Lists
- Use defined names or Excel Tables for holiday ranges
- Consider creating a master holiday list for the organization
- Limit holiday ranges to relevant years only

## Alternative Methods

### Using WORKDAY Function
For calculating end dates given start date and working days:
```
=WORKDAY(start_date, working_days, holidays)
```

### Manual Calculation (Without Functions)
For educational purposes or custom logic:
```
=SUMPRODUCT((WEEKDAY(ROW(INDIRECT(start_date&":"&end_date)),2)<6)*1) - COUNTIF(holidays, ">="&start_date, holidays, "<="&end_date)
```
