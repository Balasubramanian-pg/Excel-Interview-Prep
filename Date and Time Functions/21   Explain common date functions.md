# Common Date Functions in Excel

This guide explains the essential date functions in Excel for working with dates, calculating date differences, extracting date components, and performing date arithmetic.

## Function Syntax and Usage

### TODAY Function
```
=TODAY()
```
Returns the current date. Updates automatically when the workbook is opened or recalculated.

**Parameters:**
- No arguments required
- Returns serial number representing current date
- Format displays as date, but underlying value is numeric

### NOW Function
```
=NOW()
```
Returns the current date and time. Updates automatically when the workbook is opened or recalculated.

**Parameters:**
- No arguments required
- Returns serial number with decimal for time portion
- Format displays as date and time

### DATE Function
```
=DATE(year, month, day)
```
Creates a date from individual year, month, and day components.

**Parameters:**
- `year`: 1900-9999 (Excel adjusts years <1900)
- `month`: 1-12 (Excel handles values outside this range intelligently)
- `day`: 1-31 (Excel handles values outside this range intelligently)

### YEAR, MONTH, DAY Functions
```
=YEAR(date)
=MONTH(date)  
=DAY(date)
```
Extract specific components from a date.

**Parameters:**
- `date`: Date value or reference to date cell
- Returns numeric values for each component

### WEEKDAY Function
```
=WEEKDAY(date, [return_type])
```
Returns the day of the week as a number.

**Parameters:**
- `date`: Date value or reference
- `return_type`: Optional, determines numbering system:
  - `1` (default): Sunday=1 through Saturday=7
  - `2`: Monday=1 through Sunday=7
  - `3`: Monday=0 through Sunday=6

### EOMONTH Function
```
=EOMONTH(start_date, months)
```
Returns the last day of the month a specified number of months from start_date.

**Parameters:**
- `start_date`: Starting date for calculation
- `months`: Number of months before/after start_date
  - `0`: End of current month
  - Positive: Future months
  - Negative: Past months

### EDATE Function
```
=EDATE(start_date, months)
```
Returns the date that is the specified number of months before or after start_date.

**Parameters:**
- `start_date`: Starting date for calculation
- `months`: Number of months before/after start_date
  - Positive: Future dates
  - Negative: Past dates

## Worked Examples

**Current Date:**
```
=TODAY()
```
Returns: Current system date (e.g., 3/15/2024)

**Current Date and Time:**
```
=NOW()
```
Returns: Current system date and time (e.g., 3/15/2024 14:30)

**Create Specific Date:**
```
=DATE(2024, 12, 25)
```
Returns: 12/25/2024

**Extract Date Components:**
```
=YEAR("12/25/2024")
```
Returns: `2024`

```
=MONTH("12/25/2024")
```
Returns: `12`

```
=DAY("12/25/2024")
```
Returns: `25`

**Day of Week:**
```
=WEEKDAY("12/25/2024", 2)
```
Returns: `3` (Wednesday, using Monday=1 system)

**End of Month:**
```
=EOMONTH("3/15/2024", 0)
```
Returns: 3/31/2024 (last day of March 2024)

```
=EOMONTH("3/15/2024", 1)
```
Returns: 4/30/2024 (last day of next month)

**Date Arithmetic:**
```
=EDATE("3/15/2024", 3)
```
Returns: 6/15/2024 (3 months later)

```
=EDATE("3/15/2024", -6)
```
Returns: 9/15/2023 (6 months earlier)

> [!NOTE]
> Excel stores dates as serial numbers where 1 = January 1, 1900. This allows for easy date arithmetic. TODAY() and NOW() are volatile functions that recalculate with every worksheet change.

> [!IMPORTANT]
> EOMONTH and EDATE functions require the Analysis ToolPak in Excel 2007 and earlier versions. They are built-in functions in Excel 2010 and later. These functions preserve the day component when possible (EDATE) or calculate month ends accurately (EOMONTH).

> [!TIP]
- Use DATE instead of typing dates to avoid regional format issues
- Combine functions for complex calculations: `=EOMONTH(TODAY(), -1)+1` gives first day of current month
- Use WORKDAY and NETWORKDAYS for business day calculations excluding weekends/holidays

## Practical Applications

### Age Calculation
```
=YEAR(TODAY()) - YEAR(birth_date) - IF(TODAY() < DATE(YEAR(TODAY()), MONTH(birth_date), DAY(birth_date)), 1, 0)
```

### Days Between Dates
```
=end_date - start_date
```

### First Day of Month
```
=DATE(YEAR(A1), MONTH(A1), 1)
```

### Quarter Calculation
```
="Q" & INT((MONTH(A1)-1)/3)+1
```

### Fiscal Year Calculation
```
=YEAR(A1) + IF(MONTH(A1)>=7, 1, 0)
```
(Assuming fiscal year starts July 1)

---
