# Time Functions in Excel

This guide explains the essential time functions in Excel for working with time values, extracting time components, creating time values from individual components, and performing time calculations.

## Function Syntax and Usage

### TIME Function
```
=TIME(hour, minute, second)
```
Creates a time value from individual hour, minute, and second components.

**Parameters:**
- `hour`: 0-23 (Excel handles values outside this range intelligently)
- `minute`: 0-59 (Excel handles values outside this range intelligently)
- `second`: 0-59 (Excel handles values outside this range intelligently)

### HOUR Function
```
=HOUR(time)
```
Extracts the hour component from a time value.

**Parameters:**
- `time`: Time value or reference to time cell
- Returns numeric value from 0 (12:00 AM) to 23 (11:00 PM)

### MINUTE Function
```
=MINUTE(time)
```
Extracts the minute component from a time value.

**Parameters:**
- `time`: Time value or reference to time cell
- Returns numeric value from 0 to 59

### SECOND Function
```
=SECOND(time)
```
Extracts the second component from a time value.

**Parameters:**
- `time`: Time value or reference to time cell
- Returns numeric value from 0 to 59

### NOW Function
```
=NOW()
```
Returns the current date and time. Updates automatically when the workbook is opened or recalculated.

**Parameters:**
- No arguments required
- Returns serial number with decimal representing time portion
- Format displays as date and time

## Worked Examples

**Create Specific Time:**
```
=TIME(14, 30, 0)
```
Returns: `2:30 PM` (serial value: 0.604166666666667)

```
=TIME(8, 45, 15)
```
Returns: `8:45:15 AM` (serial value: 0.364756944444444)

**Extract Time Components:**
```
=HOUR("14:30:00")
```
Returns: `14`

```
=MINUTE("14:30:00")
```
Returns: `30`

```
=SECOND("14:30:00")
```
Returns: `0`

**Current Date and Time:**
```
=NOW()
```
Returns: Current system date and time (e.g., 3/15/2024 14:30:25)

**Time Arithmetic:**
```
=TIME(14, 30, 0) + TIME(1, 15, 0)
```
Returns: `3:45 PM` (2:30 PM + 1 hour 15 minutes)

```
=TIME(23, 0, 0) + TIME(2, 0, 0)
```
Returns: `1:00 AM` (time rolls over to next day)

**Extract Time from DateTime:**
```
=NOW() - INT(NOW())
```
Returns: Current time without date component

**Create Time from Text:**
```
=TIMEVALUE("2:30 PM")
```
Returns: `0.604166666666667` (same as TIME(14,30,0))

> [!NOTE]
> Excel stores times as fractional numbers where 0.0 = 12:00:00 AM and 0.999988425925926 = 11:59:59 PM. One hour = 1/24, one minute = 1/1440, one second = 1/86400.

> [!IMPORTANT]
> The NOW() function is volatile and recalculates with every worksheet change. For static timestamps, use Ctrl+; for current date and Ctrl+Shift+; for current time, or use VBA for one-time timestamp entry.

> [!TIP]
- Use 24-hour format to avoid AM/PM confusion in calculations
- Format cells as Time to display results properly
- Combine with DATE function for complete datetime values: `=DATE(2024,3,15)+TIME(14,30,0)`

## Practical Applications

### Time Difference Calculation
```
=end_time - start_time
```
Format result cell as `[h]:mm` to display durations over 24 hours

### Business Hours Calculation
```
=IF(AND(start_time>=TIME(9,0,0), end_time<=TIME(17,0,0)), end_time-start_time, 0)
```

### Round Time to Nearest 15 Minutes
```
=TIME(HOUR(A1), ROUND(MINUTE(A1)/15, 0)*15, 0)
```

### Calculate Overtime Hours
```
=MAX(0, (end_time - start_time) - TIME(8,0,0))
```

### Time Validation
```
=AND(A1>=TIME(9,0,0), A1<=TIME(17,0,0))
```
Validates if time is within business hours

## Advanced Time Functions

### TIMEVALUE Function
```
=TIMEVALUE(time_text)
```
Converts a time in text format to Excel time serial number.

**Parameters:**
- `time_text`: Text representation of time (e.g., "2:30 PM", "14:30")

### NETWORKDAYS with Time
Calculate working time between dates excluding weekends:
```
=(NETWORKDAYS(start_datetime, end_datetime)-1)*(end_time-start_time) + IF(WEEKDAY(end_datetime,2)<6, end_time, 0) - IF(WEEKDAY(start_datetime,2)<6, start_time, 0)
```

### Time Zone Conversion
```
=datetime + TIME(timezone_difference, 0, 0)
```

## Common Time Formats

### Display Formats
- `h:mm AM/PM` - 12-hour format (2:30 PM)
- `hh:mm:ss` - 24-hour format (14:30:00)
- `[h]:mm:ss` - Duration over 24 hours (26:15:30)
- `mm:ss` - Minutes and seconds only

### Duration Calculations
For elapsed time calculations, use custom format `[h]:mm:ss` to prevent Excel from rolling over at 24 hours.
