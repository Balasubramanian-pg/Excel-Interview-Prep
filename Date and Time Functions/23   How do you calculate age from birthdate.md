# How to Calculate Age from Birthdate in Excel

This guide explains multiple methods to calculate age from a birthdate, including the hidden DATEDIF function, mathematical approaches, and the YEARFRAC function, with considerations for accuracy and use cases.

## Formula Syntax

### Method 1: DATEDIF Function (Most Accurate)
```
=DATEDIF(birthdate, TODAY(), "Y")
```

**Parameters:**
- `birthdate`: Cell reference containing the birthdate
- `TODAY()`: Current date (updates automatically)
- `"Y"`: Returns complete years between dates

### Method 2: Mathematical Calculation
```
=INT((TODAY() - birthdate) / 365.25)
```

**Parameters:**
- `TODAY() - birthdate`: Calculates total days between dates
- `365.25`: Accounts for leap years (average days per year)
- `INT()`: Returns integer portion, truncating decimal

### Method 3: YEARFRAC Function
```
=INT(YEARFRAC(birthdate, TODAY(), 1))
```

**Parameters:**
- `birthdate`: Start date for calculation
- `TODAY()`: End date for calculation
- `1`: Basis for day count (actual/actual)
- `INT()`: Converts fractional years to whole number

## Worked Examples

Given birthdate in cell A1: `March 15, 1980`
Current date: `March 15, 2024`

**Method 1 - DATEDIF:**
```
=DATEDIF(A1, TODAY(), "Y")
```
Returns: `44` (exactly 44 years)

**Method 2 - Mathematical:**
```
=INT((TODAY() - A1) / 365.25)
```
Calculation: `(44 * 365.25) / 365.25 = 44`
Returns: `44`

**Method 3 - YEARFRAC:**
```
=INT(YEARFRAC(A1, TODAY(), 1))
```
Returns: `44`

### Testing Edge Cases
**Birthdate:** February 29, 1980 (Leap Year)
**Current date:** February 28, 2024

**DATEDIF:**
```
=DATEDIF(A2, TODAY(), "Y")
```
Returns: `43` (hasn't reached birthday this year)

**Mathematical:**
```
=INT((TODAY() - A2) / 365.25)
```
Returns: `43`

**YEARFRAC:**
```
=INT(YEARFRAC(A2, TODAY(), 1))
```
Returns: `43`

> [!NOTE]
> DATEDIF is a hidden function in Excel that doesn't appear in function autocomplete but works in all versions. It's the most accurate method as it correctly handles leap years and month boundaries.

> [!IMPORTANT]
> The mathematical method using 365.25 is an approximation and may be off by a day in some cases. For precise age calculations, especially for legal or medical purposes, use DATEDIF or YEARFRAC.

> [!TIP]
> To display age with years, months, and days:
> `=DATEDIF(A1,TODAY(),"Y")&" years, "&DATEDIF(A1,TODAY(),"YM")&" months, "&DATEDIF(A1,TODAY(),"MD")&" days"`

## Advanced Age Calculations

### Age with Months and Days
```
=DATEDIF(birthdate, TODAY(), "Y") & " Years, " & DATEDIF(birthdate, TODAY(), "YM") & " Months, " & DATEDIF(birthdate, TODAY(), "MD") & " Days"
```

### Age as Decimal (for precise calculations)
```
=YEARFRAC(birthdate, TODAY(), 1)
```
Returns: `44.246` (for example)

### Age on Specific Date
```
=DATEDIF(birthdate, target_date, "Y")
```
Calculate age as of a specific past or future date.

### Age with Conditional Formatting
Highlight minors (under 18):
```
=DATEDIF(A1, TODAY(), "Y") < 18
```

## DATEDIF Function Details

### Complete Syntax
```
=DATEDIF(start_date, end_date, unit)
```

**Unit Codes:**
- `"Y"`: Complete years between dates
- `"M"`: Complete months between dates
- `"D"`: Complete days between dates
- `"YM"`: Months excluding years
- `"YD"`: Days excluding years
- `"MD"`: Days excluding years and months

### Common DATEDIF Patterns
**Years only:**
```
=DATEDIF(A1, TODAY(), "Y")
```

**Months only:**
```
=DATEDIF(A1, TODAY(), "M")
```

**Days only:**
```
=DATEDIF(A1, TODAY(), "D")
```

## Alternative Methods

### Using DATE and YEAR Functions
```
=YEAR(TODAY()) - YEAR(birthdate) - IF(TODAY() < DATE(YEAR(TODAY()), MONTH(birthdate), DAY(birthdate)), 1, 0)
```
This method explicitly checks if birthday has occurred this year.

### Using ROUNDDOWN for Age
```
=ROUNDDOWN(YEARFRAC(birthdate, TODAY(), 1), 0)
```
Alternative to INT() that may be more intuitive.

## Error Handling

### Handling Future Birthdates
```
=IF(birthdate > TODAY(), "Future date", DATEDIF(birthdate, TODAY(), "Y"))
```

### Handling Invalid Dates
```
=IF(ISDATE(birthdate), DATEDIF(birthdate, TODAY(), "Y"), "Invalid date")
```

### Blank Cell Handling
```
=IF(birthdate = "", "", DATEDIF(birthdate, TODAY(), "Y"))
```

## Performance Considerations

- **DATEDIF**: Fastest and most accurate
- **YEARFRAC**: Slightly slower but very accurate
- **Mathematical**: Fast but approximate
- Choose based on your accuracy requirements and dataset size
