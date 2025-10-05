### Formulas & Functions

**What's the difference between relative and absolute cell references?**

- **Relative references** (A1) change when you copy a formula to another cell. If you copy =A1 from B1 to B2, it becomes =A2
- **Absolute references** ($A$1) stay fixed when copied. $A$1 remains $A$1 no matter where you paste it
- **Mixed references** ($A1 or A$1) lock either the column or row while allowing the other to change

**Common functions:**

- **SUM(range)**: Adds all numbers in a range. Example: =SUM(A1:A10)
- **AVERAGE(range)**: Calculates the mean of numbers. Example: =AVERAGE(B1:B20)
- **COUNT(range)**: Counts cells containing numbers only
- **COUNTA(range)**: Counts all non-empty cells (numbers, text, dates)
- **COUNTBLANK(range)**: Counts empty cells in a range

**How do you use IF statements?**
Syntax: =IF(logical_test, value_if_true, value_if_false)
Example: =IF(A1>100, "High", "Low") returns "High" if A1 is greater than 100, otherwise "Low"

**What's the difference between COUNT and COUNTA?**

- COUNT only counts cells with numeric values
- COUNTA counts all non-empty cells including text, dates, and numbers
