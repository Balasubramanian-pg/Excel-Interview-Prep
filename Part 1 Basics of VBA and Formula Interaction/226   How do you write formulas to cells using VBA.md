### 226. **How do you write formulas to cells using VBA?**

**Method 1 - Direct Formula Assignment:**

```
Sub WriteFormula()
    ' Write formula to single cell
    Range("A1").Formula = "=SUM(B1:B10)"

    ' Write formula with absolute references
    Range("A2").Formula = "=SUM($B$1:$B$10)"

    ' Write array formula (legacy)
    Range("A3").FormulaArray = "=SUM(B1:B10*C1:C10)"
End Sub

```

**Method 2 - FormulaR1C1 (Relative References):**

```
Sub WriteFormulaR1C1()
    ' R1C1 notation - more flexible for copying
    Range("A1").FormulaR1C1 = "=SUM(R1C2:R10C2)"

    ' Relative reference (current row, column 2)
    Range("A1:A100").FormulaR1C1 = "=RC[1]*RC[2]"
    ' Multiplies column B * column C for each row

    ' Mix of absolute and relative
    Range("D1:D100").FormulaR1C1 = "=RC[-1]*R1C1"
    ' Column C * value in A1
End Sub

```

**Method 3 - Formula2 (Excel 365 Dynamic Arrays):**

```
Sub WriteFormula2()
    ' Dynamic array formulas
    Range("A1").Formula2 = "=FILTER(B:B,C:C>100)"

    ' XLOOKUP
    Range("A1").Formula2 = "=XLOOKUP(D1,B:B,C:C,""Not Found"")"

    ' LET function
    Range("A1").Formula2 = "=LET(x,B1,y,C1,x*y+x/y)"
End Sub

```

**Best Practices:**

```
Sub FormulaBestPractices()
    Dim rng As Range
    Set rng = Range("A1:A1000")

    ' Turn off calculation during bulk operations
    Application.Calculation = xlCalculationManual

    ' Write formulas
    rng.FormulaR1C1 = "=RC[1]*RC[2]"

    ' Turn calculation back on
    Application.Calculation = xlCalculationAutomatic

    ' Force calculation
    Application.Calculate
End Sub

```
