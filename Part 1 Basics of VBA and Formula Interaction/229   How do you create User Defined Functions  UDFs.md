### 229. **How do you create User Defined Functions (UDFs)?**

**Basic UDF Structure:**

```
Function MyFunction(arg1 As Double, arg2 As Double) As Double
    MyFunction = arg1 * arg2 + arg1 / arg2
End Function
' Use in Excel: =MyFunction(A1, B1)

```

**UDF with Range Arguments:**

```
Function SumPositive(rng As Range) As Double
    Dim cell As Range
    Dim total As Double

    total = 0
    For Each cell In rng
        If IsNumeric(cell.Value) Then
            If cell.Value > 0 Then
                total = total + cell.Value
            End If
        End If
    Next cell

    SumPositive = total
End Function
' Use: =SumPositive(A1:A100)

```

**UDF with Multiple Return Types:**

```
Function CalculateStats(rng As Range, statType As String) As Variant
    Select Case UCase(statType)
        Case "MEAN"
            CalculateStats = WorksheetFunction.Average(rng)
        Case "MEDIAN"
            CalculateStats = WorksheetFunction.Median(rng)
        Case "STDEV"
            CalculateStats = WorksheetFunction.StDev_S(rng)
        Case "COUNT"
            CalculateStats = WorksheetFunction.Count(rng)
        Case Else
            CalculateStats = CVErr(xlErrNA)
    End Select
End Function
' Use: =CalculateStats(A1:A100, "MEAN")

```

**UDF with Optional Arguments:**

```
Function CustomDiscount(amount As Double, _
                        Optional tier As Integer = 1, _
                        Optional bonus As Double = 0) As Double
    Dim discount As Double

    Select Case tier
        Case 1: discount = 0
        Case 2: discount = 0.05
        Case 3: discount = 0.1
        Case 4: discount = 0.15
        Case Else: discount = 0.2
    End Select

    CustomDiscount = amount * (1 - discount) - bonus
End Function
' Use: =CustomDiscount(A1) or =CustomDiscount(A1, 3) or =CustomDiscount(A1, 3, 10)

```
