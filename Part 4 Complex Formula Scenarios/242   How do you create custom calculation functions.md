### 242. **How do you create custom calculation functions?**

```
' Advanced UDF with multiple features
Function CustomCalc(sales As Range, costs As Range, _
                   Optional taxRate As Double = 0.1, _
                   Optional roundDigits As Integer = 2) As Variant

    On Error GoTo ErrorHandler

    ' Validate inputs
    If sales.Rows.Count <> costs.Rows.Count Then
        CustomCalc = CVErr(xlErrRef)
        Exit Function
    End If

    Dim result() As Variant
    ReDim result(1 To sales.Rows.Count, 1 To 1)

    Dim i As Long
    For i = 1 To sales.Rows.Count
        If IsNumeric(sales.Cells(i, 1).Value) And _
           IsNumeric(costs.Cells(i, 1).Value) Then

            Dim profit As Double
            profit = sales.Cells(i, 1).Value - costs.Cells(i, 1).Value

            Dim afterTax As Double
            afterTax = profit * (1 - taxRate)

            result(i, 1) = Round(afterTax, roundDigits)
        Else
            result(i, 1) = CVErr(xlErrValue)
        End If
    Next i

    ' Return array for multiple results
    CustomCalc = result
    Exit Function

ErrorHandler:
    CustomCalc = CVErr(xlErrValue)
End Function

' Use: =CustomCalc(A2:A10, B2:B10, 0.15, 2)

```

**UDF with Worksheet Functions:**

```
Function AdvancedLookup(lookupValue As Variant, _
                       searchRange As Range, _
                       returnRange As Range, _
                       Optional defaultValue As Variant = "Not Found") As Variant

    On Error GoTo ErrorHandler

    ' Use WorksheetFunction for robust lookup
    Dim result As Variant

    ' Try XLOOKUP first (Excel 365)
    On Error Resume Next
    result = Application.WorksheetFunction.XLookup(lookupValue, _
                                                   searchRange, _
                                                   returnRange, _
                                                   defaultValue)

    If Err.Number <> 0 Then
        ' Fall back to INDEX-MATCH
        Err.Clear
        Dim matchResult As Variant
        matchResult = Application.WorksheetFunction.Match(lookupValue, _
                                                         searchRange, 0)

        If Not IsError(matchResult) Then
            result = Application.WorksheetFunction.Index(returnRange, matchResult)
        Else
            result = defaultValue
        End If
    End If
    On Error GoTo 0

    AdvancedLookup = result
    Exit Function

ErrorHandler:
    AdvancedLookup = defaultValue
End Function

' Use: =AdvancedLookup(A2, $D$2:$D$100, $E$2:$E$100, "N/A")

```

**UDF with Array Return:**

```
Function GetStats(dataRange As Range) As Variant
    ' Returns array of statistics: Count, Sum, Average, Min, Max, StdDev

    Dim result(1 To 6, 1 To 2) As Variant

    On Error GoTo ErrorHandler

    With Application.WorksheetFunction
        result(1, 1) = "Count"
        result(1, 2) = .Count(dataRange)

        result(2, 1) = "Sum"
        result(2, 2) = .Sum(dataRange)

        result(3, 1) = "Average"
        result(3, 2) = .Average(dataRange)

        result(4, 1) = "Min"
        result(4, 2) = .Min(dataRange)

        result(5, 1) = "Max"
        result(5, 2) = .Max(dataRange)

        result(6, 1) = "StdDev"
        result(6, 2) = .StDev_S(dataRange)
    End With

    GetStats = result
    Exit Function

ErrorHandler:
    GetStats = CVErr(xlErrValue)
End Function

' Use: Select 6 rows x 2 columns, type =GetStats(A1:A100), press Ctrl+Shift+Enter

```
