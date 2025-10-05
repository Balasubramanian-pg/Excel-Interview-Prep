### 233. **How do you loop through ranges and apply formulas?**

```
Sub LoopAndApplyFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cell As Range
    Dim lastRow As Long

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Method 1: Loop through each cell
    For Each cell In ws.Range("D1:D" & lastRow)
        ' Formula references the same row
        cell.Formula = "=B" & cell.Row & "*C" & cell.Row
    Next cell

    ' Method 2: Using For loop with row counter
    Dim i As Long
    For i = 2 To lastRow
        ws.Cells(i, 5).Formula = "=IF(A" & i & ">100,B" & i & "*1.1,B" & i & ")"
    Next i

    ' Method 3: Apply formula to entire range at once (faster)
    ws.Range("F2:F" & lastRow).FormulaR1C1 = "=RC[-4]*RC[-3]"

    ' Method 4: Conditional formula application
    For i = 2 To lastRow
        If ws.Cells(i, 1).Value = "Active" Then
            ws.Cells(i, 6).Formula = "=B" & i & "*C" & i
        Else
            ws.Cells(i, 6).Value = 0
        End If
    Next i
End Sub

```

**Advanced Looping with Array Formulas:**

```
Sub LoopWithArrays()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim dataArray As Variant
    Dim resultArray() As Variant
    Dim lastRow As Long
    Dim i As Long

    ' Turn off calculation for performance
    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Read data into array (faster than cell-by-cell)
    dataArray = ws.Range("A1:C" & lastRow).Value

    ' Resize result array
    ReDim resultArray(1 To UBound(dataArray, 1), 1 To 1)

    ' Process in memory
    For i = 1 To UBound(dataArray, 1)
        If IsNumeric(dataArray(i, 2)) And IsNumeric(dataArray(i, 3)) Then
            resultArray(i, 1) = dataArray(i, 2) * dataArray(i, 3)
        Else
            resultArray(i, 1) = ""
        End If
    Next i

    ' Write results back (single operation)
    ws.Range("D1").Resize(UBound(resultArray, 1), 1).Value = resultArray

    ' Alternative: Apply formula to range instead
    ws.Range("E1:E" & lastRow).Formula = "=B1*C1"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

```
