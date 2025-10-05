### 249. **How do you create formulas for pivot-like aggregations?**

```
Sub CreatePivotLikeFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Assume data in columns A (Category), B (SubCategory), C (Amount)

    ' Create summary sheet
    Dim summaryWs As Worksheet
    Set summaryWs = Worksheets.Add
    summaryWs.Name = "Aggregation Summary"

    ' Get unique categories
    summaryWs.Range("A1").Value = "Category"
    summaryWs.Range("B1").Value = "Total"
    summaryWs.Range("C1").Value = "Count"
    summaryWs.Range("D1").Value = "Average"

    ' Excel 365: Use UNIQUE to get categories
    If Val(Application.Version) >= 16 Then
        summaryWs.Range("A2").Formula2 = "=SORT(UNIQUE(FILTER(" & ws.Name & "!A:A," & ws.Name & "!A:A<>"""")))"

        ' Corresponding aggregations
        summaryWs.Range("B2").Formula2 = "=SUMIF(" & ws.Name & "!$A:$A,A2," & ws.Name & "!$C:$C)"
        summaryWs.Range("C2").Formula2 = "=COUNTIF(" & ws.Name & "!$A:$A,A2)"
        summaryWs.Range("D2").Formula2 = "=AVERAGEIF(" & ws.Name & "!$A:$A,A2," & ws.Name & "!$C:$C)"

        ' Copy formulas down automatically with spill
        ' The # reference will expand with unique values
    Else
        ' Older Excel: Manual approach
        ' Would need to list categories manually or use advanced array formulas
        MsgBox "For full automation, Excel 365 is recommended", vbInformation
    End If

    ' Two-way aggregation (Category by SubCategory)
    summaryWs.Range("F1").Value = "Category \ SubCategory"

    ' Get unique subcategories across top
    If Val(Application.Version) >= 16 Then
        summaryWs.Range("G1").Formula2 = "=TRANSPOSE(SORT(UNIQUE(FILTER(" & ws.Name & "!B:B," & ws.Name & "!B:B<>""""))))"

        ' Matrix formula for intersections
        summaryWs.Range("G2").Formula2 = _
            "=SUMIFS(" & ws.Name & "!$C:$C," & ws.Name & "!$A:$A,$F2," & ws.Name & "!$B:$B,G$1)"
    End If

    summaryWs.Columns.AutoFit
End Sub

```

**Advanced Cross-Tab with VBA:**

```
Sub CreateCrossTabulation()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim outputWs As Worksheet
    Set outputWs = Worksheets.Add
    outputWs.Name = "CrossTab"

    ' Get unique row and column values
    Dim rowCategories As Object
    Set rowCategories = CreateObject("Scripting.Dictionary")

    Dim colCategories As Object
    Set colCategories = CreateObject("Scripting.Dictionary")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    Dim i As Long
    For i = 2 To lastRow
        If Not rowCategories.exists(ws.Cells(i, 1).Value) Then
            rowCategories.Add ws.Cells(i, 1).Value, Nothing
        End If

        If Not colCategories.exists(ws.Cells(i, 2).Value) Then
            colCategories.Add ws.Cells(i, 2).Value, Nothing
        End If
    Next i

    ' Write headers
    outputWs.Range("A1").Value = "Category"

    Dim col As Long
    col = 2
    Dim key As Variant

    For Each key In colCategories.Keys
        outputWs.Cells(1, col).Value = key
        col = col + 1
    Next key

    ' Write row categories
    Dim row As Long
    row = 2
    For Each key In rowCategories.Keys
        outputWs.Cells(row, 1).Value = key
        row = row + 1
    Next key

    ' Create SUMIFS formulas for each cell
    For row = 2 To rowCategories.Count + 1
        For col = 2 To colCategories.Count + 1
            outputWs.Cells(row, col).Formula = _
                "=SUMIFS(" & ws.Name & "!$C:$C," & _
                ws.Name & "!$A:$A,$A" & row & "," & _
                ws.Name & "!$B:$B," & outputWs.Cells(1, col).Address(False, True) & ")"
        Next col
    Next row

    ' Add totals
    Dim totalCol As Long
    totalCol = colCategories.Count + 2

    outputWs.Cells(1, totalCol).Value = "Total"
    outputWs.Cells(1, totalCol).Font.Bold = True

    For row = 2 To rowCategories.Count + 1
        outputWs.Cells(row, totalCol).Formula = _
            "=SUM(" & outputWs.Cells(row, 2).Address & ":" & _
            outputWs.Cells(row, totalCol - 1).Address & ")"
    Next row

    ' Add row totals
    Dim totalRow As Long
    totalRow = rowCategories.Count + 2
    outputWs.Cells(totalRow, 1).Value = "Total"
    outputWs.Cells(totalRow, 1).Font.Bold = True

    For col = 2 To totalCol
        outputWs.Cells(totalRow, col).Formula = _
            "=SUM(" & outputWs.Cells(2, col).Address & ":" & _
            outputWs.Cells(totalRow - 1, col).Address & ")"
    Next col

    ' Format
    outputWs.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    outputWs.Rows(1).Font.Bold = True
    outputWs.Columns.AutoFit
End Sub

```
