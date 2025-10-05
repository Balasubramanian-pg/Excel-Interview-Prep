### 248. **How do you create self-updating formulas?**

```
Sub CreateSelfUpdatingFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Create dynamic named range that expands with data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Method 1: Dynamic named range with OFFSET
    ws.Names.Add Name:="DynamicData", _
                 RefersTo:="=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)"

    ' Use in formula
    ws.Range("B1").Formula = "=SUM(DynamicData)"
    ws.Range("B2").Formula = "=AVERAGE(DynamicData)"
    ws.Range("B3").Formula = "=COUNTA(DynamicData)"

    ' Method 2: Table-based (best practice)
    Dim tbl As ListObject
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C" & lastRow), , xlYes)
    tbl.Name = "DataTable"
    On Error GoTo 0

    ' Formulas automatically adjust with table
    ws.Range("E1").Formula = "=SUM(DataTable[Amount])"
    ws.Range("E2").Formula = "=AVERAGE(DataTable[Amount])"

    ' Method 3: Excel 365 dynamic arrays (spill ranges)
    If Val(Application.Version) >= 16 Then
        ws.Range("F1").Formula2 = "=FILTER(A:A,A:A<>"""")"
        ws.Range("G1").Formula = "=SUM(F1#)"  ' # references spill range
    End If

    ' Method 4: INDEX with COUNTA for last value
    ws.Range("H1").Formula = "=INDEX(A:A,COUNTA(A:A))"  ' Last value
    ws.Range("H2").Formula = "=INDEX(A:A,COUNTA(A:A)-1)"  ' Second to last
End Sub

```

**Auto-Expanding Summary Report:**

```
Sub CreateAutoExpandingReport()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Create summary that automatically includes all data
    Dim summaryWs As Worksheet
    Set summaryWs = Worksheets.Add
    summaryWs.Name = "Summary"

    ' Headers
    summaryWs.Range("A1:C1").Value = Array("Metric", "Formula", "Value")
    summaryWs.Range("A1:C1").Font.Bold = True

    ' Auto-expanding formulas
    summaryWs.Range("A2").Value = "Count"
    summaryWs.Range("B2").Formula = "=COUNTA(" & ws.Name & "!A:A)-1"
    summaryWs.Range("C2").Formula = "=B2"

    summaryWs.Range("A3").Value = "Sum"
    summaryWs.Range("B3").Formula = "=SUBTOTAL(9," & ws.Name & "!B:B)"
    summaryWs.Range("C3").Formula = "=B3"

    summaryWs.Range("A4").Value = "Average"
    summaryWs.Range("B4").Formula = "=SUBTOTAL(1," & ws.Name & "!B:B)"
    summaryWs.Range("C4").Formula = "=B4"

    summaryWs.Range("A5").Value = "Max"
    summaryWs.Range("B5").Formula = "=SUBTOTAL(4," & ws.Name & "!B:B)"
    summaryWs.Range("C5").Formula = "=B5"

    summaryWs.Range("A6").Value = "Min"
    summaryWs.Range("B6").Formula = "=SUBTOTAL(5," & ws.Name & "!B:B)"
    summaryWs.Range("C6").Formula = "=B6"

    summaryWs.Range("A7").Value = "Std Dev"
    summaryWs.Range("B7").Formula = "=STDEV.S(" & ws.Name & "!B:B)"
    summaryWs.Range("C7").Formula = "=B7"

    ' Last Updated timestamp (volatile)
    summaryWs.Range("A9").Value = "Last Updated:"
    summaryWs.Range("B9").Formula = "=NOW()"
    summaryWs.Range("B9").NumberFormat = "yyyy-mm-dd hh:mm:ss"

    summaryWs.Columns("A:C").AutoFit
End Sub

```
