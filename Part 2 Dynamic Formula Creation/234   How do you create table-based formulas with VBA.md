### 234. **How do you create table-based formulas with VBA?**

```
Sub CreateTableFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim tbl As ListObject
    Dim lastRow As Long

    ' Create table if doesn't exist
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    On Error Resume Next
    Set tbl = ws.ListObjects("SalesTable")
    On Error GoTo 0

    If tbl Is Nothing Then
        Set tbl = ws.ListObjects.Add(xlSrcRange, ws.Range("A1:C" & lastRow), , xlYes)
        tbl.Name = "SalesTable"
        tbl.TableStyle = "TableStyleMedium2"
    End If

    ' Add calculated column using structured references
    If tbl.ListColumns.Count = 3 Then
        tbl.ListColumns.Add
        tbl.ListColumns(4).Name = "Total"
    End If

    ' Use structured references in formulas
    tbl.ListColumns("Total").DataBodyRange.Formula = _
        "=[@Quantity]*[@Price]"

    ' Add another calculated column with conditional logic
    tbl.ListColumns.Add
    tbl.ListColumns(5).Name = "Discount"
    tbl.ListColumns("Discount").DataBodyRange.Formula = _
        "=IF([@Total]>1000,[@Total]*0.1,0)"

    ' Final total column
    tbl.ListColumns.Add
    tbl.ListColumns(6).Name = "Net"
    tbl.ListColumns("Net").DataBodyRange.Formula = _
        "=[@Total]-[@Discount]"

    ' Add totals row
    tbl.ShowTotals = True
    tbl.ListColumns("Total").TotalsCalculation = xlTotalsCalculationSum
    tbl.ListColumns("Discount").TotalsCalculation = xlTotalsCalculationSum
    tbl.ListColumns("Net").TotalsCalculation = xlTotalsCalculationSum
End Sub

```
