### 251. **How do you create formula-based dashboards?**

```
Sub CreateFormulaDashboard()
    Dim dashWs As Worksheet
    Set dashWs = Worksheets.Add
    dashWs.Name = "Dashboard"

    Dim dataWs As Worksheet
    Set dataWs = Worksheets("Data")  ' Assume data sheet exists

    ' Title
    With dashWs.Range("A1:J1")
        .Merge
        .Value = "Sales Dashboard"
        .Font.Size = 18
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' KPI Cards
    Dim kpiRow As Long
    kpiRow = 3

    ' Total Sales
    CreateKPICard dashWs, "B" & kpiRow, "Total Sales", _
        "=SUM(Data!C:C)", "$#,##0", RGB(68, 114, 196)

    ' Average Sale
    CreateKPICard dashWs, "E" & kpiRow, "Average Sale", _
        "=AVERAGE(Data!C:C)", "$#,##0", RGB(112, 173, 71)

    ' Transaction Count
    CreateKPICard dashWs, "H" & kpiRow, "Transactions", _
        "=COUNTA(Data!A:A)-1", "#,##0", RGB(255, 192, 0)

    ' Growth Rate
    kpiRow = kpiRow + 4
    CreateKPICard dashWs, "B" & kpiRow, "Growth vs Last Month", _
        "=(SUM(Data!C:C)-SUM(LastMonth!C:C))/SUM(LastMonth!C:C)", "0.0%", RGB(237, 125, 49)

    ' Top Product
    CreateKPICard dashWs, "E" & kpiRow, "Top Product", _
        "=INDEX(Data!B:B,MATCH(MAX(Data!C:C),Data!C:C,0))", "@", RGB(165, 165, 165)

    ' Conversion Rate
    CreateKPICard dashWs, "H" & kpiRow, "Conversion Rate", _
        "=COUNTIF(Data!D:D,""Closed"")/COUNTA(Data!D:D)", "0.0%", RGB(68, 114, 196)

    ' Dynamic Date Range
    dashWs.Range("B11").Value = "Showing data from:"
    dashWs.Range("C11").Formula = "=MIN(Data!A:A)"
    dashWs.Range("C11").NumberFormat = "yyyy-mm-dd"

    dashWs.Range("E11").Value = "to:"
    dashWs.Range("F11").Formula = "=MAX(Data!A:A)"
    dashWs.Range("F11").NumberFormat = "yyyy-mm-dd"

    ' Top 5 Products Table
    dashWs.Range("B13").Value = "Top 5 Products"
    dashWs.Range("B13").Font.Bold = True
    dashWs.Range("B13").Font.Size = 14

    If Val(Application.Version) >= 16 Then
        ' Use dynamic arrays
        dashWs.Range("B14").Formula2 = _
            "=TAKE(SORT(UNIQUE(Data!B:B),1,-1,SUMIF(Data!B:B,UNIQUE(Data!B:B),Data!C:C),-1),5)"
    End If

    ' Format
    dashWs.Columns.AutoFit
    dashWs.Tab.Color = RGB(68, 114, 196)
End Sub

Sub CreateKPICard(ws As Worksheet, topLeftCell As String, title As String, _
                  formula As String, numFormat As String, color As Long)

    Dim rng As Range
    Set rng = ws.Range(topLeftCell).Resize(3, 2)

    ' Border
    With rng.Borders
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .Color = color
    End With

    ' Title
    With rng.Cells(1, 1).Resize(1, 2)
        .Merge
        .Value = title
        .Font.Bold = True
        .Interior.Color = color
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With

    ' Value
    With rng.Cells(2, 1).Resize(2, 2)
        .Merge
        .Formula = formula
        .NumberFormat = numFormat
        .Font.Size = 24
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

```

---
