### 246. **How do you audit and trace formulas programmatically?**

```
Sub AuditFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cell As Range
    Set cell = Selection.Cells(1, 1)

    If cell.HasFormula Then
        ' Get precedents (cells this formula depends on)
        On Error Resume Next
        Dim precedents As Range
        Set precedents = cell.Precedents

        If Not precedents Is Nothing Then
            Debug.Print "Precedents of " & cell.Address & ":"
            Debug.Print precedents.Address(External:=True)
            precedents.Interior.Color = RGB(255, 255, 200)  ' Highlight yellow
        Else
            Debug.Print "No direct precedents"
        End If
        On Error GoTo 0

        ' Get dependents (cells that depend on this cell)
        On Error Resume Next
        Dim dependents As Range
        Set dependents = cell.Dependents

        If Not dependents Is Nothing Then
            Debug.Print "Dependents of " & cell.Address & ":"
            Debug.Print dependents.Address
            dependents.Interior.Color = RGB(200, 255, 200)  ' Highlight green
        Else
            Debug.Print "No direct dependents"
        End If
        On Error GoTo 0

        ' Show formula in message box
        MsgBox "Formula: " & cell.Formula & vbCrLf & vbCrLf & _
               "Result: " & cell.Value, vbInformation, "Formula Audit"
    Else
        MsgBox "Selected cell does not contain a formula", vbInformation
    End If
End Sub

```

**Comprehensive Formula Audit Report:**

```
Sub CreateFormulaAuditReport()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim reportWs As Worksheet
    Set reportWs = Worksheets.Add
    reportWs.Name = "Formula Audit_" & Format(Now, "hhmmss")

    ' Headers
    reportWs.Range("A1:F1").Value = Array("Cell", "Formula", "Value", "Has Error", _
                                          "Precedents", "Dependents")
    reportWs.Range("A1:F1").Font.Bold = True

    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaRange Is Nothing Then
        MsgBox "No formulas found", vbInformation
        Exit Sub
    End If

    Dim cell As Range
    Dim reportRow As Long
    reportRow = 2

    Application.ScreenUpdating = False

    For Each cell In formulaRange
        reportWs.Cells(reportRow, 1).Value = cell.Address
        reportWs.Cells(reportRow, 2).Value = "'" & cell.Formula  ' Prefix with ' to show as text
        reportWs.Cells(reportRow, 3).Value = cell.Text
        reportWs.Cells(reportRow, 4).Value = IsError(cell.Value)

        ' Get precedents
        On Error Resume Next
        Dim prec As Range
        Set prec = cell.Precedents
        If Not prec Is

```

```
        If Not prec Is Nothing Then
            reportWs.Cells(reportRow, 5).Value = prec.Address(External:=True)
        Else
            reportWs.Cells(reportRow, 5).Value = "None"
        End If
        Set prec = Nothing
        On Error GoTo 0

        ' Get dependents
        On Error Resume Next
        Dim deps As Range
        Set deps = cell.Dependents
        If Not deps Is Nothing Then
            reportWs.Cells(reportRow, 6).Value = deps.Address
        Else
            reportWs.Cells(reportRow, 6).Value = "None"
        End If
        Set deps = Nothing
        On Error GoTo 0

        ' Highlight errors
        If IsError(cell.Value) Then
            reportWs.Rows(reportRow).Interior.Color = RGB(255, 200, 200)
        End If

        reportRow = reportRow + 1
    Next cell

    ' Auto-fit columns
    reportWs.Columns("A:F").AutoFit

    ' Add summary
    reportWs.Range("H1").Value = "Summary"
    reportWs.Range("H2").Value = "Total Formulas:"
    reportWs.Range("I2").Value = reportRow - 2
    reportWs.Range("H3").Value = "Formulas with Errors:"
    reportWs.Range("I3").Formula = "=COUNTIF(D:D,TRUE)"
    reportWs.Range("H4").Value = "Unique Formula Types:"
    reportWs.Range("I4").Formula = "=COUNTA(UNIQUE(B:B))-1"

    Application.ScreenUpdating = True

    MsgBox "Audit complete. Report created in new sheet.", vbInformation
End Sub

```
