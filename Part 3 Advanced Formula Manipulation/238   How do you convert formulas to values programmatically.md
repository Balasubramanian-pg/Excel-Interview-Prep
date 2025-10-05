### 238. **How do you convert formulas to values programmatically?**

```
Sub ConvertFormulasToValues()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Method 1: Convert specific range
    Dim rng As Range
    Set rng = Range("A1:A100")

    rng.Value = rng.Value  ' This converts formulas to values

    ' Method 2: Convert only formula cells
    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not formulaRange Is Nothing Then
        formulaRange.Value = formulaRange.Value
    End If

    ' Method 3: Selective conversion with conditions
    Dim cell As Range
    For Each cell In Range("B1:B1000")
        If cell.HasFormula Then
            If IsNumeric(cell.Value) Then
                ' Only convert if result is numeric
                cell.Value = cell.Value
            End If
        End If
    Next cell

    ' Method 4: Convert but keep backup
    Dim backupSheet As Worksheet
    Set backupSheet = Worksheets.Add
    backupSheet.Name = "Backup_" & Format(Now, "yyyymmdd_hhmmss")
    ws.UsedRange.Copy backupSheet.Range("A1")

    ' Now convert originals
    ws.UsedRange.Value = ws.UsedRange.Value
End Sub

```

**Convert Formulas to Values with Logging:**

```
Sub ConvertWithLogging()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim logSheet As Worksheet
    Set logSheet = Worksheets.Add
    logSheet.Name = "Conversion Log"

    logSheet.Range("A1:C1").Value = Array("Cell", "Original Formula", "Converted Value")

    Dim logRow As Long
    logRow = 2

    Dim cell As Range
    For Each cell In ws.Range("A1:Z1000")
        If cell.HasFormula Then
            ' Log the conversion
            logSheet.Cells(logRow, 1).Value = cell.Address
            logSheet.Cells(logRow, 2).Value = "'" & cell.Formula
            logSheet.Cells(logRow, 3).Value = cell.Value

            ' Convert
            cell.Value = cell.Value

```

```
            logRow = logRow + 1
        End If
    Next cell

    MsgBox "Converted " & (logRow - 2) & " formulas to values. Check log sheet."
End Sub

```
