### 250. **How do you optimize formula performance with VBA?**

```
Sub OptimizeFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' Optimization 1: Replace SUMIF with direct SUM where possible
    Dim cell As Range
    For Each cell In ws.UsedRange.SpecialCells(xlCellTypeFormulas)
        Dim formula As String
        formula = cell.Formula

        ' Replace inefficient patterns
        If InStr(formula, "SUMIF") > 0 And InStr(formula, "*") > 0 Then
            ' Check if can be replaced with SUMPRODUCT
            Debug.Print "Consider optimizing: " & cell.Address
        End If
    Next cell

    ' Optimization 2: Convert entire column references to specific ranges
    Dim formulaRange As Range
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)

    Dim lastDataRow As Long
    lastDataRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    For Each cell In formulaRange
        formula = cell.Formula

        ' Replace A:A with A1:A[lastrow]
        If InStr(formula, "A:A") > 0 Then
            cell.Formula = Replace(formula, "A:A", "A1:A" & lastDataRow)
        End If

        ' Similar for other columns
        If InStr(formula, "B:B") > 0 Then
            cell.Formula = Replace(cell.Formula, "B:B", "B1:B" & lastDataRow)
        End If
    Next cell

    ' Optimization 3: Replace volatile functions where possible
    For Each cell In formulaRange
        formula = cell.Formula

        ' Replace OFFSET with INDEX where possible
        If InStr(formula, "OFFSET") > 0 Then
            Debug.Print "Volatile function (OFFSET) in: " & cell.Address
        End If

        If InStr(formula, "INDIRECT") > 0 Then
            Debug.Print "Volatile function (INDIRECT) in: " & cell.Address
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Formula optimization scan complete. Check Immediate window for details.", vbInformation
End Sub

```

**Convert Array Formulas to Regular Formulas:**

```
Sub ConvertArrayFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cell As Range
    Dim convertCount As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    For Each cell In ws.UsedRange
        If cell.HasArray And Not cell.HasFormula Then
            ' This is part of an array formula but not the top-left cell
            ' Skip it
        ElseIf cell.HasArray And cell.HasFormula Then
            ' This is the top-left cell of an array formula
            Dim arrayFormula As String
            arrayFormula = cell.FormulaArray

            Dim arrayRange As Range
            Set arrayRange = cell.CurrentArray

            ' Check if it's a simple array that can be converted
            If InStr(arrayFormula, "SUM(IF(") > 0 Then
                ' Can potentially convert to SUMIFS
                Debug.Print "Array formula in " & cell.Address & " may be convertible to SUMIFS"
            End If

            ' Example conversion: =SUM(IF(A:A="X",B:B,0)) to =SUMIF(A:A,"X",B:B)
            If InStr(arrayFormula, "=SUM(IF(") > 0 And InStr(arrayFormula, ",0))") > 0 Then
                ' Parse and rebuild as SUMIF (simplified logic)
                ' In practice, this requires more sophisticated parsing
                Debug.Print "Can convert: " & arrayFormula
            End If
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub

```
