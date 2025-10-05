### 237. **How do you handle formula errors programmatically?**

```
Sub HandleFormulaErrors()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim cell As Range

    For Each cell In ws.Range("A1:A100")
        If cell.HasFormula Then
            ' Check for errors
            If IsError(cell.Value) Then
                Select Case CVErr(cell.Value)
                    Case CVErr(xlErrDiv0)
                        Debug.Print cell.Address & ": Division by zero"
                        cell.Formula = "=IFERROR(" & Mid(cell.Formula, 2) & ",0)"

                    Case CVErr(xlErrNA)
                        Debug.Print cell.Address & ": #N/A error"
                        cell.Formula = "=IFNA(" & Mid(cell.Formula, 2) & ",""Not Found"")"

                    Case CVErr(xlErrName)
                        Debug.Print cell.Address & ": #NAME? error (invalid formula)"

                    Case CVErr(xlErrNull)
                        Debug.Print cell.Address & ": #NULL! error"

                    Case CVErr(xlErrNum)
                        Debug.Print cell.Address & ": #NUM! error"

                    Case CVErr(xlErrRef)
                        Debug.Print cell.Address & ": #REF! error (invalid reference)"

                    Case CVErr(xlErrValue)
                        Debug.Print cell.Address & ": #VALUE! error"
                End Select
            End If
        End If
    Next cell
End Sub

```

**Add Error Handling to All Formulas:**

```
Sub AddErrorHandlingToAllFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim formulaCell As Range
    Dim formulaRange As Range

    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not formulaRange Is Nothing Then
        Application.Calculation = xlCalculationManual
        Application.ScreenUpdating = False

        For Each formulaCell In formulaRange
            Dim currentFormula As String
            currentFormula = formulaCell.Formula

            ' Only add if not already wrapped
            If Left(currentFormula, 9) <> "=IFERROR(" Then
                formulaCell.Formula = "=IFERROR(" & Mid(currentFormula, 2) & ","""")"
            End If
        Next formulaCell

        Application.Calculation = xlCalculationAutomatic
        Application.ScreenUpdating = True
    End If
End Sub

```
