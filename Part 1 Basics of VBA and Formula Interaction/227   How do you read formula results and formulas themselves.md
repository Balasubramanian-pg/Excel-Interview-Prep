### 227. **How do you read formula results and formulas themselves?**

```
Sub ReadFormulas()
    Dim cell As Range
    Set cell = Range("A1")

    ' Get the formula as text
    Debug.Print cell.Formula
    Debug.Print cell.FormulaR1C1

    ' Get the calculated value
    Debug.Print cell.Value
    Debug.Print cell.Value2  ' More precise, no currency/date formatting

    ' Check if cell contains formula
    If cell.HasFormula Then
        MsgBox "Cell has formula: " & cell.Formula
    End If

    ' Get formula for entire range
    Dim formulaRange As Range
    Set formulaRange = Range("A1:A10")

    ' Check if any cells have formulas
    On Error Resume Next
    Set formulaRange = formulaRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If Not formulaRange Is Nothing Then
        MsgBox "Found formulas in " & formulaRange.Address
    End If
End Sub

```

**Read Different Value Types:**

```
Sub ReadValueTypes()
    Dim cell As Range
    Set cell = Range("A1")

    ' Standard value
    Debug.Print cell.Value

    ' Display format
    Debug.Print cell.Text

    ' Value without formatting
    Debug.Print cell.Value2

    ' For dates
    If IsDate(cell.Value) Then
        Debug.Print "Date: " & cell.Value
        Debug.Print "Serial: " & cell.Value2
    End If

    ' For formulas with errors
    If IsError(cell.Value) Then
        Debug.Print "Error type: " & CVErr(cell.Value)
    End If
End Sub

```
