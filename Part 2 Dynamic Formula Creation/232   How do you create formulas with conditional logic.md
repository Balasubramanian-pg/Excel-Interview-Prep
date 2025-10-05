### 232. **How do you create formulas with conditional logic?**

```
Sub ConditionalFormulaCreation()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim useComplexFormula As Boolean
    useComplexFormula = True  ' Could be based on user input or data conditions

    If useComplexFormula Then
        ' Complex nested IF
        Range("A1").Formula = "=IFS(B1>100,""High"",B1>50,""Medium"",B1>0,""Low"",TRUE,""None"")"
    Else
        ' Simple IF
        Range("A1").Formula = "=IF(B1>50,""High"",""Low"")"
    End If

    ' Choose formula based on Excel version
    If Val(Application.Version) >= 16 Then
        ' Use XLOOKUP for Excel 365/2021
        Range("C1").Formula = "=XLOOKUP(A1,D:D,E:E,""Not Found"")"
    Else
        ' Use VLOOKUP for older versions
        Range("C1").Formula = "=IFERROR(VLOOKUP(A1,D:E,2,FALSE),""Not Found"")"
    End If

    ' Dynamic calculation method
    Dim calcMethod As String
    calcMethod = Range("Settings!A1").Value

    Select Case calcMethod
        Case "Average"
            Range("Result").Formula = "=AVERAGE(Data)"
        Case "Median"
            Range("Result").Formula = "=MEDIAN(Data)"
        Case "Weighted"
            Range("Result").Formula = "=SUMPRODUCT(Data,Weights)/SUM(Weights)"
    End Select
End Sub

```
