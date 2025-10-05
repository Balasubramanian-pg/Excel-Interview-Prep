### 247. **How do you find and replace within formulas?**

```
Sub FindReplaceInFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim findText As String
    Dim replaceText As String

    findText = InputBox("Find text in formulas:", "Find")
    If findText = "" Then Exit Sub

    replaceText = InputBox("Replace with:", "Replace")

    Dim cell As Range
    Dim formulaRange As Range
    Dim changeCount As Long

    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaRange Is Nothing Then
        MsgBox "No formulas found", vbInformation
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    changeCount = 0

    For Each cell In formulaRange
        Dim originalFormula As String
        originalFormula = cell.Formula

        If InStr(1, originalFormula, findText, vbTextCompare) > 0 Then
            ' Replace in formula
            cell.Formula = Replace(originalFormula, findText, replaceText, , , vbTextCompare)
            changeCount = changeCount + 1

            ' Log the change
            Debug.Print "Changed " & cell.Address & ":"
            Debug.Print "  From: " & originalFormula
            Debug.Print "  To:   " & cell.Formula
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    MsgBox "Replaced " & changeCount & " occurrences in formulas.", vbInformation
End Sub

```

**Advanced Formula Find/Replace with Backup:**

```
Sub AdvancedFormulaReplace()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Create backup
    Dim backupWs As Worksheet
    ws.Copy After:=ws
    Set backupWs = ActiveSheet
    backupWs.Name = "Backup_" & ws.Name & "_" & Format(Now, "yyyymmdd_hhmmss")

    ' Define replacements as a dictionary
    Dim replacements As Object
    Set replacements = CreateObject("Scripting.Dictionary")

    ' Add replacement pairs
    replacements.Add "VLOOKUP", "XLOOKUP"
    replacements.Add "Sheet1!", "Data!"
    replacements.Add "$A$1:$A$100", "$A$1:$A$1000"
    replacements.Add "0.1", "0.15"  ' Update tax rate example

    Dim formulaRange As Range
    On Error Resume Next
    Set formulaRange = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0

    If formulaRange Is Nothing Then Exit Sub

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim cell As Range
    Dim key As Variant
    Dim changeLog As String

    For Each cell In formulaRange
        Dim modified As Boolean
        modified = False
        Dim newFormula As String
        newFormula = cell.Formula

        ' Apply all replacements
        For Each key In replacements.Keys
            If InStr(1, newFormula, key, vbTextCompare) > 0 Then
                newFormula = Replace(newFormula, key, replacements(key), , , vbTextCompare)
                modified = True
            End If
        Next key

        If modified Then
            changeLog = changeLog & cell.Address & ": " & cell.Formula & " → " & newFormula & vbCrLf
            cell.Formula = newFormula
        End If
    Next cell

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Show change log
    If changeLog <> "" Then
        Dim logWs As Worksheet
        Set logWs = Worksheets.Add
        logWs.Name = "ChangeLog_" & Format(Now, "hhmmss")
        logWs.Range("A1").Value = "Change Log"
        logWs.Range("A2").Value = changeLog
        logWs.Columns("A").AutoFit
    End If

    MsgBox "Formula replacement complete. Check ChangeLog sheet for details.", vbInformation
End Sub

```
