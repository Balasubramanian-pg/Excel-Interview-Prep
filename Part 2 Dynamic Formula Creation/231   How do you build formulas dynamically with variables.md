### 231. **How do you build formulas dynamically with variables?**

```
Sub DynamicFormulaCreation()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Method 1: String concatenation
    Dim formulaString As String
    Dim startRow As Long, endRow As Long

    startRow = 1
    endRow = 100

    formulaString = "=SUM(A" & startRow & ":A" & endRow & ")"
    Range("B1").Formula = formulaString

    ' Method 2: With column variables
    Dim col1 As String, col2 As String
    col1 = "A"
    col2 = "B"

    formulaString = "=SUMIF(" & col1 & ":" & col1 & ",""West""," & _
                    col2 & ":" & col2 & ")"
    Range("C1").Formula = formulaString

    ' Method 3: Using Cells reference
    Dim targetCell As Range
    Set targetCell = Cells(1, 1)

    formulaString = "=SUM(" & targetCell.Address & ":" & _
                    Cells(100, 1).Address & ")"
    Range("D1").Formula = formulaString

    ' Method 4: Building complex formulas
    Dim criteria As String
    criteria = "West"
    Dim threshold As Double
    threshold = 1000

    formulaString = "=SUMIFS(C:C,A:A,""" & criteria & """,B:B,"">""&" & threshold & ")"
    Range("E1").Formula = formulaString
End Sub

```

**Dynamic Formula with Named Ranges:**

```
Sub DynamicWithNames()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Create named range dynamically
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Names.Add Name:="SalesData", _
                 RefersTo:=ws.Range("A1:A" & lastRow)

    ' Use named range in formula
    Range("B1").Formula = "=SUM(SalesData)"
    Range("B2").Formula = "=AVERAGE(SalesData)"
    Range("B3").Formula = "=MAX(SalesData)"

    ' Dynamic named range with OFFSET
    ws.Names.Add Name:="DynamicRange", _
                 RefersTo:="=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)"

    Range("C1").Formula = "=SUM(DynamicRange)"
End Sub

```
