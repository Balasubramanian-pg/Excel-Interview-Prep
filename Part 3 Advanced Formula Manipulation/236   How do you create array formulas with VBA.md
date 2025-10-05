### 236. **How do you create array formulas with VBA?**

```
Sub CreateArrayFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Legacy array formula (pre-365)
    Range("A1").FormulaArray = "=SUM(B1:B10*C1:C10)"

    ' Multi-cell array formula
    Range("D1:D10").FormulaArray = "=B1:B10*C1:C10"

    ' Array formula with IF
    Range("E1").FormulaArray = "=SUM(IF(A1:A100=""West"",B1:B100,0))"

    ' Excel 365 dynamic array (no CSE needed)
    If Val(Application.Version) >= 16 Then
        Range("F1").Formula2 = "=FILTER(A:B,C:C>100)"
        Range("G1").Formula2 = "=SORT(A:B,2,-1)"
        Range("H1").Formula2 = "=UNIQUE(A:A)"
    End If

    ' Check if formula is array formula
    Dim cell As Range
    Set cell = Range("A1")

    If cell.HasArray Then
        MsgBox "Cell contains array formula: " & cell.FormulaArray
        MsgBox "Array covers: " & cell.CurrentArray.Address
    End If
End Sub

```

**Create Dynamic Array Formulas:**

```
Sub CreateDynamicArrayFormulas()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' SEQUENCE
    Range("A1").Formula2 = "=SEQUENCE(10,5,1,1)"

    ' RANDARRAY
    Range("B1").Formula2 = "=RANDARRAY(10,3,1,100,TRUE)"

    ' FILTER with multiple criteria
    Range("C1").Formula2 = "=FILTER(Data,(Category=""A"")*(Amount>100))"

    ' SORT by multiple columns
    Range("D1").Formula2 = "=SORT(Data,{2,3},{1,-1})"

    ' SORTBY
    Range("E1").Formula2 = "=SORTBY(Names,Scores,-1)"

    ' XLOOKUP returning array
    Range("F1").Formula2 = "=XLOOKUP(A:A,LookupTable[ID],LookupTable[[Name]:[Amount]])"

    ' Combination formulas
    Range("G1").Formula2 = "=SORT(UNIQUE(FILTER(A:A,B:B>100)))"

    ' Handle spill range
    Dim spillRange As Range
    On Error Resume Next
    Set spillRange = Range("A1").SpillParent.SpillingToRange
    On Error GoTo 0

    If Not spillRange Is Nothing Then
        MsgBox "Spill range: " & spillRange.Address
    End If
End Sub

```
