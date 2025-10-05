### 230. **How do you use WorksheetFunction in VBA?**

```
Sub UseWorksheetFunctions()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Basic functions
    Dim total As Double
    total = WorksheetFunction.Sum(Range("A1:A10"))

    Dim average As Double
    average = WorksheetFunction.Average(Range("A1:A10"))

    Dim maxVal As Double
    maxVal = WorksheetFunction.Max(Range("A1:A10"))

    ' VLOOKUP
    Dim lookupResult As Variant
    On Error Resume Next
    lookupResult = WorksheetFunction.VLookup("SearchValue", _
                                             Range("A:B"), _
                                             2, _
                                             False)
    If Err.Number <> 0 Then
        lookupResult = "Not Found"
        Err.Clear
    End If
    On Error GoTo 0

    ' XLOOKUP (Excel 365)
    Dim xlookupResult As Variant
    xlookupResult = WorksheetFunction.XLookup("SearchValue", _
                                              Range("A:A"), _
                                              Range("B:B"), _
                                              "Not Found")

    ' SUMIF
    Dim conditionalSum As Double
    conditionalSum = WorksheetFunction.SumIf(Range("A:A"), ">100", Range("B:B"))

    ' COUNTIFS
    Dim conditionalCount As Long
    conditionalCount = WorksheetFunction.CountIfs(Range("A:A"), "West", _
                                                  Range("B:B"), ">1000")

    ' TEXT function
    Dim formatted As String
    formatted = WorksheetFunction.Text(Now, "yyyy-mm-dd")
End Sub

```

**Advanced WorksheetFunction Uses:**

```
Sub AdvancedWorksheetFunctions()
    ' Array functions
    Dim matchPosition As Long
    matchPosition = WorksheetFunction.Match("Value", Range("A:A"), 0)

    ' INDEX-MATCH combination
    Dim result As Variant
    result = WorksheetFunction.Index(Range("C:C"), _
             WorksheetFunction.Match("Value", Range("A:A"), 0))

    ' TRANSPOSE
    Dim transposed As Variant
    transposed = WorksheetFunction.Transpose(Range("A1:A10").Value)

    ' UNIQUE (Excel 365)
    Dim uniqueValues As Variant
    uniqueValues = WorksheetFunction.Unique(Range("A1:A100").Value)

    ' FILTER (Excel 365)
    Dim filtered As Variant
    filtered = WorksheetFunction.Filter(Range("A:B").Value, _
                                        Range("C:C").Value, _
                                        ">100")

    ' Statistical functions
    Dim correlation As Double
    correlation = WorksheetFunction.Correl(Range("A:A"), Range("B:B"))

    Dim percentile As Double
    percentile = WorksheetFunction.Percentile_Inc(Range("A:A"), 0.95)
End Sub

```
