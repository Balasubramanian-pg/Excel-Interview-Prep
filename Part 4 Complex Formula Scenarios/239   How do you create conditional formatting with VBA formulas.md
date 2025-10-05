### 239. **How do you create conditional formatting with VBA formulas?**

```
Sub CreateConditionalFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng As Range
    Set rng = ws.Range("A1:A100")

    ' Clear existing conditional formatting
    rng.FormatConditions.Delete

    ' Method 1: Formula-based condition
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1>100")
        .Interior.Color = RGB(255, 200, 200)  ' Light red
        .Font.Bold = True
    End With

    ' Method 2: Multiple conditions
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND($A1>50,$A1<=100)")
        .Interior.Color = RGB(255, 255, 200)  ' Light yellow
    End With

    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$A1<=50")
        .Interior.Color = RGB(200, 255, 200)  ' Light green
    End With

    ' Method 3: Entire row formatting based on column value
    Set rng = ws.Range("A1:E100")
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=$E1=""Complete""")
        .Interior.Color = RGB(200, 255, 200)
        .StopIfTrue = False
    End With

    ' Method 4: Alternate row shading
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=MOD(ROW(),2)=0")
        .Interior.Color = RGB(240, 240, 240)  ' Light gray
    End With

    ' Method 5: Highlight duplicates
    Set rng = ws.Range("A1:A100")
    With rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=COUNTIF($A$1:$A$100,$A1)>1")
        .Interior.Color = RGB(255, 150, 150)  ' Red
        .Font.Color = RGB(255, 255, 255)  ' White
    End With
End Sub

```

**Advanced Conditional Formatting:**

```
Sub AdvancedConditionalFormatting()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Highlight top 10 values
    Dim rng As Range
    Set rng = ws.Range("B2:B100")

    With rng.FormatConditions.Add(Type:=xlExpression, _
         Formula1:="=B2>=LARGE($B$2:$B$100,10)")
        .Interior.Color = RGB(146, 208, 80)
        .Font.Bold = True
    End With

    ' Heat map with gradient
    Set rng = ws.Range("C2:C100")
    With rng.FormatConditions.AddColorScale(ColorScaleType:=3)
        .ColorScaleCriteria(1).Type = xlConditionValueLowestValue
        .ColorScaleCriteria(1).FormatColor.Color = RGB(255, 0, 0)  ' Red

        .ColorScaleCriteria(2).Type = xlConditionValuePercentile
        .ColorScaleCriteria(2).Value = 50
        .ColorScaleCriteria(2).FormatColor.Color = RGB(255, 255, 0)  ' Yellow

        .ColorScaleCriteria(3).Type = xlConditionValueHighestValue
        .ColorScaleCriteria(3).FormatColor.Color = RGB(0, 255, 0)  ' Green
    End With

    ' Data bars
    Set rng = ws.Range("D2:D100")
    With rng.FormatConditions.AddDatabar
        .BarColor.Color = RGB(0, 112, 192)
        .BarFillType = xlDataBarFillGradient
        .Direction = xlLTR
        .ShowValue = True
    End With

    ' Icon sets based on percentiles
    Set rng = ws.Range("E2:E100")
    With rng.FormatConditions.AddIconSetCondition
        .IconSet = ThisWorkbook.IconSets(xl3TrafficLights1)
        .IconCriteria(2).Type = xlConditionValuePercent
        .IconCriteria(2).Value = 33
        .IconCriteria(3).Type = xlConditionValuePercent
        .IconCriteria(3).Value = 67
    End With

    ' Date-based formatting
    Set rng = ws.Range("F2:F100")
    ' Overdue dates in red
    With rng.FormatConditions.Add(Type:=xlExpression, _
         Formula1:="=AND(F2<TODAY(),F2<>"""")")
        .Interior.Color = RGB(255, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Dates within next 7 days in yellow
    With rng.FormatConditions.Add(Type:=xlExpression, _
         Formula1:="=AND(F2>=TODAY(),F2<=TODAY()+7)")
        .Interior.Color = RGB(255, 255, 0)
    End With
End Sub

```
