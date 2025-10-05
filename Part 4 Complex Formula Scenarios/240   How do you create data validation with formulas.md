### 240. **How do you create data validation with formulas?**

```
Sub CreateDataValidation()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim rng As Range

    ' Method 1: List validation from range
    Set rng = ws.Range("A2:A100")
    With rng.Validation
        .Delete  ' Clear existing validation
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=$G$2:$G$10"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "Select Category"
        .InputMessage = "Choose from the dropdown list"
        .ErrorTitle = "Invalid Entry"
        .ErrorMessage = "Please select a valid category"
    End With

    ' Method 2: Custom formula validation
    Set rng = ws.Range("B2:B100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=B2>=A2"
        .ErrorMessage = "Value must be greater than or equal to column A"
    End With

    ' Method 3: Date validation
    Set rng = ws.Range("C2:C100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=AND(C2>=TODAY(),C2<=TODAY()+365)"
        .ErrorMessage = "Date must be between today and one year from now"
    End With

    ' Method 4: Prevent duplicates
    Set rng = ws.Range("D2:D100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=COUNTIF($D$2:$D$100,D2)=1"
        .ErrorMessage = "Duplicate values are not allowed"
    End With

    ' Method 5: Dependent dropdown
    Set rng = ws.Range("E2:E100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=INDIRECT($A2)"  ' A2 contains the category name
        .IgnoreBlank = True
        .InCellDropdown = True
    End With
End Sub

```

**Advanced Data Validation:**

```
Sub AdvancedDataValidation()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Complex conditional validation
    Dim rng As Range
    Set rng = ws.Range("F2:F100")

    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             AlertStyle:=xlValidAlertStop, _
             Formula1:="=IF($E2=""High"",F2>=1000,IF($E2=""Medium"",F2>=500,F2>=0))"
        .ErrorMessage = "Value doesn't meet requirements based on priority"
    End With

    ' Searchable dropdown with dynamic filter
    Set rng = ws.Range("G2:G100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             AlertStyle:=xlValidAlertWarning, _
             Formula1:="=OFFSET(Products,0,0,COUNTA(Products),1)"
    End With

    ' Multiple criteria validation
    Set rng = ws.Range("H2:H100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             Formula1:="=AND(H2>=$G2,H2<=1.5*$G2,H2<=10000)"
        .ErrorMessage = "Value must be between column G and 150% of column G, max 10000"
    End With

    ' Validation based on another sheet
    Set rng = ws.Range("I2:I100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateList, _
             Formula1:="=ValidList!$A$2:$A$100"
        .IgnoreBlank = True
    End With

    ' Time-based validation (business hours only)
    Set rng = ws.Range("J2:J100")
    With rng.Validation
        .Delete
        .Add Type:=xlValidateCustom, _
             Formula1:="=AND(HOUR(J2)>=9,HOUR(J2)<17,WEEKDAY(J2,2)<=5)"
        .ErrorMessage = "Time must be during business hours (9 AM - 5 PM, weekdays)"
    End With
End Sub

```
