### 241. **How do you create cascading dropdowns with VBA?**

```
Sub CreateCascadingDropdowns()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Assume we have:
    ' Categories in column A of "Lists" sheet
    ' Items for each category in columns B, C, D, etc.

    ' First dropdown: Category
    With ws.Range("A2:A100").Validation
        .Delete
        .Add Type:=xlValidateList, _
             Formula1:="=Lists!$A$2:$A$10"
        .InCellDropdown = True
        .InputTitle = "Category"
        .InputMessage = "Select a category"
    End With

    ' Second dropdown: Subcategory (depends on category)
    With ws.Range("B2:B100").Validation
        .Delete
        .Add Type:=xlValidateList, _
             Formula1:="=INDIRECT(A2)"  ' A2 must be a named range
        .InCellDropdown = True
        .InputTitle = "Subcategory"
        .InputMessage = "Select a subcategory"
    End With

    ' Third dropdown: Item (depends on subcategory)
    With ws.Range("C2:C100").Validation
        .Delete
        .Add Type:=xlValidateList, _
             Formula1:="=INDIRECT(B2&""_Items"")"
        .InCellDropdown = True
        .InputTitle = "Item"
        .InputMessage = "Select an item"
    End With
End Sub

```

**Dynamic Cascading with Event Handling:**

```
' This goes in the worksheet module (not a standard module)
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Me

    On Error GoTo ErrorHandler
    Application.EnableEvents = False

    ' When category changes, clear dependent cells
    If Not Intersect(Target, ws.Range("A:A")) Is Nothing Then
        If Target.Row > 1 Then
            ws.Cells(Target.Row, 2).ClearContents  ' Clear subcategory
            ws.Cells(Target.Row, 3).ClearContents  ' Clear item

            ' Update subcategory dropdown
            Dim categoryValue As String
            categoryValue = Target.Value

            If categoryValue <> "" Then
                With ws.Cells(Target.Row, 2).Validation
                    .Delete
                    .Add Type:=xlValidateList, _
                         Formula1:="=INDIRECT(""" & categoryValue & """)"
                    .InCellDropdown = True
                End With
            End If
        End If
    End If

    ' When subcategory changes, clear dependent cells
    If Not Intersect(Target, ws.Range("B:B")) Is Nothing Then
        If Target.Row > 1 Then
            ws.Cells(Target.Row, 3).ClearContents  ' Clear item

            Dim subcategoryValue As String
            subcategoryValue = Target.Value

            If subcategoryValue <> "" Then
                With ws.Cells(Target.Row, 3).Validation
                    .Delete
                    .Add Type:=xlValidateList, _
                         Formula1:="=INDIRECT(""" & subcategoryValue & "_Items"")"
                    .InCellDropdown = True
                End With
            End If
        End If
    End If

ExitHandler:
    Application.EnableEvents = True
    Exit Sub

ErrorHandler:
    MsgBox "Error: " & Err.Description
    Resume ExitHandler
End Sub

```
