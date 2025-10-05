### 243. **How do you create volatile UDFs?**

```
' Volatile function - recalculates every time Excel calculates
Function CurrentUser() As String
    Application.Volatile  ' Makes function volatile
    CurrentUser = Environ("USERNAME")
End Function

Function LastCalculated() As String
    Application.Volatile
    LastCalculated = Format(Now, "yyyy-mm-dd hh:mm:ss")
End Function

Function RandomBetweenUnique(bottom As Long, top As Long) As Long
    Application.Volatile
    RandomBetweenUnique = Int((top - bottom + 1) * Rnd + bottom)
End Function

' Non-volatile version for comparison
Function StaticDate() As String
    ' Does NOT use Application.Volatile
    StaticDate = Format(Date, "yyyy-mm-dd")
    ' Only recalculates when cell or its precedents change
End Function

```

**Conditional Volatility:**

```
Function SmartRefresh(value As Variant, forceRefresh As Boolean) As Variant
    ' Only volatile if forceRefresh is TRUE
    If forceRefresh Then
        Application.Volatile
    End If

    ' Your calculation here
    SmartRefresh = value * 1.1  ' Example calculation
End Function

' Use: =SmartRefresh(A1, FALSE)  ' Not volatile
'      =SmartRefresh(A1, TRUE)   ' Volatile

```
