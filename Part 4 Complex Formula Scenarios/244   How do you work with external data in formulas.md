### 244. **How do you work with external data in formulas?**

```
Sub CreateFormulasWithExternalData()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Method 1: Link to another workbook
    Dim externalPath As String
    externalPath = "C:\Data\ExternalData.xlsx"

    Range("A1").Formula = "='[" & Dir(externalPath) & "]Sheet1'!A1"

    ' Method 2: Create formula with external reference
    Dim formulaString As String
    formulaString = "=VLOOKUP(A1,'[ExternalData.xlsx]Sheet1'!$A:$B,2,FALSE)"
    Range("B1").Formula = formulaString

    ' Method 3: Use INDIRECT with external reference (requires workbook open)
    Range("C1").Formula = "=INDIRECT(""'[ExternalData.xlsx]Sheet1'!A1"")"

    ' Method 4: Import data then use formulas
    Dim externalWb As Workbook
    Set externalWb = Workbooks.Open(externalPath)

    ' Copy data
    externalWb.Sheets("Sheet1").Range("A1:B100").Copy _
        Destination:=ws.Range("E1")

    ' Create formulas referencing imported data
    ws.Range("D1:D100").Formula = "=VLOOKUP(A1,$E:$F,2,FALSE)"

    externalWb.Close SaveChanges:=False
End Sub

```

**Query External Database:**

```
Sub CreateFormulasFromDatabase()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    ' Import data from database using ADO
    Dim conn As Object
    Set conn = CreateObject("ADODB.Connection")

    Dim rs As Object
    Set rs = CreateObject("ADODB.Recordset")

    ' Connection string (example for SQL Server)
    Dim connString As String
    connString = "Provider=SQLOLEDB;Data Source=ServerName;" & _
                 "Initial Catalog=DatabaseName;Integrated Security=SSPI;"

    conn.Open connString

    ' Execute query
    Dim sql As String
    sql = "SELECT Category, SUM(Amount) as Total FROM Sales GROUP BY Category"

    rs.Open sql, conn

    ' Import to worksheet
    ws.Range("A1").CopyFromRecordset rs

    ' Create formulas based on imported data
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ws.Range("C2:C" & lastRow).Formula = "=B2/$B$" & lastRow

    rs.Close
    conn.Close

    Set rs = Nothing
    Set conn = Nothing
End Sub

```
