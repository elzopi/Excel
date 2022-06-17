Attribute VB_Name = "ProcForecast"
Sub ProcForecast()
'Best used when first column has value on last row and first row has a value in the last column

Dim wks As Worksheet
Dim Psheet As Worksheet
Dim LastRow As Long
Dim LastColumn As Long
Dim StartCell As Range
Dim objTable As ListObject
Dim tblName As String
Dim cCols As Collection
Dim i As Integer
Dim j As Integer
Dim Prange As Range
Dim Pcache As PivotCache

' build collection/matrix for exclusion, using ad-hoc list below plus $ columns found

Set cCols = New Collection

cCols.Add "Region"
cCols.Add "PM Manager"
cCols.Add "Proj Type"
cCols.Add "% Inv"
cCols.Add "Un-Ute Hrs Prev Qrts"
cCols.Add "Managing Dept"
cCols.Add "Curr"
cCols.Add "Proj Rate"
cCols.Add "Adj Rate USD"
cCols.Add "Proj XRate"
cCols.Add "Curr XRate"
cCols.Add "Subsidiary"
cCols.Add "Subsid Base Curr"


Set wks = ActiveSheet

' Convert to range
' check if there's already a table present
If wks.ListObjects.Count > 0 Then

    ActiveSheet.ListObjects(1).Unlist

End If

Set StartCell = Range("A1")

'Find Last Row and Column
  LastRow = wks.Cells(wks.Rows.Count, StartCell.Column).End(xlUp).Row
  LastColumn = wks.Cells(StartCell.Row, wks.Columns.Count).End(xlToLeft).Column

'Select Range
  wks.Range(StartCell, wks.Cells(LastRow, LastColumn)).Select

Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)

objTable.TableStyle = "TableStyleMedium3"

' Hide columns with headers USD$
tblName = objTable.Name


    For Each tbl In wks.ListObjects
        
        'Save Excel defined Table name to cell in column A
'        Range("A1").Cells(i, 1).Value = tbl.Name

        'Iterate through columns in Excel defined Table
        For j = 1 To tbl.Range.Columns.Count

            'Save header name to cell next to table name
            strHdr = tbl.Range.Cells(1, j)
'            Range("A1").Cells(i, j + 1).Value = tbl.Range.Cells(1, j)
            If InStr(1, strHdr, "$", vbTextCompare) Then
                
                cCols.Add strHdr
                
'                Range(tblName & "[" & strHdr & "]").EntireColumn.Hidden = True
'                Debug.Print "Contains: " & strHdr
            End If
            
        'Continue with next column
        Next j

        'Add 1 to variable i
        i = i + 1
    
    'Continue with next Excel defined Table
    Next tbl

' Process the header collection built from constants and $ found
For i = 1 To cCols.Count
    
    Debug.Print cCols(i)
    Range(tblName & "[" & cCols(i) & "]").EntireColumn.Hidden = True

Next i

' Range(tblName & "[Ttl Invc USD$]").EntireColumn.Hidden = True

'Insert Blank Pivot Table
'Define Pivot Cache
'Insert a New Blank Worksheet

On Error Resume Next
Application.DisplayAlerts = False
Worksheets("PivotTable").Delete
Sheets.Add Before:=ActiveSheet
ActiveSheet.Name = "PivotTable"
Application.DisplayAlerts = True
Set Psheet = Worksheets("PivotTable")
Set DSheet = Worksheets("Data")
Set Psheet = Worksheets("PivotTable")
Set Prange = wks.Cells(1, 1).Resize(LastRow, LastColumn)
Set Pcache = ActiveWorkbook.PivotCaches.Create _
(SourceType:=xlDatabase, SourceData:=Prange). _
CreatePivotTable(TableDestination:=Psheet.Cells(2, 2), _
TableName:="PivotProjects")
Set PTable = Pcache.CreatePivotTable _
(TableDestination:=Psheet.Cells(1, 1), TableName:="PivotProjects")

Call BuildPivot
Call AdjustWidths

End Sub

Sub BuildPivot()

'
' BuildPivot Macro
'

'
    Range("B5").Select
    Application.CommandBars("Help").Visible = False
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("Project Manager")
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("PA #")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Customer Name and Project Name")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotProjects").RowAxisLayout xlTabularRow
    ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Customer Name and Project Name").Subtotals = Array(False, False, False, False, _
        False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("PA #").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Region").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Project Manager"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("PM Manager").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Proj Type").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Status").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Health").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Opp ID").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Budg Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Est Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Act Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Bklg Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("% Comp").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("% Inv").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Budg USD$"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Invc USD$"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("TtL Rev USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Ttl Bklg USD$"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M1 Est Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M1 Est USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M1 Actl Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M1 Act USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M2 Est Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M2 Est USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M2 Act Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M2 Actl USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M3 Est Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M3 Est USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M3 Act Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("M3 Act USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Est Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Est USD$").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Act Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Act USD$").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Bklg Hrs").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q Bklg USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+1 Est Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+1 Est USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+2 Est Hrs").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+2 Est USD$").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+3 & Beyond Est Hrs"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Q+3 & Beyond USD$"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Un-Ute Hrs Prev Qrts"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Managing Dept"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Curr").Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Proj Rate").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Adj Rate USD").Subtotals _
        = Array(False, False, False, False, False, False, False, False, False, False, False, False _
        )
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Proj XRate").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Curr XRate").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Subsidiary").Subtotals = _
        Array(False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotProjects").PivotFields("Subsid Base Curr"). _
        Subtotals = Array(False, False, False, False, False, False, False, False, False, False, _
        False, False)
    Columns("B:B").ColumnWidth = 13.17
    
' Pivot fields to report
    
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Ttl Budg Hrs"), "Count of Ttl Budg Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Count of Ttl Budg Hrs")
        .Caption = "Sum of Ttl Budg Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Ttl Est Hrs"), "Count of Ttl Est Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Count of Ttl Est Hrs")
        .Caption = "Sum of Ttl Est Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Ttl Act Hrs"), "Count of Ttl Act Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Count of Ttl Act Hrs")
        .Caption = "Sum of Ttl Act Hrs"
        .Function = xlSum
    End With
    
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Sum of Ttl Est Hrs"), "Sum of Ttl Est Hrs", xlSum
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Ttl Bklg Hrs"), "Count of Ttl Bklg Hrs", xlCount
    
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Count of Ttl Bklg Hrs")
        .Caption = "Sum of Ttl Bklg Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("% Comp"), "Count of % Comp", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("Count of % Comp")
        .Caption = "Sum of % Comp"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q Est Hrs"), "Count of Q Est Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("Count of Q Est Hrs")
        .Caption = "Sum of Q Est Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q Act Hrs"), "Count of Q Act Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("Count of Q Act Hrs")
        .Caption = "Sum of Q Act Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q Bklg Hrs"), "Count of Q Bklg Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields("Count of Q Bklg Hrs" _
        )
        .Caption = "Sum of Q Bklg Hrs"
        .Function = xlSum
    End With
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q+1 Est Hrs"), "Sum of Q+1 Est Hrs", xlSum
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q+2 Est Hrs"), "Sum of Q+2 Est Hrs", xlSum
    ActiveSheet.PivotTables("PivotProjects").AddDataField ActiveSheet.PivotTables( _
        "PivotProjects").PivotFields("Q+3 & Beyond Est Hrs"), _
        "Count of Q+3 & Beyond Est Hrs", xlCount
    With ActiveSheet.PivotTables("PivotProjects").PivotFields( _
        "Count of Q+3 & Beyond Est Hrs")
        .Caption = "Sum of Q+3 & Beyond Est Hrs"
        .Function = xlSum
    End With
End Sub

Sub AdjustWidths()
'
' AdjustWidths Macro
'
    Columns("C:C").Select
    Selection.ColumnWidth = 26
    Columns("D:N").Select
    Columns("D:N").Select
    Selection.ColumnWidth = 10
    Range("C6").Select
End Sub

Sub ListTables()

'From: https://www.get-digital-help.com/list-all-tables-and-corresponding-headers-in-a-workbook-vba/

Dim tbl As ListObject
Dim WS As Worksheet
Dim i As Single, j As Single
Dim strHdr As String

'Insert new worksheet and save to object WS
Set WS = Sheets.Add

'Save 1 to variable i
i = 1

'Go through each worksheet in the worksheets object collection
For Each WS In Worksheets

    'Go through all Excel defined Tables located in the current WS worksheet object
    For Each tbl In WS.ListObjects
        
        'Save Excel defined Table name to cell in column A
        Range("A1").Cells(i, 1).Value = tbl.Name

        'Iterate through columns in Excel defined Table
        For j = 1 To tbl.Range.Columns.Count

            'Save header name to cell next to table name
            strHdr = tbl.Range.Cells(1, j)
            Range("A1").Cells(i, j + 1).Value = tbl.Range.Cells(1, j)
            If InStr(1, strHdr, "$", vbTextCompare) Then
                Debug.Print "Contains: " & strHdr
            End If
            
        'Continue with next column
        Next j

        'Add 1 to variable i
        i = i + 1
    
    'Continue with next Excel defined Table
    Next tbl

'Continue with next worksheet
Next WS

'Exit macro
End Sub
