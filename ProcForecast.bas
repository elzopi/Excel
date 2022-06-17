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

' WIP unhide all columns

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
