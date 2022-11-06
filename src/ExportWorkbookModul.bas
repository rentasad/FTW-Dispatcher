Attribute VB_Name = "ExportWorkbookModul"
Public Const FILE_FORMAT_NUMBER_XLSX As Long = 51
Public Const FILE_FORMAT_NUMBER_XLSB As Long = 50
Public Const FILE_FORMAT_STRING_XLSX As String = ".xlsx"
Public Const FILE_FORMAT_STRING_XLSB As String = ".xlsb"

Public Sub exportDispatchingTable()
TableModul.disableUpdates
Dim ftwDispatchingHelperWorkbook As Workbook
Dim newAirlineDispatchingWorkbook As Workbook
Dim exportPath As String
Dim airlineName As String

Dim dispatchingWorksheet As Worksheet
Dim airlineplanesWorksheet As Worksheet

Dim newDispatchingWorksheet As Worksheet
Dim newAirlinePlanesWorksheet As Worksheet

exportPath = ConfigTable.Cells(19, 2)
airlineName = ConfigTable.Cells(21, 2)

Dim filepath As String
filepath = exportPath & airlineName & "-Dispatching.xlsx"
Set ftwDispatchingHelperWorkbook = ActiveWorkbook

Set dispatchingWorksheet = ftwDispatchingHelperWorkbook.Worksheets("Dispatching")
Set airlineplanesWorksheet = ftwDispatchingHelperWorkbook.Worksheets("Airline Planes")
Set newAirlineDispatchingWorkbook = Workbooks.Add
dispatchingWorksheet.Activate
dispatchingWorksheet.Select
dispatchingWorksheet.Copy Before:=newAirlineDispatchingWorkbook.Sheets(1)
ftwDispatchingHelperWorkbook.Activate
airlineplanesWorksheet.Activate
airlineplanesWorksheet.Copy After:=newAirlineDispatchingWorkbook.Sheets(1)
' Delete emptySheet
'newAirlineDispatchingWorkbook.Sheets("Tabelle1").
newAirlineDispatchingWorkbook.Activate
For Each pt In ActiveWorkbook.Worksheets("Dispatching").PivotTables
         pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create _
            (SourceType:=xlDatabase, SourceData:="Uebersicht_Airline_Flugzeuge")
Next pt


Set newAirlinePlanesWorksheet = newAirlineDispatchingWorkbook.Worksheets("Airline Planes")
newAirlineDispatchingWorkbook.Sheets("Airline Planes").Activate
Dim test As Variant
Set test = newAirlinePlanesWorksheet.QueryTables
ActiveWorkbook.Queries("Uebersicht_Airline_Flugzeuge").Delete

delete_file_if_exists filepath
newAirlineDispatchingWorkbook.SaveAs filepath, FILE_FORMAT_NUMBER_XLSX

TableModul.enableUpdates

End Sub

Sub delete_file_if_exists(ByVal filename As String)
    Dim aFile As String
    
    If Len(Dir$(filename)) > 0 Then
    ' File Exists, Delete it
    Kill filename
    End If
End Sub

