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
Dim homeDispatchingWorksheet As Worksheet

Dim newDispatchingWorksheet As Worksheet
Dim newAirlinePlanesWorksheet As Worksheet

exportPath = ConfigTable.Cells(19, 2)
airlineName = ConfigTable.Cells(21, 2)

Dim filepath As String
filepath = exportPath & airlineName & "-Dispatching.xlsx"
Set ftwDispatchingHelperWorkbook = ActiveWorkbook

Set dispatchingWorksheet = ftwDispatchingHelperWorkbook.Worksheets("Dispatching")
Set airlineplanesWorksheet = ftwDispatchingHelperWorkbook.Worksheets("Airline Planes")
Set homeDispatchingWorksheet = ftwDispatchingHelperWorkbook.Worksheets("HOME Dispatchingtable")
homeDispatchingWorksheet.Visible = True
' Copy Sheets to new Workbook
Set newAirlineDispatchingWorkbook = Workbooks.Add
dispatchingWorksheet.Activate
dispatchingWorksheet.Select
dispatchingWorksheet.Copy Before:=newAirlineDispatchingWorkbook.Sheets(1)
'ftwDispatchingHelperWorkbook.Activate
'    airlineplanesWorksheet.Activate
'    airlineplanesWorksheet.Copy After:=newAirlineDispatchingWorkbook.Sheets(1)
ftwDispatchingHelperWorkbook.Activate
    homeDispatchingWorksheet.Activate
    homeDispatchingWorksheet.Copy After:=newAirlineDispatchingWorkbook.Sheets(1)

ftwDispatchingHelperWorkbook.Activate
homeDispatchingWorksheet.Visible = False
' *************************
' *  COPY only VALUES     *
' *************************
airlineplanesWorksheet.Activate
airlineplanesWorksheet.ListObjects.Item("Uebersicht_Airline_Flugzeuge").Range.Select
airlineplanesWorksheet.ListObjects.Item("Uebersicht_Airline_Flugzeuge").Range.Copy


newAirlineDispatchingWorkbook.Sheets("Tabelle1").Activate
newAirlineDispatchingWorkbook.Sheets("Tabelle1").Name = "Airline Planes"


Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Dim objTable As ListObject
Set objTable = ActiveSheet.ListObjects.Add(xlSrcRange, Selection, , xlYes)
objTable.Name = "Uebersicht_Airline_Flugzeuge"



For Each pt In ActiveWorkbook.Worksheets("Dispatching").PivotTables
         pt.ChangePivotCache ActiveWorkbook.PivotCaches.Create _
            (SourceType:=xlDatabase, SourceData:="Uebersicht_Airline_Flugzeuge")
Next pt


Set newAirlinePlanesWorksheet = newAirlineDispatchingWorkbook.Worksheets("Airline Planes")
newAirlineDispatchingWorkbook.Sheets("Airline Planes").Activate

Dim test As Variant
' DELETE unused Queries
'Set test = newAirlinePlanesWorksheet.QueryTables
'ActiveWorkbook.Queries("Uebersicht_Airline_Flugzeuge").Delete

' DELETE Active Connections
'Dim connectionCount As Long
'connectionCount = newAirlineDispatchingWorkbook.Connections.Count
'If (connectionCount = 1) Then
'    newAirlineDispatchingWorkbook.Connections.Item(1).Delete
'End If

'Dim xConnect As Object
'For Each xConnect In newAirlineDispatchingWorkbook.Connections
'If xConnect.Name <> "ThisWorkbookDataModel" Then xConnect.Delete
'Next xConnect




' Hide unused Sheets
newAirlineDispatchingWorkbook.Sheets("Airline Planes").Visible = False
'newAirlineDispatchingWorkbook.Sheets("Tabelle1").Visible = False
newAirlineDispatchingWorkbook.Sheets("Dispatching").Activate
delete_file_if_exists filepath
newAirlineDispatchingWorkbook.SaveAs filepath, FILE_FORMAT_NUMBER_XLSX
ControllerTable.Activate
TableModul.enableUpdates
MsgBox "The Generation was completed successfully.", 64, "Generation successfull"
End Sub




' Delete File if exist
Sub delete_file_if_exists(ByVal filename As String)
    Dim aFile As String
    
    If Len(Dir$(filename)) > 0 Then
    ' File Exists, Delete it
    Kill filename
    End If
End Sub

