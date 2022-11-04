Attribute VB_Name = "Modul1"
Sub PrintWorkbookPath()

ControllerTable.Cells(8, 2) = Application.ActiveWorkbook.FullName
End Sub
Sub RefreshAllDatas()
    ThisWorkbook.RefreshAll
    Const airplanePivotTableName As String = "AirplanePivotTable"
    Dim pivots As Variant
    distanceTable.PivotTables(airplanePivotTableName).RefreshTable
End Sub
