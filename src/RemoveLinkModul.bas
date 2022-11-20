Attribute VB_Name = "RemoveLinkModul"
' ***************************************************************************
' * SOURCE: https://www.denisreis.com/excel-vba-verknuepfungen-entfernen/   *
' ***************************************************************************
Sub removeWorkbookLinksInWorkbook(ByVal Workbook As Workbook)
Dim ArrayV As Variant
Dim CntLine As Long
ArrayV = Workbook.LinkSources(Type:=xlLinkTypeExcelLinks)
For CntLine = 1 To UBound(ArrayV)

Workbook.BreakLink Name:=ArrayV(CntLine), Type:=xlLinkTypeExcelLinks

Next CntLine
End Sub
