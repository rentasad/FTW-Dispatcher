Attribute VB_Name = "RouteGeneratorModule"
'@Folder("Modules.Airport")
' @TODO: PAX CARGO COMBO Kombinationen berücksichtigen

Public Const COLUMN_DEPARTURE As Long = 1
Public Const COLUMN_ARRIVAL As Long = 2
Public Const COLUMN_TERMINAL_PAX As Long = 3
Public Const COLUMN_TERMINAL_CARGO As Long = 4
Public Const COLUMN_LONGEST_RWY As Long = 5
Public Const COLUMN_DISTANCE_NM As Long = 6
Public Const COLUMN_RANGE_VALID As Long = 7


Public Const FIELD_PIVOT_COLUMN_MIN_RUNWAY As String = "$I$2"
Public Const FIELD_PIVOT_COLUMN_MAX_RANGE As String = "$J$2"
Public Const FIELD_PIVOT_COLUMN_SEATS As String = "$K$2"
Public Const FIELD_PIVOT_COLUMN_CARGO As String = "$L$2"
Public Const FIELD_PIVOT_COLUMN_LONG_RANGE_PLANE As String = "$M$2"


Public Function getTableLengthFromRoutestable() As Long
Dim tbl As ListObject
Set tbl = ActiveSheet.ListObjects("Routes")
    'MsgBox tbl.Range.Rows.Count
    'MsgBox tbl.HeaderRowRange.Rows.Count
    'MsgBox tbl.DataBodyRange.Rows.Count
    getTableLengthFromRoutestable = tbl.Range.Rows.Count
    
    
Set tbl = Nothing
End Function

Public Sub formatRouteTable()
    Dim tbl As ListObject
    distanceTable.Activate
    Set tbl = ActiveSheet.ListObjects("Routes")
    ' Runway LENGTH
    With tbl.DataBodyRange.Range("E1:E" & getTableLengthFromRoutestable)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, _
                Operator:=xlLessEqual, _
        Formula1:="=" & FIELD_PIVOT_COLUMN_MIN_RUNWAY
            .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
    End With
    ' DISTANCE/RANGE
    With tbl.DataBodyRange.Range("F1:F" & getTableLengthFromRoutestable)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, _
                Operator:=xlGreater, _
        Formula1:="=" & FIELD_PIVOT_COLUMN_MAX_RANGE
            .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
    End With
    ' TERMINAL/PAX
    With tbl.DataBodyRange.Range("C1:C" & getTableLengthFromRoutestable)
            .FormatConditions.Delete
            .FormatConditions.Add Type:=xlCellValue, _
                Operator:=xlLess, _
        Formula1:="=" & FIELD_PIVOT_COLUMN_SEATS
            .FormatConditions(1).Interior.Color = RGB(255, 199, 206)
    End With
End Sub

' FORMEL FÜR RangeValid-Spalte:
' =WENN(
' ODER(
   ' UND(PIVOTDATENZUORDNEN("Summe von Long Range Plane?";$I$1)=1;
           ' [@[DISTANCE NM]]>=500;
            ' [@[DISTANCE NM]]<=PIVOTDATENZUORDNEN("Summe von Max Range (nm)";$I$1)
   ' );UND(
    ' PIVOTDATENZUORDNEN("Summe von Long Range Plane?";$I$1)=0;
    ' [@[DISTANCE NM]]<=PIVOTDATENZUORDNEN("Summe von Max Range (nm)";$I$1)
' ));1;0
'
' )

Public Sub writeDistanceToDistanceTable()
TableModul.disableUpdates
distanceTable.Range("A2:F999999").ClearContents
distanceTable.Cells(1, COLUMN_DEPARTURE) = "DEPARTURE"
distanceTable.Cells(1, COLUMN_ARRIVAL) = "DESTINATION"
distanceTable.Cells(1, COLUMN_TERMINAL_PAX) = "TERMINAL PAX"
distanceTable.Cells(1, COLUMN_TERMINAL_CARGO) = "TERMINAL CARGO"
distanceTable.Cells(1, COLUMN_DISTANCE_NM) = "DISTANCE NM"
distanceTable.Cells(1, COLUMN_LONGEST_RWY) = "LONGEST RUNWAY"
distanceTable.Cells(1, COLUMN_RANGE_VALID) = "RANGE VALID"

Dim dictAirports As Scripting.Dictionary
Set dictAirports = AirportObjectAccessModul.getAirportDictionary
Const KM_TO_NM_CONVERSION_CONSTANT As Double = 0.539956803


Dim distanceKm1 As Double
Dim distanceNM1 As Double
Dim minimumPax As Variant
Dim minimumCargo As Variant
Dim longestRunway As Long
Dim line As Long
line = 2
For Each airportIcao1 In dictAirports.Keys()
    For Each airportIcao2 In dictAirports.Keys()
        If Not (airportIcao1 Like airportIcao2) Then
        
            
        
        
            Dim airportDep As AirportObject
            Dim airportDest As AirportObject
            Set airportDep = dictAirports.Item(airportIcao1)
            Set airportDest = dictAirports.Item(airportIcao2)
            
            If (isAirportCombinationValid(airportDep, airportDest)) Then
            
            
                distanceKm1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, Km)
                distanceNM1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, NM)
                longestRunway = airportDest.maxRunwayLength
                minimumPax = getMinimumPax(airportDep, airportDest)
                minimumCargo = getMinimumCargo(airportDep, airportDest)
                
                
                distanceTable.Cells(line, COLUMN_DEPARTURE) = airportDep.icao
                distanceTable.Cells(line, COLUMN_ARRIVAL) = airportDest.icao
                distanceTable.Cells(line, COLUMN_TERMINAL_PAX) = minimumPax
                distanceTable.Cells(line, COLUMN_TERMINAL_CARGO) = minimumCargo
                distanceTable.Cells(line, COLUMN_DISTANCE_NM) = Round(distanceNM1, 0)
                distanceTable.Cells(line, COLUMN_LONGEST_RWY) = longestRunway
                
                line = line + 1
            End If
            
        End If
    Next
Next

' MakeTable

'distanceTable.Activate
'distanceTable.ListObjects.Add(xlSrcRange, Range("A1:F" & line), , xlYes).Name = "Routes"
formatRouteTable
TableModul.enableUpdates
End Sub



' Check if Combination of CARGO/PAX/COMBO Terminaltype is valid
Public Function isAirportCombinationValid(ByVal airportDep As AirportObject, ByVal airportDest As AirportObject) As Boolean
    
    Const TERMINAL_TYPE_PAX As String = "PAX"
    Const TERMINAL_TYPE_CARGO As String = "CARGO"
    Const TERMINAL_TYPE_COMBO As String = "COMBO"

    Dim isValid As Boolean
    Const IS_EQUAL As Long = 0
    Const NOT_EQUAL As Long = -1
    
    Dim depIcao As String
    depIcao = airportDep.icao
    Dim destIcao As String
    destIcao = airportDest.icao
'    If (StrComp(depIcao, "EBOS") = 0) And (StrComp(destIcao, "EDDS") = 0) Then
'        Debug.Print ("STOP")
'    End If
    
    
    Dim departureTerminalType As String
    Dim destinationTerminalType As String
    departureTerminalType = UCase(airportDep.terminalType)
    destinationTerminalType = UCase(airportDest.terminalType)
    
    If (StrComp(departureTerminalType, TERMINAL_TYPE_PAX) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_PAX) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_PAX) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_PAX) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_CARGO) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_COMBO) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_CARGO) = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, TERMINAL_TYPE_CARGO) = IS_EQUAL) And (StrComp(destinationTerminalType, TERMINAL_TYPE_CARGO) = IS_EQUAL) Then
        isValid = True
    Else
        isValid = False
    End If
    
    isAirportCombinationValid = isValid
End Function
' return minimum pax  value
Public Function getMinimumPax(ByVal airportDep As AirportObject, ByVal airportDest As AirportObject) As Variant
    Dim minPax As Variant
    minPax = Null
    If (airportDep.terminalSize >= airportDest.terminalSize) Then
        minPax = airportDest.terminalSize
    Else
        minPax = airportDep.terminalSize
    End If
    getMinimumPax = minPax
        
End Function

' Return minimum cargo size
Public Function getMinimumCargo(ByVal airportDep As AirportObject, ByVal airportDest As AirportObject) As Variant
    Dim minCargo As Variant
    minCargo = Null
    If (airportDep.cargoSize >= airportDest.cargoSize) Then
        minCargo = airportDest.cargoSize
    Else
        minCargo = airportDep.cargoSize
    End If

    getMinimumCargo = minCargo
End Function

Sub test()
    Dim line As Long
    line = 10
    distanceTable.ListObjects.Add(xlSrcRange, Range("A1:E" & line), , xlYes).Name = "Routes"
End Sub

