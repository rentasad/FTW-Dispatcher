Attribute VB_Name = "RouteGeneratorModule"
'@Folder("Modules.Airport")
' @TODO: PAX CARGO COMBO Kombinationen berücksichtigen

Public Const COLUMN_DEPARTURE As Long = 1
Public Const COLUMN_ARRIVAL As Long = 2
Public Const COLUMN_TERMINAL_PAX As Long = 3
Public Const COLUMN_TERMINAL_CARGO As Long = 4
Public Const COLUMN_DISTANCE_NM As Long = 5


Public Sub writeDistanceToDistanceTable()
TableModul.disableUpdates
distanceTable.Range("A1:E999999").ClearContents
distanceTable.Cells(1, COLUMN_DEPARTURE) = "DEPARTURE"
distanceTable.Cells(1, COLUMN_ARRIVAL) = "DESTINATION"
distanceTable.Cells(1, COLUMN_TERMINAL_PAX) = "TERMINAL_PAX"
distanceTable.Cells(1, COLUMN_TERMINAL_CARGO) = "TERMINAL_CARGO"
distanceTable.Cells(1, COLUMN_DISTANCE_NM) = "DISTANCE_NM"

Dim dictAirports As Scripting.Dictionary
Set dictAirports = AirportObjectAccessModul.getAirportDictionary
Const KM_TO_NM_CONVERSION_CONSTANT As Double = 0.539956803


Dim distanceKm1 As Double
Dim distanceNM1 As Double
Dim minimumPax As Variant
Dim minimumCargo As Variant

Dim line As Long
line = 2
For Each airportIcao1 In dictAirports.Keys()
    For Each airportIcao2 In dictAirports.Keys()
        If Not (airportIcao1 Like airportIcao2) Then
        
            
        
        
            Dim airportDep As airportObject
            Dim airportDest As airportObject
            Set airportDep = dictAirports.Item(airportIcao1)
            Set airportDest = dictAirports.Item(airportIcao2)
            
            If (isAirportCombinationValid(airportDep, airportDest)) Then
            
            
                distanceKm1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, Km)
                distanceNM1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, NM)
                
                minimumPax = getMinimumPax(airportDep, airportDest)
                minimumCargo = getMinimumCargo(airportDep, airportDest)
                
                
                distanceTable.Cells(line, COLUMN_DEPARTURE) = airportDep.icao
                distanceTable.Cells(line, COLUMN_ARRIVAL) = airportDest.icao
                distanceTable.Cells(line, COLUMN_TERMINAL_PAX) = minimumPax
                distanceTable.Cells(line, COLUMN_TERMINAL_CARGO) = minimumCargo
                distanceTable.Cells(line, 5) = Round(distanceNM1, 0)
                
                line = line + 1
            End If
            
        End If
    Next
Next

' MakeTable



distanceTable.Activate
distanceTable.ListObjects.Add(xlSrcRange, Range("A1:E" & line), , xlYes).Name = "Routes"
TableModul.enableUpdates
End Sub

' Check if Combination of CARGO/PAX/COMBO Terminaltype is valid
Public Function isAirportCombinationValid(ByVal airportDep As airportObject, ByVal airportDest As airportObject) As Boolean
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
    departureTerminalType = airportDep.terminalType
    destinationTerminalType = airportDest.terminalType
    
    If (StrComp(departureTerminalType, "Pax") = IS_EQUAL) And (StrComp(destinationTerminalType, "Pax") = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, "Pax") = IS_EQUAL) And (StrComp(destinationTerminalType, "Combo") = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, "Combo") = IS_EQUAL) And (StrComp(destinationTerminalType, "Combo") = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, "Combo") = IS_EQUAL) And (StrComp(destinationTerminalType, "Combo") = IS_EQUAL) Then
        isValid = True
    ElseIf (StrComp(departureTerminalType, "Cargo") = IS_EQUAL) And (StrComp(destinationTerminalType, "Cargo") = IS_EQUAL) Then
        isValid = True
    Else
        isValid = False
    End If
    
    isAirportCombinationValid = isValid
End Function
' return minimum pax  value
Public Function getMinimumPax(ByVal airportDep As airportObject, ByVal airportDest As airportObject) As Variant
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
Public Function getMinimumCargo(ByVal airportDep As airportObject, ByVal airportDest As airportObject) As Variant
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

