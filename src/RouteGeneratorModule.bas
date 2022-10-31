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

Dim line As Long
line = 2
For Each airportIcao1 In dictAirports.Keys()
    For Each airportIcao2 In dictAirports.Keys()
        If Not (airportIcao1 Like airportIcao2) Then
            Dim airportDep As airportObject
            Dim airportDest As airportObject
            Set airportDep = dictAirports.Item(airportIcao1)
            Set airportDest = dictAirports.Item(airportIcao2)
            
            distanceKm1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, Km)
            distanceNM1 = DistanceCaluculationModul.dDistance(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, NM)
            
            distanceTable.Cells(line, COLUMN_DEPARTURE) = airportDep.icao
            distanceTable.Cells(line, COLUMN_ARRIVAL) = airportDest.icao
            distanceTable.Cells(line, COLUMN_TERMINAL_PAX) = airport
            distanceTable.Cells(line, COLUMN_TERMINAL_CARGO) = ""
            distanceTable.Cells(line, 5) = distanceNM1
            
            line = line + 1
            
        End If
    Next
Next
'distanceTable.Range("C:D" & line).NumberFormat = "#_ ;-#"
distanceTable.Activate
TableModul.enableUpdates
End Sub


