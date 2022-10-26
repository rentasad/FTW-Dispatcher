Attribute VB_Name = "AirportModul"
Const COLUMN_ICAO As Long = 1
'ICA    name    latitude_deg    longitude_deg   longest_runway
Const COLUMN_NAME As Long = 2
Const COLUMN_LATITUDE As Long = 3
Const COLUMN_LONGITUDE As Long = 4
Const COLUMN_LONGEST_RUNWAY_ As Long = 5

Public Sub updateAirportDbFromMysqlDb()
Dim rs As ADODB.Recordset
Dim connection As ADODB.connection
Dim mySqlConfig As mySqlConfigObject
Set mySqlConfig = DatabaseConnectionModul.getMySqlConfigObjectFromConfigSheet

Dim oConn As ADODB.connection
    If oConn Is Nothing Then
        Dim connectionString As String
        connectionString = DatabaseConnectionModul.getConnectionStringFromMySqlConfigObject(mySqlConfig)
        
        Set oConn = New ADODB.connection
        oConn.Open connectionString
    End If
    Set rs = CreateObject("ADODB.Recordset")
    
Dim query As String
query = "SELECT " _
    & "airports.ident, " _
    & "airports.name, " _
    & "airports.latitude_deg, " _
    & "airports.longitude_deg, " _
    & "max(length_ft) As longest_runway " _
    & "FROM airports " _
    & "LEFT JOIN runways " _
    & "    ON runways.airport_ref  = airports.id " _
    & "Group BY " _
    & "    airports.ident, " _
    & "    airports.name, " _
    & "    airports.latitude_deg, " _
    & "    airports.longitude_deg "


Set connection = DatabaseConnectionModul.getConnection(mySqlConfig)

Set rs = CreateObject("ADODB.Recordset")
    'Debug.Print sqlQuery
    rs.Open query, oConn
Dim line As Long
line = 1
AirportData.Range("A:E").Clear
AirportData.Cells(line, 1) = "ICAO"
AirportData.Cells(line, 2) = "Name"
AirportData.Cells(line, 3) = "Latitude"
AirportData.Cells(line, 4) = "Longitude"
AirportData.Cells(line, 5) = "Longest_Runway"

Do While Not rs.EOF
      line = line + 1
      AirportData.Cells(line, 1) = rs.Fields().Item("ident")
      AirportData.Cells(line, 2) = rs.Fields().Item("name")
      AirportData.Cells(line, 3) = rs.Fields().Item("latitude_deg")
      AirportData.Cells(line, 4) = rs.Fields().Item("longitude_deg")
      AirportData.Cells(line, 5) = rs.Fields().Item("longest_runway")
      rs.MoveNext

Loop

End Sub
  
Public Sub writeDistanceToDistanceTable()
distanceTable.Range("A1:D999999").ClearContents
distanceTable.Cells(1, 1) = "DEPARTURE"
distanceTable.Cells(1, 2) = "DESTINATION"
distanceTable.Cells(1, 3) = "DISTANCE_KM"
distanceTable.Cells(1, 4) = "DISTANCE_NM"

Dim dictAirports As Scripting.Dictionary
Set dictAirports = AirportModul.getAirportDictionary
Const KM_TO_NM_CONVERSION_CONSTANT As Double = 0.539956803


Dim distanceKm As Double
Dim distanceNM As Double
Dim line As Long
line = 2

For Each airportIcao1 In dictAirports.Keys()
    For Each airportIcao2 In dictAirports.Keys()
        If Not (airportIcao1 Like airportIcao2) Then
            Dim airportDep As AirportObject
            Dim airportDest As AirportObject
            Set airportDep = dictAirports.Item(airportIcao1)
            Set airportDest = dictAirports.Item(airportIcao2)
            distanceKm = DistanceCaluculationModul.GetDistanceCoord(airportDep.latitude, airportDep.longitude, airportDest.latitude, airportDest.longitude, "K")
            distanceNM = distanceKm * KM_TO_NM_CONVERSION_CONSTANT
            
            distanceTable.Cells(line, 1) = airportDep.icao
            distanceTable.Cells(line, 2) = airportDest.icao
            distanceTable.Cells(line, 3) = distanceKm
            distanceTable.Cells(line, 4) = distanceNM
            
            line = line + 1
            
        End If
    Next
Next
distanceTable.Range("C1:D" & line).NumberFormat = "#_ ;-#"
distanceTable.Activate
End Sub





Public Function getAirportDictionary() As Scripting.Dictionary

Dim line As Integer
Dim nextLineAvaible As Boolean
Dim icao As String
Dim airportName As String
Dim latitude As Double
Dim longitude As Double
Dim longesRwy As Long


nextLineAvaible = True
Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
line = 2
Dim test As String
test = Range("A2")

Dim airport As AirportObject

Do While nextLineAvaible

    If Not IsEmpty(VaAirportTable.Range("A" & line)) And (IsError(VaAirportTable.Cells(line, 2)) = False) Then
     
        Set airport = New AirportObject
        
        icao = VaAirportTable.Cells(line, 1)
        airportName = VaAirportTable.Cells(line, 2)
        latitude = VaAirportTable.Cells(line, 3)
        longitude = VaAirportTable.Cells(line, 4)
        longesRwy = VaAirportTable.Cells(line, 5)
        
        airport.airportName = airportName
        airport.icao = icao
        airport.latitude = latitude
        airport.longitude = longitude
        airport.maxRunwayLength = longesRwy
        
        Dim key As String
        Dim value As Variant
        key = icao
        Set value = airport
        
        dict.Add key, value
    Else
    nextLineAvaible = False
    
     
    End If
line = line + 1
Loop
Set getAirportDictionary = dict

End Function

