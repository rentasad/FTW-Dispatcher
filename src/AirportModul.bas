Attribute VB_Name = "AirportModul"
'@Folder "Modules.Airport"

 Public Const COLUMN_ICAO As Long = 1
 Public Const COLUMN_TERMINAL_TYPE As Long = 2
 Public Const COLUMN_TERMINAL_SIZE As Long = 3
 Public Const COLUMN_CARGO_SIZE As Long = 4
 Public Const COLUMN_LATITUDE As Long = 5
 Public Const COLUMN_LONGITUDE As Long = 6
 Public Const COLUMN_MAX_RUNWAY_LENGTH As Long = 7
 Public Const COLUMN_AIRPORT_NAME As Long = 8


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

Dim lon As String
Dim lat As String

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







