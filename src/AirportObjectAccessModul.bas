Attribute VB_Name = "AirportObjectAccessModul"
'@Folder("Modules.Airport")


Public Function getNewExcelConnection() As connection
    Dim cn As Object
    Set cn = CreateObject("ADODB.Connection")
    With cn
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionString = "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    Set getNewConnection = cn
End Function


Public Function getAirportDictionary() As Scripting.Dictionary

Dim line As Integer
Dim nextLineAvaible As Boolean
Dim icao As String
Dim airportName As String
Dim latitude As String
Dim longitude As String
Dim longestRwy As Long
Dim terminalType As String
Dim terminalSize As Long
Dim cargoSize As Long


nextLineAvaible = True
Dim dict As Scripting.Dictionary
    Set dict = New Scripting.Dictionary
    
line = 2
Dim test As String
test = Range("A2")

Dim airport As airportObject
Dim verifyColumn As Variant

Const NOT_IN_DATABASE_STRING As String = "Not in Database"

Do While nextLineAvaible

    If Not IsEmpty(VAFboTable.Range("A" & line)) And (IsError(VAFboTable.Cells(line, 2)) = False) Then
     
        
        ' This string has value "NOT IN DATABASE" if airport was not found in global Airport database
        verifyColumn = VAFboTable.Cells(line, AirportModul.COLUMN_MAX_RUNWAY_LENGTH)
        icao = VAFboTable.Cells(line, AirportModul.COLUMN_ICAO)
        resultCompare = StrComp(verifyColumn, NOT_IN_DATABASE_STRING)
        If (resultCompare = -1) Then
            Set airport = New airportObject
            
            
            airportName = VAFboTable.Cells(line, AirportModul.COLUMN_AIRPORT_NAME)
            latitude = VAFboTable.Cells(line, AirportModul.COLUMN_LATITUDE)
            longitude = VAFboTable.Cells(line, AirportModul.COLUMN_LONGITUDE)
            longestRwy = VAFboTable.Cells(line, AirportModul.COLUMN_MAX_RUNWAY_LENGTH)
            terminalType = VAFboTable.Cells(line, AirportModul.COLUMN_TERMINAL_TYPE)
            If (IsNumeric(VAFboTable.Cells(line, AirportModul.COLUMN_TERMINAL_SIZE))) Then
                terminalSize = VAFboTable.Cells(line, AirportModul.COLUMN_TERMINAL_SIZE)
                airport.terminalSize = terminalSize
            End If
            If (IsNumeric(VAFboTable.Cells(line, AirportModul.COLUMN_CARGO_SIZE))) Then
                cargoSize = VAFboTable.Cells(line, AirportModul.COLUMN_CARGO_SIZE)
                airport.cargoSize = cargoSize
            End If
            airport.airportName = airportName
            airport.icao = icao
            airport.latitude = Replace(latitude, ".", ",")
            airport.longitude = Replace(longitude, ".", ",")
            airport.maxRunwayLength = longestRwy
            airport.terminalType = terminalType
            
            
            
            Dim key As String
            Dim value As Variant
            key = icao
            Set value = airport
            dict.Add key, value
        Else
            ' Airport not in Database -> Entry Skipped
            If (existAirportInManualInputTable(icao)) Then
                Set airport = getAirportObjectFromManualInputSheet(icao)
                
            End If
            
            
        End If
        
    Else
    nextLineAvaible = False
    
     
    End If
line = line + 1
Loop
Set getAirportDictionary = dict

End Function
Public Function getAirportObjectFromManualInputSheet(ByVal icao As String)
    If (existAirportInManualInputTable(icao)) Then
    Dim ao As airportObject
    Set ao = New airportObject
    
    Dim connection As Object
    Dim rs As ADODB.Recordset
    Dim counter As Long
    Set connection = CreateObject("ADODB.Connection")
    With connection
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionString = "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    
    Dim query As String
    query = "SELECT * FROM [Manual Inputs$] WHERE [ICAO] = '" & icao & "'"
    Set rs = connection.Execute(query)
    ao.airportName = rs(COLUMN_AIRPORT_NAME - 1)
    Dim cargoSize As Variant
    cargoSize = rs(COLUMN_CARGO_SIZE - 1)
    If Not (IsNull(cargoSize)) Then
        ao.cargoSize = cargoSize
    End If
    
    ao.icao = rs(COLUMN_ICAO - 1)
    ao.longitude = rs(COLUMN_LONGITUDE - 1)
    ao.latitude = rs(COLUMN_LATITUDE - 1)
    ao.maxRunwayLength = rs(COLUMN_MAX_RUNWAY_LENGTH - 1)
    ao.terminalType = rs(COLUMN_TERMINAL_TYPE - 1)
    ao.terminalSize = rs(COLUMN_TERMINAL_SIZE - 1)
    
    rs.Close
    connection.Close
    Set rs = Nothing
    Set connection = Nothing
    
    Set getAirportObjectFromManualInputSheet = ao
    End If
End Function


' Return true if Airport exist in Manual Inputs Sheet
Public Function existAirportInManualInputTable(ByVal icao As String) As Boolean
    Dim counter As Long
    Dim connection As Object
    Set connection = CreateObject("ADODB.Connection")
    With connection
        .Provider = "Microsoft.ACE.OLEDB.12.0"
        .connectionString = "Data Source=" & ThisWorkbook.Path & "\" & ThisWorkbook.Name & ";" & _
        "Extended Properties=""Excel 12.0 Xml;HDR=YES"";"
        .Open
    End With
    
    Dim query As String
    query = "SELECT count(ICAO) as anzahl FROM [Manual Inputs$] WHERE [ICAO] = '" & icao & "'"
    Set rs = connection.Execute(query)
    
    counter = rs.Fields().Item("anzahl")
    rs.Close
    connection.Close
      Set rs = Nothing
    Set connection = Nothing
    
    existAirportInManualInputTable = counter > 0

End Function


