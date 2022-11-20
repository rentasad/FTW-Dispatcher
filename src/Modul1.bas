Attribute VB_Name = "Modul1"
Sub downloadAirportDatabase()
Dim urlAirports As String
Dim urlRunways As String
Dim targetPathRunways As String
Dim targetPathAirports As String


urlAirports = "https://davidmegginson.github.io/ourairports-data/airports.csv"
urlRunways = "https://davidmegginson.github.io/ourairports-data/runways.csv"
targetPathAirports = ConfigTable.Cells(17, 2) & "downloads\airports.csv"
targetPathRunways = ConfigTable.Cells(17, 2) & "downloads\runways.csv"


downloadFile urlAirports, targetPathAirports
downloadFile urlRunways, targetPathRunways


End Sub


Public Sub downloadFile(ByVal url, ByVal targetPath)

Set objXmlHttpReq = CreateObject("Microsoft.XMLHTTP")
     objXmlHttpReq.Open "GET", url, False, "username", "password"
     objXmlHttpReq.send

     If objXmlHttpReq.Status = 200 Then
          Set objStream = CreateObject("ADODB.Stream")
          objStream.Open
          objStream.Type = 1
          objStream.Write objXmlHttpReq.responseBody
          objStream.SaveToFile targetPath, 2
          objStream.Close
     End If
End Sub
