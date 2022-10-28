Attribute VB_Name = "DistanceCaluculationModul"
Enum Measurement
    Miles = 3958.756
    Km = 6371
    NM = 3440.065
    Meters = 6371000
End Enum

'Source http://www.odosmatthewscoding.com/2019/06/how-to-calculate-distance-between-two.html
Function dDistance(ByRef lat1 As Double, _
                   ByRef lon1 As Double, _
                   ByRef lat2 As Double, _
                   ByRef lon2 As Double, _
                   m As Measurement) As Double
    Dim func As Object
    If lat1 = lat2 And lon1 = lon2 Then
        dDistance = 0
        Exit Function
    End If
    Set func = Application.WorksheetFunction
    dDistance = func.Acos(Cos(func.Radians(90 - lat1)) * _
                    Cos(func.Radians(90 - lat2)) + _
                    Sin(func.Radians(90 - lat1)) * _
                    Sin(func.Radians(90 - lat2)) * _
                    Cos(func.Radians(lon1 - lon2))) * m
End Function


' SOURCE: https://www.exceldemy.com/excel-vba-calculate-distance-between-two-addresses-or-coordinates/#Download_Practice_Workbook
Public Function getDistanceByApi(ByVal lat1 As String, ByVal lon1 As String, ByVal lat2 As String, ByVal lon2 As String)
Dim startlocation As String
Dim destination As String
startlocation = lat1 & ", " & lon1
destination = CStr(lat2) & ", " & CStr(lon2)

Debug.Print ("Startlocation: " & startlocation)
Debug.Print ("Destination:" & destination)

Dim keyvalue As String

Dim Initial_Value As String
Dim Second_Value As String
Dim Destination_Value As String
Dim mitHTTP As Object
Dim mitUrl As String

keyvalue = ConfigTable.Cells(15, 2)

    Initial_Value = "https://dev.virtualearth.net/REST/v1/Routes/DistanceMatrix?origins="
    Second_Value = "&destinations="
    Destination_Value = "&travelMode=transit&o=xml&key=" & keyvalue & "&distanceUnit=km"

    Set mitHTTP = CreateObject("MSXML2.ServerXMLHTTP")

    mitUrl = Initial_Value & startlocation & Second_Value & destination & Destination_Value
    mitHTTP.Open "GET", mitUrl, False
    mitHTTP.SetRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
    mitHTTP.Send ("")
    Dim response As String
    response = mitHTTP.ResponseText
    Debug.Print (response)
    
    getDistanceByApi = Round(Round(WorksheetFunction.FilterXML(mitHTTP.ResponseText, "//TravelDistance"), 3), 0)
End Function


Private Function DegreesToRadians(degrees As Double) As Double
   
    Const PI As Single = 3.14159265358979
    DegreesToRadians = degrees * PI / 180
End Function



' SOURCE:https://analystcave.com/excel-calculate-distances-between-addresses/

Public Function GetDistanceCoord(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double, ByVal unit As String) As Double
    Dim theta As Double: theta = lon1 - lon2
    Dim dist As Double: dist = Math.Sin(deg2rad(lat1)) * Math.Sin(deg2rad(lat2)) + Math.Cos(deg2rad(lat1)) * Math.Cos(deg2rad(lat2)) * Math.Cos(deg2rad(theta))
    dist = WorksheetFunction.Acos(dist)
    dist = rad2deg(dist)
    dist = dist * 60 * 1.1515
    If unit = "K" Then
        dist = dist * 1.609344
    ElseIf unit = "N" Then
        dist = dist * 0.8684
    End If
    GetDistanceCoord = dist
End Function


 
Function deg2rad(ByVal deg As Double) As Double
    deg2rad = (deg * WorksheetFunction.PI / 180#)
End Function
 
Function rad2deg(ByVal rad As Double) As Double
    rad2deg = rad / WorksheetFunction.PI * 180#
End Function

