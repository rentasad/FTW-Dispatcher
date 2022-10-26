Attribute VB_Name = "DistanceCaluculationModul"
Private Function degreesToRadians(degrees As Double) As Double
    Const PI As Single = 3.14159265358979
    degreesToRadians = degrees * PI / 180
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

