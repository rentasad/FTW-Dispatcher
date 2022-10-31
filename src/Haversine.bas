Attribute VB_Name = "Haversine"
'@Folder "Modules.Utilities"
Public Function getDistance(ByVal dbl_Longitude1 As Double, dbl_Longitude2 As Double, dbl_Latitude1 As Double, dbl_Latitude2 As Double) As Double


    
    Dim dbl_dLat As Double
    Dim dbl_dLon As Double
    Dim dbl_a As Double
    Dim dbl_P As Double

    dbl_P = WorksheetFunction.PI / 180
    dbl_dLat = dbl_P * (dbl_Latitude2 - dbl_Latitude1)      'to radians
    dbl_dLon = dbl_P * (dbl_Longitude2 - dbl_Longitude1)    'to radians

    dbl_a = Sin(dbl_dLat / 2) * Sin(dbl_dLat / 2) + _
            Cos(dbl_Latitude1 * dbl_P) * Cos(dbl_Latitude2 * dbl_P) * Sin(dbl_dLon / 2) * Sin(dbl_dLon / 2)

    Dim c As Double
    Dim dbl_Distance_KM As Double
    c = 2 * WorksheetFunction.Atan2(Sqr(1 - dbl_a), Sqr(dbl_a))  ' *** swapped arguments to Atan2
    dbl_Distance_KM = 6371 * c

    Debug.Print dbl_Distance_KM
    getDistance = dbl_Distance_KM
End Function
