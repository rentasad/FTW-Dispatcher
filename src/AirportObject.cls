VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AirportObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private mstricao As String
Private mstrairportName As String
Private mdbllongitude As String
Private mdbllatitude As String
Private mlngmaxRunwayLength As Long
Property Get icao() As String
    icao = mstricao
End Property
Property Let icao(rData As String)
    mstricao = rData
End Property
Property Get airportName() As String
    airportName = mstrairportName
End Property
Property Let airportName(rData As String)
    mstrairportName = rData
End Property
Property Get longitude() As String
    longitude = mdbllongitude
End Property
Property Let longitude(rData As String)
    mdbllongitude = rData
End Property
Property Get latitude() As String
    latitude = mdbllatitude
End Property
Property Let latitude(rData As String)
    mdbllatitude = rData
End Property
Property Get maxRunwayLength() As Long
    maxRunwayLength = mlngmaxRunwayLength
End Property
Property Let maxRunwayLength(rData As Long)
    mlngmaxRunwayLength = rData
End Property



