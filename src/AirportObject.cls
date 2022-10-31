VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AirportObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Objects.AirportObjects"

Private mstricao As String
Private mstrairportName As String
Private terminalTypeString As String
Private terminalSizeLong As Long
Private mcargoSizeLong As Long
Private mdbllongitude As String
Private mdbllatitude As String
Private mlngmaxRunwayLength As Long




Public Property Get icao() As String
    icao = mstricao
End Property
Public Property Let icao(rData As String)
    mstricao = rData
End Property
Public Property Get airportName() As String
    airportName = mstrairportName
End Property
Public Property Let airportName(rData As String)
    mstrairportName = rData
End Property
Public Property Get longitude() As String
    longitude = mdbllongitude
End Property
Public Property Let longitude(rData As String)
    mdbllongitude = rData
End Property
Public Property Get latitude() As String
    latitude = mdbllatitude
End Property
Public Property Let latitude(rData As String)
    mdbllatitude = rData
End Property
Public Property Get maxRunwayLength() As Long
    maxRunwayLength = mlngmaxRunwayLength
End Property
Public Property Let maxRunwayLength(rData As Long)
    mlngmaxRunwayLength = rData
End Property


Public Property Get terminalType() As String
terminalType = terminalTypeString
End Property

Public Property Let terminalType(ByVal vNewValue As String)
terminalTypeString = vNewValue
End Property
Public Property Get terminalSize() As Long
terminalSize = terminalSizeLong
End Property

Public Property Let terminalSize(ByVal vNewValue As Long)
terminalSizeLong = vNewValue
End Property
Public Property Get cargoSize() As Long
cargoSize = cargoSizeLong
End Property

Public Property Let cargoSize(ByVal vNewValue As Long)
cargoSizeLong = vNewValue
End Property

