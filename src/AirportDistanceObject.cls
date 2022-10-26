VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AirportDistanceObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mstrIcaoFrom As String
Private mstrIcaoTo As String
Private mdbldistanceMiles As Double
Private mdbldistanceMeters As Double
Property Get IcaoFrom() As String
    IcaoFrom = mstrIcaoFrom
End Property
Property Let IcaoFrom(rData As String)
    mstrIcaoFrom = rData
End Property
Property Get IcaoTo() As String
    IcaoTo = mstrIcaoTo
End Property
Property Let IcaoTo(rData As String)
    mstrIcaoTo = rData
End Property
Property Get distanceMiles() As Double
    distanceMiles = mdbldistanceMiles
End Property
Property Let distanceMiles(rData As Double)
    mdbldistanceMiles = rData
End Property
Property Get distanceMeters() As Double
    distanceMeters = mdbldistanceMeters
End Property
Property Let distanceMeters(rData As Double)
    mdbldistanceMeters = rData
End Property
