VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MySqlConfigObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Objects"
Private mstrhostname As String
Private mstrdatabase As String
Private mstrusername As String
Private mstrpassword As String
Private mlngport As Long
Public Property Get hostname() As String
    hostname = mstrhostname
End Property
Public Property Let hostname(rData As String)
    mstrhostname = rData
End Property
Public Property Get database() As String
    database = mstrdatabase
End Property
Public Property Let database(rData As String)
    mstrdatabase = rData
End Property
Public Property Get username() As String
    username = mstrusername
End Property
Public Property Let username(rData As String)
    mstrusername = rData
End Property
Public Property Get password() As String
    password = mstrpassword
End Property
Public Property Let password(rData As String)
    mstrpassword = rData
End Property
Public Property Get port() As Long
    port = mlngport
End Property
Public Property Let port(rData As Long)
    mlngport = rData
End Property
