VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cConnectionTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"No"
Public Key As String

'local variable(s) to hold property value(s)
Private mvarIPAddress As String 'local copy
Private mvarUID As String 'local copy
Private mvarPort As String 'local copy
'local variable(s) to hold property value(s)
Private mvarWinsockIndex As Long 'local copy
Public Property Let WinsockIndex(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.WinsockIndex = 5
    mvarWinsockIndex = vData
End Property


Public Property Get WinsockIndex() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.WinsockIndex
    WinsockIndex = mvarWinsockIndex
End Property



Public Property Let Port(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Port = 5
    mvarPort = vData
End Property


Public Property Get Port() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Port
    Port = mvarPort
End Property



Public Property Let UID(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.UID = 5
    mvarUID = vData
End Property


Public Property Get UID() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.UID
    UID = mvarUID
End Property



Public Property Let IPAddress(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.IPAddress = 5
    mvarIPAddress = vData
End Property


Public Property Get IPAddress() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.IPAddress
    IPAddress = mvarIPAddress
End Property



