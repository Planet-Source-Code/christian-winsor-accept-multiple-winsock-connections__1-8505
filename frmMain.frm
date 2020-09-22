VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   Caption         =   "Messaging Server"
   ClientHeight    =   3360
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   6960
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3360
   ScaleWidth      =   6960
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5160
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   20
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0946
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar sbMain 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   3030
      Width           =   6960
      _ExtentX        =   12277
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9208
            Picture         =   "frmMain.frx":0E4A
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock wsMain 
      Index           =   0
      Left            =   0
      Top             =   2880
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Frame fraFunctions 
      Height          =   2895
      Left            =   5160
      TabIndex        =   1
      Top             =   -60
      Width           =   1755
      Begin VB.CommandButton comKick 
         Caption         =   "Kick User"
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   180
         Width           =   1455
      End
   End
   Begin MSComctlLib.ListView lvClient 
      Height          =   2835
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   5001
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "UID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "IP Address"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Port Connected"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuServ 
      Caption         =   "Service"
      Begin VB.Menu mnuServStart 
         Caption         =   "Start Server"
      End
      Begin VB.Menu mnuServStop 
         Caption         =   "Stop Server"
         Enabled         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public cConnectionCol As New cConnectionTrackCol
Public cConnection As New cConnectionTrack

Private Sub comKick_Click()
On Error Resume Next
If lvClient.SelectedItem Then
wsMain(cConnectionCol(lvClient.SelectedItem.Key).WinsockIndex).Close
Unload wsMain(cConnectionCol(lvClient.SelectedItem.Key).WinsockIndex)
lvClient.ListItems.Remove (lvClient.SelectedItem.Index)
'remove the item from the collection using the tag from the winsock control _
that we set earlier
cConnectionCol.Remove lvClient.SelectedItem.Key
End If
End Sub

Private Sub Form_Load()
frmMain.Caption = "Messaging Server " & App.Major & "." & App.Minor
sbMain.Panels(1).Text = "0 Users Online"
sbMain.Panels(2).Text = "Status: Offline"
End Sub

Private Sub Form_Resize()
On Error Resume Next
'resize the controls when the form resizes
lvClient.Width = frmMain.ScaleWidth - fraFunctions.Width
lvClient.Height = frmMain.ScaleHeight - sbMain.Height
fraFunctions.Left = lvClient.Left + lvClient.Width + 5
fraFunctions.Height = frmMain.ScaleHeight + 50 - sbMain.Height
'resize the column headers when the form gets resized
lvClient.ColumnHeaders.Item(1).Width = lvClient.Width / 3
lvClient.ColumnHeaders.Item(2).Width = lvClient.Width / 3
lvClient.ColumnHeaders.Item(3).Width = lvClient.Width / 3
End Sub

Private Sub mnuFileExit_Click()
End
End Sub

Private Sub mnuServStart_Click()
'set the winsock properties
wsMain(0).Protocol = sckTCPProtocol
wsMain(0).LocalPort = 6500
wsMain(0).Listen
mnuServStart.Enabled = False
mnuServStop.Enabled = True
If wsMain(0).State = sckListening Then
    sbMain.Panels(2).Text = "Status: Online"
    sbMain.Panels(2).Picture = ImageList1.ListImages(1).Picture
End If
End Sub

Private Sub LookForOpenPort(ByVal requestID As Long)
Dim sPortDecided As String
Dim itmX As ListItem

'randomize the seed for generation of ports
Randomize
'generate a random port between 3000 and 7000
sPortDecided = Str(Int((7000 - 3000 + 1) * Rnd + 3000))
'look through the collection to see if the port already exists
For Each cConnection In cConnectionCol
    'if the port already exists then call the sub again to get another port number
    If cConnection.Port = sPortDecided Then
        LookForOpenPort requestID
    'the port is not in use and we can continue to connect the user
    Else
        'load a new instance of the winsock control
        Load wsMain(wsMain.UBound + 1)
        'we set the winsock instances tag property to match the key we use when adding _
        it to the collection class, this is so we can search for it later when it _
        needs to be removed
        wsMain(wsMain.UBound).Tag = "c40" & sPortDecided
        'set the port to the random port number that was generated to accept the _
        connection on
        wsMain(wsMain.UBound).LocalPort = CLng(sPortDecided)
        'accept the connection on this port using the new instance of the winsock _
        control
        wsMain(wsMain.UBound).Accept requestID
        'add this connection to the collection so that we can track it
        cConnectionCol.Add wsMain.UBound, "", "", sPortDecided, "c40" & sPortDecided
        'fill in the row for the listview with the information from the client _
        we use a unique key so that we can remove it later using the key
        Set itmX = lvClient.ListItems.Add(, "c40" & sPortDecided)
        itmX.ListSubItems.Add , "c40" & sPortDecided & "1", wsMain(wsMain.UBound).RemoteHostIP
        itmX.ListSubItems.Add , "c40" & sPortDecided & "2", sPortDecided
        If cConnectionCol.Count = 0 Then
            sbMain.Panels(1).Text = "0 Users Online"
        End If
        If cConnectionCol.Count = 1 Then
            sbMain.Panels(1).Text = "1 User Online"
        End If
        If cConnectionCol.Count > 1 Then
            sbMain.Panels(1).Text = cConnectionCol.Count & " Users Online"
        End If
        Exit Sub
    End If
DoEvents
Next
'this if statement is for the first connection run, this is because there are no items _
in the collection class at this point so we make the first connection
If cConnectionCol.Count = 0 Then
    'load the new isntance of the winsock control for the connection to be accepted on
    Load wsMain(wsMain.UBound + 1)
    'set the port that was generated
    wsMain(wsMain.UBound).LocalPort = CLng(sPortDecided)
    'set the tag to the unique key we will use, used for removing from the collection _
    when the connection is closed
    wsMain(wsMain.UBound).Tag = "c40" & sPortDecided
    'accept the connection on the specified port in the new instance of the control
    wsMain(wsMain.UBound).Accept requestID
    'add this connection to the class
    cConnectionCol.Add wsMain.UBound, "", "", sPortDecided, "c40" & sPortDecided
    'add the connection info to the listview control
    Set itmX = lvClient.ListItems.Add(, "c40" & sPortDecided)
    itmX.ListSubItems.Add , "c40" & sPortDecided & "1", wsMain(wsMain.UBound).RemoteHostIP
    itmX.ListSubItems.Add , "c40" & sPortDecided & "2", sPortDecided
    If cConnectionCol.Count = 0 Then
        sbMain.Panels(1).Text = "0 Users Online"
    End If
    If cConnectionCol.Count = 1 Then
        sbMain.Panels(1).Text = "1 User Online"
    End If
    If cConnectionCol.Count > 1 Then
        sbMain.Panels(1).Text = cConnectionCol.Count & " Users Online"
    End If
    Exit Sub
End If
End Sub

Private Sub mnuServStop_Click()
Dim iCounter As Integer
On Error Resume Next
For iCounter = 1 To wsMain.UBound
    wsMain(iCounter).Close
    Unload wsMain(iCounter)
DoEvents
Next iCounter
wsMain(0).Close
Set cConnectionCol = Nothing
Set cConnection = Nothing
lvClient.ListItems.Clear
mnuServStop.Enabled = False
mnuServStart.Enabled = True
If wsMain(0).State = sckClosed Then
    sbMain.Panels(2).Text = "Status: Offline"
    sbMain.Panels(2).Picture = ImageList1.ListImages(2).Picture
End If
End Sub

Private Sub wsMain_ConnectionRequest(Index As Integer, ByVal requestID As Long)
'calls the sub to look through the collection to find an open port
LookForOpenPort requestID
End Sub

Private Sub wsMain_Error(Index As Integer, ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Dim iIndexToRemove As Integer
Debug.Print Description
'if the connection is lost then remove the connection from the collection and unload _
the winsock control instance it was using
If Number = 10054 Then
    'remove the item from the collection using the tag from the winsock control _
    that we set earlier
    cConnectionCol.Remove wsMain(Index).Tag
    'call the function to find out what index number the corresponding tag refers to
    iIndexToRemove = FindListItemToRemove(wsMain(Index).Tag)
    'remove the listitem from the listview control
    lvClient.ListItems.Remove (iIndexToRemove)
    'unload the instance of the winsock control
    Unload wsMain(Index)
End If
If cConnectionCol.Count = 0 Then
    sbMain.Panels(1).Text = "0 Users Online"
End If
If cConnectionCol.Count = 1 Then
    sbMain.Panels(1).Text = "1 User Online"
End If
If cConnectionCol.Count > 1 Then
    sbMain.Panels(1).Text = cConnectionCol.Count & " Users Online"
End If

If Number <> 10054 Then
    sbMain.Panels(2).Text = "Status: Error " & Number & " : " & Description
End If
End Sub

Private Function FindListItemToRemove(Tag As String) As Integer
Dim iCounter As Integer
'cycle through the listview control to find the index that we want to get rid of
For iCounter = 1 To lvClient.ListItems.Count
    'if the current listitem in the listview control has a matching tag then we _
    know that this is the item that we want to remove
    If lvClient.ListItems(iCounter).Key = Tag Then
        'return the index value through the functions return
        FindListItemToRemove = iCounter
        'exit the function since we already found the listitem we were looking for
        Exit Function
    End If
DoEvents
Next iCounter
End Function
