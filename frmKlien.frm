VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmKlien 
   Caption         =   "Aplikasi Klien"
   ClientHeight    =   2625
   ClientLeft      =   9975
   ClientTop       =   2085
   ClientWidth     =   6465
   LinkTopic       =   "Form1"
   ScaleHeight     =   2625
   ScaleWidth      =   6465
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   690
      TabIndex        =   2
      Top             =   2130
      Width           =   975
   End
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   690
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   4080
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   1485
      Left            =   690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   525
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Pesan"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmKlien"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    'set up the Winsock1 to connect to the local computer
    Winsock1.RemoteHost = "127.0.0.1"
    Winsock1.RemotePort = 11111
    Winsock1.Connect
End Sub

Private Sub cmdSend_Click()
    'send the data thats in the text box and
    'clear it to prepare for the next chat message
    Winsock1.SendData txtChat.Text
    DoEvents
    
    txtMain.Text = txtMain.Text & vbCrLf & txtChat.Text
    txtChat.Text = ""
End Sub

Private Sub Winsock1_Connect()
    'we are connected!
    MsgBox "Connected"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String
    
    ' get the data from the socket
    Winsock1.GetData strData
    ' display it in the textbox
    txtMain.Text = txtMain.Text & vbCrLf & strData
    ' scroll the box down
    txtMain.SelStart = Len(txtMain.Text)
End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    ' an error has occured somewhere, so let the user know
    MsgBox "Error: " & Description
    ' close the socket, ready to go again
    Winsock1.Close
End Sub

