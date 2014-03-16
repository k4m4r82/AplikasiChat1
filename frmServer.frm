VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmServer 
   Caption         =   "Aplikasi Server"
   ClientHeight    =   2610
   ClientLeft      =   3180
   ClientTop       =   2085
   ClientWidth     =   6510
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   6510
   Begin VB.TextBox txtChat 
      Height          =   285
      Left            =   690
      TabIndex        =   1
      Top             =   120
      Width           =   5655
   End
   Begin VB.CommandButton cmdSend 
      Caption         =   "Send"
      Height          =   375
      Left            =   690
      TabIndex        =   0
      Top             =   2130
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3600
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox txtMain 
      Height          =   1485
      Left            =   690
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   525
      Width           =   5655
   End
   Begin VB.Label Label1 
      Caption         =   "Pesan"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   450
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdSend_Click()
    Winsock1.SendData txtChat.Text
    DoEvents
    
    txtMain.Text = txtMain.Text & vbCrLf & txtChat.Text
    txtChat.Text = ""
End Sub

Private Sub Form_Load()
    Winsock1.LocalPort = 11111
    Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
    'reset the socket, and accept the new connection
    Winsock1.Close
    Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
    Dim strData As String

    'get the data and display it in the textbox
    Winsock1.GetData strData
    txtMain.Text = txtMain.Text & vbCrLf & strData
    txtMain.SelStart = Len(txtMain.Text)
End Sub

