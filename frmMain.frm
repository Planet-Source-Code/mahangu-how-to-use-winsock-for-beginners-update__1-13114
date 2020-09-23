VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Winsock Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Close Winscok Connection"
      Height          =   495
      Left            =   2760
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Send Data"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2160
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open Winsock Connection"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Connect "127.0.0.1", "100"

End Sub

Private Sub Command2_Click()
Winsock1.SendData ("Test")

End Sub

Private Sub Command3_Click()
Winsock1.Close
End Sub

Private Sub Form_Load()
Winsock1.LocalPort = 100
Winsock1.Listen 'Tells winsock to listen for Data


End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Dim requestID
If Winsock1.State <> sckConnected Then
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim StrData ' Declares the data string (can be placed in the global declarations)
Winsock1.GetData StrData 'Tells winsock to get the data and put it in the data string
MsgBox StrData 'Shows the data in a message box

End Sub

