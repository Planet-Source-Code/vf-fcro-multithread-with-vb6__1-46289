VERSION 5.00
Begin VB.Form FX 
   ClientHeight    =   3300
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4350
   LinkTopic       =   "Form1"
   ScaleHeight     =   3300
   ScaleWidth      =   4350
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit APP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   2
      Top             =   2520
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2445
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Top             =   0
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SEND"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   2880
      Width           =   1095
   End
End
Attribute VB_Name = "FX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WS1 As New MTWSCK
Dim WithEvents WS2 As Winsock
Attribute WS2.VB_VarHelpID = -1





Private Sub Command1_Click()
WS2.SendData Text2
DoEvents
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Caption = "Main Thread Id:" & GetCurrentThreadId
WS1.StartThread


Set WS2 = New Winsock
WS2.Close
WS2.LocalPort = 9000
WS2.Listen

'Start Multithread with SOCKET on it!

End Sub

Private Sub Form_Unload(Cancel As Integer)
WS2.Close
WS1.ExitThread
End Sub

Private Sub WS2_ConnectionRequest(ByVal requestID As Long)
WS2.Close
WS2.Accept requestID 'Close Listen & Establish connection!
End Sub

Private Sub WS2_DataArrival(ByVal bytesTotal As Long)
Dim S As String
S = Space(bytesTotal)
WS2.GetData S

Text1 = ">" & S & vbCrLf & Text1
End Sub
