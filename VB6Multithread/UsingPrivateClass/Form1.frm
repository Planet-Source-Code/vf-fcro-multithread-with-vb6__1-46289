VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   810
   ScaleWidth      =   4680
   Begin VB.CommandButton Command1 
      Caption         =   "Test Event!"
      Height          =   615
      Left            =   1560
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents K As THR





Private Sub Command1_Click()
K.TestEvent 122
End Sub

Private Sub Form_Load()
Caption = "Thread Id:" & GetCurrentThreadId
Set K = New THR
K.StartThread
End Sub

Private Sub Form_Unload(Cancel As Integer)
K.EndThread
End Sub

Private Sub K_CallBackEvent(ByVal Value As Long)
MsgBox "EVENT ON START", , "Thread Id:" & GetCurrentThreadId
End Sub
