VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   7245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   2
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Sync Call Threads"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   5
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Async Call Threads"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   2880
      Width           =   1935
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Destroy Threads / Create New"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   3
      Top             =   2880
      Width           =   2655
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2595
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   7215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim WithEvents MTT1 As MTConnector
Attribute MTT1.VB_VarHelpID = -1
Dim WithEvents MTT2 As MTConnector
Attribute MTT2.VB_VarHelpID = -1


Private Sub Command1_Click()
List1.Clear
Randomize
Dim U As Long
For U = 0 To 4

MTT1.CallThread CLng(Rnd * 212), CLng(Rnd * 188)
MTT2.CallThread CLng(Rnd * 212), CLng(Rnd * 188)
Next U

'Look at results of Async calling!
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
MTT1.EndThread
MTT2.EndThread

MTT1.StartThread
MTT2.StartThread
Label1 = "Created 2 FreeThreads with Id:" & MTT1.ThreadId & "," & MTT2.ThreadId


End Sub

Private Sub Command4_Click()
List1.Clear
Randomize
Dim U As Long
For U = 0 To 4

MTT1.CallBackThread CLng(Rnd * 212), CLng(Rnd * 188)
MTT2.CallBackThread CLng(Rnd * 212), CLng(Rnd * 188)
Next U
End Sub

Private Sub Form_Load()
Set MTT1 = New MTConnector
Set MTT2 = New MTConnector
Caption = "Main Thread Id:" & GetCurrentThreadId
End Sub

Private Sub Form_Unload(Cancel As Integer)
MTT1.EndThread
MTT2.EndThread
End Sub

Private Sub MTT1_ThreadCallBackUs(ByVal Reason As Long, ByVal Message As Long)
List1.AddItem "SYNC CALL" & vbTab & "Reason:" & Reason & ",Message:" & Message & ", THREAD ID:" & GetCurrentThreadId

End Sub

Private Sub MTT1_ThreadCallUs(ByVal Reason As Long, ByVal Message As Long)
List1.AddItem "ASYNC CALL" & vbTab & "Reason:" & Reason & ",Message:" & Message & ", THREAD ID:" & GetCurrentThreadId
End Sub

Private Sub MTT2_ThreadCallBackUs(ByVal Reason As Long, ByVal Message As Long)
List1.AddItem "SYNC CALL" & vbTab & "Reason:" & Reason & ",Message:" & Message & ", THREAD ID:" & GetCurrentThreadId
End Sub

Private Sub MTT2_ThreadCallUs(ByVal Reason As Long, ByVal Message As Long)
List1.AddItem "ASYNC CALL" & vbTab & "Reason:" & Reason & ",Message:" & Message & ", THREAD ID:" & GetCurrentThreadId
End Sub
