VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2205
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3435
   LinkTopic       =   "Form1"
   ScaleHeight     =   2205
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      Height          =   255
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.PictureBox Picture3 
      Height          =   255
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   4
      Top             =   1440
      Width           =   735
   End
   Begin VB.PictureBox Picture2 
      Height          =   255
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   3
      Top             =   1080
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Run Blink"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      Height          =   255
      Left            =   2640
      ScaleHeight     =   195
      ScaleWidth      =   675
      TabIndex        =   0
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label4 
      BackColor       =   &H000000FF&
      Caption         =   "Thread:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label3 
      BackColor       =   &H000000FF&
      Caption         =   "Thread:"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H000000FF&
      Caption         =   "Thread:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H000000FF&
      Caption         =   "Thread:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MYCLASS(3) As New MultiClass


Private Sub Command1_Click()
Dim u As Long
For u = 0 To UBound(MYCLASS)
MYCLASS(u).DoSomeAction
Next u
End Sub

Private Sub Form_Load()
Label1 = "New Thread Id:" & MYCLASS(0).RunThread
Set MYCLASS(0).BlinkObject = Picture1
Label2 = "New Thread Id:" & MYCLASS(1).RunThread
Set MYCLASS(1).BlinkObject = Picture2
Label3 = "New Thread Id:" & MYCLASS(2).RunThread
Set MYCLASS(2).BlinkObject = Picture3
Label4 = "New Thread Id:" & MYCLASS(3).RunThread
Set MYCLASS(3).BlinkObject = Picture4


MYCLASS(0).COLOR1 = &HFF&
MYCLASS(0).COLOR2 = &HC22211
MYCLASS(0).SPEED = 100

MYCLASS(1).COLOR1 = &HFF22&
MYCLASS(1).COLOR2 = &HC211CC
MYCLASS(1).SPEED = 80

MYCLASS(2).COLOR1 = &H338822
MYCLASS(2).COLOR2 = &HCC22&
MYCLASS(2).SPEED = 50

MYCLASS(3).COLOR1 = &H990099
MYCLASS(3).COLOR2 = &H55CC&
MYCLASS(3).SPEED = 20
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim u As Long
For u = 0 To UBound(MYCLASS)
MYCLASS(u).ExitThread 'Na UnLoad uvijek Terminiraj jer ne znaš jel je još aktivan ili ne!
Set MYCLASS(u) = Nothing
Next u
End Sub

