VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Ordinary Threads"
   ClientHeight    =   1875
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2625
   LinkTopic       =   "Form1"
   ScaleHeight     =   1875
   ScaleWidth      =   2625
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Test Thread Safe Exit"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   2415
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Create [*] Nested Threads"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create [3] Standard Threads"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim THREADID As Long
Dim THREADHANDLE As Long

Private Sub Command1_Click()
THREADHANDLE = StartThreadOnAddress(AddressOf Thread, AddressOf ThreadExceptionError, 0, THREADID)
THREADHANDLE = StartThreadOnAddress(AddressOf Thread, AddressOf ThreadExceptionError, 0, THREADID)
THREADHANDLE = StartThreadOnAddress(AddressOf Thread, AddressOf ThreadExceptionError, 0, THREADID)
End Sub


Private Sub Command2_Click()
THREADHANDLE = StartThreadOnAddress(AddressOf NestedThread, AddressOf ThreadExceptionError, 0, THREADID)
End Sub


Private Sub Command3_Click()
THREADHANDLE = StartThreadOnAddress(AddressOf ThreadError, AddressOf ThreadExceptionError2, 0, THREADID)
End Sub
