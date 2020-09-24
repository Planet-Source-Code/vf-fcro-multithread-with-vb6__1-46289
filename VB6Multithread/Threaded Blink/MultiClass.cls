VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MultiClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements FreeWaitThreadInterface
Dim THREADER As New WaitTypeThreader
Public BlinkObject As PictureBox

Public COLOR1 As Long
Public COLOR2 As Long
Public SPEED As Long

Public Function RunThread() As Long
Dim TID As Long
THREADER.StartThread Me, False, 0, RunThread
End Function
Public Sub ExitThread()
THREADER.ForceExitThread
End Sub
Public Sub DoSomeAction()
THREADER.CallThread 0, 0
End Sub



'OVO TI JE IMPLEMENTIRANI THREADING INTERFACE [VIDI IMPLEMENTS pod GENERAL]
'SVE PROCEDURE MORAJU BITI TU!
'DAKLE OVE PRAZNE NE BRIŠI!
Private Sub FreeWaitThreadInterface_ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub
Private Sub FreeWaitThreadInterface_ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)
Dim X As Long
Do
BlinkObject.BackColor = COLOR1
RefreshWindow BlinkObject.hwnd

Sleep SPEED

BlinkObject.BackColor = COLOR2
RefreshWindow BlinkObject.hwnd

Sleep SPEED

X = X + 1
Loop While X < 100

End Sub
Private Function FreeWaitThreadInterface_ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long
End Function
Private Sub FreeWaitThreadInterface_ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
Private Sub FreeWaitThreadInterface_ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
Private Sub FreeWaitThreadInterface_ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
Private Sub FreeWaitThreadInterface_ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub
