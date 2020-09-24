VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTConnector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements FreeWaitThreadInterface
Dim MT As New WaitTypeThreader '-->THREADER COMMUNICATOR


Dim TID As Long
Dim TH As Long
Event ThreadCallUs(ByVal Reason As Long, ByVal Message As Long)
Event ThreadCallBackUs(ByVal Reason As Long, ByVal Message As Long)

Property Get ThreadId() As Long
ThreadId = TID
End Property

Public Sub StartThread()
TH = MT.StartThread(Me, False, 0, TID, THREAD_PRIORITY_NORMAL)
End Sub
Public Sub EndThread()
MT.ForceExitThread
End Sub

Public Sub CallThread(ByVal Reason As Long, ByVal Message As Long)
MT.CallThread Reason, Message
End Sub

Public Sub CallBackThread(ByVal Reason As Long, ByVal Message As Long)
MT.CallBackThread Reason, Message
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)
RaiseEvent ThreadCallUs(Reason, Message)
End Sub

Private Function FreeWaitThreadInterface_ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long
RaiseEvent ThreadCallBackUs(Reason, Message)
End Function

Private Sub FreeWaitThreadInterface_ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub
