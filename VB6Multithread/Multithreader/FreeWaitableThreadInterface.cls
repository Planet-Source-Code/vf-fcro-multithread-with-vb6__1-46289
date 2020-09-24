VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeWaitThreadInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
Public Sub ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)
End Sub
Public Function ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long
End Function
Public Sub ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub
Public Sub ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub
Public Sub ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
Public Sub ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
