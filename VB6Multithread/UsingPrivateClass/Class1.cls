VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "THR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements VBFreeThreading.FreeWaitThreadInterface
Dim THREADER As New WaitTypeThreader


Event CallBackEvent(ByVal Value As Long)
Dim WithEvents CLS As PrivateTest
Attribute CLS.VB_VarHelpID = -1

Public Sub StartThread()
Dim TID As Long
Call THREADER.StartThread(Me, True, 0, TID) '2nd parameter works only with compiled EXE,[if not THREAD WILL EXIT WITH AN ERROR BECAUSE OLE/COM COULDN'T FIND INTERNAL STRUCTURE!]
'in IDE set it to FALSE!
'Warning:DONT COMPRESS YOUR EXE FILE BECAUSE ALL PRIVATE OLE/COM STUFF AND CONTEXTS WILL NOT
'WORK BECAUSE VB COMPILE & SET FIRST ENTRY POINT INSTRUCTION WITH POINTER TO THAT STRUCTURE!
'PUSH XXXXXXX ;<----------Address
'CALL MSVBVM60.RTThunderMain
'If you compress your file set 2nd param to FALSE [like IDE],but you'll not able
'to use private objects & evented object!


End Sub
Public Sub EndThread()
THREADER.ForceExitThread
End Sub
Public Sub TestEvent(ByVal Value As Long)
THREADER.CallThread Value, 0
End Sub



Private Sub CLS_OnCall(ByVal Value As Long)
RaiseEvent CallBackEvent(Value)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)
CLS.Expose Reason
End Sub

Private Function FreeWaitThreadInterface_ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long
End Function

Private Sub FreeWaitThreadInterface_ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
Set CLS = Nothing
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
Set CLS = Nothing
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
Set CLS = New PrivateTest 'CREATE PRIVATE CLASS inside new thread!!!!!!!!!!!
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
End Sub

