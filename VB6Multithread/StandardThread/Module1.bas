Attribute VB_Name = "Module1"
Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)



Public Sub Thread()
MsgBox "New Thread Id:" & GetCurrentThreadId, , "Info"
End Sub
Public Function ThreadExceptionError(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
'Error in thread
ExitThread 0
End Function

Public Sub NestedThread()
MsgBox "New Thread Id:" & GetCurrentThreadId, , "Info"
Dim TID As Long
'Create Thread inside the thread
StartThreadOnAddress AddressOf NestedThread, AddressOf ThreadExceptionError, 0, TID
End Sub





Public Sub ThreadError()
MsgBox "New Thread Id:" & GetCurrentThreadId & vbCrLf & "Make an error in thread!", , "Info"
CopyMemory ByVal 2222, ByVal 59999, 100000
End Sub

Public Function ThreadExceptionError2(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
MsgBox "Thread Id:" & GetCurrentThreadId, vbCritical, "Error!"
ExitThread 0
End Function

