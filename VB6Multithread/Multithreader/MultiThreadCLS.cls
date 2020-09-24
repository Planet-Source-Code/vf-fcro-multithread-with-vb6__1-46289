VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WaitTypeThreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private ObjectProcedures() As FunctionSPointerS
Implements Bridges
Dim BRIDGE As Bridges

Private QueryM(2) As Long
'QUERYM(0)=MESSAGE
'QUERYM(1)=EVENT
'QUERYM(2)=REPLY


Private EvSync(2) As Long
'EvSysnc(0)=Buffer
'EvSync(1)=Bufferlength
'EvSync(2)=EVENT


Public Enum ThreadPriority
    THREAD_PRIORITY_IDLE = -15
    THREAD_PRIORITY_LOWEST = -2
    THREAD_PRIORITY_BELOW_NORMAL = -1
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_ABOVE_NORMAL = 1
    THREAD_PRIORITY_HIGHEST = 2
    THREAD_PRIORITY_TIME_CRITICAL = 15
End Enum


Private HMEM2 As Long
Private SHMEM2 As Long
Private BRIDGEASM2 As Long
Private SEHASM2 As Long


Private OCall2 As FreeWaitThreadInterface

Private WThrH As Long
Private WThrId As Long
Private IsErrorW As Byte
Private EnEvent As Boolean



Private Function Bridges_SehF(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
End Function

Private Sub Bridges_ThreadProcedure()
End Sub

Private Sub Bridges_WaitableThreadProcedure()

If EnEvent Then EnableEvents

OCall2.ThreadOnStart WThrId, WThrH

Dim Wmsg As msg
Dim GLPTR As Long
Dim GLSIZE As Long
Do While GetMessage(Wmsg, 0&, 0&, 0&)
If Wmsg.hwnd = 0 Then
If Wmsg.Message = &H9999& Then
Exit Do

ElseIf Wmsg.Message = &H9998& Then
GLSIZE = GlobalSize(Wmsg.wParam)
GLPTR = GlobalLock(Wmsg.wParam)
OCall2.ThreadOnAsyncDataTransfer WThrId, WThrH, Wmsg.lParam, GLPTR, GLSIZE
GlobalUnlock Wmsg.wParam
GlobalFree Wmsg.wParam


ElseIf Wmsg.Message = &H9997& Then
OCall2.ThreadOnCall WThrId, WThrH, Wmsg.lParam, Wmsg.wParam

ElseIf Wmsg.Message = &H9996& Then
Dim QRPY(2) As Long
CopyMemory QRPY(0), ByVal Wmsg.wParam, 12
QRPY(2) = OCall2.ThreadOnCallback(WThrId, WThrH, Wmsg.lParam, QRPY(0))
CopyMemory ByVal Wmsg.wParam + 8, QRPY(2), 4
SetEvent QRPY(1)

ElseIf Wmsg.Message = &H9995& Then
Dim SYQRPY(2) As Long
CopyMemory SYQRPY(0), ByVal Wmsg.wParam, 12
OCall2.ThreadOnSyncDataTransfer WThrId, WThrH, Wmsg.lParam, SYQRPY(0), SYQRPY(1)
SetEvent SYQRPY(2)


End If
End If
Call TranslateMessage(Wmsg)
Call DispatchMessage(Wmsg)
Loop



OCall2.ThreadOnExit WThrId, WThrH
WThrH = 0
WThrId = 0
CloseHandle EvSync(2)
CloseHandle QueryM(1)
EExitThread 0
End Sub


Private Function Bridges_SehW(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
'IsErrorW = 1
'OCall2.ThreadOnExitError WThrId, WThrH

Dim TEIP As Long
CopyMemory TEIP, ByVal pContext + &HB8&, 4
If TEIP = 0 Then
OCall2.ThreadOnExit WThrId, WThrH
Else
IsErrorW = 1
OCall2.ThreadOnExitError WThrId, WThrH
End If

CoUninitialize
OleUninitialize
WThrH = 0
WThrId = 0
IsErrorW = 0
SetEvent EvSync(2)
SetEvent QueryM(1)
CloseHandle EvSync(2)
CloseHandle QueryM(1)
EExitThread 0
End Function

Public Function StartThread(ThreadObject As FreeWaitThreadInterface, ByVal EnableEventsObject As Boolean, ByVal StackSize As Long, ByRef ThreadId As Long, Optional ByVal PriorityThread As ThreadPriority = THREAD_PRIORITY_NORMAL) As Long
If WThrH <> 0 Or IsErrorW = 1 Then StartThread = -1: Exit Function
If HMEM2 <> 0 Then GlobalUnlock HMEM2: GlobalFree HMEM2: GlobalUnlock SHMEM2: GlobalFree SHMEM2
EnEvent = EnableEventsObject

QueryM(1) = CreateEvent(ByVal 0&, 0, 0, "Y" & ObjPtr(Me))
If QueryM(1) = 0 Then StartThread = -1: Exit Function
EvSync(2) = CreateEvent(ByVal 0&, 0, 0, "E" & ObjPtr(Me))
If EvSync(2) = 0 Then CloseHandle QueryM(1): StartThread = -1: Exit Function

Set OCall2 = ThreadObject
ObjectProcedures = GetObjectFunctionsPointers(BRIDGE, 4)
HMEM2 = GlobalAlloc(0, CalculateSpaceForDelegation(0))
BRIDGEASM2 = GlobalLock(HMEM2)
DelegateFunction BRIDGEASM2, BRIDGE, ObjectProcedures(1).FunctionAddress, 1

SHMEM2 = GlobalAlloc(0, CalculateSpaceForDelegation(4))
SEHASM2 = GlobalLock(SHMEM2)
DelegateFunction SEHASM2, BRIDGE, ObjectProcedures(3).FunctionAddress, 4

StartThread = CreateMT(BRIDGEASM2, StackSize, ThreadId, GUID1, GUID2, "X" & ObjPtr(Me), SEHASM2)
SetThreadPriority StartThread, PriorityThread
WThrH = StartThread
WThrId = ThreadId
End Function

Public Property Get Priority() As Long
If WThrH <> 0 Then Priority = GetThreadPriority(WThrH)
End Property
Public Property Let Priority(ByVal PriorityThread As ThreadPriority)
If WThrH <> 0 Then SetThreadPriority WThrH, PriorityThread
End Property


Public Function SyncSendDataToThread(ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long) As Long
If WThrH <> 0 Then
EvSync(0) = lpBufferData
EvSync(1) = BufferLength
If PostThreadMessage(WThrId, &H9995&, VarPtr(EvSync(0)), Reason) = 0 Then SyncSendDataToThread = -1: Exit Function
If WaitForSingleObject(EvSync(2), &HFFFFFFFF) <> 0 Then SyncSendDataToThread = -1
Else
SyncSendDataToThread = -1
End If
End Function

Public Function AsyncSendDataToThread(ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long) As Long
If WThrH <> 0 Then
AsyncSendDataToThread = SendData(WThrId, lpBufferData, BufferLength, Reason, &H9998&)
Else
AsyncSendDataToThread = -1
End If
End Function

Public Function CallThread(ByVal Reason As Long, ByVal Message As Long) As Long
If WThrH <> 0 Then
If PostThreadMessage(WThrId, &H9997&, Message, Reason) = 0 Then CallThread = -1
Else
CallThread = -1
End If
End Function

Public Function CallBackThread(ByVal Reason As Long, ByVal Message As Long) As Long
Dim ret As Long
If WThrH <> 0 Then
QueryM(0) = Message
If PostThreadMessage(WThrId, &H9996&, VarPtr(QueryM(0)), Reason) = 0 Then CallBackThread = -1: Exit Function
ret = WaitForSingleObject(QueryM(1), &HFFFFFFFF)
If ret <> 0 Then
CallBackThread = -1
Else
CallBackThread = QueryM(2)
End If
Else
CallBackThread = -1
End If
End Function
Public Function ExitThread() As Long
If WThrH <> 0 Then
PostThreadMessage WThrId, &H9999&, 0, 0
WaitForSingleObject WThrH, &HFFFFFFFF
Else
ExitThread = -1
End If
End Function

Public Function ForceExitThread() As Long
If WThrH <> 0 Then
ThreadOut WThrH, WThrId
GlobalUnlock HMEM2
GlobalFree HMEM2
HMEM2 = 0
GlobalUnlock SHMEM2
GlobalFree SHMEM2
SHMEM2 = 0
Else
ForceExitThread = -1
End If
End Function


Private Sub Class_Initialize()
Set BRIDGE = Me
End Sub

Private Sub Class_Terminate()
If WThrH <> 0 Then ForceExitThread
End Sub
