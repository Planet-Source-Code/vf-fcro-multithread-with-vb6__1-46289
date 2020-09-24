VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeTypeThreader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ObjectProcedures() As FunctionSPointerS
Implements Bridges
Dim BRIDGE As Bridges


Private SHMEM As Long
Private HMEM As Long
Private BRIDGEASM As Long
Private SEHASM As Long

Private OCall As FreeThreadInterface

Private ThrH As Long
Private ThrId As Long
Private IsErrorF As Byte
Private EnEvent As Boolean


Private Sub Bridges_WaitableThreadProcedure()
End Sub
Private Function Bridges_SehW(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
End Function
Private Sub Bridges_ThreadProcedure()

If EnEvent Then EnableEvents
OCall.ThreadStart ThrId, ThrH
ThrH = 0
ThrId = 0
EExitThread 0
End Sub
Private Function Bridges_SehF(ByVal pExcept As Long, ByVal pFrame As Long, ByVal pContext As Long, ByVal pDispatch As Long) As Long
IsErrorF = 1
OCall.ThreadError ThrId, ThrH

CoUninitialize
OleUninitialize
ThrH = 0
ThrId = 0
IsErrorF = 0
EExitThread 0
End Function
Public Property Get Priority() As Long
If ThrH <> 0 Then Priority = GetThreadPriority(ThrH)
End Property
Public Property Let Priority(ByVal PriorityThread As ThreadPriority)
If ThrH <> 0 Then SetThreadPriority ThrH, PriorityThread
End Property

Public Function StartThread(ThreadObject As FreeThreadInterface, ByVal EnableEventsObject As Boolean, ByVal StackSize As Long, ByRef ThreadId As Long, Optional ByVal PriorityThread As ThreadPriority = THREAD_PRIORITY_NORMAL) As Long
If ThrH <> 0 Or IsErrorF = 1 Then StartThread = -1: Exit Function

If HMEM <> 0 Then GlobalUnlock HMEM: GlobalFree HMEM: GlobalUnlock SHMEM: GlobalFree SHMEM
EnEvent = EnableEventsObject

Set OCall = ThreadObject
ObjectProcedures = GetObjectFunctionsPointers(BRIDGE, 4)
HMEM = GlobalAlloc(0, CalculateSpaceForDelegation(0))
BRIDGEASM = GlobalLock(HMEM)
DelegateFunction BRIDGEASM, BRIDGE, ObjectProcedures(0).FunctionAddress, 1

SHMEM = GlobalAlloc(0, CalculateSpaceForDelegation(4))
SEHASM = GlobalLock(SHMEM)
DelegateFunction SEHASM, BRIDGE, ObjectProcedures(2).FunctionAddress, 4
StartThread = CreateMT(BRIDGEASM, StackSize, ThreadId, GUID1, GUID2, "X" & ObjPtr(Me), SEHASM)
SetThreadPriority StartThread, PriorityThread
ThrH = StartThread
ThrId = ThreadId
End Function

Public Function ForceExitThread() As Long
If ThrH <> 0 Then
ThreadOut ThrH, ThrId
GlobalUnlock HMEM
GlobalFree HMEM
HMEM = 0
GlobalUnlock SHMEM
GlobalFree SHMEM
SHMEM = 0
Else
ForceExitThread = -1
End If
End Function

Private Sub Class_Initialize()
Set BRIDGE = Me
End Sub

Private Sub Class_Terminate()
If ThrH <> 0 Then ForceExitThread
End Sub
