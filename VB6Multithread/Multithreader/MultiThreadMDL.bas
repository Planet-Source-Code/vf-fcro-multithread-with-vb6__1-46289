Attribute VB_Name = "MultiThreadMDL"
Option Explicit
 
 Const MAXIMUM_SUPPORTED_EXTENSION = 512
 Const SIZE_OF_80387_REGISTERS = 80

Public Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(SIZE_OF_80387_REGISTERS - 1) As Byte
  
    Cr0NpxState As Long
End Type


Public Type CONTEXT

    ContextFlags As Long

    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long
    
   FloatSave As FLOATING_SAVE_AREA

    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long



    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long



    Ebp As Long
    Eip As Long
    SegCs As Long
    EFlags As Long
    Esp As Long
    SegSs As Long


   ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION - 1) As Byte

End Type


Public Type GUID
  dwData1 As Long
  wData2 As Integer
  wData3 As Integer
  abData4(7) As Byte
End Type


Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type msg
    hwnd As Long
    Message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Declare Sub EnableEvents Lib "multit" ()

Public Declare Function SetThreadPriority Lib "kernel32" (ByVal hThread As Long, ByVal nPriority As Long) As Long
Public Declare Function GetThreadPriority Lib "kernel32" (ByVal hThread As Long) As Long
'Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As Long) As Long
Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (lpEventAttributes As Any, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'Declare Function OpenEvent Lib "kernel32" Alias "OpenEventA" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal lpName As String) As Long
Declare Function SetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Declare Function GlobalSize Lib "kernel32" (ByVal HMEM As Long) As Long
Public Declare Function CoMarshalInterThreadInterfaceInStream Lib "ole32.dll" (riid As Any, pUnk As Any, ByRef ppStm As Long) As Long
Declare Function CoGetInterfaceAndReleaseStream Lib "ole32.dll" (ByRef pStm As Long, ByVal iid As Long, ppv As Any) As Long
Declare Function ResumeThread Lib "kernel32" (ByVal hThread As Long) As Long
Declare Function SetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Declare Function GetThreadContext Lib "kernel32" (ByVal hThread As Long, lpContext As CONTEXT) As Long
Declare Function SuspendThread Lib "kernel32" (ByVal hThread As Long) As Long
 Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
 Declare Function TranslateMessage Lib "user32" (lpMsg As msg) As Long
Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetHandleInformation Lib "kernel32" (ByVal hObject As Long, lpdwFlags As Long) As Long
Public Declare Sub CoUninitialize Lib "ole32.dll" ()
Public Declare Sub OleUninitialize Lib "ole32.dll" ()
Declare Sub EExitThread Lib "kernel32" Alias "ExitThread" (ByVal dwExitCode As Long)
Declare Function DisableThreadLibraryCalls Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function TerminateThread Lib "kernel32" (ByVal hThread As Long, ByVal dwExitCode As Long) As Long
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function CreateMT Lib "multit" (ByVal StartAddress As Long, ByVal StackSize As Long, ByRef ThreadId As Long, OBJGUID As GUID, DINF As GUID, ByVal EventName As String, ByVal SehProc As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)
Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Declare Function GlobalFree Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Declare Function GlobalLock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Declare Function GlobalUnlock Lib "kernel32.dll" (ByVal HMEM As Long) As Long
Declare Sub DbgBreakPoint Lib "ntdll.dll" ()
Declare Function CLSIDFromString Lib "ole32" (lpsz As Any, pclsid As Any) As Long
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As Long) As Long
Declare Function GetSystemDirectory Lib "kernel32" _
   Alias "GetSystemDirectoryA" _
  (ByVal lpBuffer As String, _
   ByVal nSize As Long) As Long

Public LstInterface As String




Public GUID1 As GUID
Public GUID2 As GUID


Public Type FunctionSPointerS
FunctionPtr As Long
FunctionAddress As Long
End Type
Public Function GetSystemDir() As String
Dim nSize As Long
GetSystemDir = Space(256)
nSize = GetSystemDirectory(GetSystemDir, 256&)
GetSystemDir = Left(GetSystemDir, nSize)
End Function
Public Function CalculateSpaceForDelegation(ByVal NumberOfParameters As Byte) As Long
CalculateSpaceForDelegation = 31 + NumberOfParameters * 3
End Function

Public Sub DelegateFunction(ByVal CallingADR As Long, Obj As Object, ByVal MethodAddress As Long, ByVal NumberOfParameters As Byte)
Dim TmpA As Long
Dim u As Long
TmpA = CallingADR
CopyMemory ByVal CallingADR, &H68EC8B55, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
CallingADR = CallingADR + 4

Dim StackP As Byte
StackP = 4 + 4 * NumberOfParameters

For u = 1 To NumberOfParameters
CopyMemory ByVal CallingADR, CInt(&H75FF), 2
CallingADR = CallingADR + 2
CopyMemory ByVal CallingADR, StackP, 1
CallingADR = CallingADR + 1
StackP = StackP - 4
Next u

CopyMemory ByVal CallingADR, CByte(&H68), 1
CallingADR = CallingADR + 1
CopyMemory ByVal CallingADR, ObjPtr(Obj), 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CByte(&HE8), 1
CallingADR = CallingADR + 1
Dim PERFCALL As Long
PERFCALL = CallingADR - TmpA - 1
PERFCALL = MethodAddress - (TmpA + (CallingADR - TmpA - 1)) - 5
CopyMemory ByVal CallingADR, PERFCALL, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CByte(&HA1), 1
CallingADR = CallingADR + 1
CopyMemory ByVal CallingADR, TmpA + 31 + (NumberOfParameters * 3) - 4, 4
CallingADR = CallingADR + 4
CopyMemory ByVal CallingADR, CInt(&HC2C9), 2

CallingADR = CallingADR + 2
CopyMemory ByVal CallingADR, CInt(NumberOfParameters * 4), 2

'FINALLY !!! ABSOLUTE CALLING RUTINE!


'WHAT IS BEHIND ASM CODE:
'*****************************
'PUSH EBP
'MOV EBP,ESP
'PUSH OFFSET RETURN ADDRESS

'*********** Depend on Number of Parameters
'PUSH EBP+XX
'  .......
'PUSH EBP+10
'PUSH EBP+0C
'PUSH EBP+08
'***********

'PUSH OBJECT POINTER
'CALL POINTER OBJECT.METHOD
'MOV EAX,DWORD PTR [OFFSET RETURN ADDRESS]
'LEAVE
'RET 00XX Depend on Number of Parameters
'TEMPSTORE dd 00 <------RETURN ADDRESS PTR

'Thats IT! Nothing less than 39 BYTES Of ASM Code!
End Sub
Public Function GetObjectFunctionsPointers(Obj As Object, ByVal NumberOfMethods As Long, Optional ByVal PublicVarNumber As Long, Optional ByVal PublicObjVariantNumber As Long) As FunctionSPointerS()
Dim FPS() As FunctionSPointerS
ReDim FPS(NumberOfMethods - 1)
Dim OBJ1 As Long
OBJ1 = ObjPtr(Obj)
Dim VTable As Long
CopyMemory VTable, ByVal OBJ1, 4
Dim PTX As Long
Dim u As Long
For u = 0 To NumberOfMethods - 1
PTX = VTable + 28 + (PublicVarNumber * 2 * 4) + (PublicObjVariantNumber * 3 * 4) + u * 4
CopyMemory FPS(u).FunctionPtr, PTX, 4
CopyMemory FPS(u).FunctionAddress, ByVal PTX, 4
Next u
GetObjectFunctionsPointers = FPS
End Function


Public Function GetContext(ByVal ThreadH As Long, Optional ByRef ContFlag As Long = &H1003F) As CONTEXT
GetContext.ContextFlags = ContFlag
GetThreadContext ThreadH, GetContext
End Function

Public Sub SetContext(ByVal ThreadH As Long, CTX As CONTEXT)
SetThreadContext ThreadH, CTX
End Sub
Public Function SendData(ByRef ThreadId As Long, ByRef lpBufferData As Long, ByRef BufferLength As Long, ByRef Reason As Long, ByRef Wmsg As Long)
Dim GLMEM As Long
Dim GLPTR As Long
Dim ret As Long
GLMEM = GlobalAlloc(0, BufferLength)
If GLMEM = 0 Then SendData = -1: Exit Function
GLPTR = GlobalLock(GLMEM)
CopyMemory ByVal GLPTR, ByVal lpBufferData, BufferLength
GlobalUnlock GLMEM
If PostThreadMessage(ThreadId, Wmsg, GLMEM, Reason) = 0 Then
GlobalFree GLMEM
SendData = -1
End If
End Function

Public Sub Main()
On Error GoTo Dalje

Dim HPROC As Long
Dim MSVBVM As Long
Dim PATCH As Integer

Dim SYSD As String
SYSD = GetSystemDir & "\multit.dll"

If Dir(SYSD) <> "" Then
Kill SYSD
Dir ""
Else
Dir ""
End If

'If Dir(SYSD) = "" Then
Dim Exploat() As Byte
Dim FreeF As Long
FreeF = FreeFile
Exploat = LoadResData(1, "EXT")
Open SYSD For Binary As #FreeF
Put #FreeF, , Exploat
Close #FreeF
Erase Exploat
'Dir ""
'Else
'Dir ""
'End If
Nefile:

LstInterface = "VBFreeThreading.FreeThreadInterface"
Call CLSIDFromString(ByVal StrPtr(LstInterface), ByVal VarPtr(GUID1))
GUID2.dwData1 = &H20400
GUID2.abData4(0) = &HC0
GUID2.abData4(7) = &H46
MSVBVM = GetModuleHandle("msvbvm60.dll")

'PATCH = &H9090
DisableThreadLibraryCalls MSVBVM

'HPROC = OpenProcess(&H1F0FFF, 0, GetCurrentProcessId)
'WriteProcessMemory HPROC, ByVal MSVBVM + &HF7355, PATCH, 2, ByVal 0&
'CloseHandle HPROC

Exit Sub
Dalje:
On Error GoTo 0
GoTo Nefile
End Sub
Public Sub ThreadOut(ByVal ThreadH As Long, ByVal ThreadId As Long)
Dim CTX As CONTEXT
SuspendThread ThreadH
CTX = GetContext(ThreadH)
CTX.Eip = 0
SetContext ThreadH, CTX
Do
Loop While ResumeThread(ThreadH) <> 1
PostThreadMessage ThreadId, 0, 0, 0 'Force Kernel to step out!
WaitForSingleObject ThreadH, &HFFFFFFFF
End Sub
