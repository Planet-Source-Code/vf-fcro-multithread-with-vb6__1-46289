VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VBFreeThreading.FreeWaitThreadInterface
Dim THREADER As New WaitTypeThreader
Dim TID As Long
Dim TH As Long

Dim HDLG As Long 'Dialog Box
Dim TBOX As Long


Sub StartThread()
TH = THREADER.StartThread(Me, False, 0, TID, THREAD_PRIORITY_NORMAL)
End Sub

Sub EndThread()
THREADER.ForceExitThread
End Sub

Sub SendDataThread(ByVal Data As String, ByVal AsyncS As Boolean)
Dim SData() As Byte
SData = StrConv(Data, vbFromUnicode)
If AsyncS = True Then
THREADER.AsyncSendDataToThread 0, VarPtr(SData(0)), Len(Data)
Else
THREADER.SyncSendDataToThread 0, VarPtr(SData(0)), Len(Data)
End If
End Sub

Sub GetDataThread()
THREADER.CallThread 0, 0
End Sub

Private Sub SetInBox(ByVal lpBufferData As Long, ByVal BufferLength As Long)
Dim S As String
S = Space(BufferLength)
CopyMemory ByVal S, ByVal lpBufferData, BufferLength
SetWindowText TBOX, S
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
SetInBox lpBufferData, BufferLength
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)
Dim TLen As Long
TLen = GetWindowTextLength(TBOX)
Dim S As String
S = Space(TLen)
GetWindowText TBOX, S, TLen + 1
Form1.Text1 = S
End Sub

Private Function FreeWaitThreadInterface_ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long

End Function

Private Sub FreeWaitThreadInterface_ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
DestroyWindow TBOX
DestroyWindow HDLG
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
DestroyWindow TBOX
DestroyWindow HDLG
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
Dim EXSTY As Long
Dim STY As Long
STY = WindowsStyles.WS_DLGFRAME Or WindowsStyles.ws_VISIBLE Or _
WindowsStyles.WS_BORDER Or WindowsStyles.WS_SYSMENU Or WindowsStyles.WS_MINIMIZEBOX Or _
WindowsStyles.WS_POPUP

EXSTY = WindowsExStyles.WS_EX_DLGMODALFRAME Or WindowsExStyles.WS_EX_WINDOWEDGE

HDLG = CreateWindowEx(EXSTY, "#32768", "", STY, 0, 0, 206, 224, Form1.hwnd, 0, 0, ByVal 0&)
SetWindowText HDLG, "Thread Id:" & GetCurrentThreadId

STY = WindowsStyles.ws_VISIBLE Or _
WindowsStyles.WS_CHILD Or EditStyles.ES_MULTILINE Or EditStyles.ES_WANTRETURN Or _
WindowsStyles.ws_hscroll Or WindowsStyles.ws_VSCROLL

EXSTY = WindowsExStyles.WS_EX_CLIENTEDGE

TBOX = CreateWindowEx(EXSTY, "Edit", "", STY, 0, 0, 200, 199, HDLG, 0, 0, ByVal 0&)


End Sub

Private Sub FreeWaitThreadInterface_ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)
SetInBox lpBufferData, BufferLength
End Sub
