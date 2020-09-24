VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MTWSCK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements VBFreeThreading.FreeWaitThreadInterface
Dim THREADER As New VBFreeThreading.WaitTypeThreader

Dim hDlg As Long 'Dialog Box
Dim TBOX As Long
Dim TOUTBOX As Long

Dim TID As Long
Dim TH As Long
Dim WithEvents MTWINSOCK As Winsock 'WINSOCK inside the THREAD!
Attribute MTWINSOCK.VB_VarHelpID = -1




Public Sub StartThread()
TH = THREADER.StartThread(Me, True, 0, TID)
End Sub
Public Sub ExitThread()
THREADER.ForceExitThread
End Sub






Private Sub FreeWaitThreadInterface_ThreadOnAsyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)

End Sub

Private Sub FreeWaitThreadInterface_ThreadOnCall(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long)

End Sub

Private Function FreeWaitThreadInterface_ThreadOnCallback(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal Message As Long) As Long

End Function

Private Sub FreeWaitThreadInterface_ThreadOnExit(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
SetWindowLong hDlg, -4, DIALOGP
DestroyWindow TBOX
DestroyWindow hDlg

MTWINSOCK.Close
Set MTWINSOCK = Nothing
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnExitError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
SetWindowLong hDlg, -4, DIALOGP
DestroyWindow TBOX
DestroyWindow hDlg

MTWINSOCK.Close
Set MTWINSOCK = Nothing
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
'CREATE DIALOG + OTHER WINDOWS + MULTITHREADER OBJECT {WINSOCK}


Dim EXSTY As Long
Dim STY As Long
STY = WindowsStyles.WS_DLGFRAME Or WindowsStyles.ws_VISIBLE Or _
WindowsStyles.WS_BORDER Or WindowsStyles.WS_SYSMENU Or WindowsStyles.WS_MINIMIZEBOX Or _
WindowsStyles.WS_POPUP

EXSTY = WindowsExStyles.WS_EX_DLGMODALFRAME Or WindowsExStyles.WS_EX_WINDOWEDGE

hDlg = CreateWindowEx(EXSTY, "#32770", "Thread Id:" & GetCurrentThreadId, STY, 0, 0, 206, 280, FX.hwnd, 0, 0, ByVal 0&)
DIALOGP = SetWindowLong(hDlg, -4, AddressOf DialogProc)

STY = WindowsStyles.ws_VISIBLE Or _
WindowsStyles.WS_CHILD Or EditStyles.ES_MULTILINE Or EditStyles.ES_WANTRETURN Or _
WindowsStyles.ws_hscroll Or WindowsStyles.ws_VSCROLL Or EditStyles.ES_READONLY

EXSTY = WindowsExStyles.WS_EX_CLIENTEDGE
TBOX = CreateWindowEx(EXSTY, "Edit", "", STY, 0, 0, 200, 199, hDlg, 0, 0, ByVal 0&)

STY = WindowsStyles.ws_VISIBLE Or WindowsStyles.WS_CHILD
TOUTBOX = CreateWindowEx(EXSTY, "Edit", "", STY, 0, 201, 200, 20, hDlg, 0, 0, ByVal 0&)
OUTGOINGTEXTHWND = TOUTBOX

STY = WindowsStyles.WS_CHILD Or WindowsStyles.ws_VISIBLE
Call CreateWindowEx(EXSTY, "Button", "Send", STY, 63, 225, 70, 25, hDlg, 0, 0, ByVal 0&)


Set MTWINSOCK = New Winsock 'CREATE MULTITHREAD WINSOCK OBJECT [WITH EVENTS!!!]
MTWINSOCK.Connect "localhost", 9000

Set SOBJ = MTWINSOCK
End Sub

Private Sub FreeWaitThreadInterface_ThreadOnSyncDataTransfer(ByVal ThreadId As Long, ByVal ThreadHandle As Long, ByVal Reason As Long, ByVal lpBufferData As Long, ByVal BufferLength As Long)

End Sub




Private Sub MTWINSOCK_Connect()
MsgBox "CONNECTION BEETWEEN 2 WINSOCK FROM 2 DIFFERENT THREADS ESTABLISHED!", vbExclamation, "Success!"
End Sub

Private Sub MTWINSOCK_DataArrival(ByVal bytesTotal As Long)
Dim INCOMMINGT As String
Dim JOINTXT As String
Dim TextLen As Long
TextLen = GetWindowTextLength(TBOX)
JOINTXT = Space(TextLen)
GetWindowText TBOX, JOINTXT, TextLen + 1
MTWINSOCK.GetData INCOMMINGT
SetWindowText TBOX, ">" & INCOMMINGT & vbCrLf & JOINTXT
End Sub

