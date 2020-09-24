Attribute VB_Name = "Module1"
Declare Function GetCurrentThreadId Lib "kernel32" () As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function DefDlgProc Lib "user32" Alias "DefDlgProcA" (ByVal hDlg As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal ByteLen As Long)

Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, wParam As Any, lParam As Any) As Long
Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hinstance As Long, lpParam As Any) As Long
Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Enum EditStyles
ES_AUTOHSCROLL = &H80&
ES_AUTOVSCROLL = &H40&
ES_CENTER = &H1&
ES_CONTINUOUS = (&H80000000)
ES_DISABLENOSCROLL = &H2000
ES_DISPLAY_REQUIRED = (&H2)
ES_EX_NOCALLOLEINIT = &H1000000
ES_LEFT = &H0&
ES_LOWERCASE = &H10&
ES_MULTILINE = &H4&
ES_NOHIDESEL = &H100&
ES_NOIME = &H80000
ES_NOOLEDRAGDROP = &H8
ES_NUMBER = &H2000&
ES_OEMCONVERT = &H400&
ES_PASSWORD = &H20&
ES_READONLY = &H800&
ES_RIGHT = &H2&
ES_SAVESEL = &H8000
ES_SELECTIONBAR = &H1000000
ES_SELFIME = &H40000
ES_SUNKEN = &H4000
ES_UPPERCASE = &H8&
ES_VERTICAL = &H400000
ES_WANTRETURN = &H1000&
End Enum



Public Enum WindowsStyles
 WS_MAXIMIZEBOX = &H10000
 WS_MINIMIZEBOX = &H20000
 WS_THICKFRAME = &H40000
 WS_SYSMENU = &H80000
 ws_hscroll = &H100000
 ws_VSCROLL = &H200000
 WS_DLGFRAME = &H400000
 WS_BORDER = &H800000
 WS_MAXIMIZE = &H1000000
' WS_CLIPCHILDREN = &H2000000
' WS_CLIPSIBLINGS = &H4000000
 WS_DISABLED = &H8000000
 ws_VISIBLE = &H10000000
 WS_MINIMIZE = &H20000000
 WS_CHILD = &H40000000
 WS_POPUP = &H80000000
End Enum

Public Enum WindowsExStyles
WS_EX_DLGMODALFRAME = &H1&
WS_EX_NOPARENTNOTIFY = &H4&
WS_EX_TOPMOST = &H8&
WS_EX_ACCEPTFILES = &H10&
WS_EX_TRANSPARENT = &H20&
WS_EX_MDICHILD = &H40&
WS_EX_TOOLWINDOW = &H80&
WS_EX_WINDOWEDGE = &H100&
WS_EX_CLIENTEDGE = &H200&
WS_EX_CONTEXTHELP = &H400&
WS_EX_RIGHT = &H1000&
WS_EX_RTLREADING = &H2000&
WS_EX_LEFTSCROLLBAR = &H4000&
WS_EX_CONTROLPARENT = &H10000
WS_EX_STATICEDGE = &H20000
WS_EX_APPWINDOW = &H40000
WS_EX_LAYERED = &H80000
End Enum

Public SOBJ As Winsock
Public OUTGOINGTEXTHWND As Long

Public DIALOGP As Long
Public Sub GetHILO(ByVal Value As Long, ByRef HI As Long, ByRef LO As Long)
CopyMemory LO, Value, 2
CopyMemory HI, ByVal (VarPtr(Value) + 2), 2
End Sub

Public Function DialogProc(ByVal hwnd As Long, ByVal umsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Select Case umsg


'IT's SHAME...
'WE ARE NOT ABLE TO USE FORMS,SO WE WILL USE PURE DIALOGS WITH 2 EDITBOX and 1 BUTTON!
Case &H111  'WM_COMMAND
Dim CTID As Long
Dim NMSG As Long
GetHILO wParam, NMSG, CTID
If lParam <> 0 And NMSG = 0 Then 'BTN_CLICK

Dim OUTTEXT As String
Dim TextLen As Long
TextLen = GetWindowTextLength(OUTGOINGTEXTHWND)
OUTTEXT = Space(TextLen)
GetWindowText OUTGOINGTEXTHWND, OUTTEXT, TextLen + 1
SOBJ.SendData OUTTEXT

DoEvents

End If


End Select
DialogProc = DefDlgProc(hwnd, umsg, wParam, lParam)
End Function

