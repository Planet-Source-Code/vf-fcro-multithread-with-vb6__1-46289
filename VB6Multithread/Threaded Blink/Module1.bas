Attribute VB_Name = "Module1"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As Any, ByVal bErase As Long) As Long


Public Sub RefreshWindow(ByVal hwnd As Long)
InvalidateRect hwnd, ByVal 0&, 0
End Sub
