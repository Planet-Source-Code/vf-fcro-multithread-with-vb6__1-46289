VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeThreadInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Sub ThreadStart(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub

Public Sub ThreadError(ByVal ThreadId As Long, ByVal ThreadHandle As Long)
End Sub
