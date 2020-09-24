VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FreeThreaderStandard"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Function StartThreadOnAddress(ByVal StartAddress As Long, ByVal SehProcedureAddress As Long, ByVal StackSize As Long, ByRef ThreadId As Long) As Long
StartThreadOnAddress = CreateMT(StartAddress, StackSize, ThreadId, GUID1, GUID2, "X" & StartAddress, SehProcedureAddress)
End Function
