Attribute VB_Name = "Internal"
Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Declare Function VirtualQuery Lib "kernel32" (lpAddress As Any, lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Declare Function FindFromPTR Lib "multit" (ByVal BeginSearchAddress As Long, ByVal LengthOfSearch As Long, SearchPattern As Any, ByVal PatternLength As Long) As Long
Declare Sub EbLoadRuntime Lib "msvbvm60" (ByVal BaseAddress As Long, ByVal BaseInit As Long)


Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long

End Type
Public Function FindBaseInit() As Long
Dim PTRN(1) As Long
Dim ret As Long
Dim MHDL As Long
Dim MEMB As MEMORY_BASIC_INFORMATION

PTRN(0) = &H21354256
PTRN(1) = &H2A231C
'MHDL = GetModuleHandle(vbNullString)
MHDL = &H400000
Do
VirtualQuery ByVal MHDL, MEMB, Len(MEMB)
If MEMB.AllocationBase = 0 Then Exit Function
ret = FindFromPTR(MEMB.BaseAddress, MEMB.RegionSize, PTRN(0), 8)
If ret <> -1 Then Exit Do
MHDL = MHDL + MEMB.RegionSize
Loop
FindBaseInit = MHDL + ret - 1
End Function
