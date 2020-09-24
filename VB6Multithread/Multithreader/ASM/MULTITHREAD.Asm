    .586                      ; force 32 bit code
      .model flat, stdcall      ; memory model & calling convention
      option casemap :none      ; case sensitive

include MULTITHREAD.INC

ExitProc PROTO 
EnterProc PROTO :DWORD

.data



Ct dd 0
Guids dd 0,0,0,0

MSVB db "msvbvm60.dll",0
LDLIB db "EbLoadRunTime",0
CCLIB db "EbCreateContext",0
SCLIB db "EbSetContextWorkerThread",0
LDPROC dd 0,0,0

.code





;Param1=AddressToCall,Param2=StackSize,Param3=
CreateMT proc Param1:DWORD,Param2:DWORD,Param3:DWORD,Param4:DWORD,Param5:DWORD,Param6:DWORD,Param7:DWORD

assume FS: nothing

sub esp,4
push esi
push edi
push ecx
mov eax,DWORD PTR fs:[018h]
add eax,0e14h
mov Ct,eax
lea esi,[ebp+014h]
mov edi,OFFSET Guids
cld
mov ecx,4
rep movsd

invoke CreateEvent,NULL,0,0,Param6
mov DWORD PTR [ebp-4],eax
invoke CreateThread,0,Param2,EnterProc,Param1,0,Param3
push eax
invoke WaitForSingleObject,DWORD PTR [ebp-4],0FFFFFFFFh
invoke CloseHandle,DWORD PTR [EBP-4]
pop eax
pop ecx
pop edi
pop esi
ret

CreateMT endp


EnterProc proc Param1:DWORD
assume FS: nothing	

sub esp,4

push Guids+0ch
push FS:[0]
mov FS:[0],esp


mov edi,DWORD PTR fs:[018h]
add edi,0e14h
mov esi,Ct
cld
mov ecx,1
rep movsd


invoke CoInitialize,0
invoke OleInitialize,0

lea eax,[ebp-4]
;push eax
invoke CoCreateInstance,Guids,0,017h,Guids+4,eax


invoke OpenEvent,01f0003h,0,Guids+8
invoke SetEvent,eax

call Param1


invoke CoUninitialize
invoke OleUninitialize
invoke ExitThread,0



	ret

EnterProc endp

EnableEvents proc

invoke TlsGetValue,3
mov eax,DWORD PTR [eax+028h] ;GET CX
lea ebx,[eax+04ch]

push DWORD PTR [ebx]
push ebx
call LDPROC+4 ;EbCreateContext
call LDPROC+8 ;EbSetContextWorkerThread


;INIT RUNTIME!
mov eax,040003ch
mov eax,DWORD PTR [eax]
lea eax,[eax+0400000h]
mov eax,DWORD PTR [eax+028h]
lea eax,[eax+0400001h]
mov eax,DWORD PTR [eax]
add eax,030h
push DWORD PTR[eax]
push 0400000h
call LDPROC

ret
EnableEvents endp




DllEntry proc hInstance:DWORD, reason:DWORD, reserved1:DWORD
sub esp,4
invoke GetModuleHandle,addr MSVB
mov DWORD PTR [ebp-4],eax
invoke GetProcAddress,eax,addr LDLIB
mov LDPROC,eax
invoke GetProcAddress,DWORD PTR [ebp-4],addr CCLIB
mov LDPROC+4,eax
invoke GetProcAddress,DWORD PTR [ebp-4],addr SCLIB
mov LDPROC+8,eax
mov  eax,1
ret
DllEntry Endp

end DllEntry