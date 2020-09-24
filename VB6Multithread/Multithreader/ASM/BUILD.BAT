@echo off
\masm32\bin\ml /c /coff /Cp MULTITHREAD.asm
\masm32\bin\link /DLL /DEF:MULTITHREAD.def /SUBSYSTEM:WINDOWS /LIBPATH:\masm32\lib MULTITHREAD.obj
pause