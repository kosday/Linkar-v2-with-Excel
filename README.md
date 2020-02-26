# Linkar with Excel using linkaclientCOM.dll library

This demo shows how a persistent client works with an Excel frame document and shows data in a spreadsheet.

· Window: 
This window collects necessary data for login

· Result: 
It runs several operations and prints all the data in the spreadsheet

 

Register COM Library:

To use this COM library it's necessary  to register it. You can do that using RegAsm.exe tool whose location depends on the machine 
where it will run from CMD. You must run CMD as Administrator to avoid mistakes in the register.

For instance:

32-bit OS:
C:\>C:\Windows\Microsoft.NET\Framework\v4.0.30319\regasm "C:\linkar\Clients\NET_Framework\x86\LinkarClientCOM.DLL" /codebase /tlb

64-bit OS:
C:\>C:\Windows\Microsoft.NET\Framework64\v4.0.30319\regasm "C:\linkar\Clients\NET_Framework\x64\LinkarClientCOM.DLL" /codebase /tlb
 

NOTICE!

- You must run CMD as Administrator to avoid mistakes in the register

- You must have installed the 4.5 (v4.0.30319) Framework and use Regasm from that FrameWork

- There are two RegAsm versions, 32 and 64 bits (FrameWork or FrameWork64), you will have to use the same as the Linkar one you want to register, otherwise the command will return an error.

