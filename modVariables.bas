Attribute VB_Name = "modVariables"
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetProcessHeap Lib "kernel32" () As Long
Private Declare Function HeapAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function HeapReAlloc Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any, ByVal dwBytes As Long) As Long
Private Declare Function HeapSize Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Function HeapFree Lib "kernel32" (ByVal hHeap As Long, ByVal dwFlags As Long, lpMem As Any) As Long
Private Declare Sub CopyMemoryWrite Lib "kernel32" Alias "RtlMoveMemory" (ByVal Destination As Long, Source As Any, ByVal Length As Long)
Private Declare Sub CopyMemoryRead Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, ByVal Source As Long, ByVal Length As Long)

Public Function Var(vLength As Long) As Long
 Var = HeapAlloc(GetProcessHeap(), 0, vLength)
End Function

Public Function VarEx(nData As Long) As Long
Dim lRetVal As Long
 lRetVal = HeapAlloc(GetProcessHeap(), 0, LenB(nData))
  SetVar lRetVal, nData
 VarEx = lRetVal
End Function

Public Function GetVar(Address As Long, vLength As Long) As Long
 CopyMemoryRead GetVar, Address, vLength
End Function

'>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>
'>WARNING MAY CAUSE CRASH OR CAUSE SUDDEN STOP OF EXECUTION!<
'<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<
Public Function GetVarEx(Address As Long) As Long

Dim lRetVal As Long
Dim Tmp As Long

 lRetVal = HeapSize(GetProcessHeap(), 0, Address)
 
 If lRetVal <> 0 Then
  CopyMemoryRead Tmp, Address, lRetVal
  GetVarEx = Tmp
 End If
 
End Function

Public Function SetVar(Address As Long, nData As Long) As Long
 CopyMemoryWrite Address, nData, LenB(nData)
End Function

Public Function ResizeVar(Address As Long, vLength As Long) As Long
 ResizeVar = HeapReAlloc(GetProcessHeap(), 0, Address, vLength)
End Function

Public Function KillVar(Address As Long)
 HeapFree GetProcessHeap(), 0, Address
End Function
