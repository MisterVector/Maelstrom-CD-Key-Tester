Attribute VB_Name = "Hash_Func"
'Do not modify this file!
'This is part of BNHash functionality and could possibly be updated. If you don't want to lose anywork
'then it's advised that you create your own module.

Option Explicit

Public Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long
Private Declare Function SetProcessWorkingSetSize Lib "Kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long _
                                                                , ByVal dwMaximumWorkingSetSize As Long) As Long
Public Sub FreeMemory()
ews_memory
spw_memory
End Sub

Public Function ews_memory() As Long:   ews_memory = EmptyWorkingSet(GetCurrentProcess):                  End Function
Public Function spw_memory() As Long:   spw_memory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1): End Function

Public Function P_split(sIP As String) As String
  On Error Resume Next

  Dim splt() As String, i As Byte

  splt = Split(sIP, ".")
  For i = 0 To UBound(splt)
    P_split = P_split & Chr$(CStr(splt(i)))
  Next i
End Function
