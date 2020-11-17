Attribute VB_Name = "modGeneralAPI"
Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function EmptyWorkingSet Lib "psapi.dll" (ByVal hProcess As Long) As Long
Private Declare Function SetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, ByVal dwMinimumWorkingSetSize As Long _
                                                                , ByVal dwMaximumWorkingSetSize As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal numbytes As Long)
Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpszOp As String, ByVal lpszFile As String, ByVal lpszParams As String, ByVal LpszDir As String, ByVal FsShowCmd As Long) As Long

Public Type FILETIME
    dwLowDateTime       As Long
    dwHighDateTime      As Long
End Type

Public Type SYSTEMTIME
    wYear               As Integer
    wMonth              As Integer
    wDayOfWeek          As Integer
    wDay                As Integer
    wHour               As Integer
    wMinute             As Integer
    wSecond             As Integer
    wMilliseconds       As Integer
End Type
Public tpLocal As SYSTEMTIME
Public tpSystem As SYSTEMTIME

Public Sub FreeMemory()
    ews_memory
    spw_memory
End Sub

Public Function ews_memory() As Long:   ews_memory = EmptyWorkingSet(GetCurrentProcess):                  End Function
Public Function spw_memory() As Long:   spw_memory = SetProcessWorkingSetSize(GetCurrentProcess, -1, -1): End Function

Public Function GetFTTime(FT As FILETIME, Optional Shorten As Boolean = False, Optional localTime As Boolean = True) As String
    Dim LocalFT As FILETIME
    Dim SysTime As SYSTEMTIME
    Dim SetHour As String
    Dim AP      As String

    If (localTime) Then
        FileTimeToLocalFileTime FT, LocalFT
        FileTimeToSystemTime LocalFT, SysTime
    Else
        FileTimeToSystemTime FT, SysTime
    End If
  
    If (SysTime.wHour = 0) Then
        AP = "AM"
        SetHour = "12"
    ElseIf (SysTime.wHour < 12) Then
        AP = "AM"
        SetHour = Trim$(str$(SysTime.wHour))
    ElseIf (SysTime.wHour = 12) Then
        AP = "PM"
        SetHour = "12"
    Else
        AP = "PM"
        SetHour = Trim$(str$(SysTime.wHour))
    End If
  
    SysTime.wDayOfWeek = SysTime.wDayOfWeek + 1
  
    If (Shorten) Then
        GetFTTime = Format$(SysTime.wMonth, "00") & "/" & Format$(SysTime.wDay, "00") & "/" & Right$(SysTime.wYear, 2) & " " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    Else
        GetFTTime = ConvertShortToLong(WeekdayName(SysTime.wDayOfWeek, True)) & ", " & MonthName(SysTime.wMonth, True) & " " & SysTime.wDay & ", " & SysTime.wYear & " at " & SetHour & ":" & Format$(SysTime.wMinute, "00") & ":" & Format$(SysTime.wSecond, "00") & " " & AP
    End If
End Function

Private Function ConvertShortToLong(Day As String)
    Select Case Day
        Case "Mon": ConvertShortToLong = "Monday"
        Case "Tue": ConvertShortToLong = "Tuesday"
        Case "Wed": ConvertShortToLong = "Wednesday"
        Case "Thu": ConvertShortToLong = "Thursday"
        Case "Fri": ConvertShortToLong = "Friday"
        Case "Sat": ConvertShortToLong = "Saturday"
        Case "Sun": ConvertShortToLong = "Sunday"
    End Select
End Function

