Attribute VB_Name = "modWindowFunctions"
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HTCAPTION = 2
Public Const WM_NCLBUTTONDOWN = &HA1

Public Sub moveEntireForm(fm As Form, Button As Integer)
    If (Button = 1) Then
        ReleaseCapture
        SendMessage fm.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
    End If
End Sub


