Attribute VB_Name = "modBenchmark"
Public Const SECONDS_PER_HOUR As Long = 3600
Public Const SECONDS_PER_MINUTE As Long = 60

Public curSeconds As Long

Public Function returnProperTimeString(ByVal Seconds As Long)
    Dim h As String, m As String, s As String
  
    If (Seconds >= SECONDS_PER_HOUR) Then
        h = (Seconds \ SECONDS_PER_HOUR)
        Seconds = Seconds - (SECONDS_PER_HOUR * h)
    Else
        h = "0"
    End If
  
    If (Seconds >= SECONDS_PER_MINUTE) Then
        m = (Seconds \ SECONDS_PER_MINUTE)
        Seconds = Seconds - (SECONDS_PER_MINUTE * m)
    Else
        m = "00"
    End If
  
    s = IIf(Seconds > 0, Seconds, "00")
  
    If (Len(s) = 1) Then s = "0" & s
    If (Len(m) = 1) Then m = "0" & m
  
    returnProperTimeString = h & ":" & m & ":" & s
End Function

Public Sub resetBenchmark()
    curSeconds = 0

    frmMain.lblControl(TIME_ELAPSED).Caption = "0:00:00"
    frmMain.lblControl(KEYS_PER_SECOND).Caption = "0.000"
End Sub

