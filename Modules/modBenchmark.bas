Attribute VB_Name = "modBenchmark"
Public curSeconds As Long

Public Function returnProperTimeString(ByVal Seconds As Long)
  Dim h As String, m As String, s As String
  
  s = (Seconds Mod 60)
  m = ((Seconds \ 60) Mod 60)
  h = ((Seconds \ 24 \ 60) Mod 60)
  
  If Len(s) = 1 Then s = "0" & s
  If Len(m) = 1 Then m = "0" & m
  
  returnProperTimeString = h & ":" & m & ":" & s
End Function

Public Sub resetBenchmark()
  curSeconds = 0

  frmMain.lblControl(TIME_ELAPSED).Caption = "0:00:00"
  frmMain.lblControl(KEYS_PER_SECOND).Caption = "0.000"
End Sub

