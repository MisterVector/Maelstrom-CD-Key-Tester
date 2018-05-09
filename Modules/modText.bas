Attribute VB_Name = "modText"
Public Sub AddChat(ParamArray saElements() As Variant)
  Dim arrTmp() As String
  
  With frmMain.rtbChat
    .SelStart = Len(.text)
    .SelLength = 0
    .SelColor = vbWhite
    .SelText = "[" & Time() & "] "
    
    For i = 0 To UBound(saElements) Step 2
      .SelStart = Len(.text)
      .SelLength = 0
      .SelColor = saElements(i)
      
      .SelText = saElements(i + 1) & IIf(i + 1 = UBound(saElements), vbNewLine, vbNullString)
    Next i
  End With
  
  With frmMain.rtbChat
    If (UBound(Split(.text, vbNewLine)) + 1 >= 100) Then
      .text = vbNullString
    End If
  End With
End Sub

Public Sub AddChatB(ParamArray saElements() As Variant)
  Dim arrTmp() As String
  
  With frmMain.rtbChat
    .SelStart = Len(.text)
    .SelLength = 0
    .SelColor = vbWhite
    .SelText = "[" & Time() & "] "
    
    For i = 0 To UBound(saElements) Step 2
      .SelStart = Len(.text)
      .SelLength = 0
      .SelColor = saElements(i)
      
      .SelText = saElements(i + 1) & IIf(i + 1 = UBound(saElements), vbNewLine, vbNullString)
    Next i
  End With
End Sub
