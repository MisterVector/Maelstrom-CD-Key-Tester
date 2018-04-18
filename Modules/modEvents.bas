Attribute VB_Name = "modEvents"
Public Sub perfectKeyEvaluation(Index As Integer)
  Dim keepOriginalKey As Boolean, showRegular As Boolean, showExpansion As Boolean
  Dim clearSavedKeyState As Boolean

  With BNETData(Index)
    .numTested = .numTested + 1
  
    If canTestExpansion(.product) Then keepOriginalKey = True
        
    If .isExpansion Then
      AddChat TEXT_PERFECT, "Socket #" & Index & ": Found a perfect expansion key!"
      
      exportKeyToFile .cdKeyExp, .productExpansion, "Perfect"
      removeCDKeyByIndex .cdKeyExpIndex, .productExpansion
      .cdKeyExp = vbNullString
      testedExpKeys = testedExpKeys + 1
      showExpansion = True
      
      If keepOriginalKey Then
        If .TestedEXP = config.expansionTestsPerRegularKey Then
          AddChat vbYellow, "Socket #" & Index & ": Expansion tests per regular key reached. Rotating key..."
          keepOriginalKey = False
        Else
          .TestedEXP = .TestedEXP + 1
        End If
      End If
    Else
      If keepOriginalKey Then
        AddChat TEXT_PERFECT, "Socket #" & Index & ": Perfect key will be used to test expansion."
        .savedKeyState = "Perfect"
      Else
        AddChat TEXT_PERFECT, "Socket #" & Index & ": Found a perfect key!"
      End If
    End If

    If Not keepOriginalKey Then
      exportKeyToFile .cdKey, .productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, "Perfect")
      removeCDKeyByIndex .cdKeyIndex, .productRegular
      
      .cdKey = vbNullString
      .TestedEXP = 0
      
      testedNonExpKeys = testedNonExpKeys + 1
      showRegular = True
      
      clearSavedKeyState = True
    End If
    
    Dim lblKeyStateIndex As Integer
    
    If showRegular Then
      lblKeyStateIndex = getLabelByKeyState(.productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, "perfect"))
      frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
      postKeysTested .productRegular
    End If
    
    If showExpansion Then
      lblKeyStateIndex = getLabelByKeyState(.productExpansion, "perfect")
      frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
      postKeysTested .productExpansion
    End If
    
    If clearSavedKeyState Then
      .savedKeyState = vbNullString
    End If
  End With
End Sub

Public Sub voidedMutedOrJailedKeyEvaluation(Index As Integer, ByVal isMuted As Boolean, ByVal isVoided As Boolean)
  Dim keyState As String, keepOriginalKey As Boolean, showRegular As Boolean, showExpansion As Boolean
  Dim color As Long, clearSavedKeyState As Boolean
  
  If isMuted And Not isVoided Then
    keyState = "Muted"
    color = TEXT_MUTED
  End If
    
  If Not isMuted And isVoided Then
    keyState = "Voided"
    color = TEXT_VOIDED
  End If
  
  If isMuted And isVoided Then
    keyState = "Jailed"
    color = TEXT_JAILED
  End If
  
  With BNETData(Index)
    .numTested = .numTested + 1
  
    If canTestExpansion(.product) Then keepOriginalKey = True
      
    If .isExpansion Then
      AddChat color, "Socket #" & Index & ": Expansion key is " & LCase(keyState) & "."
    
      exportKeyToFile .cdKeyExp, .productExpansion, keyState
      removeCDKeyByIndex .cdKeyExpIndex, .productExpansion
      testedExpKeys = testedExpKeys + 1
      .cdKeyExp = vbNullString
      showExpansion = True
      
      If keepOriginalKey Then
        If .TestedEXP = config.expansionTestsPerRegularKey Then
          AddChat vbYellow, "Socket #" & Index & ": Expansion tests per regular key reached. Rotating key..."
          keepOriginalKey = False
        Else
          .TestedEXP = .TestedEXP + 1
        End If
      End If
    Else
      If keepOriginalKey Then
        AddChat color, "Socket #" & Index & ": " & keyState & " key will be used to test expansion."
        .savedKeyState = keyState
      Else
        AddChat color, "Socket #" & Index & ": Key is " & LCase(keyState) & "."
      End If
    End If
    
    If Not keepOriginalKey Then
      exportKeyToFile .cdKey, .productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, keyState)
      removeCDKeyByIndex .cdKeyIndex, .productRegular
    
      .TestedEXP = 0
      .cdKey = vbNullString
      
      testedNonExpKeys = testedNonExpKeys + 1
      showRegular = True
      
      clearSavedKeyState = True
    End If

    Dim lblKeyStateIndex As Integer

    If showRegular Then
      lblKeyStateIndex = getLabelByKeyState(.productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, keyState))
      frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
      postKeysTested .productRegular
    End If

    If showExpansion Then
      lblKeyStateIndex = getLabelByKeyState(.productExpansion, keyState)
      frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
      postKeysTested .productExpansion
    End If
    
    If clearSavedKeyState Then
      .savedKeyState = vbNullString
    End If
  End With
End Sub

Public Sub handleOtherKeys(Index As Integer, ID As Long, ByVal inUse As String)
  Dim keyState As String, color As Long, lblKeyStateIndex As Integer
  Dim keepOriginalKey As Boolean, showRegular As Boolean, showExpansion As Boolean
  Dim clearSavedKeyState As Boolean
  
  With BNETData(Index)
    Select Case ID
      Case &H200, &H202, &H203, &H210, &H212
      
      ' 0x200 = invalid key
      ' 0x202 = banned key
      ' 0x203 = other product
      ' 0x210 = expansion invalid
      ' 0x212 = expansion banned
      
        .numTested = .numTested + 1
      
        Select Case ID
          Case &H200, &H210: keyState = "Invalid": color = TEXT_INVALID
          Case &H202, &H212: keyState = "Banned": color = TEXT_BANNED
          Case &H203, &H213: keyState = "Other": color = TEXT_OTHER
        End Select

        If .isExpansion Then
          AddChat color, "Socket #" & Index & ": Expansion key is " & IIf(keyState = "Other", "for other product", LCase(keyState)) & "."
          exportKeyToFile .cdKeyExp, .productExpansion, keyState
          removeCDKeyByIndex .cdKeyExpIndex, .productExpansion
          .cdKeyExp = vbNullString
          testedExpKeys = testedExpKeys + 1
          showExpansion = True
          
          If canTestExpansion(.product) Then keepOriginalKey = True
          
          If keepOriginalKey Then
            If .TestedEXP = config.expansionTestsPerRegularKey Then
              AddChat vbYellow, "Socket #" & Index & ": Expansion tests per regular key reached. Rotating key..."
              keepOriginalKey = False
            Else
              .TestedEXP = .TestedEXP + 1
            End If
          End If
        Else
          AddChat color, "Socket #" & Index & ": Key is " & IIf(keyState = "other", "for other product", LCase(keyState)) & "."
        End If

        If Not keepOriginalKey Then
          exportKeyToFile .cdKey, .productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, keyState)
          removeCDKeyByIndex .cdKeyIndex, .productRegular
          
          .TestedEXP = 0
          .cdKey = vbNullString
          
          testedNonExpKeys = testedNonExpKeys + 1
          showRegular = True
          
          clearSavedKeyState = True
        End If
      
        If showRegular Then
          lblKeyStateIndex = getLabelByKeyState(.productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, keyState))
          frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
          postKeysTested .productRegular
        End If
        
        If showExpansion Then
          lblKeyStateIndex = getLabelByKeyState(.productExpansion, keyState)
          frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
          postKeysTested .productExpansion
        End If
        
        If clearSavedKeyState Then
          .savedKeyState = vbNullString
        End If
      Case &H201, &H211  'cdkey in use, expansion key in use
        .numTested = .numTested + 1
        
        If .isExpansion Then
          AddChat TEXT_IN_USE, "Socket #" & Index & ": Expansion key is in use by " & inUse & "."
  
          exportKeyToFile .cdKeyExp, .productExpansion, "In Use", inUse
          removeCDKeyByIndex .cdKeyExpIndex, .productExpansion
          testedExpKeys = testedExpKeys + 1
          showExpansion = True
          
          If canTestExpansion(.product) Then keepOriginalKey = True
          
          If keepOriginalKey Then
            If .TestedEXP = config.expansionTestsPerRegularKey Then
              AddChat vbYellow, "Socket #" & Index & ": Expansion tests per regular key reached. Rotating key..."
              keepOriginalKey = False
            Else
              .TestedEXP = .TestedEXP + 1
            End If
          End If
          
          inUse = vbNullString
        Else
          AddChat TEXT_IN_USE, "Socket #" & Index & ": Key is in use by " & inUse & "."
        End If
        
        If Not keepOriginalKey Then
          exportKeyToFile .cdKey, .productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, "In Use"), inUse
          removeCDKeyByIndex .cdKeyIndex, .productRegular
          
          .TestedEXP = 0
          .cdKey = vbNullString
          
          testedNonExpKeys = testedNonExpKeys + 1
          showRegular = True
          
          clearSavedKeyState = True
        End If
        
        If showRegular Then
          lblKeyStateIndex = getLabelByKeyState(.productRegular, IIf(.savedKeyState <> vbNullString, .savedKeyState, "inuse"))
          frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
          postKeysTested .productRegular
        End If
        
        If showExpansion Then
          lblKeyStateIndex = getLabelByKeyState(.productExpansion, "inuse")
          frmMain.lblControl(lblKeyStateIndex).Caption = frmMain.lblControl(lblKeyStateIndex).Caption + 1
          postKeysTested .productExpansion
        End If
        
        If clearSavedKeyState Then
          .savedKeyState = vbNullString
        End If
    End Select
  End With
End Sub

Public Sub connectSocket(ByVal Index As Integer)
  frmMain.sckBNCS(Index).Connect BNETData(Index).proxyIP, BNETData(Index).proxyPort
End Sub

Public Sub closeSocket(ByVal Index As Integer)
  frmMain.sckBNCS(Index).Close
End Sub

Public Function IsProxyPacket(Index As Integer, ByVal data As String) As Boolean
  Select Case Mid$(data, 1, 2)
    Case Chr(&H0) & Chr(&H5A): 'Accepted
                               frmMain.sckBNCS(Index).SendData Chr$(&H1)
                               Send0x50 Index
                               IsProxyPacket = True
                               Exit Function
    Case Chr(&H0) & Chr(&H5B): 'Denied
                               IsProxyPacket = True
    Case Chr(&H0) & Chr(&H5C): 'Rejected
                               Handle_Proxy = True
    Case Chr(&H0) & Chr(&H5D): 'Rejected
                               IsProxyPacket = True
  End Select

  If Not IsProxyPacket Then
    If InStr(data, " ") Then
      If Mid(data, 10, 3) = "200" Then
        frmMain.sckBNCS(Index).SendData Chr$(&H1)
        Send0x50 Index
        IsProxyPacket = True
      End If
    End If
  End If
End Function

