Attribute VB_Name = "modKeyFunctions"
Public totalNonExpKeys As Long
Public totalExpKeys As Long
Public testedNonExpKeys As Long
Public testedExpKeys As Long

Public Type CDKeyType
    W2BN() As String
    W2BNIndex As Long
    W2BNTested As Long
    W2BNTotal As Long
  
    D2DV() As String
    D2DVIndex As Long
    D2DVTested As Long
    D2DVTotal As Long
  
    D2XP() As String
    D2XPIndex As Long
    D2XPTested As Long
    D2XPTotal As Long
End Type
Public CDKeys As CDKeyType

Public Type ParsedKeys
    dicKeys As New Dictionary
  
    w2bnCount As Long
    d2dvCount As Long
    d2xpCount As Long
  
    badFiles As Long
    duplicateKeys As Long
    invalidKeys As Long
    skippedLines As Long
End Type

Public Type DecodedKey
    Product As String
    successful As Boolean
End Type

Public Type FoundKey
    CDKey As String
    Product As String
    keyIndex As Long
End Type

Public Sub loadCDKeys()
    Dim arrDefaultKeyFiles() As Variant, pk As ParsedKeys
  
    arrDefaultKeyFiles = Array("W2BN.txt", "D2DV.txt", "D2XP.txt")
  
    totalNonExpKeys = 0
    testedNonExpKeys = 0
    totalExpKeys = 0
    testedExpKeys = 0
      
    ReDim CDKeys.W2BN(0)
    ReDim CDKeys.D2DV(0)
    ReDim CDKeys.D2XP(0)
  
    CDKeys.W2BNIndex = -1
    CDKeys.D2DVIndex = -1
    CDKeys.D2XPIndex = -1
  
    CDKeys.W2BNTested = 0
    CDKeys.D2DVTested = 0
    CDKeys.D2XPTested = 0
  
    CDKeys.W2BNTotal = 0
    CDKeys.D2DVTotal = 0
    CDKeys.D2XPTotal = 0
  
    For i = 26 To 52
        frmMain.lblControl(i).Caption = 0
    Next i
  
    For i = 94 To 96
        frmMain.lblControl(i).Caption = "0.0%"
    Next i
  
    If (Dir$(CDKEYS_FOLDER, 16) = vbNullString) Then
        MkDir CDKEYS_FOLDER
    Else
        Dim fso As New FileSystemObject
        Dim f As Folder
    
        Set f = fso.GetFolder(CDKEYS_FOLDER)
        loadKeysFromFiles f, pk
    End If
  
    If (pk.dicKeys.count > 0) Then
        Dim w2bnIdx As Long, d2dvIdx As Long, d2xpIdx As Long
    
        w2bnIdx = 0
        d2dvIdx = 0
        d2xpIdx = 0
    
        If (pk.w2bnCount > 0) Then
            ReDim CDKeys.W2BN(pk.w2bnCount - 1)
            CDKeys.W2BNIndex = 0
            CDKeys.W2BNTotal = pk.w2bnCount
    
            frmMain.lblControl(W2BNTotal).Caption = pk.w2bnCount
        End If
    
        If (pk.d2dvCount > 0) Then
            ReDim CDKeys.D2DV(pk.d2dvCount - 1)
            CDKeys.D2DVIndex = 0
            CDKeys.D2DVTotal = pk.d2dvCount
          
            frmMain.lblControl(D2DVTotal).Caption = pk.d2dvCount
        End If
    
        If (pk.d2xpCount > 0) Then
            ReDim CDKeys.D2XP(pk.d2xpCount - 1)
            CDKeys.D2XPIndex = 0
            CDKeys.D2XPTotal = pk.d2xpCount
    
            frmMain.lblControl(D2XPTotal).Caption = pk.d2xpCount
        End If
    
        Dim key As Variant, keyProduct As Variant
  
        For Each key In pk.dicKeys.keys
            keyProduct = pk.dicKeys.Item(key)
      
            Select Case keyProduct
                Case "W2BN"
                    CDKeys.W2BN(w2bnIdx) = key
                    w2bnIdx = w2bnIdx + 1
                Case "D2DV"
                    CDKeys.D2DV(d2dvIdx) = key
                    d2dvIdx = d2dvIdx + 1
                Case "D2XP"
                    CDKeys.D2XP(d2xpIdx) = key
                    d2xpIdx = d2xpIdx + 1
            End Select
        Next
    End If
  
    totalNonExpKeys = (pk.w2bnCount + pk.d2dvCount)
    totalExpKeys = pk.d2xpCount
  
    reportProcessedKeys pk
  
    frmMain.lblControl(KEYS_TOTAL).Caption = totalNonExpKeys + totalExpKeys
    frmMain.lblControl(KEYS_TESTED).Caption = 0
    frmMain.lblControl(PERCENT_TOTAL).Caption = "0.0%"
  
    resetBenchmark
End Sub

Public Sub loadKeysFromFiles(ByVal keyFolder As Folder, pk As ParsedKeys)
    On Error Resume Next

    Dim sf As Folder

    For Each sf In keyFolder.SubFolders
        loadKeysFromFiles sf, pk
    Next
  
    Dim f As File
  
    For Each f In keyFolder.Files
        If (getFileSize(f.path) > 0) Then
            Dim arrFileLines() As String
    
            Open f.path For Input As #1
                arrFileLines = Split(Input(LOF(1), 1), vbNewLine)
            Close #1
      
            If (err.Number = 0) Then
                For i = 0 To UBound(arrFileLines)
                    processKeyLine arrFileLines(i), pk
                Next i
            Else
                err.Clear
                pk.badFiles = pk.badFiles + 1
            End If
        End If
    Next
End Sub

Public Sub deleteAllKeyFiles(ByVal keyFolder As Folder)
    On Error Resume Next

    Dim sf As Folder

    For Each sf In keyFolder.SubFolders
        deleteAllKeyFiles sf
        sf.Delete True
    Next
  
    Dim f As File
  
    For Each f In keyFolder.Files
        f.Delete True
    Next
End Sub

Public Sub processKeyLine(ByVal keyLine As String, pk As ParsedKeys)
    Dim cleanKey As String, dk As DecodedKey, lenKey As Integer

    cleanKey = cleanKeyLine(keyLine)
    lenKey = Len(cleanKey)
  
    If (lenKey = 16 Or lenKey = 26) Then
        If (isSanitizedKey(cleanKey)) Then
            cleanKey = UCase$(cleanKey)
        
            If (Not pk.dicKeys.Exists(cleanKey)) Then
                dk = Decode(cleanKey)
        
                If (dk.successful) Then
                    pk.dicKeys.Add cleanKey, dk.Product
        
                    Select Case dk.Product
                        Case "W2BN"
                            pk.w2bnCount = pk.w2bnCount + 1
                        Case "D2DV"
                            pk.d2dvCount = pk.d2dvCount + 1
                        Case "D2XP"
                            pk.d2xpCount = pk.d2xpCount + 1
                    End Select
                Else
                    pk.invalidKeys = pk.invalidKeys + 1
                End If
            Else
                pk.duplicateKeys = pk.duplicateKeys + 1
            End If
        Else
            pk.skippedLines = pk.skippedLines + 1
        End If
    Else
        If (cleanKey <> vbNullString) Then
            pk.skippedLines = pk.skippedLines + 1
        End If
    End If
End Sub

Public Function cleanKeyLine(keyLine As String) As String
    Dim parsedKeyLine As String

    parsedKeyLine = Trim$(keyLine)
  
    If (parsedKeyLine <> vbNullString) Then
        If (InStr(parsedKeyLine, " ---> ")) Then
            parsedKeyLine = left$(parsedKeyLine, InStr(parsedKeyLine, " ---> ") - 1)
        End If
    
        If (InStr(parsedKeyLine, " ")) Then
            Dim splitString() As String, longestLength As Integer, longestString As String
            Dim found As Boolean
      
            longestLength = -1
            longestString = vbNullString
      
            splitString = Split(parsedKeyLine, " ")
      
            For i = 0 To UBound(splitString)
                Dim line As String, lenLine As Integer, validLength As Boolean
        
                line = splitString(i)
                line = Replace(line, "-", vbNullString)
        
                If (line <> vbNullString) Then
                    lenLine = Len(line)
                    validLength = (lenLine = 16 Or lenLine = 26)
          
                    If (validLength) Then
                        parsedKeyLine = line
                        found = True
                        Exit For
                    Else
                        If (lenLine > longestLength) Then
                            longestLength = lenLine
                            longestString = line
                        End If
                    End If
                End If
            Next i
      
            If (Not found) Then
                parsedKeyLine = longestString
            End If
        Else
            parsedKeyLine = Replace(parsedKeyLine, "-", vbNullString)
        End If
    
        If (Len(parsedKeyLine) > 26) Then
            parsedKeyLine = left$(parsedKeyLine, 26)
        End If
    End If
  
    cleanKeyLine = parsedKeyLine
End Function

Public Function isSanitizedKey(ByVal key As String) As Boolean
    For i = 1 To Len(key)
        Dim ch As String
        ch = UCase$(Mid$(key, i, 1))
    
        If ((Asc(ch) < 65 Or Asc(ch) > 90) And Not IsNumeric(ch)) Then
            isSanitizedKey = False
            Exit Function
        End If
    Next i

    isSanitizedKey = True
End Function

Public Sub reportProcessedKeys(pk As ParsedKeys)
    If (pk.duplicateKeys > 0) Then
        AddChat vbRed, "Removed ", vbWhite, pk.duplicateKeys, vbRed, " duplicate key" & IIf(pk.duplicateKeys > 1, "s", vbNullString) & "."
    End If
  
    If (pk.invalidKeys > 0) Then
        AddChat vbRed, "Removed ", vbWhite, pk.invalidKeys, vbRed, " invalid key" & IIf(pk.invalidKeys > 1, "s", vbNullString) & "."
    End If
  
    If (pk.skippedLines > 0) Then
        AddChat vbRed, "Skipped ", vbWhite, pk.skippedLines, vbRed, " line" & IIf(pk.skippedLines > 1, "s", vbNullString) & "."
    End If
  
    If (pk.badFiles > 0) Then
        AddChat vbRed, "Skipped ", vbWhite, pk.badFiles, vbRed, " bad file" & IIf(pk.badFiles > 1, "s", vbNullString) & "."
    End If
End Sub

Public Function canTestRegularKeys() As Boolean
    If (CDKeys.W2BNIndex > -1) Then canTestRegularKeys = True
    If (CDKeys.D2DVIndex > -1) Then canTestRegularKeys = True
End Function

Public Function canTestExpansion(ByVal Product As String) As Boolean
    Select Case Product
        Case "D2DV", "D2XP":    If (CDKeys.D2XPIndex > -1) Then canTestExpansion = True
    End Select
End Function

Public Function getCDKeyFromList() As FoundKey
    Dim found As Boolean, key As String, fk As FoundKey, i As Long
  
    If (CDKeys.W2BNIndex > -1) Then
        For i = CDKeys.W2BNIndex To CDKeys.W2BNTotal - 1
            key = CDKeys.W2BN(i)
      
            If (key <> vbNullString) Then
                fk.CDKey = key
                fk.Product = "W2BN"
                fk.keyIndex = i
      
                found = True
                Exit For
            End If
        Next i
  
        If (Not found Or i = CDKeys.W2BNTotal - 1) Then
            CDKeys.W2BNIndex = -1
        Else
            CDKeys.W2BNIndex = i + 1
        End If
    ElseIf (CDKeys.D2DVIndex > -1) Then
        For i = CDKeys.D2DVIndex To CDKeys.D2DVTotal - 1
            key = CDKeys.D2DV(i)
      
            If (key <> vbNullString) Then
                fk.CDKey = key
                fk.Product = "D2DV"
                fk.keyIndex = i
      
                found = True
                Exit For
            End If
        Next i
  
        If (Not found Or i = CDKeys.D2DVTotal - 1) Then
            CDKeys.D2DVIndex = -1
        Else
            CDKeys.D2DVIndex = i + 1
        End If
    End If
  
    getCDKeyFromList = fk
End Function

Public Function getCDKeyFromListEx(ByVal Product As String) As FoundKey
    Dim fk As FoundKey, found As Boolean, key As String, i As Long

    Select Case Product
        Case "D2DV", "D2XP"
            If (CDKeys.D2XPIndex > -1) Then
                For i = CDKeys.D2XPIndex To CDKeys.D2XPTotal - 1
                    key = CDKeys.D2XP(i)
          
                    If (key <> vbNullString) Then
                        With fk
                            .CDKey = key
                            .Product = "D2XP"
                            .keyIndex = i
                        End With
            
                        found = True
                        Exit For
                    End If
                Next i
      
                If (Not found Or i = CDKeys.D2XPTotal - 1) Then
                    CDKeys.D2XPIndex = -1
                Else
                    CDKeys.D2XPIndex = i + 1
                End If
            End If
    End Select
  
    getCDKeyFromListEx = fk
End Function

Public Sub removeCDKeyByIndex(ByVal keyIndex As Long, ByVal Product As String)
    Select Case Product
        Case "W2BN":    CDKeys.W2BN(keyIndex) = vbNullString
        Case "D2DV":    CDKeys.D2DV(keyIndex) = vbNullString
        Case "D2XP":    CDKeys.D2XP(keyIndex) = vbNullString
  End Select
End Sub

Public Sub exportKeyToFile(ByVal key As String, ByVal Product As String, State As String, Optional ByVal inUseName As String = vbNullString)
    Dim dirName As String, cdKeyFile As String, stateString As String
  
    Select Case Product
        Case "W2BN": dirName = "WarCraft II"
        Case "D2DV": dirName = "Diablo II"
        Case "D2XP": dirName = "Diablo II - Lord of Destruction"
    End Select

    cdKeyFile = App.path & "\"

    If (config.cdKeyProfile <> vbNullString) Then
        cdKeyFile = cdKeyFile & "CD-Key Profiles\"
    
        If (Not DirExists(cdKeyFile)) Then
            MkDir (cdKeyFile)
        End If
    
        cdKeyFile = cdKeyFile & config.cdKeyProfile
    
        If (config.addRealmToProfile) Then
            cdKeyFile = cdKeyFile & " @ " & config.ServerRealm
        End If
    
        cdKeyFile = cdKeyFile & "\"
    
        If (Not DirExists(cdKeyFile)) Then
            MkDir (cdKeyFile)
        End If
    Else
        cdKeyFile = cdKeyFile & CDKEYS_TESTED_DEFAULT_FOLDER & "\"
  
        If (Not DirExists(cdKeyFile)) Then
            MkDir (cdKeyFile)
        End If
    End If

    If (config.addDateToTested) Then
        Dim dateFolder As String
  
        cdKeyFile = cdKeyFile & Format$(Now, " mmmm d, yyyy") & "\"
    
        If (Not DirExists(cdKeyFile)) Then
            MkDir (cdKeyFile)
        End If
    End If
  
    cdKeyFile = cdKeyFile & dirName & "\"
  
    If (Not DirExists(cdKeyFile)) Then
        MkDir (cdKeyFile)
    End If
  
    cdKeyFile = cdKeyFile & State & ".txt"
  
    Open cdKeyFile For Append As #1
        Print #1, UCase$(key) & IIf(inUseName <> vbNullString, " ---> " & inUseName, vbNullString)
    Close #1
End Sub

Public Function getLabelByKeyState(ByVal Product As String, ByVal State As String) As Integer
    Dim labelConstant As Integer

    Select Case LCase$(State)
        Case "perfect"
            Select Case Product
                Case "W2BN": labelConstant = W2BNPerfect
                Case "D2DV": labelConstant = D2DVPerfect
                Case "D2XP": labelConstant = D2XPPerfect
            End Select
        Case "inuse"
            Select Case Product
                Case "W2BN": labelConstant = W2BNInUse
                Case "D2DV": labelConstant = D2DVInUse
                Case "D2XP": labelConstant = D2XPInUse
            End Select
        Case "muted"
            Select Case Product
                Case "W2BN": labelConstant = W2BNMuted
                Case "D2DV": labelConstant = D2DVMuted
                Case "D2XP": labelConstant = D2XPMuted
            End Select
        Case "voided"
            Select Case Product
                Case "W2BN": labelConstant = W2BNVoided
                Case "D2DV": labelConstant = D2DVVoided
                Case "D2XP": labelConstant = D2XPVoided
            End Select
        Case "jailed"
            Select Case Product
                Case "W2BN": labelConstant = W2BNJailed
                Case "D2DV": labelConstant = D2DVJailed
                Case "D2XP": labelConstant = D2XPJailed
            End Select
        Case "other"
            Select Case Product
                Case "W2BN": labelConstant = W2BNOther
                Case "D2DV": labelConstant = D2DVOther
                Case "D2XP": labelConstant = D2XPOther
            End Select
        Case "banned"
            Select Case Product
                Case "W2BN": labelConstant = W2BNBanned
                Case "D2DV": labelConstant = D2DVBanned
                Case "D2XP": labelConstant = D2XPBanned
            End Select
        Case "invalid"
            Select Case Product
                Case "W2BN": labelConstant = W2BNInvalid
                Case "D2DV": labelConstant = D2DVInvalid
                Case "D2XP": labelConstant = D2XPInvalid
            End Select
    End Select
  
    getLabelByKeyState = labelConstant
End Function

Public Sub postKeysTested(ByVal Product As String)
    Dim keysTested As Long, keysTotal As Long, lblKeysPercentIndex As Integer
  
    Select Case Product
        Case "W2BN"
            CDKeys.W2BNTested = CDKeys.W2BNTested + 1
            keysTested = CDKeys.W2BNTested
            keysTotal = CDKeys.W2BNTotal
            lblKeysPercentIndex = W2BNPercent
        Case "D2DV"
            CDKeys.D2DVTested = CDKeys.D2DVTested + 1
            keysTested = CDKeys.D2DVTested
            keysTotal = CDKeys.D2DVTotal
            lblKeysPercentIndex = D2DVPercent
        Case "D2XP"
            CDKeys.D2XPTested = CDKeys.D2XPTested + 1
            keysTested = CDKeys.D2XPTested
            keysTotal = CDKeys.D2XPTotal
            lblKeysPercentIndex = D2XPPercent
    End Select
  
    frmMain.lblControl(lblKeysPercentIndex).Caption = Format$(((keysTested / keysTotal) * 100), "0.0") & "%"
    frmMain.lblControl(KEYS_TESTED).Caption = testedNonExpKeys + testedExpKeys
    frmMain.lblControl(PERCENT_TOTAL).Caption = Format$(((testedNonExpKeys + testedExpKeys) / (totalNonExpKeys + totalExpKeys)) * 100, "0.0") & "%"
End Sub

Public Sub sendKeysBack()
    Dim arrProducts() As Variant, Product As Variant
    Dim fso As New FileSystemObject, f As Folder
    
    arrProducts = Array("W2BN", "D2DV", "D2XP")
    Set f = fso.GetFolder(CDKEYS_FOLDER)
  
    deleteAllKeyFiles f
  
    For Each Product In arrProducts
        Dim arrCDKeys() As String, hasKeys As Boolean
    
        hasKeys = False
    
        Select Case Product
            Case "W2BN"
                If (CDKeys.W2BNTotal > 0) Then
                    arrCDKeys = CDKeys.W2BN
                    hasKeys = True
                End If
            Case "D2DV"
                If (CDKeys.D2DVTotal > 0) Then
                    arrCDKeys = CDKeys.D2DV
                    hasKeys = True
                End If
            Case "D2XP"
                If (CDKeys.D2XPTotal > 0) Then
                    arrCDKeys = CDKeys.D2XP
                    hasKeys = True
                End If
        End Select
    
        Open App.path & "\" & CDKEYS_FOLDER & "\" & Product & ".txt" For Output As #1
        If (hasKeys) Then
            For i = 0 To UBound(arrCDKeys)
                If (arrCDKeys(i) <> vbNullString) Then
                    Print #1, arrCDKeys(i)
                End If
            Next i
        End If
        Close #1
    Next
End Sub

Public Sub wipeCDKeysFromTesting()
    For i = 0 To loadedSockets - 1
        With BNETData(i)
            .CDKey = vbNullString
            .cdKeyExp = vbNullString
            
            .cdKeyIndex = 0
            .cdKeyExpIndex = 0
            
            .numTested = 0
            .TestedEXP = 0
            
            .Product = vbNullString
            .productRegular = vbNullString
            .productExpansion = vbNullString
            
            .savedKeyState = vbNullString
            
            .isExpansion = False
        End With
    Next i
End Sub

Public Function assignKeys(index As Integer) As Boolean
    With BNETData(index)
        If (.CDKey = vbNullString) Then
            If (canTestRegularKeys()) Then
                Dim fk As FoundKey
      
                fk = getCDKeyFromList()
        
                .CDKey = fk.CDKey
                .cdKeyIndex = fk.keyIndex
        
                .Product = fk.Product
                .productRegular = fk.Product

                .isExpansion = False
            Else
                assignKeys = False
                Exit Function
            End If
        Else
            If (canTestExpansion(.Product)) Then
                Dim fkEx As FoundKey
                fkEx = getCDKeyFromListEx(.Product)

                .cdKeyExp = fkEx.CDKey
                .cdKeyExpIndex = fkEx.keyIndex

                .Product = fkEx.Product
                .productExpansion = fkEx.Product

                .isExpansion = True
            Else
                .isExpansion = False
            End If
        End If
    End With

    assignKeys = True
End Function

Public Sub restoreKeysToList()
    For i = 0 To UBound(BNETData)
        With BNETData(i)
            If (.CDKey <> vbNullString And .Product <> vbNullString) Then
                repopulateKeyList .Product, .CDKey
        
              .CDKey = vbNullString
              .Product = vbNullString
              .productRegular = vbNullString
            End If
      
            If (.cdKeyExp <> vbNullString And .productExpansion <> vbNullString) Then
                repopulateKeyList .productExpansion, .cdKeyExp
        
                .cdKeyExp = vbNullString
                .productExpansion = vbNullString
            End If
        End With
    Next i
End Sub

Private Sub repopulateKeyList(ByVal Product As String, ByVal key As String)
    Dim arrKeys() As String
  
    Select Case Product
        Case "W2BN"
            arrKeys = CDKeys.W2BN
      
            If (CDKeys.W2BNIndex = -1) Then
                CDKeys.W2BNIndex = 0
            End If
        Case "D2DV"
            arrKeys = CDKeys.D2DV
    
            If (CDKeys.D2DVIndex = -1) Then
                CDKeys.D2DVIndex = 0
            End If
        Case "D2XP"
            arrKeys = CDKeys.D2XP
      
            If (CDKeys.D2XPIndex = -1) Then
                CDKeys.D2XPIndex = 0
            End If
    End Select
  
    For i = 0 To UBound(arrKeys)
        If (arrKeys(i) = vbNullString) Then
            arrKeys(i) = key
            Exit For
        End If
    Next i
End Sub

Public Function Decode(ByVal CDKey As String) As DecodedKey
    Dim publicVal As Long, Product As Long, dk As DecodedKey
  
    Product = -1
  
    kd_quick CDKey, 0, 0, publicVal, Product, vbNullString, 0

    If (Product > 0) Then
        Select Case Product
            Case &H4
                dk.Product = "W2BN"
                dk.successful = True
            Case &H6, &H7, &H18
                dk.Product = "D2DV"
                dk.successful = True
            Case &HA, &HC, &H19
                dk.Product = "D2XP"
                dk.successful = True
        End Select
    End If
  
    Decode = dk
End Function
