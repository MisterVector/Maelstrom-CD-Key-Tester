Attribute VB_Name = "modBNET"
Public Sub Recv0x25(Index As Integer)
    packet(Index).InsertDWORD packet(Index).GetDWORD
    packet(Index).sendPacket &H25
End Sub

Public Sub Send0x50(Index As Integer)
    With packet(Index)
        .InsertDWORD &H0
        .InsertNonNTString "68XI" & StrReverse(BNETData(Index).product)
        .InsertDWORD getVerByte(BNETData(Index).product)
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertDWORD &H0
        .InsertNTString "USA"
        .InsertNTString "United States"
        .sendPacket &H50
    End With
End Sub

Public Sub Recv0x50(Index As Integer)
    Dim tempFT As FILETIME, mpqFileTime As String, mpqFileName As String
    Dim checksumFormula As String
    
    packet(Index).Skip 4              'Logon type
    
    BNETData(Index).ClientToken = GetTickCount
    BNETData(Index).ServerToken = packet(Index).GetDWORD
    packet(Index).Skip 4    'UDPValue
    
    tempFT.dwLowDateTime = packet(Index).GetDWORD
    tempFT.dwHighDateTime = packet(Index).GetDWORD
    
    mpqFileTime = Hash_Filetime.GetFTTime(tempFT)
    mpqFileName = packet(Index).getNTString
    
    checksumFormula = packet(Index).getNTString
    
    Send0x51 Index, mpqFileTime, mpqFileName, checksumFormula
End Sub

Public Sub Send0x51(Index As Integer, ByVal mpqFileTime As String, ByVal mpqFileName As String, ByVal checksumFormula As String)
    Dim CDKeyHash(1)      As String * 20
    Dim ProdVal(1)        As Long
    Dim PubVal(1)         As Long
    Dim hashFiles()       As String
    Dim EXEVersion        As Long
    Dim EXEchecksum       As Long
    Dim EXEInfoString     As String
  
    With BNETData(Index)
        Dim lockdownFileName As String, hsr As HashSearchResult
    
        If (.product = "W2BN") Then
            lockdownFileName = left(mpqFileName, Len(mpqFileName) - 4) & ".dll"
        End If
    
        hsr = getHashes(.product, lockdownFileName)
    
        If (Not hsr.hashesExist) Then
            AddChatB vbRed, "Socket #" & Index & ": Check revision failed for " & .product & "!"
            AddChatB vbRed, "Reason: " & hsr.errorMessage & "."
            stopTesting vbYellow, "Address the issue and then hit ""Start"" again."
            Exit Sub
        End If
    
        hashFiles = hsr.hashes
        EXEInfoString = String$(crev_max_result, Chr$(0))
  
        If ((Hash_Lib.decode_hash_cdkey(.cdKey, .ClientToken, .ServerToken, PubVal(0), ProdVal(0), CDKeyHash(0)) = 0)) Then
            closeSocket Index
            frmMain.tmrCheckFailed(Index).Enabled = False
        
            AddChatB vbRed, "Socket #" & Index & ": Key (" & .product & ") failed to decode. Rotating key..."
            .cdKey = vbNullString
      
            testedNonExpKeys = testedNonExpKeys + 1
            postKeysTested .product
      
            If (assignKeys(Index)) Then
                frmMain.tmrReconnect(Index).Enabled = True
            Else
                AddChat vbRed, "Socket #" & Index & ": The key list has run out."
                markSocketDead Index
        
                If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                    AddChatB vbYellow, "All keys have been tested."
                    stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                    Exit Sub
                ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                    AddChatB vbYellow, "All non-expansion keys have been tested."
                    stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                    Exit Sub
                End If
      
                If (socketsAvailable = 0) Then
                    AddChatB vbYellow, "No more connections could be made."
                    stopTesting vbYellow, "Add more proxies and then click ""Reload Proxies""."
                    Exit Sub
                End If
            End If
      
            Exit Sub
        End If
  
        If (.cdKeyExp <> vbNullString And (.product = "W3XP" Or .product = "D2XP")) Then
            If (Hash_Lib.decode_hash_cdkey(.cdKeyExp, .ClientToken, .ServerToken, PubVal(1), ProdVal(1), CDKeyHash(1)) = 0) Then
                closeSocket Index
                frmMain.tmrCheckFailed(Index).Enabled = False
          
                AddChat vbRed, "Socket #" & Index & ": " & " Expansion key (" & .productExpansion & ") failed to decode. Rotating key..."
                .cdKeyExp = vbNullString
        
                testedExpKeys = testedExpKeys + 1
                postKeysTested .productExpansion
        
                If (assignKeys(Index)) Then
                    frmMain.tmrReconnect(Index).Enabled = True
                Else
                    AddChat vbRed, "Socket #" & Index & ": The key list has run out."
                    markSocketDead Index
          
                    If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                        AddChatB vbYellow, "All keys have been tested."
                        stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                        Exit Sub
                    ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                        AddChatB vbYellow, "All non-expansion keys have been tested."
                        stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                        Exit Sub
                    End If
        
                    If (socketsAvailable = 0) Then
                        AddChatB vbYellow, "No more connections could be made."
                        stopTesting vbYellow, "Add more proxies and then click ""Reload Proxies""."
                        Exit Sub
                    End If
                End If
      
                Exit Sub
            End If
        End If
    
        Dim result As Long
    
        result = Hash_Lib.check_revision(mpqFileTime, IIf(lockdownFileName <> vbNullString, lockdownFileName, mpqFileName), checksumFormula, App.path & "\CheckRevisionFromWarden.ini", .product, EXEVersion, EXEchecksum, EXEInfoString)
    End With

    With packet(Index)
        .InsertDWORD BNETData(Index).ClientToken
        .InsertDWORD EXEVersion
        .InsertDWORD EXEchecksum
        .InsertDWORD IIf(ProdVal(1), &H2, &H1)
        .InsertDWORD &H0
        
        .InsertDWORD Len(BNETData(Index).cdKey)
        .InsertDWORD ProdVal(0)
        .InsertDWORD PubVal(0)
        .InsertDWORD &H0
        .InsertNonNTString CDKeyHash(0)

        If (ProdVal(1) <> 0) Then
            .InsertDWORD Len(BNETData(Index).cdKeyExp)
            .InsertDWORD ProdVal(1)
            .InsertDWORD PubVal(1)
            .InsertDWORD &H0
            .InsertNonNTString CDKeyHash(1)
        End If

        .InsertNTString KillNull(EXEInfoString)
        .InsertNTString KeyTesterName
        .sendPacket &H51

        BNETData(Index).EXEInfoString = KillNull(EXEInfoString)
        BNETData(Index).EXEVersion = EXEVersion
        BNETData(Index).EXEchecksum = EXEchecksum
    End With
End Sub

Public Sub Recv0x51(Index As Integer)
    Dim statusCode As Long, product As String
  
    statusCode = packet(Index).GetDWORD
    product = BNETData(Index).product
    FreeMemory
  
    If (statusCode = &H0) Then
        If (product = "WAR3" Or product = "W3XP") Then
            Send0x53 Index
        Else
            Send0x3A Index
        End If
    Else
        frmMain.tmrCheckFailed(Index).Enabled = False
        closeSocket Index
    
        Dim inUse As String
  
        If (statusCode = &H201 Or statusCode = &H211) Then
            inUse = packet(Index).getNTString
    
            If (inUse = vbNullString) Then
                inUse = "Anonymous key owner"
            End If
        End If
    
        Select Case statusCode
            Case &H100
                frmMain.lblStart_EmulateClick
            
                msgBoxResult = MsgBox("The hashes for " & product & " are out of date. ", vbOKOnly & vbExclamation, PROGRAM_NAME)
    
                Exit Sub
            Case &H101
                AddChatB vbRed, "The version byte for " & product & " was invalid."
                AddChatB vbRed, "Attempting to update version byte..."
            
                frmMain.tmrBenchmark.Enabled = False
            
                For i = 0 To UBound(BNETData)
                    frmMain.tmrCheckFailed(i).Enabled = False
                    frmMain.tmrReconnect(i).Enabled = False
                    closeSocket i
                Next i
            
                frmMain.sckBNLS.Connect config.bnlsServer, 9367
                requestProduct = product
                frmMain.tmrCheckBNLS.Enabled = True
            
                Exit Sub
            Case &H102
                frmMain.lblStart_EmulateClick
          
                msgBoxResult = MsgBox("The hashes for " & product & " are too new.", vbOKOnly & vbExclamation, PROGRAM_NAME)
    
                Exit Sub
        End Select
        
        Call handleOtherKeys(Index, statusCode, inUse)
        
        If (Not assignKeys(Index)) Then
            Dim testEnded As Boolean
        
            AddChat vbRed, "Socket #" & Index & ": The key list has run out."
            markSocketDead Index
          
            If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                AddChatB vbYellow, "All keys have been tested."
                stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                Exit Sub
            ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                AddChatB vbYellow, "All non-expansion keys have been tested."
                stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
                Exit Sub
            End If
        
            If (socketsAvailable = 0) Then
                AddChatB vbYellow, "No more connections could be made."
                stopTesting vbYellow, "Add more proxies and then click ""Reload Proxies""."
                Exit Sub
            End If
        
            Exit Sub
        End If
        
        frmMain.tmrReconnect(Index).Enabled = True
    End If
End Sub

Public Sub Send0x3A(Index As Integer)
    Dim HashCode As String * 20
    double_hash_password config.password, BNETData(Index).ClientToken, _
                        BNETData(Index).ServerToken, HashCode

    With packet(Index)
        .InsertDWORD BNETData(Index).ClientToken
        .InsertDWORD BNETData(Index).ServerToken
        .InsertNonNTString HashCode
        .InsertNTString config.name
        .sendPacket &H3A
    End With
End Sub

Public Sub Recv0x3A(Index As Integer)
    Select Case packet(Index).GetDWORD
        Case &H0: 'Success
            Send0x14 Index
            Send0xAC Index
        Case &H1: 'Creating account
            AddChatB vbYellow, config.name & "@" & config.ServerRealm & " does not exist. Maelstrom will create it."
    
            For i = 0 To UBound(BNETData)
                If (i <> Index) Then
                    closeSocket i
                    frmMain.tmrCheckFailed(i).Enabled = False
                    frmMain.tmrReconnect(i).Enabled = False
                End If
            Next i
                         
            Send0x3D Index
        Case &H2: 'Bad password
    
            AddChat vbRed, "Invalid password for " & config.name & "@" & config.ServerRealm & "."
            stopTesting vbYellow, "Change the password and then click ""Start"" again."
        Case &H6: 'Account closed
            AddChat vbRed, "The account " & config.name & "@" & config.ServerRealm & " has been banned."
            stopTesting vbYellow, "Change the account name and then click ""Start"" again."
    End Select
End Sub

Public Sub Send0x3D(Index As Integer)
    Dim password_hash As String * 20
  
    hash_password config.password, password_hash
  
    With packet(Index)
        .InsertNonNTString password_hash
        .InsertNTString config.name
        .sendPacket &H3D
    End With
End Sub

Public Sub Recv0x3D(Index As Integer)
    Dim result As Long
  
    With packet(Index)
        result = .GetDWORD
    End With
  
    If (result = &H0) Then
        AddChatB vbGreen, "Created the account " & config.name & "@" & config.ServerRealm & "!"
        stopTesting vbYellow, "Click ""Start"" to start testing again."
    Else
        Dim reason As String
        reason = accountIdToReason(result, False)
    
        AddChatB vbRed, "Unable to create the account " & config.name & "@" & config.ServerRealm & "!"
        AddChatB vbRed, "Reason: " & reason & "."
  
        stopTesting vbYellow, "Fix the issue with the account and click ""Start"" again."
    End If
End Sub

'// SID_NEWS
Public Sub Send0x46(Index As Integer)
    packet(Index).InsertDWORD &HFFFFFFFF
    packet(Index).sendPacket &H46
End Sub

'// SID_NEWS
Public Sub Recv0x46(Index As Integer)
    Dim isVoided As Boolean, isMuted As Boolean
    Dim dumpedPacket As String, product As String

    frmMain.tmrCheckFailed(Index).Enabled = False
    closeSocket Index

    dumpedPacket = packet(Index).getPacket

    If (InStr(dumpedPacket, "Your account is muted.") > 0) Then isMuted = True
    If (InStr(dumpedPacket, "Your account has had all chat privileges suspended.") > 0) Then isVoided = True
  
    If (isMuted Or isVoided) Then
        Call voidedMutedOrJailedKeyEvaluation(Index, isMuted, isVoided)
    Else
        Call perfectKeyEvaluation(Index)
    End If

    If (Not assignKeys(Index)) Then
        Dim testEnded As Boolean
  
        AddChat vbRed, "Socket #" & Index & ": The key list has run out."
        markSocketDead Index
    
        If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
            AddChatB vbYellow, "All keys have been tested."
            stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
            Exit Sub
        ElseIf (testedNonExpKeys = totalNonExpKeys) Then
            AddChatB vbYellow, "All non-expansion keys have been tested."
            stopTesting vbYellow, "Add more keys and then click ""Reload CD-Keys""."
            Exit Sub
        End If
  
        If (socketsAvailable = 0) Then
            AddChatB vbYellow, "No more connections could be made."
            stopTesting vbYellow, "Add more proxies and then click ""Reload Proxies""."
            Exit Sub
        End If
  
        Exit Sub
    End If
    
    frmMain.tmrReconnect(Index).Enabled = True
End Sub

Public Sub Send0x52(Index As Integer)
    Dim saltHash As String: saltHash = Space(Len(config.nameW3) + 65)
  
    nls_account_create BNETData(Index).nls_P, saltHash

    With packet(Index)
        .InsertNonNTString saltHash
        .sendPacket &H52
    End With
End Sub

Public Sub Recv0x52(Index As Integer)
    Dim result As Long

    With packet(Index)
        result = .GetDWORD
    End With
  
    If (result = &H0) Then
        AddChatB vbGreen, "Created the account " & config.nameW3 & "@" & config.serverRealmW3 & "!"
        stopTesting vbYellow, "Click ""Start"" to begin testing again."
    Else
        Dim reason As String
        reason = accountIdToReason(result, True)
  
        AddChatB vbRed, "Could not create the account " & config.nameW3 & "@" & config.serverRealmW3 & "!"
        AddChatB vbRed, "Reason: " & reason & "."
  
        stopTesting vbYellow, "Fix the issue with the account and click ""Start"" again."
    End If
End Sub

Public Sub Send0x53(Index As Integer)
    Dim nls_A As String

    BNETData(Index).nls_P = nls_init(config.nameW3, config.passwordW3)

    If (BNETData(Index).nls_P = 0) Then
        frmMain.lblStart_EmulateClick
        MsgBox "NLS made a bad call.", vbOKOnly & vbCritical, PROGRAM_NAME
        EndAll
        Exit Sub
    End If

    nls_A = Space(Len(config.nameW3) + 33)
    
    If (nls_account_logon(BNETData(Index).nls_P, nls_A) = 0) Then
        frmMain.lblStart_EmulateClick
        MsgBox "Unable to create NLS key.", vbOKOnly & vbCritical, PROGRAM_NAME
        EndAll
        Exit Sub
    End If

    packet(Index).InsertNonNTString left$(nls_A, Len(nls_A) - Len(config.nameW3) - 1)
    packet(Index).InsertNTString config.nameW3
    packet(Index).sendPacket &H53
End Sub

Public Sub Recv0x53(Index As Integer)
    Select Case packet(Index).GetDWORD
        Case &H0: Send0x54 Index   'Passed
        Case &H1: 'Account Not made
            AddChatB vbYellow, config.nameW3 & "@" & config.serverRealmW3 & " does not exist. Maelstrom will create it."
              
            For i = 0 To UBound(BNETData)
                If (i <> Index) Then
                    closeSocket i
                    frmMain.tmrCheckFailed(i).Enabled = False
                    frmMain.tmrReconnect(i).Enabled = False
                End If
            Next i
            
            Send0x52 Index
        Case &H5                   'Upgrade...
        Case Else: Exit Sub
    End Select
End Sub

Public Sub Send0x54(Index As Integer)
    Dim ProofHash As String * 20
    Dim salt      As String: salt = packet(Index).GetNonNTString(32)
    Dim ServerKey As String: ServerKey = packet(Index).GetNonNTString(32)

    nls_account_logon_proof BNETData(Index).nls_P, ProofHash, ServerKey, salt

    packet(Index).InsertNonNTString ProofHash
    packet(Index).sendPacket &H54
End Sub

Public Sub Recv0x54(Index As Integer)
    Select Case packet(Index).GetDWORD
        Case &H0: GoTo Continue
        Case &H1:
        Case &H2:
        Case &HF:
        Case &HE:
            GoTo Continue
    End Select
    
    Exit Sub
Continue:

    nls_free (BNETData(Index).nls_P)        'Unloads the NLS object to avoid overhead

    Send0x14 Index
    Send0xAC Index
End Sub

Public Sub Send0x14(Index As Integer)
    packet(Index).InsertNonNTString "tenb"
    packet(Index).sendPacket &H14
End Sub

Public Sub Send0xAC(Index As Integer)
    With packet(Index)
        .InsertNTString config.name
        .InsertByte &H0
        .sendPacket &HA

        .InsertDWORD &H2
        .InsertNTString config.homeChannel
        .sendPacket &HC
    End With
End Sub

Public Sub Recv0x0A(Index As Integer)
    Send0x46 Index
  
    With BNETData(Index)
        .nls_P = 0
        .ServerToken = 0
    End With
End Sub
