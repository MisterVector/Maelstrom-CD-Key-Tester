Attribute VB_Name = "modBNET"
Public Sub Recv0x25(index As Integer)
    packet(index).InsertDWORD packet(index).GetDWORD
    packet(index).sendPacket &H25
End Sub

Public Sub Send0x50(index As Integer)
    With packet(index)
        .InsertDWORD &H0
        .InsertNonNTString "68XI" & StrReverse(BNETData(index).Product)
        .InsertDWORD getVerByte(BNETData(index).Product)
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

Public Sub Recv0x50(index As Integer)
    Dim tempFT As FILETIME, mpqFileTime As String, mpqFileName As String
    Dim checksumFormula As String
    
    packet(index).Skip 4              'Logon type
    
    BNETData(index).ClientToken = GetTickCount
    BNETData(index).ServerToken = packet(index).GetDWORD
    packet(index).Skip 4    'UDPValue
    
    tempFT.dwLowDateTime = packet(index).GetDWORD
    tempFT.dwHighDateTime = packet(index).GetDWORD
    
    mpqFileTime = GetFTTime(tempFT)
    mpqFileName = packet(index).getNTString
    
    checksumFormula = packet(index).getNTString
    
    Send0x51 index, mpqFileTime, mpqFileName, checksumFormula
End Sub

Public Sub Send0x51(index As Integer, ByVal mpqFileTime As String, ByVal mpqFileName As String, ByVal checksumFormula As String)
    Dim CDKeyHash(1)      As String * 20
    Dim ProdVal(1)        As Long
    Dim PubVal(1)         As Long
    Dim hashFiles()       As String
    Dim EXEVersion        As Long
    Dim EXEchecksum       As Long
    Dim exeInfoString     As String
  
    With BNETData(index)
        Dim lockdownFileName As String, hsr As HashSearchResult
    
        If (.Product = "W2BN") Then
            lockdownFileName = left$(mpqFileName, Len(mpqFileName) - 4) & ".dll"
        End If
    
        hsr = getHashes(.Product, lockdownFileName)
    
        If (Not hsr.hashesExist) Then
            AddChatB vbRed, "Socket #" & index & ": Check revision failed for " & .Product & "!"
            AddChatB vbRed, "Reason: " & hsr.errorMessage & "."
            stopTesting vbYellow, "Address the issue and then hit ", vbWhite, "Start", vbYellow, " again."
            Exit Sub
        End If
    
        hashFiles = hsr.hashes
        exeInfoString = String$(crev_max_result, Chr$(0))
  
        If ((modBNETAPI.kd_quick(.CDKey, .ClientToken, .ServerToken, PubVal(0), ProdVal(0), CDKeyHash(0), 20) = 0)) Then
            closeSocket index
            frmMain.tmrCheckFailed(index).Enabled = False
        
            AddChatB vbRed, "Socket #" & index & ": Key (" & .Product & ") failed to decode. Rotating key..."
            .CDKey = vbNullString
      
            testedNonExpKeys = testedNonExpKeys + 1
            postKeysTested .Product
      
            If (assignKeys(index)) Then
                frmMain.tmrReconnect(index).Enabled = True
            Else
                AddChat vbRed, "Socket #" & index & ": The key list has run out."
                markSocketDead index
        
                If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                    AddChatB vbYellow, "All keys have been tested."
                    stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                    Exit Sub
                ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                    AddChatB vbYellow, "All non-expansion keys have been tested."
                    stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                    Exit Sub
                End If
      
                If (socketsAvailable = 0) Then
                    AddChatB vbYellow, "No more connections could be made."
                    stopTesting vbYellow, "Add more proxies and then click ", vbWhite, "Reload Proxies", vbYellow, "."
                    Exit Sub
                End If
            End If
      
            Exit Sub
        End If
  
        If (.cdKeyExp <> vbNullString And .Product = "D2XP") Then
            If ((modBNETAPI.kd_quick(.cdKeyExp, .ClientToken, .ServerToken, PubVal(1), ProdVal(1), CDKeyHash(1), 20) = 0)) Then
                closeSocket index
                frmMain.tmrCheckFailed(index).Enabled = False
          
                AddChat vbRed, "Socket #" & index & ": " & " Expansion key (" & .productExpansion & ") failed to decode. Rotating key..."
                .cdKeyExp = vbNullString
        
                testedExpKeys = testedExpKeys + 1
                postKeysTested .productExpansion
        
                If (assignKeys(index)) Then
                    frmMain.tmrReconnect(index).Enabled = True
                Else
                    AddChat vbRed, "Socket #" & index & ": The key list has run out."
                    markSocketDead index
          
                    If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                        AddChatB vbYellow, "All keys have been tested."
                        stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                        Exit Sub
                    ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                        AddChatB vbYellow, "All non-expansion keys have been tested."
                        stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                        Exit Sub
                    End If
        
                    If (socketsAvailable = 0) Then
                        AddChatB vbYellow, "No more connections could be made."
                        stopTesting vbYellow, "Add more proxies and then click ", vbWhite, "Reload Proxies", vbYellow, "."
                        Exit Sub
                    End If
                End If
      
                Exit Sub
            End If
        End If
    
        modBNETAPI.check_revision mpqFileTime, IIf(lockdownFileName <> vbNullString, lockdownFileName, mpqFileName), checksumFormula, App.path & "\VersionCheck.ini", .Product, EXEVersion, EXEchecksum, exeInfoString
    End With

    With packet(index)
        .InsertDWORD BNETData(index).ClientToken
        .InsertDWORD EXEVersion
        .InsertDWORD EXEchecksum
        .InsertDWORD IIf(ProdVal(1), &H2, &H1)
        .InsertDWORD &H0
        
        .InsertDWORD Len(BNETData(index).CDKey)
        .InsertDWORD ProdVal(0)
        .InsertDWORD PubVal(0)
        .InsertDWORD &H0
        .InsertNonNTString CDKeyHash(0)

        If (ProdVal(1) <> 0) Then
            .InsertDWORD Len(BNETData(index).cdKeyExp)
            .InsertDWORD ProdVal(1)
            .InsertDWORD PubVal(1)
            .InsertDWORD &H0
            .InsertNonNTString CDKeyHash(1)
        End If

        .InsertNTString KillNull(exeInfoString)
        .InsertNTString KEY_TESTER_NAME
        .sendPacket &H51
    End With
End Sub

Public Sub Recv0x51(index As Integer)
    Dim statusCode As Long, Product As String
  
    statusCode = packet(index).GetDWORD
    Product = BNETData(index).Product
    FreeMemory
  
    If (statusCode = &H0) Then
        Send0x3A index
    Else
        frmMain.tmrCheckFailed(index).Enabled = False
        closeSocket index
    
        Dim inUse As String
  
        If (statusCode = &H201 Or statusCode = &H211) Then
            inUse = packet(index).getNTString
    
            If (inUse = vbNullString) Then
                inUse = "Anonymous key owner"
            End If
        End If
    
        Select Case statusCode
            Case &H100
                frmMain.lblStart_EmulateClick
            
                msgBoxResult = MsgBox("The hashes for " & Product & " are out of date. ", vbOKOnly Or vbExclamation, PROGRAM_TITLE)
    
                Exit Sub
            Case &H101
                AddChatB vbRed, "The version byte for " & Product & " was invalid."
                AddChatB vbRed, "Attempting to update version byte..."
            
                frmMain.tmrBenchmark.Enabled = False
            
                For i = 0 To UBound(BNETData)
                    frmMain.tmrCheckFailed(i).Enabled = False
                    frmMain.tmrReconnect(i).Enabled = False
                    closeSocket i
                Next i
            
                frmMain.sckBNLS.Connect config.bnlsServer, 9367
                requestProduct = Product
                frmMain.tmrCheckBNLS.Enabled = True
            
                Exit Sub
            'Case &H102
                'frmMain.lblStart_EmulateClick
          
                'msgBoxResult = MsgBox("The hashes for " & product & " are too new.", vbOKOnly Or vbExclamation, PROGRAM_TITLE)
    
                'Exit Sub
        End Select
        
        Call handleOtherKeys(index, statusCode, inUse)
        
        If (Not assignKeys(index)) Then
            Dim testEnded As Boolean
        
            AddChat vbRed, "Socket #" & index & ": The key list has run out."
            markSocketDead index
          
            If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
                AddChatB vbYellow, "All keys have been tested."
                stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                Exit Sub
            ElseIf (testedNonExpKeys = totalNonExpKeys) Then
                AddChatB vbYellow, "All non-expansion keys have been tested."
                stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
                Exit Sub
            End If
        
            If (socketsAvailable = 0) Then
                AddChatB vbYellow, "No more connections could be made."
                stopTesting vbYellow, "Add more proxies and then click ", vbWhite, "Reload Proxies", vbYellow, "."
                Exit Sub
            End If
        
            Exit Sub
        End If
        
        frmMain.tmrReconnect(index).Enabled = True
    End If
End Sub

Public Sub Send0x3A(index As Integer)
    Dim HashCode As String * 20
    
    HashCode = doubleHashPassword(config.password, BNETData(index).ClientToken, BNETData(index).ServerToken)
    
    With packet(index)
        .InsertDWORD BNETData(index).ClientToken
        .InsertDWORD BNETData(index).ServerToken
        .InsertNonNTString HashCode
        .InsertNTString config.name
        .sendPacket &H3A
    End With
End Sub

Public Sub Recv0x3A(index As Integer)
    Select Case packet(index).GetDWORD
        Case &H0: 'Success
            Send0x14 index
            Send0xAC index
        Case &H1: 'Creating account
            AddChatB vbWhite, config.name & "@" & config.ServerRealm, vbYellow, " does not exist. Maelstrom will create it."
    
            For i = 0 To UBound(BNETData)
                If (i <> index) Then
                    closeSocket i
                    frmMain.tmrCheckFailed(i).Enabled = False
                    frmMain.tmrReconnect(i).Enabled = False
                End If
            Next i
                         
            Send0x3D index
        Case &H2: 'Bad password
    
            AddChat vbRed, "Invalid password for ", vbWhite, config.name & "@" & config.ServerRealm, vbRed, "."
            stopTesting vbYellow, "Change the password and then click ", vbWhite, "Start", vbYellow, " again."
        Case &H6: 'Account closed
            AddChat vbRed, "The account ", vbWhite, config.name & "@" & config.ServerRealm, vbRed, " has been banned."
            stopTesting vbYellow, "Change the account name and then click ", vbWhite, "Start", vbYellow, " again."
    End Select
End Sub

Public Sub Send0x3D(index As Integer)
    Dim password_hash As String * 20
  
    password_hash = hashPassword(config.password)

    With packet(index)
        .InsertNonNTString password_hash
        .InsertNTString config.name
        .sendPacket &H3D
    End With
End Sub

Public Sub Recv0x3D(index As Integer)
    Dim Result As Long
  
    With packet(index)
        Result = .GetDWORD
    End With
  
    If (Result = &H0) Then
        AddChatB vbGreen, "Created the account ", vbWhite, config.name & "@" & config.ServerRealm, vbGreen, "!"
        stopTesting vbYellow, "Click ", vbWhite, "Start", vbYellow, " to start testing again."
    Else
        Dim reason As String
        reason = accountIdToReason(Result)
    
        AddChatB vbRed, "Unable to create the account ", vbWhite, config.name & "@" & config.ServerRealm, vbRed, "!"
        AddChatB vbRed, "Reason: " & reason & "."
  
        stopTesting vbYellow, "Fix the issue with the account and click ", vbWhite, "Start", vbYellow, " again."
    End If
End Sub

'// SID_NEWS
Public Sub Send0x46(index As Integer)
    packet(index).InsertDWORD &HFFFFFFFF
    packet(index).sendPacket &H46
End Sub

'// SID_NEWS
Public Sub Recv0x46(index As Integer)
    Dim isVoided As Boolean, isMuted As Boolean
    Dim dumpedPacket As String, Product As String

    frmMain.tmrCheckFailed(index).Enabled = False
    closeSocket index

    dumpedPacket = packet(index).getPacket

    If (InStr(dumpedPacket, "Your account is muted.") > 0) Then isMuted = True
    If (InStr(dumpedPacket, "Your account has had all chat privileges suspended.") > 0) Then isVoided = True
  
    If (isMuted Or isVoided) Then
        Call voidedMutedOrJailedKeyEvaluation(index, isMuted, isVoided)
    Else
        Call perfectKeyEvaluation(index)
    End If

    If (Not assignKeys(index)) Then
        Dim testEnded As Boolean
  
        AddChat vbRed, "Socket #" & index & ": The key list has run out."
        markSocketDead index
    
        If ((testedNonExpKeys + testedExpKeys) = (totalNonExpKeys + totalExpKeys)) Then
            AddChatB vbYellow, "All keys have been tested."
            stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
            Exit Sub
        ElseIf (testedNonExpKeys = totalNonExpKeys) Then
            AddChatB vbYellow, "All non-expansion keys have been tested."
            stopTesting vbYellow, "Add more keys and then click ", vbWhite, "Reload CD-Keys", vbYellow, "."
            Exit Sub
        End If
  
        If (socketsAvailable = 0) Then
            AddChatB vbYellow, "No more connections could be made."
            stopTesting vbYellow, "Add more proxies and then click ", vbWhite, "Reload Proxies", vbYellow, "."
            Exit Sub
        End If
  
        Exit Sub
    End If
    
    frmMain.tmrReconnect(index).Enabled = True
End Sub

Public Sub Send0x14(index As Integer)
    packet(index).InsertNonNTString "tenb"
    packet(index).sendPacket &H14
End Sub

Public Sub Send0xAC(index As Integer)
    With packet(index)
        .InsertNTString config.name
        .InsertByte &H0
        .sendPacket &HA

        .InsertDWORD &H2
        .InsertNTString config.homeChannel
        .sendPacket &HC
    End With
End Sub

Public Sub Recv0x0A(index As Integer)
    Send0x46 index
  
    With BNETData(index)
        .nls_P = 0
        .ServerToken = 0
    End With
End Sub
