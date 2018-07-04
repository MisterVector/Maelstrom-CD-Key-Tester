Attribute VB_Name = "modOtherCode"
Public Function loadProxies() As ProxiesLoaded
    Dim proxyVersions() As Variant, dicTemp As New Dictionary, proxyCount As Long
    Dim maxProxiesReached As Boolean, pl As ProxiesLoaded
  
    proxyCount = 0
  
    proxyVersions = Array("SOCKS4", "HTTP")
  
    For i = 0 To UBound(proxyVersions)
        Dim proxyVersion As String, proxyFile As String
  
        proxyVersion = proxyVersions(i)
        proxyFile = App.path & "\" & proxyVersion & ".txt"
  
        If (Dir$(proxyFile) = vbNullString) Then
            Open proxyFile For Output As #1
            Close #1
        End If
    
        If (getFileSize(proxyFile) > 0) Then
            Dim tProxies() As String
      
            Open proxyFile For Input As #1
                tProxies = Split(Input(LOF(1), 1), vbNewLine)
            Close #1

            Dim proxyIndex As Long
            proxyIndex = 0

            For ii = 0 To UBound(tProxies)
                If (tProxies(ii) <> vbNullString) Then
                    tProxies(ii) = Trim(tProxies(ii))
                    
                    If (InStr(tProxies(ii), ":")) Then
                        Dim IP As String, port As String
                
                        IP = Split(tProxies(ii), ":")(0)
                        port = Split(tProxies(ii), ":")(1)
            
                        If (IsNumeric(port) And IsNumeric(Replace(IP, ".", vbNullString))) Then
                            If (port <= 65535 And port > 0) Then
                                If (Not dicTemp.Exists(IP)) Then
                                    If (proxyCount = MAX_PROXIES) Then
                                        pl.maxProxiesReached = True
                                        Exit For
                                    End If
                      
                                    Dim proxyLine As String
                        
                                    proxyLine = IP & "|" & port & "|" & proxyVersion & "|" & proxyIndex
                                    proxyIndex = proxyIndex + 1
                                    dicTemp.Add IP, proxyLine
                        
                                    proxyCount = proxyCount + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next ii
      
            If (pl.maxProxiesReached) Then
                Exit For
            End If
        End If
    Next i
  
    pl.loadedCount = proxyCount

    If (dicTemp.count > 0) Then
        Dim arrProxies() As clsProxyType, idx As Long, line As String, p As Variant
        Dim proxy As clsProxyType
    
        ReDim arrProxies(dicTemp.count - 1)
    
        For Each p In dicTemp.Items
            Dim data() As String
            data = Split(p, "|")
    
            Set arrProxies(idx) = New clsProxyType
    
            With arrProxies(idx)
                .setIP (data(0))
                .setPort (data(1))
                .setVersion (data(2))
                .setIndex (data(3))
            End With
      
            idx = idx + 1
        Next
    
        proxies.setConnectionsPerProxy config.socketsPerProxy
        proxies.setProxies arrProxies
    End If
  
    Set dicTemp = Nothing
  
    loadProxies = pl
End Function

Public Sub setupHashFiles()
    hashes.w2bnHashes(0) = App.path & "\Binaries\W2BN\Warcraft II BNE.exe"
    hashes.w2bnHashes(1) = App.path & "\Binaries\W2BN\storm.dll"
    hashes.w2bnHashes(2) = App.path & "\Binaries\W2BN\Battle.snp"
    hashes.w2bnHashes(3) = App.path & "\Binaries\W2BN\W2BN.bin"
    
    hashes.d2dvHashes(0) = App.path & "\Binaries\D2DV\Game.exe"
    
    hashes.d2xpHashes(0) = App.path & "\Binaries\D2XP\Game.exe"
    
    hashes.war3Hashes(0) = App.path & "\Binaries\WAR3\Warcraft III.exe"
    
    hashes.lockdownPath = App.path & "\Lockdown\"
    hashes.checkRevisionInfo = App.path & "\CheckRevisionFromWarden.ini"
End Sub

Public Function getHashes(ByVal product As String, Optional lockdownFileName As String = vbNullString) As HashSearchResult
    Dim hashFiles() As String, isLockdown As Boolean
    Dim result As HashSearchResult
  
    Select Case product
        Case "W2BN"
            hashFiles = hashes.w2bnHashes
            isLockdown = True
        Case "D2DV"
            hashFiles = hashes.d2dvHashes
        Case "D2XP"
            hashFiles = hashes.d2xpHashes
        Case "WAR3", "W3XP"
            hashFiles = hashes.war3Hashes
    End Select
  
    If (product = "D2DV" Or product = "D2XP") Then
        If (Dir$(hashes.checkRevisionInfo) = vbNullString) Then
        result.errorMessage = "missing CheckRevisionFromWarden.ini"
        getHashes = result
        Exit Function
        End If
    End If
  
    For i = 0 To UBound(hashFiles)
        If (Dir$(hashFiles(i)) = vbNullString) Then
        result.errorMessage = "missing hash files for " & product
        getHashes = result
        Exit Function
        End If
    Next i
  
    If (isLockdown And Len(lockdownFileName) >= 8 And LCase(left(lockdownFileName, 8)) = "lockdown") Then
        If (Dir$(hashes.lockdownPath & lockdownFileName) = vbNullString) Then
        result.errorMessage = "missing lockdown files for " & product
        getHashes = result
        Exit Function
        End If
    End If
  
    result.hashes = hashFiles
    result.hashesExist = True

    getHashes = result
End Function

Public Sub calculateAvailableSockets()
    Dim totalProxyConnections As Long, totalCDKeys As Long
  
    totalProxyConnections = proxies.countProxies() * config.socketsPerProxy
    totalCDKeys = totalNonExpKeys - testedNonExpKeys
  
    If (totalProxyConnections >= config.sockets And totalCDKeys >= config.sockets) Then
        socketsAvailable = config.sockets
    Else
        socketsAvailable = IIf(totalProxyConnections < totalCDKeys, totalProxyConnections, totalCDKeys)
    End If
  
    frmMain.lblControl(SOCKETS_AVAILABLE).Caption = socketsAvailable
    frmMain.lblControl(SOCKETS_TOTAL).Caption = config.sockets
  
    If (socketsAvailable < config.sockets) Then
        AddChat vbRed, "Insufficient keys or proxies. Not all sockets are available."
        AddChat vbRed, "Available sockets after loading is ", vbWhite, socketsAvailable, vbRed, " of ", vbWhite, config.sockets, vbRed, "."
    End If
End Sub

Public Function accountIdToReason(ByVal ID As Long, ByVal isWar3 As Boolean) As String
    Dim reason As String

    If (isWar3) Then
        Select Case ID
            Case &H4: reason = "username already exists."
            Case &H7: reason = "username is too short or blank."
            Case &H8: reason = "username contains an illegal character."
            Case &H9: reason = "username contains an illegal word."
            Case &HA: reason = "username contains too few alphanumeric characters."
            Case &HB: reason = "username contains adjacent punctuation characters."
            Case &HC: reason = "username contains too many punctuation characters."
            Case Else: reason = "username already exists."
        End Select
    Else
        Select Case ID
            Case &H1: reason = "username is too short"
            Case &H2: reason = "username contains invalid characters"
            Case &H3: reason = "username contained a banned word"
            Case &H4: reason = "username already exists"
            Case &H5: reason = "username is still being created"
            Case &H6: reason = "username does not contain enough alphanumeric characters"
            Case &H7: reason = "username contained adjacent punctuation characters"
            Case &H8: reason = "username contained too many punctuation characters"
        End Select
    End If

    accountIdToReason = reason
End Function

Public Function loadConfig() As Dictionary
    Dim dicErrors As New Dictionary, tempValue As String, error As Boolean
    
    config.name = ReadINI("Main", "Name", "Config.ini")
    
    If (Len(config.name) < 3) Then
        dicErrors.Add CONFIG_USERNAME, config.name
    End If
    
    config.nameW3 = ReadINI("Main", "NameW3", "Config.ini")
    
    If (Len(config.nameW3) < 3) Then
        dicErrors.Add CONFIG_USERNAMEW3, config.nameW3
    End If
    
    config.password = ReadINI("Main", "Password", "Config.ini")
    
    If (Len(config.password) < 1) Then
        dicErrors.Add CONFIG_PASSWORD, config.password
    End If
    
    config.passwordW3 = ReadINI("Main", "PasswordW3", "Config.ini")
    
    If (Len(config.passwordW3) < 1) Then
        dicErrors.Add CONFIG_PASSWORDW3, config.passwordW3
    End If
    
    config.homeChannel = ReadINI("Main", "HomeChannel", "Config.ini")
    
    If (Len(config.homeChannel) < 1) Then
        dicErrors.Add CONFIG_HOME_CHANNEL, config.homeChannel
    End If
    
    config.server = ReadINI("Main", "Server", "Config.ini")
    
    If (Len(config.server) < 3 Or Not isValidServerAddress(config.server)) Then
        dicErrors.Add CONFIG_SERVER, config.server
    Else
        config.serverIP = getProperGateway(config.server)
    
        Dim sr As ServerRealm
    
        sr = serverToRealm(config.serverIP)
    
        config.ServerRealm = sr.realm
        config.serverRealmW3 = sr.realmW3
        
        If (config.serverIP = vbNullString) Then
            dicErrors.Add CONFIG_SERVER, config.server
        End If
    End If
  
    tempValue = ReadINI("Main", "Sockets", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_SOCKETS & "f", DEFAULT_SOCKETS
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue > 0 And tempValue <= MAX_SOCKETS) Then
                config.sockets = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_SOCKETS, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "SocketsPerProxy", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_SOCKETS_PER_PROXY & "f", DEFAULT_SOCKETS_PER_PROXY
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue > 0 And tempValue <= MAX_SOCKETS_PER_PROXY) Then
                config.socketsPerProxy = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_SOCKETS_PER_PROXY, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "BNLSServer", "Config.ini")
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_BNLS_SERVER & "f", DEFAULT_BNLS_SERVER
    Else
        If (Len(tempValue) < 3) Then
            dicErrors.Add CONFIG_BNLS_SERVER, tempValue
        Else
            config.bnlsServer = tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "TestCountPerProxy", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_TEST_COUNT_PER_PROXY & "f", DEFAULT_TEST_COUNT_PER_PROXY
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue >= 0 And tempValue <= MAX_TEST_COUNT_PER_PROXY) Then
                config.testCountPerProxy = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_TEST_COUNT_PER_PROXY, tempValue
        End If
    End If
  
  tempValue = ReadINI("Main", "ExpansionTestsPerRegularKey", "Config.ini")
  error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_EXP_TESTS_PER_REG_KEY & "f", DEFAULT_EXP_TESTS_PER_REG_KEY
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue > 0 And tempValue <= MAX_EXP_TESTS_PER_REG_KEY) Then
                config.expansionTestsPerRegularKey = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_EXP_TESTS_PER_REG_KEY, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "ReconnectTime", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_RECONNECT_TIME & "f", DEFAULT_RECONNECT_TIME
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue > 0 And tempValue <= MAX_RECONNECT_TIME) Then
                config.reconnectTime = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_RECONNECT_TIME, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "CheckFailure", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_CHECK_FAILURE & "f", DEFAULT_CHECK_FAILURE
    Else
        If (IsNumericB(tempValue)) Then
            If (tempValue > 0 And tempValue <= MAX_CHECK_FAILURE) Then
                config.checkFailure = tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_CHECK_FAILURE, tempValue
        End If
    End If
  
    config.cdKeyProfile = ReadINI("Main", "CDKeyProfile", "Config.ini")
  
    If (Len(config.cdKeyProfile) > 0 And Not isValidCDKeyProfile(config.cdKeyProfile)) Then
        dicErrors.Add CONFIG_CDKEY_PROFILE, config.cdKeyProfile
    End If
  
    tempValue = UCase(ReadINI("Main", "AddDateToTested", "Config.ini"))
    
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_ADD_DATE_TO_TESTED & "f", DEFAULT_ADD_DATE_TO_TESTED
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_ADD_DATE_TO_TESTED, DEFAULT_ADD_DATE_TO_TESTED
        Else
            If (tempValue = "Y") Then
                config.addDateToTested = True
            End If
        End If
    End If

    tempValue = UCase(ReadINI("Main", "SkipFailedProxies", "Config.ini"))
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_SKIP_FAILED_PROXIES & "f", DEFAULT_SKIP_FAILED_PROXIES
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_SKIP_FAILED_PROXIES, DEFAULT_SKIP_FAILED_PROXIES
        Else
            If (tempValue = "Y") Then
                config.skipFailedProxies = True
            End If
        End If
    End If
  
    tempValue = UCase(ReadINI("Main", "SaveGoodProxies", "Config.ini"))
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_SAVE_GOOD_PROXIES & "f", DEFAULT_SAVE_GOOD_PROXIES
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_SAVE_GOOD_PROXIES, DEFAULT_SAVE_GOOD_PROXIES
        Else
            If (tempValue = "Y") Then
                config.saveGoodProxies = True
            End If
        End If
    End If

    tempValue = UCase(ReadINI("Main", "AddRealmToProfile", "Config.ini"))
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_ADD_REALM_TO_PROFILE & "f", DEFAULT_ADD_REALM_TO_PROFILE
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_ADD_REALM_TO_PROFILE, DEFAULT_ADD_REALM_TO_PROFILE
        Else
            If (tempValue = "Y") Then
                config.addRealmToProfile = True
            End If
        End If
    End If
  
    tempValue = UCase(ReadINI("Main", "SaveWindowPosition", "Config.ini"))
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_SAVE_WINDOW_POSITION & "f", DEFAULT_SAVE_WINDOW_POSITION
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_SAVE_WINDOW_POSITION, DEFAULT_SAVE_WINDOW_POSITION
        Else
            If (tempValue = "Y") Then
                config.saveWindowPosition = True
            End If
        End If
    End If
  
    tempValue = ReadINI("Main", "W2BNVerByte", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_VERBYTE_W2BN & "f", Hex(DEFAULT_VERBYTE_W2BN)
    Else
        If (IsNumeric("&H" & tempValue)) Then
            If (("&H" & tempValue) > 0 And ("&H" & tempValue) <= MAX_VERBYTE) Then
                config.W2BNVerByte = "&H" & tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_VERBYTE_W2BN, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "D2DVVerByte", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_VERBYTE_D2DV & "f", Hex(DEFAULT_VERBYTE_D2DV)
    Else
        If (IsNumeric("&H" & tempValue)) Then
            If (("&H" & tempValue) > 0 And ("&H" & tempValue) <= MAX_VERBYTE) Then
                config.D2DVVerByte = "&H" & tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_VERBYTE_D2DV, tempValue
        End If
    End If
  
    tempValue = ReadINI("Main", "WAR3VerByte", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_VERBYTE_WAR3 & "f", Hex(DEFAULT_VERBYTE_WAR3)
    Else
        If (IsNumeric("&H" & tempValue)) Then
            If (("&H" & tempValue) > 0 And ("&H" & tempValue) <= MAX_VERBYTE) Then
                config.WAR3VerByte = "&H" & tempValue
                error = False
            End If
        End If
    
        If (error) Then
            dicErrors.Add CONFIG_VERBYTE_WAR3, tempValue
        End If
    End If
  
    Set loadConfig = dicErrors
End Function

Public Sub writeConfig()
    With config
        WriteINI "Main", "Name", .name, "Config.ini"
        WriteINI "Main", "Password", .password, "Config.ini"
        WriteINI "Main", "NameW3", .nameW3, "Config.ini"
        WriteINI "Main", "PasswordW3", .passwordW3, "Config.ini"
        WriteINI "Main", "Server", .server, "Config.ini"
        WriteINI "Main", "BNLSServer", .bnlsServer, "Config.ini"
        WriteINI "Main", "HomeChannel", .homeChannel, "Config.ini"
        WriteINI "Main", "Sockets", .sockets, "Config.ini"
        WriteINI "Main", "SocketsPerProxy", .socketsPerProxy, "Config.ini"
        WriteINI "Main", "ExpansionTestsPerRegularKey", .expansionTestsPerRegularKey, "Config.ini"
        WriteINI "Main", "TestCountPerProxy", .testCountPerProxy, "Config.ini"
        WriteINI "Main", "CheckFailure", .checkFailure, "Config.ini"
        WriteINI "Main", "ReconnectTime", .reconnectTime, "Config.ini"
        WriteINI "Main", "CDKeyProfile", .cdKeyProfile, "Config.ini"
        WriteINI "Main", "AddDateToTested", IIf(.addDateToTested, "Y", "N"), "Config.ini"
        WriteINI "Main", "SaveGoodProxies", IIf(.saveGoodProxies, "Y", "N"), "Config.ini"
        WriteINI "Main", "SkipFailedProxies", IIf(.skipFailedProxies, "Y", "N"), "Config.ini"
        WriteINI "Main", "AddRealmToProfile", IIf(.addRealmToProfile, "Y", "N"), "Config.ini"
        WriteINI "Main", "SaveWindowPosition", IIf(.saveWindowPosition, "Y", "N"), "Config.ini"
        WriteINI "Main", "W2BNVerByte", Hex(.W2BNVerByte), "Config.ini"
        WriteINI "Main", "D2DVVerByte", Hex(.D2DVVerByte), "Config.ini"
        WriteINI "Main", "WAR3VerByte", Hex(.WAR3VerByte), "Config.ini"
    End With
End Sub

Public Sub makeDefaultValues()
    config.sockets = DEFAULT_SOCKETS
    config.socketsPerProxy = DEFAULT_SOCKETS_PER_PROXY
    config.expansionTestsPerRegularKey = DEFAULT_EXP_TESTS_PER_REG_KEY
    config.testCountPerProxy = DEFAULT_TEST_COUNT_PER_PROXY
    config.checkFailure = DEFAULT_CHECK_FAILURE
    config.reconnectTime = DEFAULT_RECONNECT_TIME
    
    config.addDateToTested = DEFAULT_ADD_DATE_TO_TESTED
    config.saveGoodProxies = DEFAULT_SAVE_GOOD_PROXIES
    config.skipFailedProxies = DEFAULT_SKIP_FAILED_PROXIES
    config.addRealmToProfile = DEFAULT_ADD_REALM_TO_PROFILE
    config.saveWindowPosition = DEFAULT_SAVE_WINDOW_POSITION
    
    config.bnlsServer = DEFAULT_BNLS_SERVER
    config.W2BNVerByte = DEFAULT_VERBYTE_W2BN
    config.D2DVVerByte = DEFAULT_VERBYTE_D2DV
    config.WAR3VerByte = DEFAULT_VERBYTE_WAR3
End Sub

Public Function IsNumericB(ByVal text As String) As Boolean
    Dim textLength As Integer
    textLength = Len(text)
  
    If (textLength > 0) Then
        For i = 1 To textLength
            Dim ch As String
      
            ch = UCase(Mid(text, i, 1))
      
            If (Not IsNumeric(ch)) Then
                IsNumericB = False
                Exit Function
            End If
        Next i
    
        IsNumericB = True
    Else
        IsNumericB = False
    End If
End Function

Public Function productToId(ByVal product As String) As Byte
    Select Case product
        Case "W2BN": productToId = &H3
        Case "D2DV": productToId = &H4
        Case "D2XP": productToId = &H5
        Case "WAR3": productToId = &H7
        Case "W3XP": productToId = &H8
    End Select
End Function

Public Function idToProduct(ByVal ID As Byte) As String
    Select Case product
        Case &H3: idToProduct = "W2BN"
        Case &H4: idToProduct = "D2DV"
        Case &H5: idToProduct = "D2XP"
        Case &H7: idToProduct = "WAR3"
        Case &H8: idToProduct = "W3XP"
    End Select
End Function

Public Sub setupConnectionData(ByVal newSockets As Integer)
    If (loadedSockets > 0 And loadedSockets <> newSockets) Then
        For i = 0 To loadedSockets - 1
            If (i > 0) Then
                Unload frmMain.sckBNCS(i)
                Unload frmMain.tmrCheckFailed(i)
                Unload frmMain.tmrReconnect(i)
            End If
        Next i
    End If
  
    If (loadedSockets = 0 Or loadedSockets <> newSockets) Then
        ReDim BNETData(newSockets - 1)
        ReDim packet(newSockets - 1)
    
        For i = 0 To newSockets - 1
            If (i > 0) Then
                Load frmMain.sckBNCS(i)
                Load frmMain.tmrCheckFailed(i)
                Load frmMain.tmrReconnect(i)
            End If
      
            Set packet(i) = New clsPacket
            packet(i).setDetails frmMain.sckBNCS(i), PacketType.BNCS
        Next i
  
        loadedSockets = newSockets
        FreeMemory
    End If
  
    For i = 0 To newSockets - 1
        frmMain.tmrCheckFailed(i).Interval = config.checkFailure
        frmMain.tmrReconnect(i).Interval = config.reconnectTime
    Next i
  
    frmMain.tmrCheckBNLS.Interval = config.checkFailure
End Sub

Public Sub markSocketDead(ByVal Index As Integer)
    socketsAvailable = socketsAvailable - 1
    frmMain.lblControl(SOCKETS_AVAILABLE).Caption = frmMain.lblControl(SOCKETS_AVAILABLE).Caption - 1
End Sub

Public Function getVerByte(ByVal product As String) As Byte
    Select Case product
        Case "W2BN": getVerByte = config.W2BNVerByte
        Case "D2DV", "D2XP": getVerByte = config.D2DVVerByte
        Case "WAR3", "W3XP": getVerByte = config.WAR3VerByte
    End Select
End Function

Public Sub initializeGatewayList()
    Dim IPs() As String

    gateways = Array("uswest.battle.net", "useast.battle.net", "europe.battle.net", "asia.battle.net", _
                     "connect-usw.classic.blizzard.com", "connect-use.classic.blizzard.com", _
                     "connect-eur.classic.blizzard.com", "connect-kor.classic.blizzard.com")

    For Each gateway In gateways
        IPs = Split(Resolve(gateway))

        If (UBound(IPs) > -1) Then
            dicGatewayIPs.Add gateway, IPs
        End If
    Next
End Sub

Public Function serverToRealm(serverIP As String) As ServerRealm
    Dim foundGateway As String, gateway As Variant, sr As ServerRealm
  
    For Each gateway In dicGatewayIPs.Keys
        If (gateway = serverIP) Then
            foundGateway = gateway
        Else
            Dim IP As Variant, IPs As Variant
            IPs = dicGatewayIPs.Item(gateway)
      
            For Each IP In IPs
                If (serverIP = IP) Then
                    foundGateway = gateway
                    Exit For
                End If
            Next
      
            If (foundGateway <> vbNullString) Then
                Exit For
            End If
        End If
    Next
  
    Select Case foundGateway
        Case "uswest.battle.net"
            sr.realm = "USWest"
            sr.realmW3 = "Lordaeron"
        Case "useast.battle.net"
            sr.realm = "USEast"
            sr.realmW3 = "Azeroth"
        Case "europe.battle.net"
            sr.realm = "Europe"
            sr.realmW3 = "Northrend"
        Case "asia.battle.net"
            sr.realm = "Asia"
            sr.realmW3 = "Kalimdor"
    End Select
  
    serverToRealm = sr
End Function

Public Function isValidServerAddress(address As String) As Boolean
    Dim realm As Variant
  
    For Each gateway In dicGatewayIPs.Keys
        If (address = gateway) Then
            isValidServerAddress = True
            Exit Function
        Else
            Dim IPs As Variant
            IPs = dicGatewayIPs.Item(gateway)
    
            For Each IP In IPs
                If (address = IP) Then
                    isValidServerAddress = True
                    Exit Function
                End If
            Next
        End If
    Next
  
    isValidServerAddress = False
End Function

Public Function getProperGateway(ByVal gateway As String) As String
    Dim gatewayList() As String
  
    If (IsNumericB(Replace(gateway, ".", vbNullString))) Then
        getProperGateway = gateway
        Exit Function
    End If

    gatewayList = dicGatewayIPs.Item(gateway)
    getProperGateway = gatewayList(CInt(Rnd * UBound(gatewayList)))
End Function

Public Function isValidCDKeyProfile(ByVal cdKeyProfile As String) As Boolean
    Dim ch As String

    For i = 1 To Len(cdKeyProfile)
        ch = Mid(cdKeyProfile, i, 1)
    
        If (ch = "\" Or ch = "\" Or ch = "*" Or ch = """" Or ch = "?" Or _
            ch = ":" Or ch = ">" Or ch = "<" Or ch = "|") Then
            isValidCDKeyProfile = False
            Exit Function
        End If
    Next i
  
    isValidCDKeyProfile = True
End Function

Public Function KillNull(ByVal text As String) As String
    Dim pos As Integer
  
    pos = InStr(text, Chr$(0))
  
    KillNull = IIf(pos > 0, Mid(text, 1, pos - 1), text)
End Function

'// Custom file operations
Public Function DirExists(dirName As String) As Boolean
    On Error GoTo ErrorHandler
    DirExists = GetAttr(dirName) And vbDirectory
ErrorHandler:
    ' if an error occurs, this function returns False
End Function

Public Function getFileSize(path As String) As Long
    On Error GoTo oops:
  
    Open path For Input As #1
        getFileSize = CLng(LOF(1))
    Close #1
  
oops:
    If (getFileSize > 0) Then Exit Function
    getFileSize = 0
End Function

Public Sub stopTesting(ByVal color As Long, ByVal msg As String)
    isTesting = False
  
    frmMain.tmrCheckBNLS.Enabled = False
    frmMain.tmrBenchmark.Enabled = False
  
    frmMain.lblReloadProxies.Enabled = True
    frmMain.lblReloadCDKeys.Enabled = True
    frmMain.lblConfig.Enabled = True

    frmMain.sckBNLS.Close
  
    For i = 0 To UBound(BNETData)
        frmMain.tmrCheckFailed(i).Enabled = False
        frmMain.tmrReconnect(i).Enabled = False
        closeSocket i
    Next i
  
    frmMain.lblStart.Caption = "Start"
  
    AddChatB color, msg
End Sub

Public Sub EndAll()
    isClosing = True
  
    sendKeysBack
  
    If (config.saveGoodProxies And config.skipFailedProxies) Then
        proxies.saveGoodProxies
    End If
  
    If (hasConfig And config.saveWindowPosition) Then
        WriteINI "Window", "Top", frmMain.top, "Config.ini"
        WriteINI "Window", "Left", frmMain.left, "Config.ini"
    End If
  
    Dim oFrm As Form
  
    For Each oFrm In Forms
        Unload oFrm
    Next
End Sub

Public Sub checkForQuitShortcut(fm As Form, key As Integer, shift As Integer)
    If (key = 115 And shift = 4) Then
        If (fm Is frmMain) Then
            EndAll
        Else
            Unload fm
        End If
    End If
End Sub

Public Function P_split(sIP As String) As String
    On Error Resume Next

    Dim splt() As String, i As Byte

    splt = Split(sIP, ".")
    
    For i = 0 To UBound(splt)
        P_split = P_split & Chr$(CStr(splt(i)))
    Next i
End Function

