Attribute VB_Name = "modOtherCode"
Public Type PortLong
    n As Long
End Type

Public Type PortBytes
    n1 As Byte
    n2 As Byte
    n3 As Byte
    n4 As Byte
End Type

Public Function loadProxies() As ProxiesLoaded
    Dim proxyVersions() As Variant, dicTemp As New Dictionary, proxyCount As Long
    Dim maxProxiesReached As Boolean, pl As ProxiesLoaded
  
    proxyCount = 0
  
    proxyVersions = Array("SOCKS4", "SOCKS5", "HTTP")
  
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
                    tProxies(ii) = Trim$(tProxies(ii))
                    
                    If (InStr(tProxies(ii), ":")) Then
                        Dim IP As String, Port As String, parts() As String
                        
                        parts = Split(tProxies(ii), ":")
                        IP = parts(0)
                        Port = parts(1)
                        
                        If (IsNumeric(Port) And IsNumeric(Replace(IP, ".", vbNullString))) Then
                            If (Port <= 65535 And Port > 0) Then
                                If (Not dicTemp.Exists(IP)) Then
                                    If (proxyCount = MAX_PROXIES) Then
                                        pl.maxProxiesReached = True
                                        Exit For
                                    End If
                      
                                    Dim proxyLine As String
                        
                                    proxyLine = IP & "|" & Port & "|" & proxyVersion & "|" & proxyIndex
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
            Dim Data() As String
            Data = Split(p, "|")
    
            Set arrProxies(idx) = New clsProxyType
    
            With arrProxies(idx)
                .setIP (Data(0))
                .setPort (Data(1))
                .setVersion (Data(2))
                .setIndex (Data(3))
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
    
    hashes.lockdownPath = App.path & "\Lockdown\"
    hashes.checkRevisionInfo = App.path & "\VersionCheck.ini"
End Sub

Public Function getHashes(ByVal Product As String, Optional lockdownFileName As String = vbNullString) As HashSearchResult
    Dim hashFiles() As String, isLockdown As Boolean
    Dim Result As HashSearchResult
  
    Select Case Product
        Case "W2BN"
            hashFiles = hashes.w2bnHashes
            isLockdown = True
        Case "D2DV"
            hashFiles = hashes.d2dvHashes
        Case "D2XP"
            hashFiles = hashes.d2xpHashes
    End Select
  
    If (Product = "D2DV" Or Product = "D2XP") Then
        If (Dir$(hashes.checkRevisionInfo) = vbNullString) Then
            Result.errorMessage = "missing VersionCheck.ini"
            getHashes = Result
            Exit Function
        End If
    End If
  
    For i = 0 To UBound(hashFiles)
        If (Dir$(hashFiles(i)) = vbNullString) Then
        Result.errorMessage = "missing hash files for " & Product
        getHashes = Result
        Exit Function
        End If
    Next i
  
    If (isLockdown And Len(lockdownFileName) >= 8 And LCase$(left$(lockdownFileName, 8)) = "lockdown") Then
        If (Dir$(hashes.lockdownPath & lockdownFileName) = vbNullString) Then
        Result.errorMessage = "missing lockdown files for " & Product
        getHashes = Result
        Exit Function
        End If
    End If
  
    Result.hashes = hashFiles
    Result.hashesExist = True

    getHashes = Result
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

Public Function accountIdToReason(ByVal ID As Long) As String
    Dim reason As String

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

    accountIdToReason = reason
End Function

Public Function loadConfig() As Dictionary
    Dim dicErrors As New Dictionary, tempValue As String, error As Boolean
    
    config.name = ReadINI("Main", "Name", "Config.ini")
    
    If (Len(config.name) < 3) Then
        dicErrors.Add CONFIG_USERNAME, config.name
    End If
    
    config.password = ReadINI("Main", "Password", "Config.ini")
    
    If (Len(config.password) < 1) Then
        dicErrors.Add CONFIG_PASSWORD, config.password
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
        config.ServerRealm = serverToRealm(config.serverIP)
        
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
  
    tempValue = UCase$(ReadINI("Main", "AddDateToTested", "Config.ini"))
    
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

    tempValue = UCase$(ReadINI("Main", "SkipFailedProxies", "Config.ini"))
  
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
  
    tempValue = UCase$(ReadINI("Main", "SaveGoodProxies", "Config.ini"))
  
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

    tempValue = UCase$(ReadINI("Main", "AddRealmToProfile", "Config.ini"))
  
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
  
    tempValue = UCase$(ReadINI("Main", "SaveWindowPosition", "Config.ini"))
  
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
  
    tempValue = UCase$(ReadINI("Main", "CheckUpdateOnStartup", "Config.ini"))
    
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_CHECK_UPDATE_ON_STARTUP & "f", DEFAULT_CHECK_UPDATE_ON_STARTUP
    Else
        If (tempValue <> "Y" And tempValue <> "N") Then
            dicErrors.Add CONFIG_CHECK_UPDATE_ON_STARTUP, DEFAULT_CHECK_UPDATE_ON_STARTUP
        Else
            If (tempValue = "Y") Then
                config.checkUpdateOnStartup = True
            End If
        End If
    End If
  
    tempValue = ReadINI("Main", "W2BNVerByte", "Config.ini")
    error = True
  
    If (tempValue = vbNullString) Then
        dicErrors.Add CONFIG_VERBYTE_W2BN & "f", Hex$(DEFAULT_VERBYTE_W2BN)
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
        dicErrors.Add CONFIG_VERBYTE_D2DV & "f", Hex$(DEFAULT_VERBYTE_D2DV)
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
  
    Set loadConfig = dicErrors
End Function

Public Sub writeConfig()
    With config
        WriteINI "Main", "Name", .name, "Config.ini"
        WriteINI "Main", "Password", .password, "Config.ini"
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
        WriteINI "Main", "CheckUpdateOnStartup", IIf(.checkUpdateOnStartup, "Y", "N"), "Config.ini"
        WriteINI "Main", "W2BNVerByte", Hex$(.W2BNVerByte), "Config.ini"
        WriteINI "Main", "D2DVVerByte", Hex$(.D2DVVerByte), "Config.ini"
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
    config.checkUpdateOnStartup = DEFAULT_CHECK_UPDATE_ON_STARTUP
    
    config.bnlsServer = DEFAULT_BNLS_SERVER
    config.W2BNVerByte = DEFAULT_VERBYTE_W2BN
    config.D2DVVerByte = DEFAULT_VERBYTE_D2DV
End Sub

Public Function IsNumericB(ByVal text As String) As Boolean
    Dim textLength As Integer
    textLength = Len(text)
  
    If (textLength > 0) Then
        For i = 1 To textLength
            Dim ch As String
      
            ch = Mid$(text, i, 1)
      
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

Public Function productToId(ByVal Product As String) As Byte
    Select Case Product
        Case "W2BN": productToId = &H3
        Case "D2DV": productToId = &H4
        Case "D2XP": productToId = &H5
    End Select
End Function

Public Function idToProduct(ByVal ID As Byte) As String
    Select Case Product
        Case &H3: idToProduct = "W2BN"
        Case &H4: idToProduct = "D2DV"
        Case &H5: idToProduct = "D2XP"
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

Public Sub markSocketDead(ByVal index As Integer)
    socketsAvailable = socketsAvailable - 1
    frmMain.lblControl(SOCKETS_AVAILABLE).Caption = frmMain.lblControl(SOCKETS_AVAILABLE).Caption - 1
End Sub

Public Function getVerByte(ByVal Product As String) As Byte
    Select Case Product
        Case "W2BN": getVerByte = config.W2BNVerByte
        Case "D2DV", "D2XP": getVerByte = config.D2DVVerByte
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

Public Function serverToRealm(serverIP As String) As String
    Dim foundGateway As String, gateway As Variant, realm As String
  
    For Each gateway In dicGatewayIPs.keys
        If (gateway = serverIP) Then
            foundGateway = gateway
            Exit For
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
        Case "uswest.battle.net", "connect-usw.classic.blizzard.com"
            realm = "USWest"
        Case "useast.battle.net", "connect-use.classic.blizzard.com"
            realm = "USEast"
        Case "europe.battle.net", "connect-eur.classic.blizzard.com"
            realm = "Europe"
        Case "asia.battle.net", "connect-kor.classic.blizzard.com"
            realm = "Asia"
    End Select
  
    serverToRealm = realm
End Function

Public Function isValidServerAddress(Address As String) As Boolean
    Dim realm As Variant
  
    For Each gateway In dicGatewayIPs.keys
        If (Address = gateway) Then
            isValidServerAddress = True
            Exit Function
        Else
            Dim IPs As Variant
            IPs = dicGatewayIPs.Item(gateway)
    
            For Each IP In IPs
                If (Address = IP) Then
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
        ch = Mid$(cdKeyProfile, i, 1)
    
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
  
    KillNull = IIf(pos > 0, Mid$(text, 1, pos - 1), text)
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

Public Sub stopTesting(ParamArray saElements() As Variant)
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

Public Sub checkForQuitShortcut(fm As Form, key As Integer, Shift As Integer)
    If (key = 115 And Shift = 4) Then
        If (fm Is frmMain) Then
            EndAll
        Else
            Unload fm
        End If
    End If
End Sub

Public Function portToBytes(Port As Long) As String
    Dim pLong As PortLong, pBytes As PortBytes

    pLong.n = Port
    
    LSet pBytes = pLong
    
    portToBytes = Chr$(pBytes.n2) & Chr$(pBytes.n1)
End Function

Public Function P_split(sIP As String) As String
    On Error Resume Next

    Dim splt() As String, i As Byte

    splt = Split(sIP, ".")
    
    For i = 0 To UBound(splt)
        P_split = P_split & Chr$(splt(i))
    Next i
End Function

Public Function makeCompatibleDate(ByVal dateTimeString As String) As Date
    dateTimeString = Replace(dateTimeString, "T", " ")
    dateTimeString = Replace(dateTimeString, "Z", vbNullString)
    
    makeCompatibleDate = dateTimeString
End Function

Public Function checkProgramUpdate(ByVal manualUpdateCheck As Boolean) As Boolean
    On Error GoTo err
    
    Dim text As String, status As Integer, requestReleaseTime As Date, releaseTime As Date, requestVersion As String, Version As String
    Dim isoRequestReleaseTime As String, isoReleaseTime As String
    Dim jsonResponse As Dictionary, jsonContents As Dictionary
    Dim updateMsg As String, msgBoxResult As Integer
    Dim xml As Object
    
    Set xml = CreateObject("MSXML2.XMLHTTP")

    xml.Open "GET", PROGRAM_UPDATE_URL, False
    xml.setRequestHeader "User-Agent", "MaelstromCDKeyTester/" & PROGRAM_VERSION
    xml.send
    
    text = xml.responseText
    
    Set jsonResponse = JSON.parse(text)
    status = jsonResponse.Item("status")
    
    If (status = 1) Then
        Set jsonContents = jsonResponse.Item("contents")
        
        isoRequestReleaseTime = jsonContents.Item("request_release_time")
        requestVeresion = jsonContents.Item("request_version")
        isoReleaseTime = jsonContents.Item("release_time")
        Version = jsonContents.Item("version")

        requestReleaseTime = makeCompatibleDate(isoRequestReleaseTime)
        releaseTime = makeCompatibleDate(isoReleaseTime)
        
        If (releaseTime > requestReleaseTime) Then
            updateMsg = "There is a new update for " & PROGRAM_NAME & "!" & vbNewLine & vbNewLine & "Your version: " & PROGRAM_VERSION & " new version: " & Version & vbNewLine & vbNewLine _
                      & "Would you like to view the changelog and download the latest update?"
        
            msgBoxResult = MsgBox(updateMsg, vbYesNo Or vbInformation, "New version for " & PROGRAM_TITLE)
    
            If (msgBoxResult = vbYes) Then
                ShellExecute 0, "open", UPDATE_SUMMARY_URL, vbNullString, vbNullString, 4
            End If
        Else
            If (manualUpdateCheck) Then
                MsgBox "There is no new version at this time.", vbOKOnly Or vbInformation, PROGRAM_TITLE
            End If
        End If
        
        checkProgramUpdate = True
        Exit Function
    End If

err:
    Set xml = Nothing
End Function
