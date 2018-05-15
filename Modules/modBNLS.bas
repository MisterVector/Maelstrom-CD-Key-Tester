Attribute VB_Name = "modBNLS"
Public Sub SendBNLS0x10()
    With bnlsPacket
        .InsertDWORD productToId(requestProduct)
        .sendPacket &H10
    End With
End Sub

Public Sub RecvBNLS0x10()
    frmMain.sckBNLS.Close
    frmMain.tmrCheckBNLS.Enabled = False
  
    Dim productId As Long

    With bnlsPacket
        productId = .GetDWORD
    
        If (productId <> &H0) Then
            Dim verByte As Long, productName As String, configName As String, updated As Boolean
    
            verByte = .GetDWORD

            Select Case requestProduct
                Case "W2BN"
                    updated = (verByte <> config.W2BNVerByte)
                    productName = "Warcraft II"
                
                    If (updated) Then
                        config.W2BNVerByte = verByte
                        configName = "W2BNVerByte"
                    End If
                Case "D2DV"
                    updated = (verByte <> config.D2DVVerByte)
                    productName = "Diablo II"
                
                    If (updated) Then
                        config.D2DVVerByte = verByte
                        configName = "D2DVVerByte"
                    End If
                Case "WAR3"
                    updated = (verByte <> config.WAR3VerByte)
                    productName = "Warcraft III"
                
                    If (updated) Then
                        config.WAR3VerByte = verByte
                        configName = "WAR3VerByte"
                    End If
            End Select
            
            If (updated) Then
                Dim verByteString As String
            
                verByteString = "0x" & IIf(Len(Hex(verByte)) = 1, "0", vbNullString) & Hex(verByte)
                WriteINI "Main", configName, Hex(verByte), "Config.ini"
              
                AddChatB vbGreen, "Updated the version byte for " & productName & " to: " & verByteString & "."
            Else
                msgBoxResult = MsgBox("The version byte for " & productName & " could not be updated.", vbOKOnly & vbExclamation, PROGRAM_NAME)
            End If
            
            stopTesting vbYellow, "Click ""Start"" to begin testing again."
        Else
            AddChatB vbRed, "Failed to update version byte for product: " & requestProduct & "."
        End If
    End With
End Sub
