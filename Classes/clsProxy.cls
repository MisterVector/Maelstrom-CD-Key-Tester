VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProxy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private arrProxies() As clsProxyType
Private proxyIndex As Long
Private connectionsPerProxy As Integer

Private Sub Class_Initialize()
  ReDim arrProxies(0)
  Set arrProxies(0) = New clsProxyType
  proxyIndex = 0
End Sub

Public Sub setConnectionsPerProxy(ByVal perProxy As Integer)
  connectionsPerProxy = perProxy
End Sub

Public Sub setProxies(proxies() As clsProxyType)
  arrProxies = proxies
  proxyIndex = 0
End Sub

Public Sub decrementProxyUse(Index As Long)
  arrProxies(Index).decrementCurrentlyUsing
End Sub

Public Sub onProxyFail(Index As Long)
  With arrProxies(Index)
    Dim numMarks
    numMarks = .getDeathMarks()
    
    If numMarks < config.socketsPerProxy Then
      .incrementDeathMarks
    Else
      .resetDeathMarks
    
      .setIP (vbNullString)
      .setPort (0)
      .setVersion (vbNullString)
      .setIndex (0)
      .resetCurrentlyUsing
    End If
  End With
End Sub

Public Function countProxies() As Long
  Dim count As Long: count = 0
  
  For i = 0 To UBound(arrProxies)
    If arrProxies(i).getIP() <> vbNullString Then
      count = count + 1
    End If
  Next i
  
  countProxies = count
End Function

Public Function getProxy() As clsProxyType
  Dim found As Boolean, proxy As New clsProxyType
  
  For i = proxyIndex To UBound(arrProxies)
    If arrProxies(i).getIP() <> vbNullString And _
        arrProxies(i).getCurrentlyUsing() < connectionsPerProxy Then
      proxyIndex = i
      found = True
      Exit For
    End If
  Next i
  
  If Not found And proxyIndex > 0 Then
    For i = 0 To proxyIndex - 1
      If arrProxies(i).getIP() <> vbNullString And _
          arrProxies(i).getCurrentlyUsing() < connectionsPerProxy Then
        proxyIndex = i
        Exit For
      End If
    Next i
  End If
  
  With arrProxies(proxyIndex)
    proxy.setIP (.getIP)
    proxy.setPort (.getPort)
    proxy.setVersion (.getVersion)
    proxy.setIndex (.getIndex)
    .incrementCurrentlyUsing
  End With
  
  If checkIfAlive() Then
    proxyIndex = proxyIndex + 1
    
    If proxyIndex > UBound(arrProxies) Then
      proxyIndex = 0
    End If
  Else
    proxyIndex = -1
  End If
  
  Set getProxy = proxy
End Function

Public Function canAcquireProxy() As Boolean
  canAcquireProxy = proxyIndex > -1
End Function

Public Function checkIfAlive() As Boolean
  Dim liveCount As Long

  For i = 0 To UBound(arrProxies)
    If arrProxies(i).getIP() <> vbNullString And _
        arrProxies(i).getCurrentlyUsing < connectionsPerProxy Then
      liveCount = liveCount + 1
    End If
  Next i
  
  checkIfAlive = (liveCount > 0)
End Function

Public Sub resetProxies()
  proxyIndex = 0
  
  For i = 0 To UBound(arrProxies)
    arrProxies(i).resetCurrentlyUsing
  Next i
End Sub
 
Public Sub saveGoodProxies()
  Dim dicSocks4 As New Dictionary, dicHttp As New Dictionary
  Dim IP As String, port As Long, version As String
  
  For i = 0 To UBound(arrProxies)
    If arrProxies(i).getIP() <> vbNullString Then
      IP = arrProxies(i).getIP()
      port = arrProxies(i).getPort()
      version = arrProxies(i).getVersion()
    
      If arrProxies(i).getVersion() = "SOCKS4" Then
        dicSocks4.Add IP, port
      End If
      
      If arrProxies(i).getVersion() = "HTTP" Then
        dicHttp.Add IP, port
      End If
    End If
  Next i
  
  If dicSocks4.count > 0 Then
    Open App.path & "\SOCKS4.txt" For Output As #1
    
    For Each key In dicSocks4.Keys
      port = dicSocks4.Item(key)
      
      Print #1, key & ":" & port
    Next
    
    Close #1
  End If
  
  If dicHttp.count > 0 Then
    Open App.path & "\HTTP.txt" For Output As #1
    
    For Each key In dicHttp.Keys
      port = dicHttp.Item(key)
      
      Print #1, key & ":" & port
    Next
    
    Close #1
  End If
End Sub
