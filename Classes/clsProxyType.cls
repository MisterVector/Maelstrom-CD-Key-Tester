VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProxyType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private IP As String
Private port As Long
Private version As String
Private Index As Long
Private currentlyUsing As Integer
Private deathMarks As Integer

Public Sub setIP(s As String)
    IP = s
End Sub

Public Function getIP() As String
    getIP = IP
End Function

Public Function setPort(p As Long)
    port = p
End Function

Public Function getPort() As Long
    getPort = port
End Function

Public Sub setVersion(s As String)
    version = s
End Sub

Public Function getVersion() As String
    getVersion = version
End Function

Public Sub setIndex(i As Long)
    Index = i
End Sub

Public Function getIndex() As Long
    getIndex = Index
End Function

Public Sub incrementCurrentlyUsing()
    currentlyUsing = currentlyUsing + 1
End Sub

Public Sub decrementCurrentlyUsing()
    currentlyUsing = currentlyUsing - 1
End Sub

Public Sub resetCurrentlyUsing()
    currentlyUsing = 0
End Sub

Public Function getCurrentlyUsing() As Integer
    getCurrentlyUsing = currentlyUsing
End Function

Public Sub incrementDeathMarks()
    deathMarks = deathMarks + 1
End Sub

Public Function getDeathMarks() As Integer
    getDeathMarks = deathMarks
End Function

Public Sub resetDeathMarks()
    deathMarks = 0
End Sub
