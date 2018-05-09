VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 1  'vbDataSource
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPacket"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private sIn As String, sOut As String, sLen As Long
Private socket As Winsock, pType As PacketType

Public Sub setDetails(sck As Winsock, pt As PacketType)
  Set socket = sck
  pType = pt
End Sub

Public Sub SetData(ByVal sData As String)
  sIn = sData
  sLen = 1
End Sub

Public Sub clearIn()
  sIn = vbNullString
End Sub

Public Sub Skip(Bytes As Long)
  sLen = sLen + Bytes
End Sub

Public Sub InsertByte(ByVal value As Byte)
  sOut = sOut & Chr$(value)
End Sub

Public Sub InsertWORD(ByVal value As Integer)
  Dim NewValue As String * 2
  CopyMemory ByVal NewValue, value, 2
  sOut = sOut & NewValue
End Sub

Public Sub InsertDWORD(ByVal value As Long)
  Dim NewValue As String * 4
  CopyMemory ByVal NewValue, value, 4
  sOut = sOut & NewValue
End Sub

Public Sub InsertNTString(ByVal value As String)
  sOut = sOut & value & vbNullChar
End Sub

Public Sub InsertNonNTString(ByVal value As String)
  sOut = sOut & value
End Sub

Public Sub sendPacket(ID As Byte)
  Dim packetLengthField As String * 2
  Dim packetTotalLength As Integer, packetHeader As String
  
  If (pType = PacketType.BNCS) Then
    packetTotalLength = Len(sOut) + 4
  Else
    packetTotalLength = Len(sOut) + 3
  End If
  
  CopyMemory ByVal packetLengthField, packetTotalLength, 2
  
  If (pType = PacketType.BNCS) Then
    packetHeader = Chr$(&HFF) & Chr$(ID) & packetLengthField
  Else
    packetHeader = packetLengthField & Chr$(ID)
  End If

  If (socket.State = sckConnected) Then
    socket.SendData packetHeader & sOut
  End If

  sOut = vbNullString
End Sub

Public Function GetByte() As Byte
  GetByte = Asc(Mid$(sIn, sLen, 1))
  sLen = sLen + 1
End Function

Public Function GetWORD() As Long
  Dim NewValue As Long
  CopyMemory NewValue, ByVal Mid$(sIn, sLen, 2), 2
  GetWORD = NewValue
  sLen = sLen + 2
End Function

Public Function GetDWORD() As Long
  Dim NewValue As Long
  CopyMemory NewValue, ByVal Mid$(sIn, sLen, 4), 4
  GetDWORD = NewValue
  sLen = sLen + 4
End Function

Public Function GetNonNTString(ByVal Length As Integer) As String
  GetNonNTString = Mid$(sIn, sLen, Length)
  sLen = sLen + Length
End Function

Public Function getNTString() As String
  On Error Resume Next
  getNTString = Split(Mid$(sIn, sLen), vbNullChar)(0)
  sLen = sLen + Len(getNTString) + 1
End Function

Public Function getPacket() As String
  getPacket = sIn
End Function

