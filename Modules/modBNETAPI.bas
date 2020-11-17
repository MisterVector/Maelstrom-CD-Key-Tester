Attribute VB_Name = "modBNETAPI"
Public Declare Function check_revision Lib "VersionCheck.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Long, ByRef Checksum As Long, ByVal Result As String) As Long
Public Declare Function crev_max_result Lib "VersionCheck.dll" () As Long

Public Declare Function kd_quick Lib "bncsutil.dll" _
    (ByVal CDKey As String, ByVal ClientToken As Long, ByVal ServerToken As Long, _
    PublicValue As Long, Product As Long, ByVal HashBuffer As String, ByVal BufferLen As Long) As Long

' Old Logon System
' [!] You should use doubleHashPassword and hashPassword instead of their
'     _Raw counterparts.  (See below for those functions.)
Public Declare Sub doubleHashPassword_Raw Lib "bncsutil.dll" Alias "doubleHashPassword" _
    (ByVal Password As String, ByVal ClientToken As Long, ByVal ServerToken As Long, _
    ByVal outBuffer As String)
Public Declare Sub hashPassword_Raw Lib "bncsutil.dll" Alias "hashPassword" _
    (ByVal Password As String, ByVal outBuffer As String)

'OLS Password Hashing
Public Function doubleHashPassword(Password As String, ByVal ClientToken&, ByVal ServerToken&) As String
    Dim Hash As String * 20
    doubleHashPassword_Raw Password, ClientToken, ServerToken, Hash
    doubleHashPassword = Hash
End Function

Public Function hashPassword(Password As String) As String
    Dim Hash As String * 20
    hashPassword_Raw Password, Hash
    hashPassword = Hash
End Function
