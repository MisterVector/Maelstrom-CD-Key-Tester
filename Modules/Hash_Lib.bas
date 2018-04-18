Attribute VB_Name = "Hash_Lib"
'Do not modify this file!
'This is part of BNHash functionality and could possibly be updated. If you don't want to lose anywork
'then it's advised that you create your own module.

'LibBnet.dll by Rob

Public Declare Function nls_init Lib "libbnet.dll" (ByVal sUsername As String, ByVal sPassword As String) As Long
Public Declare Sub nls_free Lib "libbnet.dll" (ByVal lNLSPointer As Long)
Public Declare Function nls_account_logon Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Function nls_account_create Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String) As Long
Public Declare Sub nls_account_logon_proof Lib "libbnet.dll" (ByVal lNLSPointer As Long, ByVal sBufferOut As String, ByVal sServerKey As String, ByVal sSalt As String)
Public Declare Function checkRevision_ld Lib "libbnet.dll" Alias "checkrevision_ld" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sReturnDigest As String, ByVal sLockdownFile As String, ByVal sVideoFile As String) As Long
Public Declare Function checkRevision Lib "libbnet.dll" Alias "checkrevision" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sExeInfo As String, ByVal sMPQName As String) As Long
Public Declare Function decode_hash_cdkey Lib "libbnet.dll" (ByVal sCDKey As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByRef lPublicValue As Long, ByRef lProductID As Long, ByVal sBufferOut As String) As Long
Public Declare Sub double_hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByVal sBufferOut As String)
Public Declare Sub hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal sBufferOut As String)

Public Declare Function check_revision Lib "CheckRevisionFromWarden.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef version As Long, ByRef Checksum As Long, ByVal result As String) As Long
Public Declare Function crev_max_result Lib "CheckRevisionFromWarden.dll" () As Long

