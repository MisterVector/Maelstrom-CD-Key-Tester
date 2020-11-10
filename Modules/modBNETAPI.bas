Attribute VB_Name = "modBNETAPI"
Public Declare Function nls_init Lib "libbnet.dll" (ByVal sUsername As String, ByVal sPassword As String) As Long
Public Declare Sub nls_free Lib "libbnet.dll" (ByVal lNLSPointer As Long)
Public Declare Function checkRevision_ld Lib "libbnet.dll" Alias "checkrevision_ld" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sReturnDigest As String, ByVal sLockdownFile As String, ByVal sVideoFile As String) As Long
Public Declare Function checkRevision Lib "libbnet.dll" Alias "checkrevision" (ByVal sFile1 As String, ByVal sFile2 As String, ByVal sFile3 As String, ByVal sValueString As String, ByRef lVersion As Long, ByRef lChecksum As Long, ByVal sExeInfo As String, ByVal sMPQName As String) As Long
Public Declare Sub double_hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal lClientToken As Long, ByVal lServerToken As Long, ByVal sBufferOut As String)
Public Declare Sub hash_password Lib "libbnet.dll" (ByVal sPassword As String, ByVal sBufferOut As String)

Public Declare Function check_revision Lib "VersionCheck.dll" (ByVal ArchiveTime As String, ByVal ArchiveName As String, ByVal Seed As String, ByVal INIFile As String, ByVal INIHeader As String, ByRef Version As Long, ByRef Checksum As Long, ByVal result As String) As Long
Public Declare Function crev_max_result Lib "VersionCheck.dll" () As Long

Public Declare Function kd_quick Lib "bncsutil.dll" _
    (ByVal cdKey As String, ByVal ClientToken As Long, ByVal ServerToken As Long, _
    PublicValue As Long, Product As Long, ByVal HashBuffer As String, ByVal BufferLen As Long) As Long

