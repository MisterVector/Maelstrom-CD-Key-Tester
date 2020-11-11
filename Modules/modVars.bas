Attribute VB_Name = "modVars"
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, _
    ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Public Const PROGRAM_VERSION                  As String = "4.2.3"
Public Const KEY_TESTER_NAME                  As String = "Key Tester"
Public Const PROGRAM_NAME                     As String = "Maelstrom CD-Key Tester v" & PROGRAM_VERSION & " by Vector"
Public Const RELEASES_URL                     As String = "https://github.com/MisterVector/Maelstrom-CD-Key-Tester-Legacy/releases"

Public Const DEFAULT_VERBYTE_D2DV             As Long = &HE
Public Const DEFAULT_VERBYTE_W2BN             As Long = &H4F

Public Const DEFAULT_SOCKETS                  As Integer = 250
Public Const DEFAULT_SOCKETS_PER_PROXY        As Integer = 4
Public Const DEFAULT_BNLS_SERVER              As String = "jbls.codespeak.org"
Public Const DEFAULT_BNLS_PORT                As Integer = 9367
Public Const DEFAULT_EXP_TESTS_PER_REG_KEY    As Integer = 8
Public Const DEFAULT_TEST_COUNT_PER_PROXY     As Integer = 25
Public Const DEFAULT_CHECK_FAILURE            As Integer = 10000
Public Const DEFAULT_RECONNECT_TIME           As Integer = 13000
Public Const DEFAULT_ADD_DATE_TO_TESTED       As Boolean = False
Public Const DEFAULT_SAVE_GOOD_PROXIES        As Boolean = False
Public Const DEFAULT_SKIP_FAILED_PROXIES      As Boolean = True
Public Const DEFAULT_ADD_REALM_TO_PROFILE     As Boolean = False
Public Const DEFAULT_SAVE_WINDOW_POSITION     As Boolean = True
Public Const DEFAULT_CHECK_UPDATE_ON_STARTUP  As Boolean = True

Public Const MAX_SOCKETS                      As Integer = 32767
Public Const MAX_SOCKETS_PER_PROXY            As Integer = 8
Public Const MAX_EXP_TESTS_PER_REG_KEY        As Integer = 32767
Public Const MAX_TEST_COUNT_PER_PROXY         As Integer = 32767
Public Const MAX_RECONNECT_TIME               As Integer = 32767
Public Const MAX_CHECK_FAILURE                As Integer = 32767

Public Const MAX_VERBYTE                      As Long = &H7FFFFFFF
Public Const MAX_PROXIES                      As Long = 100000

Public Const TEXT_PERFECT                     As Long = vbGreen
Public Const TEXT_IN_USE                      As Long = &HE74516
Public Const TEXT_MUTED                       As Long = &HFFFF&
Public Const TEXT_VOIDED                      As Long = &H80FF&
Public Const TEXT_JAILED                      As Long = &HA4E1&
Public Const TEXT_OTHER                       As Long = &HBCBCBE
Public Const TEXT_BANNED                      As Long = &HFF
Public Const TEXT_INVALID                     As Long = &HFFFFFF

Public Const FRM_BACK_COLOR                   As Long = &H404040
Public Const TXT_BACK_COLOR                   As Long = &HE0E0E0
Public Const TXT_ERROR_COLOR                  As Long = &HC0&
Public Const TXT_WARN_COLOR                   As Long = &HD6060
Public Const TXT_FILL_COLOR                   As Long = &HED5750

Public Const CDKEYS_FOLDER                    As String = "CD-Keys"
Public Const CDKEYS_TESTED_DEFAULT_FOLDER     As String = "Tested CD-Keys"

Public dicGatewayIPs As New Dictionary
Public requestProduct As String
Public updateString As String
Public timesTillClear As Integer

Public Type HashSearchResult
    hashes()     As String
    hashesExist  As Boolean
    errorMessage As String
End Type

Public Type ServerRealm
    realm As String
End Type

Public Type ProxiesLoaded
    loadedCount As Long
    maxProxiesReached As Boolean
End Type

Public loadedSockets As Integer
Public socketsAvailable As Integer

Public hasTestedThisSession As Boolean

Public packet() As clsPacket
Public bnlsPacket As New clsPacket

Public isTesting As Boolean
Public isClosing As Boolean

Public proxies As New clsProxy

Public hasProxies As Boolean
Public hasKeys As Boolean
Public hasConfig As Boolean

Public manualUpdateCheck As Boolean

Public Type ConfigData
    name As String
    password As String
    server As String
    bnlsServer As String
    bnlsPort As Long
    homeChannel As String
    testCountPerProxy As Integer
    socketsPerProxy As Integer
    sockets As Integer
    reconnectTime As Integer
    checkFailure As Integer
    
    addDateToTested As Boolean
    skipFailedProxies As Boolean
    saveGoodProxies As Boolean
    addRealmToProfile As Boolean
    saveWindowPosition As Boolean
    checkUpdateOnStartup As Boolean
    
    expansionTestsPerRegularKey As Integer
    cdKeyProfile As String
    
    'These values are generated  by Maelstrom and not loaded from the config
    serverIP As String
    ServerRealm As String
    
    'D2XP uses the same VerByte as D2DV
    W2BNVerByte As Long
    D2DVVerByte As Long
End Type
Public config As ConfigData

Public Enum PacketType
    BNLS
    BNCS
End Enum

Public Type BNETDataType
    ' Date from/to BNCS
    ServerToken     As Long
    ClientToken     As Long
    
    ' Data for Battle.Net
    cdKey          As String
    cdKeyIndex     As Long
    product        As String
    proxyIP        As String
    proxyPort      As Long
    proxyIndex     As Long
    
    'Custom variables
    numTested      As Integer
    proxyVersion   As String
    
    isValidated    As Boolean
     
    'For expansion key data
    TestedEXP        As Integer
    isExpansion      As Boolean
    cdKeyExp         As String
    cdKeyExpIndex    As Long
    savedKeyState    As String
    
    productRegular   As String
    productExpansion As String
    
    'Data for SOCKS5 proxies
    acceptedAuth     As Boolean
    
    ' For NLS
    nls_P          As Long
End Type
Public BNETData() As BNETDataType

Private Type HashDependencies
    w2bnHashes(3)     As String
    d2dvHashes(0)     As String
    d2xpHashes(0)     As String
    lockdownPath      As String
    checkRevisionInfo As String
End Type
Public hashes As HashDependencies

